#!/bin/bash

# Configuration variables
CLIENT_MAX_BODY_SIZE="20M"  # Adjust this value based on your maximum expected file size
VENV_PATH="$HOME/venv"  # Path for the virtual environment

# Check if domain name is provided
if [ $# -eq 0 ]; then
    echo "Please provide a domain name as an argument."
    echo "Usage: $0 your_domain.com"
    exit 1
fi

DOMAIN=$1

# Update system and install dependencies
sudo apt update
sudo apt install -y python3 python3-pip python3-venv nginx \
    libjpeg-dev zlib1g-dev libfreetype6-dev liblcms2-dev libopenjp2-7-dev \
    libwebp-dev libpng-dev

# Get the current user's home directory
USER_HOME=$(eval echo ~$USER)

# Create and activate virtual environment
python3 -m venv $VENV_PATH
source $VENV_PATH/bin/activate

# Install Flask, Gunicorn, and app requirements in the virtual environment
pip install flask gunicorn
if [ -f "$USER_HOME/app/requirements.txt" ]; then
    # Remove 'zipfile' from requirements.txt if it exists
    sed -i '/^zipfile$/d' "$USER_HOME/app/requirements.txt"
    
    pip install -r "$USER_HOME/app/requirements.txt"
    if [ $? -ne 0 ]; then
        echo "Error: Failed to install requirements. Please check your requirements.txt file."
        exit 1
    fi
else
    echo "Warning: requirements.txt not found in ~/app/. Skipping additional package installation."
fi

# Verify Flask app
cd "$USER_HOME/app"
if ! python -c "import app; print(app.app)"; then
    echo "Error: Failed to import Flask app. Please check your app code."
    exit 1
fi


# Check if port 8000 is in use
if sudo lsof -i :8000; then
    echo "Warning: Port 8000 is already in use. You may need to stop the existing process."
fi

# Configure Nginx
sudo tee /etc/nginx/sites-available/flask_app << EOT
server {
    listen 80;
    server_name $DOMAIN;

    client_max_body_size $CLIENT_MAX_BODY_SIZE;

    location / {
        proxy_pass http://127.0.0.1:8000;
        proxy_set_header Host \$host;
        proxy_set_header X-Real-IP \$remote_addr;
        proxy_set_header X-Forwarded-For \$proxy_add_x_forwarded_for;
        proxy_set_header X-Forwarded-Proto \$scheme;
    }
}
EOT

# Enable the Nginx site
sudo ln -s /etc/nginx/sites-available/flask_app /etc/nginx/sites-enabled/

# Remove default Nginx site
sudo rm /etc/nginx/sites-enabled/default

# Start Nginx
sudo systemctl start nginx
sudo systemctl enable nginx

# Set up systemd service for Flask app
sudo tee /etc/systemd/system/flask_app.service << EOT
[Unit]
Description=Gunicorn instance to serve Flask app
After=network.target

[Service]
User=$USER
WorkingDirectory=$USER_HOME/app
ExecStart=$VENV_PATH/bin/gunicorn --workers 3 --bind 127.0.0.1:8000 --timeout 120 app:app
Restart=always
Environment=PATH=$VENV_PATH/bin
StandardOutput=journal
StandardError=journal

[Install]
WantedBy=multi-user.target
EOT

# Reload systemd, start and enable Flask app service
sudo systemctl daemon-reload
sudo systemctl start flask_app
sudo systemctl enable flask_app

# Check if the service started successfully
if ! sudo systemctl is-active --quiet flask_app; then
    echo "Error: Flask app service failed to start. Checking logs..."
    sudo journalctl -u flask_app.service -n 50
    exit 1
fi

# Verify Nginx configuration
sudo nginx -t

# Restart Nginx to apply changes
sudo systemctl restart nginx

# Final checks
echo "Performing final checks..."

# Check Nginx error logs
echo "Recent Nginx error logs:"
sudo tail -n 20 /var/log/nginx/error.log

# Final status check
if sudo systemctl is-active --quiet flask_app && sudo systemctl is-active --quiet nginx; then
    echo "Deployment completed. Your Flask app should now be running at http://$DOMAIN"
    echo "If you encounter issues, please review the logs and checks above."
else
    echo "Error: Services are not running as expected. Please review the logs and checks above."
fi

# Provide information about file upload configuration
echo "Note: This deployment is configured to handle file uploads up to $CLIENT_MAX_BODY_SIZE."
echo "If you need to adjust this limit, modify the CLIENT_MAX_BODY_SIZE variable at the top of this script"
echo "and run the script again."

# Deactivate the virtual environment
deactivate