#!/bin/bash

# Check if domain name is provided
if [ $# -eq 0 ]; then
    echo "Please provide the domain name used during deployment."
    echo "Usage: $0 your_domain.com"
    exit 1
fi

DOMAIN=$1
VENV_PATH="$HOME/venv"  # Path for the virtual environment

echo "Starting cleanup process for $DOMAIN..."

# Stop and disable Flask app service
echo "Stopping and disabling Flask app service..."
sudo systemctl stop flask_app
sudo systemctl disable flask_app

# Remove Flask app service file
echo "Removing Flask app service file..."
sudo rm /etc/systemd/system/flask_app.service

# Reload systemd
sudo systemctl daemon-reload

# Remove Nginx configuration
echo "Removing Nginx configuration..."
sudo rm /etc/nginx/sites-available/flask_app
sudo rm /etc/nginx/sites-enabled/flask_app

# Restore default Nginx site
echo "Restoring default Nginx site..."
sudo ln -s /etc/nginx/sites-available/default /etc/nginx/sites-enabled/default

# Restart Nginx
echo "Restarting Nginx..."
sudo systemctl restart nginx

# Remove the virtual environment
echo "Removing virtual environment..."
rm -rf $VENV_PATH

# Remove Nginx
echo "Removing Nginx..."
sudo apt remove -y nginx

# Remove Python venv package
echo "Removing Python venv package..."
sudo apt remove -y python3-venv

# Clean up apt cache
echo "Cleaning up apt cache..."
sudo apt autoremove -y
sudo apt clean

echo "Cleanup process completed."
echo "Note: This script did not remove system updates or other dependencies that might have been installed."
echo "If you want to completely revert the system, consider using a snapshot or backup from before the deployment."