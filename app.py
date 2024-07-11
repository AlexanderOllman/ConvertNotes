from flask import Flask, render_template, request, send_file, jsonify
import os
import zipfile
import shutil
import threading

app = Flask(__name__)
UPLOAD_FOLDER = 'imports'
EXPORT_FOLDER = 'exports'
PROGRESS = {'total': 0, 'processed': 0, 'successful': 0, 'unsuccessful': [], 'done': False}

# Ensure the directories exist
if os.path.exists(UPLOAD_FOLDER):
    shutil.rmtree(UPLOAD_FOLDER)
if os.path.exists(EXPORT_FOLDER):
    shutil.rmtree(EXPORT_FOLDER)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

def count_files(directory):
    return sum([len(files) for r, d, files in os.walk(directory)])

def convert_to_note_threaded(docx_path, export_dir):
    global PROGRESS
    try:
        convert_to_note(docx_path, export_dir)
        PROGRESS['successful'] += 1
    except Exception as e:
        PROGRESS['unsuccessful'].append(os.path.basename(docx_path))
    finally:
        PROGRESS['processed'] += 1

def convert_all_files_threaded(imports, exports):
    global PROGRESS
    # Create the 'exports' directory if it does not exist
    if not os.path.exists(exports):
        os.makedirs(exports)

    # Get all .docx files in the specified directory
    docx_files = [f for f in os.listdir(imports) if f.endswith('.docx')]
    PROGRESS['total'] = len(docx_files)
    
    for docx in docx_files:
        docx_path = os.path.join(imports, docx)
        convert_to_note_threaded(docx_path, exports)
    
    PROGRESS['done'] = True

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    global PROGRESS
    PROGRESS = {'total': 0, 'processed': 0, 'successful': 0, 'unsuccessful': [], 'done': False}
    
    if 'file' not in request.files:
        return 'No file part'
    
    file = request.files['file']
    
    if file.filename == '':
        return 'No selected file'
    
    if file and file.filename.endswith('.zip'):
        file_path = os.path.join(UPLOAD_FOLDER, file.filename)
        file.save(file_path)
        
        # Extract the zip file into the imports folder
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            for member in zip_ref.namelist():
                # Extract only files (not directories)
                if not member.endswith('/'):
                    member_path = os.path.join(UPLOAD_FOLDER, os.path.basename(member))
                    with zip_ref.open(member) as source, open(member_path, "wb") as target:
                        shutil.copyfileobj(source, target)
        
        # Delete the original uploaded zip file
        os.remove(file_path)

        imports_count = count_files(UPLOAD_FOLDER)
        PROGRESS['total'] = imports_count
        
        # Start the file conversion in a separate thread
        thread = threading.Thread(target=convert_all_files_threaded, args=(UPLOAD_FOLDER, EXPORT_FOLDER))
        thread.start()

        return jsonify({"status": "Processing started", "total_files": imports_count})
    else:
        return 'Invalid file type, please upload a .zip file'

@app.route('/progress')
def progress():
    return jsonify(PROGRESS)

@app.route('/download')
def download():
    exports_count = count_files(EXPORT_FOLDER)
    output_zip_path = 'exports.zip'
    
    with zipfile.ZipFile(output_zip_path, 'w') as zipf:
        for root, _, files in os.walk(EXPORT_FOLDER):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), EXPORT_FOLDER))
    
    # Clear the imports and exports folders
    shutil.rmtree(UPLOAD_FOLDER)
    shutil.rmtree(EXPORT_FOLDER)
    os.makedirs(UPLOAD_FOLDER, exist_ok=True)
    os.makedirs(EXPORT_FOLDER, exist_ok=True)
    
    return send_file(output_zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
