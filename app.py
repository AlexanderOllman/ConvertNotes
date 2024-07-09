from flask import Flask, render_template, request, send_file, jsonify
import threading
import os
import zipfile
import shutil
from PIL import Image 
import base64
import datetime
import re
import hashlib
import aspose.words as aw
import logging

# Setup logging
log_filename = os.path.expanduser('~/logs.txt')
logging.basicConfig(filename=log_filename, level=logging.ERROR, format='%(asctime)s %(message)s')

def log_exception(e):
    logging.error(e, exc_info=True)

def get_hash(image_base64):
    try:
        etag = hashlib.md5()
        etag.update(base64.b64decode(image_base64))
        localHash = etag.hexdigest()
        return localHash
    except Exception as e:
        log_exception(e)
        return None

def image_to_data(image_bytes):
    try:
        image_base64 = base64.b64encode(image_bytes).decode('utf-8')
        hash = get_hash(image_base64)
        return hash, image_base64
    except Exception as e:
        log_exception(e)
        return None, None

def extract_datetime_from_filename(filename):
    try:
        match = re.match(r'(.+)_(\d{6})_(\d{6}).*\.docx', filename)
        if match:
            _, date_str, time_str = match.groups()
            date_time = datetime.datetime.strptime(date_str + time_str, '%y%m%d%H%M%S')
            return date_time.strftime('%Y%m%dT%H%M%S') + 'Z'
        else:
            return None
    except Exception as e:
        log_exception(e)
        return None

def process_document(document):
    try:
        doc = aw.Document(document)
        sections = doc.get_child_nodes(aw.NodeType.ANY, True)
        shapes_count = doc.get_child_nodes(aw.NodeType.SHAPE, True).count
        shapeIndex = 0
        document_array = []

        for section in sections:
            section_type = aw.Node.node_type_to_string(section.node_type)
            if section_type == "Shape":
                if shapeIndex < shapes_count - 1: 
                    shape = section.as_shape()
                    if (shape.has_image):
                        image_bytes = shape.image_data.image_bytes
                        document_array.append({
                            "type": "image",
                            "content": image_bytes,
                            "size": [shape.height, shape.width]
                        })
                    shapeIndex += 1 
            elif section_type == "Paragraph":
                raw_text = section.get_text().strip()
                if raw_text != "":
                    if not "Aspose.Word" in raw_text:
                        document_array.append({
                            "type": "text",
                            "content": raw_text
                        })
        return document_array
    except Exception as e:
        log_exception(e)
        return []

def format_image(image):
    try:
        image_bytes = image["content"]
        [height, width] = image["size"]
        height = round(height)
        width = round(width)
        hash, image = image_to_data(image_bytes)
        tag = f'<en-media hash="{hash}" type="image/png" style="--en-naturalWidth:{width}; --en-naturalHeight:{height};" />'
        
        resource = f'''<resource>
<data encoding="base64">
{image}
</data>
<mime>image/png</mime>
<width>{width}</width>
<height>{height}</height>
</resource>
'''
        return tag, resource
    except Exception as e:
        log_exception(e)
        return None, None

def format_text(text):
    try:
        content = text["content"]
        tag = f"<div>{content}</div><div><br/></div>"
        return tag
    except Exception as e:
        log_exception(e)
        return ""

def get_title(text):
    try:
        input_string = text["content"]
        match = re.search(r'[\n!?]|\. ', input_string)
        if not match:
            title = input_string
        else:
            title = input_string[:match.start()]
        if title.endswith('.'):
            title = title[:-1]
        title = title.replace('/', ' or ')
        if len(title) >= 80:
            last_space_index = title.rfind(' ', 0, 80)
            if last_space_index != -1:
                title = title[:last_space_index]
            else:
                title = title[:80]
        return title
    except Exception as e:
        log_exception(e)
        return "Untitled"

def time_title(timezone_string):
    try:
        dt = datetime.datetime.strptime(timezone_string, '%Y%m%dT%H%M%S' + 'Z')
        date_string = dt.strftime('%d-%m-%Y')
        return date_string
    except Exception as e:
        log_exception(e)
        return "Untitled"

def convert_document(document):
    try:
        tags = []
        resources = []
        title = ""
        title_saved = False
        sections = process_document(document)
        for section in sections:
            if section["type"] == "image":
                image_tag, image_resource = format_image(section) 
                tags.append(image_tag)
                resources.append(image_resource)
            elif section["type"] == "text":
                if not title_saved:
                    title = get_title(section)
                    title_saved = True
                text_tag = format_text(section)
                tags.append(text_tag)
        return title, tags, resources
    except Exception as e:
        log_exception(e)
        return "", [], []

def generate_xml(timestamp, title, tags, resources):
    try:
        tag_string = "".join(tags)
        resource_string = "".join(resources)
        xml = f'''<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE en-export SYSTEM "http://xml.evernote.com/pub/evernote-export4.dtd">
<en-export export-date="{timestamp}" application="Evernote" version="10.88.4">
<note>
<title>{title}</title>
<created>{timestamp}</created>
<updated>{timestamp}</updated>
<content>
<![CDATA[<?xml version="1.0" encoding="UTF-8" standalone="no"?>
<!DOCTYPE en-note SYSTEM "http://xml.evernote.com/pub/enml2.dtd"><en-note> {tag_string} </en-note>     ]]>
</content>
{resource_string}
</note>
</en-export>
'''
        return xml
    except Exception as e:
        log_exception(e)
        return ""

def convert_to_note(document, export_dir):
    try:
        if not os.path.exists(export_dir):
            os.makedirs(export_dir)
        timestamp = extract_datetime_from_filename(document)
        title, tags, resources = convert_document(document)
        if title == "":
            title = time_title(timestamp)
        xml = generate_xml(timestamp, title, tags, resources)
        enex_filename = os.path.join(export_dir, f"{title}.enex")
        with open(enex_filename, 'w') as enex_file:
            enex_file.write(xml)
            print(f"XML content saved as {enex_filename}")
        print(f"File {document} successfully converted.")
    except Exception as e:
        log_exception(e)

def convert_all_files(imports, exports):
    try:
        if not os.path.exists(exports):
            os.makedirs(exports)
        docx_files = [f for f in os.listdir(imports) if f.endswith('.docx')]
        print(len(docx_files))
        for docx in docx_files:
            docx_path = os.path.join(imports, docx)
            convert_to_note(docx_path, exports)
        print(f"{str(len(docx_files))} successfully converted.")
    except Exception as e:
        log_exception(e)

app = Flask(__name__)
UPLOAD_FOLDER = 'imports'
EXPORT_FOLDER = 'exports'
PROGRESS = {'total': 0, 'processed': 0, 'successful': 0, 'unsuccessful': [], 'done': False}

if os.path.exists(UPLOAD_FOLDER):
    shutil.rmtree(UPLOAD_FOLDER)
if os.path.exists(EXPORT_FOLDER):
    shutil.rmtree(EXPORT_FOLDER)
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(EXPORT_FOLDER, exist_ok=True)

def count_files(directory):
    try:
        return sum([len(files) for r, d, files in os.walk(directory)])
    except Exception as e:
        log_exception(e)
        return 0

def convert_to_note_threaded(docx_path, export_dir):
    global PROGRESS
    try:
        convert_to_note(docx_path, export_dir)
        PROGRESS['successful'] += 1
    except Exception as e:
        log_exception(e)
        PROGRESS['unsuccessful'].append(os.path.basename(docx_path))
    finally:
        PROGRESS['processed'] += 1

def convert_all_files_threaded(imports, exports):
    global PROGRESS
    try:
        if not os.path.exists(exports):
            os.makedirs(exports)
        docx_files = [f for f in os.listdir(imports) if f.endswith('.docx')]
        PROGRESS['total'] = len(docx_files)
        for docx in docx_files:
            docx_path = os.path.join(imports, docx)
            convert_to_note_threaded(docx_path, exports)
        PROGRESS['done'] = True
    except Exception as e:
        log_exception(e)

@app.route('/')
def index():
    try:
        return render_template('index.html')
    except Exception as e:
        log_exception(e)
        return "An error occurred."

@app.route('/upload', methods=['POST'])
def upload_file():
    global PROGRESS
    PROGRESS = {'total': 0, 'processed': 0, 'successful': 0, 'unsuccessful': [], 'done': False}
    try:
        if 'file' not in request.files:
            return 'No file part'
        file = request.files['file']
        if file.filename == '':
            return 'No selected file'
        if file and file.filename.endswith('.zip'):
            file_path = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(file_path)
            with zipfile.ZipFile(file_path, 'r') as zip_ref:
                for member in zip_ref.namelist():
                    if not member.endswith('/'):
                        member_path = os.path.join(UPLOAD_FOLDER, os.path.basename(member))
                        with zip_ref.open(member) as source, open(member_path, "wb") as target:
                            shutil.copyfileobj(source, target)
            os.remove(file_path)
            imports_count = count_files(UPLOAD_FOLDER)
            PROGRESS['total'] = imports_count
            thread = threading.Thread(target=convert_all_files_threaded, args=(UPLOAD_FOLDER, EXPORT_FOLDER))
            thread.start()
            return jsonify({"status": "Processing started", "total_files": imports_count})
        else:
            return 'Invalid file type, please upload a .zip file'
    except Exception as e:
        log_exception(e)
        return "An error occurred."

@app.route('/progress')
def progress():
    try:
        return jsonify(PROGRESS)
    except Exception as e:
        log_exception(e)
        return jsonify({"error": "An error occurred."})

@app.route('/download')
def download():
    try:
        exports_count = count_files(EXPORT_FOLDER)
        output_zip_path = 'exports.zip'
        with zipfile.ZipFile(output_zip_path, 'w') as zipf:
            for root, _, files in os.walk(EXPORT_FOLDER):
                for file in files:
                    zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), EXPORT_FOLDER))
        shutil.rmtree(UPLOAD_FOLDER)
        shutil.rmtree(EXPORT_FOLDER)
        os.makedirs(UPLOAD_FOLDER, exist_ok=True)
        os.makedirs(EXPORT_FOLDER, exist_ok=True)
        return send_file(output_zip_path, as_attachment=True)
    except Exception as e:
        log_exception(e)
        return "An error occurred."

if __name__ == '__main__':
    app.run(debug=True)
