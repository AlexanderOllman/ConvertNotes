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
import base64
import os

def get_hash(image_base64):
    # Create a hash object
    etag = hashlib.md5()

    # Update the hash object with the base64 decoded data
    etag.update(base64.b64decode(image_base64))

    # Get the hexadecimal digest of the hash
    localHash = etag.hexdigest()

    return localHash

def image_to_data(image_bytes):
    # ext = filename.split('.')[-1]
    # prefix = f'data:image/{ext};base64,'
    # with open(filename, 'rb') as f:
    #     img = f.read()
    image_base64 = base64.b64encode(image_bytes).decode('utf-8')
    hash = get_hash(image_base64)
    return hash, image_base64

def extract_datetime_from_filename(filename):
    match = re.match(r'(.+)_(\d{6})_(\d{6}).*\.docx', filename)
    if match:
        _, date_str, time_str = match.groups()
        date_time = datetime.datetime.strptime(date_str + time_str, '%y%m%d%H%M%S')
        return date_time.strftime('%Y%m%dT%H%M%S') + 'Z'
    else:
        return None

def process_document(document):
    doc = aw.Document(document)

    #Get all sections to iterate through.
    sections = doc.get_child_nodes(aw.NodeType.ANY, True)

    #Count the number of "Shape" sections that will contain images, to remove the last one which is an Apose advert.
    shapes_count = doc.get_child_nodes(aw.NodeType.SHAPE, True).count
    shapeIndex = 0

    #Initialize array for {type, content} pair output.
    document_array = []

    #Iterate through sections.
    for section in sections:
        #Get section type to determine if section contains image or text.
        section_type = aw.Node.node_type_to_string(section.node_type)

        if section_type == "Shape":
            if shapeIndex < shapes_count - 1: 
                shape = section.as_shape()
                #Verify Shape has image. 
                if (shape.has_image):
                    #Convert image to base64 string
                    # image_base64 = base64.b64encode(shape.image_data.image_bytes).decode('utf-8')
                    image_bytes = shape.image_data.image_bytes
                    document_array.append({
                        "type": "image",
                        "content": image_bytes,
                        "size": [shape.height, shape.width]
                    })
                shapeIndex += 1 #Add to count. 

        elif section_type == "Paragraph":

            #Get text and check it is not the Apose library additions.
            raw_text = section.get_text().strip()
            if raw_text != "":
                if not "Aspose.Word" in raw_text:
                    document_array.append({
                        "type": "text",
                        "content": raw_text
                    })

    return document_array

def format_image(image):
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
    # get
    #  the image extension 
    # image_ext = base_image["ext"] 

    return tag, resource

def format_text(text):
    content = text["content"]
    tag = f"<div>{content}</div><div><br/></div>"

    return tag


def get_title(text):
    input_string = text["content"]
    
    # Find the index of the first line break, exclamation mark, question mark, or period followed by a space
    match = re.search(r'[\n!?]|\. ', input_string)
    
    # If no match is found, use the whole string as the title
    if not match:
        title = input_string
    else:
        # Split the string at the first match
        title = input_string[:match.start()]

    # Remove the final period if it exists
    if title.endswith('.'):
        title = title[:-1]
    
    # Remove any "/" characters from the title
    title = title.replace('/', ' or ')
    
    # Ensure the title is less than 30 characters long
    if len(title) >= 80:
        # Find the last space within the first 30 characters
        last_space_index = title.rfind(' ', 0, 80)
        if last_space_index != -1:
            title = title[:last_space_index]
        else:
            title = title[:80]
    
    return title

def time_title(timezone_string):
    # Parse the input string to a datetime object
    dt = datetime.datetime.strptime(timezone_string, '%Y%m%dT%H%M%S' + 'Z')
    
    # Convert the datetime object to a string in the format DD/MM/YYYY
    date_string = dt.strftime('%d-%m-%Y')
    
    return date_string

def convert_document(document):

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

def generate_xml(timestamp, title, tags, resources):

    tag_string = ""
    for tag in tags:
        tag_string += tag

    resource_string = ""
    for resource in resources:
        resource_string += resource

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

def convert_to_note(document, export_dir):
    # Ensure the export directory exists
    if not os.path.exists(export_dir):
        os.makedirs(export_dir)

    timestamp = extract_datetime_from_filename(document)
    title, tags, resources = convert_document(document)
    if title == "":
        title = time_title(timestamp)
        
    xml = generate_xml(timestamp, title, tags, resources)

    enex_filename = os.path.join(export_dir, f"{title}.enex")  # Use title as filename and save in export_dir
    with open(enex_filename, 'w') as enex_file:
        enex_file.write(xml)
        print(f"XML content saved as {enex_filename}")
    
    print(f"File {document} successfully converted.")


def convert_all_files(imports, exports):
    # Create the 'exports' directory if it does not exist
    if not os.path.exists(exports):
        os.makedirs(exports)

    # Get all .docx files in the specified directory
    docx_files = [f for f in os.listdir(imports) if f.endswith('.docx')]
    print(len(docx_files))
    for docx in docx_files:
        docx_path = os.path.join(imports, docx)
        convert_to_note(docx_path, exports)

    print(f"{str(len(docx_files))} successfully converted.")


# Example usage:
# save_docx_titles('/path/to/your/directory')

# convert_all_files("./")
# convert_to_note("Notes_240518_195459.docx")

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