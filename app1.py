import os
import zipfile
import shutil
import tempfile
from flask import Flask, render_template, request, send_file
from werkzeug.utils import secure_filename
from flask import render_template_string

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'

# Load HTML from root directory instead of the 'templates' folder
@app.route('/')
def index():
    try:
        with open('1.html', 'r') as f:
            html_content = f.read()
        return render_template_string(html_content)
    except Exception as e:
        return f"Error loading HTML: {e}", 500

if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

def extract_vba_project(file_path, extract_dir):
    """Extract the vbaProject.bin file from the Office document archive."""
    try:
        with zipfile.ZipFile(file_path, 'r') as zip_ref:
            zip_ref.extractall(extract_dir)
    except zipfile.BadZipFile:
        raise ValueError("Provided file is not a valid zip archive.")
    except Exception as e:
        raise RuntimeError(f"Error extracting VBA project: {e}")

def modify_vba_project(bin_file_path):
    """Modify the vbaProject.bin to remove the password protection."""
    try:
        with open(bin_file_path, 'rb') as file:
            data = file.read()

        # Replace the DPB and CMG entries with null bytes
        data = data.replace(b'DPB=', b'DPx=')
        data = data.replace(b'CMG=', b'CMx=')
        data = data.replace(b'GC=', b'Gx=')

        with open(bin_file_path, 'wb') as file:
            file.write(data)
    except Exception as e:
        raise RuntimeError(f"Error modifying VBA project: {e}")

def repackage_document(extract_dir, output_file, file_type):
    """Repack the extracted files into a new Office document."""
    try:
        shutil.make_archive(output_file.replace(file_type, ''), 'zip', extract_dir)
        shutil.move(f"{output_file.replace(file_type, '')}.zip", output_file)
    except Exception as e:
        raise RuntimeError(f"Error repackaging document: {e}")

def crack_vba_password(file_path):
    try:
        with tempfile.TemporaryDirectory() as temp_dir:
            file_extension = os.path.splitext(file_path)[1]
            if file_extension not in ['.docm', '.xlsm', '.pptm']:
                return False

            extract_vba_project(file_path, temp_dir)

            # Define the path to the vbaProject.bin file based on file type
            if file_extension == '.docm':
                bin_file_path = os.path.join(temp_dir, 'word', 'vbaProject.bin')
            elif file_extension == '.xlsm':
                bin_file_path = os.path.join(temp_dir, 'xl', 'vbaProject.bin')
            elif file_extension == '.pptm':
                bin_file_path = os.path.join(temp_dir, 'ppt', 'vbaProject.bin')

            if os.path.exists(bin_file_path):
                modify_vba_project(bin_file_path)
            else:
                return False

            # Repackage the document
            cracked_file = file_path.replace(file_extension, f'_cracked{file_extension}')
            repackage_document(temp_dir, cracked_file, file_extension)
            return cracked_file
    except Exception as e:
        print(f"Error cracking VBA password: {e}")
        return False

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400

    filename = secure_filename(file.filename)
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    
    try:
        file.save(file_path)
        cracked_file = crack_vba_password(file_path)
        if cracked_file:
            return send_file(cracked_file, as_attachment=True)
        return "Failed to crack the password", 400
    except Exception as e:
        return f"Error processing file: {e}", 500

if __name__ == "__main__":
    app.run(debug=True)
