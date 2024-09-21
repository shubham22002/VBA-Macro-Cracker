# VBA Password Cracker

## Overview

This is a Flask web application that allows users to upload Microsoft Office documents (specifically `.docm`, `.xlsm`, and `.pptm` files) and attempts to remove VBA project password protection from them. The modified documents can then be downloaded directly from the web interface.

## Features

- Upload Microsoft Office documents with VBA project password protection.
- Automatically extract and modify the `vbaProject.bin` file to remove password protection.
- Repackage the document for download.

## Prerequisites

- Python 3.x
- Flask
- werkzeug

## Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/yourusername/vba-password-cracker.git
   cd vba-password-cracker
   ```

2. Install the required packages:
   ```bash
   pip install Flask
   ```

3. Ensure you have the necessary file permissions for creating the `uploads` directory.

## Usage

1. Run the Flask application:
   ```bash
   python app.py
   ```

2. Open your web browser and go to `http://127.0.0.1:5000`.

3. Upload a `.docm`, `.xlsm`, or `.pptm` file via the web interface.

4. If the password is successfully removed, the modified file will be available for download.

## Code Explanation

- **index()**: Loads an HTML file from the root directory and displays it.
- **upload_file()**: Handles file uploads, processes the file to remove the VBA password, and provides the cracked file for download.
- **crack_vba_password()**: Extracts the VBA project from the uploaded document, modifies it, and repackages the document.
- **extract_vba_project()**: Extracts the `vbaProject.bin` file from the Office document archive.
- **modify_vba_project()**: Modifies the `vbaProject.bin` to remove password protection.
- **repackage_document()**: Repackages the modified files back into a new Office document.

## Error Handling

The application handles various error scenarios such as:
- Invalid file uploads.
- Issues with file extraction or modification.
- Non-Office document uploads.

## Important Notes

- This tool is intended for educational purposes only. Ensure you have the right to modify any documents before using this application.
- The effectiveness of the password removal may vary based on the complexity of the password used.

## License

This project is licensed under the MIT License. See the LICENSE file for details.

## Contributing

Contributions are welcome! Please open an issue or submit a pull request for any features or bug fixes.

## Acknowledgments

- Thanks to the Flask community for the inspiration and resources.
- Special thanks to those who have contributed to the open-source ecosystem.
