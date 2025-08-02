# Document Translator

Simple Python GUI application to translate `.odt` and `.docx` documents between languages using Google Translator.

## Features
- Supports `.odt` (OpenDocument Text) and `.docx` (Word) formats.  
- Preserves formatting in Word documents during translation.  
- Easy file selection via graphical interface.  
- Choose source and target languages from a predefined list.  
- Uses `deep-translator` library for Google Translate API.

## Requirements
- Python 3.6+  
- Packages:
  - `tkinter` (usually included with Python)  
  - `odfpy`  
  - `python-docx`  
  - `deep-translator`
    
Install dependencies with:
 - pip install odfpy python-docx deep-translator

## Usage
1. Run the application
2. Use the GUI to:  
   - Select a source `.odt` or `.docx` file.  
   - Choose the output file name.  
   - Select source and target languages.  
   - Click **Translate** to translate the document.

## Notes
- The `.docx` translator preserves text formatting by translating text parts individually.  
- For `.odt` files, formatting preservation is limited.

