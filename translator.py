import tkinter as tk
from tkinter import filedialog, messagebox
from odf.opendocument import load, OpenDocumentText
from odf.text import P
from deep_translator import GoogleTranslator
from docx import Document
import os


def translate_odt(input_path, output_path, source_lang='pl', target_lang='en'):
    doc = load(input_path)
    body = doc.text
    new_doc = OpenDocumentText()
    translator = GoogleTranslator(source=source_lang, target=target_lang)
    for element in body.childNodes:
        if element.qname[1] == 'p':
            original_text = ''.join([node.data for node in element.childNodes if node.nodeType == 3])
            if original_text.strip():
                try:
                    translated_text = translator.translate(original_text)
                except Exception as e:
                    print(f'Translation error: {e}')
                    translated_text = original_text
                new_p = P(text=translated_text)
                new_doc.text.addElement(new_p)
    new_doc.save(output_path)


def translate_docx_preserve_formatting(input_path, output_path, source_lang='pl', target_lang='en'):
    document = Document(input_path)
    translator = GoogleTranslator(source=source_lang, target=target_lang)
    for para in document.paragraphs:
        for run in para.runs:
            text = run.text.strip()
            if text:
                try:
                    translated_text = translator.translate(text)
                except Exception as e:
                    print(f'Translation error: {e}')
                    translated_text = text
                run.text = translated_text
    document.save(output_path)



def start_translation():
    input_path = entry_input.get()
    output_path = entry_output.get()
    source_lang = source_lang_var.get()
    target_lang = target_lang_var.get()

    if not (input_path.endswith('.odt') or input_path.endswith('.docx')):
        messagebox.showerror("Error", "Please select an .odt or .docx file")
        return
    if source_lang == target_lang:
        messagebox.showwarning("Warning", "Source and target languages are the same.")
        return
    try:
        if input_path.endswith('.odt'):
            translate_odt(input_path, output_path, source_lang, target_lang)
        else:
            translate_docx_preserve_formatting(input_path, output_path, source_lang, target_lang)
        messagebox.showinfo("Success", f"Translated file saved as:\n{output_path}")
    except Exception as e:
        messagebox.showerror("Error", f"An error occurred:\n{e}")


def browse_file():
    filename = filedialog.askopenfilename(
        title="Select ODT or DOCX file",
        filetypes=[("ODT and DOCX files", "*.odt *.docx")])
    if filename:
        entry_input.delete(0, tk.END)
        entry_input.insert(0, filename)
        base, ext = os.path.splitext(filename)
        suggested_output = base + "_translated" + ext
        entry_output.delete(0, tk.END)
        entry_output.insert(0, suggested_output)


languages = {
    'Polish': 'pl',
    'English': 'en',
    'German': 'de',
    'French': 'fr',
    'Spanish': 'es',
    'Italian': 'it',
    'Russian': 'ru',
    'Chinese (Simplified)': 'zh-CN',
    'Japanese': 'ja'
}

root = tk.Tk()
root.title("Document Translator (.odt and .docx)")

tk.Label(root, text="Source file (.odt or .docx):").grid(row=0, column=0, sticky="e", padx=5, pady=5)
entry_input = tk.Entry(root, width=50)
entry_input.grid(row=0, column=1, padx=5, pady=5)
btn_browse = tk.Button(root, text="Browse...", command=browse_file)
btn_browse.grid(row=0, column=2, padx=5, pady=5)

tk.Label(root, text="Output file (.odt or .docx):").grid(row=1, column=0, sticky="e", padx=5, pady=5)
entry_output = tk.Entry(root, width=50)
entry_output.grid(row=1, column=1, padx=5, pady=5)

tk.Label(root, text="Source language:").grid(row=2, column=0, sticky="e", padx=5, pady=5)
source_lang_var = tk.StringVar(value='pl')
source_menu = tk.OptionMenu(root, source_lang_var, *languages.values())
source_menu.grid(row=2, column=1, sticky="w", padx=5, pady=5)

tk.Label(root, text="Target language:").grid(row=3, column=0, sticky="e", padx=5, pady=5)
target_lang_var = tk.StringVar(value='en')
target_menu = tk.OptionMenu(root, target_lang_var, *languages.values())
target_menu.grid(row=3, column=1, sticky="w", padx=5, pady=5)

btn_translate = tk.Button(root, text="Translate", command=start_translation)
btn_translate.grid(row=4, column=1, pady=15)

root.mainloop()
