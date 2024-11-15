#!/usr/bin/env python3
import tkinter as tk
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import scrolledtext, filedialog, messagebox
from docx import Document
import zipfile

# Helper functions
def split_into_sentences(text):
    sentences = text.split('. ')
    return [sentence.strip() for sentence in sentences if sentence]

def find_duplicates(items):
    duplicates = []
    seen = set()
    for item in items:
        if item in seen:
            duplicates.append(item)
        seen.add(item)
    return duplicates

def highlight_duplicates(duplicates, text_widget):
    for duplicate in duplicates:
        start = 1.0
        while True:
            start = text_widget.search(duplicate, start, stopindex=tk.END)
            if not start:
                break
            end = f"{start}+{len(duplicate)}c"
            text_widget.tag_add("highlight", start, end)
            start = end

def process_text():
    text = text_input.get("1.0", tk.END)
    sentences = split_into_sentences(text)
    paragraphs = text.split('\n\n')


    duplicate_sentences = find_duplicates(sentences)
    duplicate_paragraphs = find_duplicates(paragraphs)


    text_input.tag_remove("highlight", "1.0", tk.END)


    highlight_duplicates(duplicate_sentences, text_input)
    highlight_duplicates(duplicate_paragraphs, text_input)


    messagebox.showinfo("Duplicates Found", f"Found {len(duplicate_sentences)} duplicate sentences and {len(duplicate_paragraphs)} duplicate paragraphs.")

def load_file_content(filepath):
    if filepath.endswith(".docx"):
        doc = Document(filepath)
        return "\n\n".join([p.text for p in doc.paragraphs if p.text.strip()])
    elif filepath.endswith(".pages"):
        with zipfile.ZipFile(filepath, 'r') as zip_ref:
            for file in zip_ref.namelist():
                if file.endswith('document.xml'):
                    with zip_ref.open(file) as f:
                        return f.read().decode('utf-8')
    else:
        messagebox.showerror("Unsupported Format", "Please select a DOCX or Pages file.")
        return ""

def open_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx"), ("Pages Documents", "*.pages")])
    if file_path:
        text = load_file_content(file_path)
        if text:
            text_input.delete("1.0", tk.END)
            text_input.insert(tk.END, text)

def on_drop(event):
    file_path = event.data
    text = load_file_content(file_path)
    if text:
        text_input.delete("1.0", tk.END)
        text_input.insert(tk.END, text)

# GUI Setup
app = TkinterDnD.Tk()  # Using TkinterDnD for drag-and-drop
app.title("Duplicate Finder")
app.geometry("600x500")


app.drop_target_register(DND_FILES)
app.dnd_bind('<<Drop>>', on_drop)


file_button = tk.Button(app, text="Upload File", command=open_file, font=("Arial", 12), bg="lightblue")
file_button.pack(pady=5)


text_input = scrolledtext.ScrolledText(app, wrap=tk.WORD, font=("Arial", 12))
text_input.pack(expand=True, fill="both", padx=10, pady=10)


find_button = tk.Button(app, text="Find Duplicates", command=process_text, font=("Arial", 12), bg="lightgreen")
find_button.pack(pady=5)


text_input.tag_configure("highlight", background="yellow", foreground="black")

app.mainloop()
