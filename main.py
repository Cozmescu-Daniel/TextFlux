import fitz  # PyMuPDF for PDFs
from googletrans import Translator
import asyncio
import customtkinter as ctk
from tkinter import filedialog
from PIL import Image, ImageTk
import io
import win32com.client as win32
import re  # For regular expression to replace numbers


# Function to extract text from a PDF
def extract_text_from_pdf(pdf_path):
    text_blocks = []

    with fitz.open(pdf_path) as pdf:
        for page_num in range(pdf.page_count):
            page = pdf.load_page(page_num)
            # Extract text blocks from the page
            page_text = page.get_text("text")  # Get plain text without formatting
            text_blocks.append(page_text)

    return text_blocks


# Function to sanitize the text by replacing numbers with 'xxxx'
def sanitize_text(text):
    return re.sub(r'\d+', 'xxxx', text)  # Replace all numbers with 'xxxx'


# Function to generate a thumbnail image from the first page of a PDF
def generate_pdf_thumbnail(pdf_path):
    with fitz.open(pdf_path) as pdf:
        page = pdf.load_page(0)  # Get the first page
        pix = page.get_pixmap()  # Get pixmap of the page
        img_bytes = pix.tobytes("png")  # Convert to PNG byte data
        img = Image.open(io.BytesIO(img_bytes))  # Open as a PIL image
        img.thumbnail((1200, 800))  # Resize to fit in GUI
        return img


# Function to translate text using Google Translate API (async)
async def translate_text(text, src_lang, target_lang):
    sanitized_text = sanitize_text(text)  # Replace numbers with 'xxxx' before translating
    translator = Translator()
    translation = await translator.translate(sanitized_text, src=src_lang, dest=target_lang)
    return translation.text


# Function to handle file selection through a file dialog
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])  # Only allow PDF files
    if file_path:
        file_path_entry.delete(0, ctk.END)  # Clear current path
        file_path_entry.insert(0, file_path)  # Insert selected file path

        # Display file preview based on its type
        display_file_preview(file_path)


# Function to display file preview
def display_file_preview(file_path):
    # Clear previous preview
    for widget in file_preview_frame.winfo_children():
        widget.destroy()  # Remove all previous widgets from the frame

    translation_textbox.delete(1.0, ctk.END)  # Clear the translation textbox

    if file_path.endswith('.pdf'):
        img_list = []
        with fitz.open(file_path) as pdf:
            for page_num in range(pdf.page_count):
                page = pdf.load_page(page_num)
                pix = page.get_pixmap()
                img_bytes = pix.tobytes("png")
                img = Image.open(io.BytesIO(img_bytes))
                img.thumbnail((1200, 800))
                img_list.append(img)

        # Add images to the scrollable frame
        for img in img_list:
            img_tk = ImageTk.PhotoImage(img)
            label = ctk.CTkLabel(file_preview_frame, image=img_tk)
            label.image = img_tk  # Keep a reference to avoid garbage collection
            label.pack(pady=5)

        # Insert text of the first 5 lines from the PDF in the translation box
        text_blocks = extract_text_from_pdf(file_path)
        translation_textbox.insert(ctk.END, '\n'.join(text_blocks[:5]))  # Display first 5 lines


# Function to start the translation
def start_translation():
    file_path = file_path_entry.get()  # Path to the file
    src_lang = src_lang_combobox.get()
    target_lang = target_lang_combobox.get()

    if not file_path:
        result_label.configure(text="Please provide a file.", text_color="white")
        return

    # Extract text based on file type (PDF)
    if file_path.endswith('.pdf'):
        text_blocks = extract_text_from_pdf(file_path)
    else:
        result_label.configure(text="Unsupported file format.", text_color="white")
        return

    if not text_blocks:
        result_label.configure(text="No text found in the file.", text_color="white")
        return

    # Concatenate all extracted text for translation
    extracted_text = " ".join(text_blocks)

    # Run the asynchronous translation function using asyncio
    loop = asyncio.get_event_loop()
    translated_text = loop.run_until_complete(translate_text(extracted_text, src_lang, target_lang))

    # Display the translated text in the translation box
    translation_textbox.delete(1.0, ctk.END)  # Clear previous translation
    translation_textbox.insert(ctk.END, translated_text)

    result_label.configure(text="Translation complete!", text_color="green")


# Function to clear translation when changing language
def change_target_language(event):
    translation_textbox.delete(1.0, ctk.END)  # Clear the translation text when language changes


# Function to open a new Outlook mail and insert the subject, body, and attachment
def open_outlook_mail():
    # Get the translated text from the translation_textbox
    translated_text = translation_textbox.get("1.0", "end-1c")  # Get all text from the textbox
    file_path = file_path_entry.get()  # Get the file path of the PDF

    # Construct the subject and body
    subject = f"INCxxxx - <issue>"
    body = f"Hi, we received {subject} saying :\n\n\"{translated_text}\""

    # Create the Outlook application object
    outlook = win32.Dispatch('Outlook.Application')

    # Create a new email item
    mail = outlook.CreateItem(0)  # 0 corresponds to the MailItem type

    # Set the subject and body
    mail.Subject = subject
    mail.Body = body

    # Attach the original PDF
    mail.Attachments.Add(file_path)  # Attach the original PDF file

    # Display the email (this will open a new draft in Outlook)
    mail.Display()


# Initialize zoom factor (removed since zoom functionality is removed)
zoom_factor = 1.0

# GUI setup with CustomTkinter (Dark Theme)
ctk.set_appearance_mode("dark")  # Enable dark mode
ctk.set_default_color_theme("blue")  # Set default color theme

root = ctk.CTk()
root.title("PDF Translator")
root.state('zoomed')  # Make the window start maximized

# Create two columns: One for file selection and translation, the other for the file preview
root.grid_columnconfigure(0, weight=1)
root.grid_columnconfigure(1, weight=2)

# File path input
file_path_label = ctk.CTkLabel(root, text="Enter file path (PDF only):")
file_path_label.grid(row=0, column=0, padx=10, pady=10)
file_path_entry = ctk.CTkEntry(root, width=300)
file_path_entry.grid(row=0, column=1, padx=10, pady=10)

# Browse button for file selection
browse_button = ctk.CTkButton(root, text="Browse", command=browse_file)
browse_button.grid(row=0, column=2, padx=10, pady=10)

# Language selection dropdowns
src_lang_label = ctk.CTkLabel(root, text="Source Language:")
src_lang_label.grid(row=1, column=0, padx=10, pady=10)
src_lang_combobox = ctk.CTkComboBox(root, values=["en", "es", "fr", "de", "it", "ro"],
                                    width=200)  # Added Romanian language
src_lang_combobox.set("en")
src_lang_combobox.grid(row=1, column=1, padx=10, pady=10)

target_lang_label = ctk.CTkLabel(root, text="Target Language:")
target_lang_label.grid(row=2, column=0, padx=10, pady=10)
target_lang_combobox = ctk.CTkComboBox(root, values=["en", "es", "fr", "de", "it", "ro"],
                                       width=200)  # Added Romanian language
target_lang_combobox.set("es")
target_lang_combobox.grid(row=2, column=1, padx=10, pady=10)

# Preview area for the file content (making it much larger now)
file_preview_frame = ctk.CTkScrollableFrame(root, label_text="File Preview", width=1000,
                                            height=300)  # Scrollable frame for content
file_preview_frame.grid(row=3, column=1, padx=10, pady=10, sticky="nsew")

# Translation area (text box for translated text)
translation_textbox = ctk.CTkTextbox(root, width=1000, height=500)  # Increased height for better text display
translation_textbox.grid(row=3, column=0, padx=10, pady=10)

# Start translation button
start_button = ctk.CTkButton(root, text="Start Translation", command=start_translation)
start_button.grid(row=4, column=0, columnspan=3, pady=20)

# Result label
result_label = ctk.CTkLabel(root, text="", text_color="white")
result_label.grid(row=5, column=0, columnspan=3, pady=10)

# Button to send translated text to Outlook
outlook_button = ctk.CTkButton(root, text="Send via mail", command=open_outlook_mail)
outlook_button.grid(row=4, column=1, columnspan=3, pady=10)

# Event binding to clear preview text when changing target language
target_lang_combobox.bind("<<ComboboxSelected>>", change_target_language)

# Run the GUI
root.mainloop()
