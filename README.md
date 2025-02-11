# PDF Translator Tool 🌍📄  

Welcome to the **TextFlux**! This application allows users to extract text from PDFs, translate it into various languages, and even send the translated content via Outlook email—all with a user-friendly graphical interface.

## Features ✨  
- **Extract text** from PDFs seamlessly.  
- **Translate** extracted text into multiple languages, including Romanian, English, French, German, and more.  
- **Preview the original PDF** within the application.  
- **Send translations via email** with one click, attaching the original PDF automatically.  
- **Modern and intuitive GUI** using CustomTkinter for a sleek user experience.  

## Technologies Used 🛠️  
- **Python** – The core programming language.  
- **CustomTkinter** – A beautiful and modern GUI framework.  
- **PyMuPDF (fitz)** – For PDF text extraction.  
- **Googletrans** – For automatic text translation.  
- **Pillow** – For rendering PDF previews.  
- **pywin32 (win32com.client)** – To automate sending emails via Outlook.  

## Requirements 📋  
To run this project, install the required dependencies:  

```bash
pip install -r requirements.txt
```
How to Use the Tool 🖥️
1️⃣ Clone the repository:
```bash
git clone https://github.com/Cozmescu-Daniel/PDFTranslator.git
```
2️⃣ Navigate to the project directory:
```bash
cd PDFTranslator
```
3️⃣ Run the application:
```bash
python main.py
```
4️⃣ Using the application:
Select a PDF file for translation.
Choose the source and target language.
Click "Start Translation" to process the text.
Click "Send Email" to open Outlook with the translated content and the original PDF attached.

## Example Screenshot 🖼️
![image](https://github.com/user-attachments/assets/6f4c0385-55d8-4bca-ba07-000f5a32b25b)

## Sending Emails via Outlook 📧
The tool can automatically open a new Outlook email with:

Subject: INCxxxx - <issue>
Content:
```Email
Hi, we received INCxxxx - <issue> saying:
"<translated content>"
Attachment: The original PDF file.
```

This action is only triggered when you click the "Send Email" button.
It is added as a particularity of the DevOps teams which deal with tickets/incidents via PDF format and work with Outlook Client mail.

## Customization ✍️
Add more languages by modifying the target_lang_combobox values.
Customize the email subject and body by editing the send_email function.
Improve translation accuracy by switching to a different translation API if needed.

## Contributing 🤝
We welcome contributions! If you’d like to enhance the tool, follow these steps:

Fork the repository.
Create a new branch for your feature.
Make your changes and commit them.
Push your changes to your fork.
Open a pull request to the main branch.

## Acknowledgements 💡
CustomTkinter for providing a modern GUI framework.
PyMuPDF for efficient PDF text extraction.
Googletrans for automatic translations.
pywin32 for Outlook email automation.

Enjoy translating your PDFs with ease! 🚀🌎
