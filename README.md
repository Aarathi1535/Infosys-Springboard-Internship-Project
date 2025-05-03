# 🏦 Automating Bank Check Extraction from Scanned PDFs

This project, **"Automatic Cheque Extraction from Scanned Document"**, streamlines the process of digitizing and extracting essential information from bank cheques provided in PDF format. The system converts scanned PDF cheques into images, preprocesses them to remove noise using OpenCV, and uses Tesseract OCR to extract text details. Extracted information is then stored in an SQLite3 database for easy access and further use.

> 📌 Developed to eliminate manual entry errors and reduce processing time in financial document digitization.

---

## 🔍 Key Features

✅ Convert scanned cheque PDFs into images  
✅ Preprocess images to enhance clarity and remove noise  
✅ Extract important textual information using OCR  
✅ Store extracted data in a local database (SQLite3)  
✅ Interactive GUI built using Tkinter for ease of use  

---

## 💻 Technologies Used

- **Python** – Core language  
- **pdf2image** – Convert PDF to image  
- **OpenCV** – Image preprocessing  
- **Pytesseract** – Optical Character Recognition  
- **SQLite3** – Lightweight local database  
- **Tkinter** – GUI development

---

## 🔁 Workflow Overview

![Dataflow](dataflow.png)

---
## 🧠 Use Case

This tool can be effectively deployed in banking sectors, accounting offices, or fintech platforms for:
- Automated cheque digitization
- Bulk document processing
- Financial data archiving
- Fraud detection systems (with extended modules)

---
