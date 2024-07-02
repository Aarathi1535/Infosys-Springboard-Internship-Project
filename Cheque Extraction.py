!pip install opencv-python
!pip install pytesseract
!pip install pdf2image
!pip install word2number

from pdf2image import convert_from_path
import cv2
import pytesseract
import re
import sqlite3
from word2number import w2n
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\hp\OneDrive\Desktop\My Desktop\tesseract.exe"

def converting_pdf_to_image():
    try:
        pdf_path = r"cheque.pdf"
        images = convert_from_path(pdf_path)
        for i, image in enumerate(images):
            image.save(f"Cheque{i + 1}.jpg", "JPEG")
    except FileNotFoundError:
        print("PDF File not found!!")

def preprocess_image(image):
    gray = cv2.cvtColor(image_open, cv2.COLOR_BGR2GRAY)
    _, thresholded_image = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    return thresholded_image

def cheque_information(text):
    amount = ''
    date = ''
    payee_name = ''
 
    for line in extracted_text.strip().split('\n'):
        line.lower()
        if "rupees " in line:
            amount = line.split("rupees ")[1].strip()
        elif "date " in line:
            date = line.split("date ")[1].strip()
            date  = date.replace(':','-')
        elif "pay. " in line:
            payee_name = line.split("pay. ")[1].strip()  
    
    return amount, date, payee_name

def database_creation(account_number,date,payee_name):
    con = sqlite3.connect("Cheque-Details.db")
    cur = con.cursor()
    cur.execute('''CREATE TABLE IF NOT EXISTS Cheque_infor
                   (id INTEGER PRIMARY KEY,
                    amount TEXT,
                    date TEXT,
                    payee_name TEXT)''')
    cur.execute("INSERT INTO Cheque_infor (amount, date, payee_name) VALUES (?, ?, ?)",
                (account_number, date, payee_name))
    con.commit()
    cur.execute("SELECT * FROM Cheque_infor")
    rows = cur.fetchall()
    for row in rows:
        print(row)
    cur.close()
    con.close()

if __name__ == '__main__':
    converting_pdf_to_image()
    image_open = cv2.imread(r"cheque.jpg")
    preprocessed_image = preprocess_image(image_open)
    extracted_text = pytesseract.image_to_string(preprocessed_image)
    amount, date, payee_name = cheque_information(extracted_text)
    print("Amount:", w2n.word_to_num(amount))
    print("Date:", date)
    print("Payee Name:", payee_name)
    database_creation(w2n.word_to_num(amount),date,payee_name)
