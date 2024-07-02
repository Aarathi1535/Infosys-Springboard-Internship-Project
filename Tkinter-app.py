import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox
import csv
import json
import xlsxwriter

from pdf2image import convert_from_path
import cv2
import pytesseract
import re
import sqlite3
from word2number import w2n

# Setting Tesseract path
pytesseract.pytesseract.tesseract_cmd = r"C:\Users\hp\OneDrive\Desktop\My Desktop\tesseract.exe"


def converting_pdf_to_image(pdf_path):
    try:
        images = convert_from_path(pdf_path)
        for i, image in enumerate(images):
            image.save(f"Cheque{i + 1}.jpg", "JPEG")
        messagebox.showinfo("Success", "PDF converted to images successfully!")
    except FileNotFoundError:
        messagebox.showerror("Error", "PDF File not found!!")


def preprocess_image(image):
    gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
    _, thresholded_image = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY | cv2.THRESH_OTSU)
    return thresholded_image


def cheque_information(extracted_text):
    amount = ''
    date = ''
    payee_name = ''

    for line in extracted_text.strip().split('\n'):
        line.lower()
        if "rupees " in line:
            amount = line.split("rupees ")[1].strip()
        elif "bay " in line:
            date = line.split("bay ")[1].strip()
            date = date.replace(':', '-')
        elif "pay. " in line:
            payee_name = line.split("pay. ")[1].strip()

    return amount, date, payee_name


def database_creation(account_number, date, payee_name):
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
    cur.close()
    con.close()


def export_to_csv(data):
    filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if filename:
        with open(filename, 'w', newline='') as csvfile:
            writer = csv.writer(csvfile)
            writer.writerow(['Amount', 'Date', 'Payee Name'])
            writer.writerow(data)
        messagebox.showinfo("Success", "Data exported to CSV successfully!")


def export_to_json(data):
    filename = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json")])
    if filename:
        with open(filename, 'w') as jsonfile:
            json.dump({'Amount': data[0], 'Date': data[1], 'Payee Name': data[2]}, jsonfile)
        messagebox.showinfo("Success", "Data exported to JSON successfully!")


def export_to_excel(data):
    filename = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
    if filename:
        workbook = xlsxwriter.Workbook(filename)
        worksheet = workbook.add_worksheet()
        headers = ['Amount', 'Date', 'Payee Name']
        for col, header in enumerate(headers):
            worksheet.write(0, col, header)
        for col, value in enumerate(data):
            worksheet.write(1, col, value)
        workbook.close()
        messagebox.showinfo("Success", "Data exported to Excel successfully!")


def browse_pdf():
    pdf_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if pdf_path:
        converting_pdf_to_image(pdf_path)


def process_cheque():
    global extracted_text
    image_open = cv2.imread(
        r"C:\Users\hp\OneDrive\Desktop\My Desktop\Infosys Springboard Internship\Web-Development-of-SweetSpot---Delivering-Delight-to-Your-Doorstep_Apr_2024\hdfc2.jpg")
    preprocessed_image = preprocess_image(image_open)
    extracted_text = pytesseract.image_to_string(preprocessed_image)
    amount, date, payee_name = cheque_information(extracted_text)
    database_creation(w2n.word_to_num(amount), date, payee_name)
    data = [w2n.word_to_num(amount), date, payee_name]
    messagebox.showinfo("Success", "Data processed successfuly!!")
    return data


def display_data():
    amount, date, payee_name = cheque_information(extracted_text)
    messagebox.showinfo("Cheque Information", f"Amount: {amount}\nDate: {date}\nPayee Name: {payee_name}")


def convert_to_csv():
    data = process_cheque()
    export_to_csv(data)


def convert_to_json():
    data = process_cheque()
    export_to_json(data)


def convert_to_excel():
    data = process_cheque()
    export_to_excel(data)


root = tk.Tk()
root.title("Cheque Processing App")
root.geometry('400x400')
btn_browse_pdf = tk.Button(root, text="Browse PDF", command=browse_pdf, height=1, width=20, bg="BLACK", fg="WHITE")
btn_browse_pdf.pack(pady=20)

btn_process_cheque = tk.Button(root, text="Process Cheque", command=process_cheque, height=1, width=20, bg="BLACK",
                               fg="WHITE")
btn_process_cheque.pack(pady=20)

btn_display_data = tk.Button(root, text="Display Data", command=display_data, height=1, width=20, bg="BLACK",
                             fg="WHITE")
btn_display_data.pack(pady=20)

btn_convert_to_csv = tk.Button(root, text="Convert to CSV", command=convert_to_csv, height=1, width=20, bg="BLACK",
                               fg="WHITE")
btn_convert_to_csv.pack(pady=20)

btn_convert_to_json = tk.Button(root, text="Convert to JSON", command=convert_to_json, height=1, width=20, bg="BLACK",
                                fg="WHITE")
btn_convert_to_json.pack(pady=20)

btn_convert_to_excel = tk.Button(root, text="Convert to Excel", command=convert_to_excel, height=1, width=20,
                                 bg="BLACK", fg="WHITE")
btn_convert_to_excel.pack(pady=20)
root.configure(background="grey")
root.mainloop()
