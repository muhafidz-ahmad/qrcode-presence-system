import pandas as pd

import streamlit as st
import xlwings as xw

import generate_qr_code
import scan_streamlit


# read and prepare the data
wb = xw.Book("data.xlsx")
worksheet = wb.sheets("Sheet1")

peserta = worksheet["A2"].expand("down").value

total_peserta = len(peserta)

scan, generate, stop = st.tabs(["Scan QR Code", "Generate QR Code", "Stop the Program"])

with stop:
    if st.button("Stop the Program"):
        st.write("Program stopped")
        st.stop()

with scan:
    "# Scan QR Code"
    if st.checkbox("Open Camera"):
        # Panggil fungsi scan_qr_code untuk memulai pemindaian
        scan_streamlit.scan(wb, worksheet, peserta)

with generate:
    "# Generate New QR Code"
    # Data yang ingin diubah menjadi QR Code
    nama = st.text_input("Write you name", "")

    if st.button("Generate QR Code"):
        # Contoh penggunaan
        filename = nama + ".png"  # Nama file untuk menyimpan QR Code

        img = generate_qr_code.generate_qr_code(nama, filename)
        img.save("qrcode/" + filename)

        # Tulis data ke excel
        total_peserta += 1
        worksheet["A" + str(total_peserta + 1)].value = nama

        st.write("QR Code {filename} telah dibuat: ")
        st.image("qrcode/" + filename, caption=filename)
