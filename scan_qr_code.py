import cv2
from pyzbar import pyzbar

import xlwings as xw
from datetime import datetime

import time

import streamlit as st

def scan_qr_code(wb, worksheet, peserta):
    # Buka kamera
    cap = cv2.VideoCapture(0)

    while True:
        # Baca frame dari kamera
        ret, frame = cap.read()

        # Pindai QR Code
        barcodes = pyzbar.decode(frame)

        # Tampilkan hasil pemindaian
        if barcodes != []:
            for barcode in barcodes:
                # Ambil data dari QR Code
                data = barcode.data.decode("utf-8")

                # Koordinat kotak QR Code
                (x, y, w, h) = barcode.rect
                
                # Masukan data presensi
                # if cv2.waitKey(1) & 0xFF == ord(' '):
                try:
                    id = peserta.index(data)

                    now = datetime.now()
                    dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

                    if worksheet['B' + str(2 + id)].value == "HADIR":
                        st.write(data + " -- SUDAH HADIR PADA -- " + str(worksheet['C' + str(2 + id)].value))
                    else:
                        worksheet['B' + str(2 + id)].value = "HADIR"
                        worksheet['C' + str(2 + id)].value = dt_string
                        st.write(data + " -- HADIR --"  + dt_string)

                    # Tampilkan kotak dan teks pada QR Code yang terdeteksi
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 255, 0), 2)
                    cv2.putText(frame, data, (x, y - 10), cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 255, 0), 2)
                    
                    time.sleep(1)
                except:
                    # Tampilkan kotak dan teks pada QR Code yang terdeteksi
                    cv2.rectangle(frame, (x, y), (x + w, y + h), (0, 0, 255), 5)
                    cv2.putText(frame, data+"TIDAK TERDAFTAR", (x, y - 10), 
                                cv2.FONT_HERSHEY_SIMPLEX, 0.5, (0, 0, 255), 2)

        # Tampilkan frame
        cv2.imshow("Scan QR Code", frame)

        # Hentikan pemindaian jika tombol 'q' ditekan
        if cv2.waitKey(1) & 0xFF == ord('q'):
            break

    # Tutup kamera dan jendela tampilan
    cap.release()
    cv2.destroyAllWindows()