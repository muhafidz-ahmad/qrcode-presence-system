import streamlit as st
from streamlit_qrcode_scanner import qrcode_scanner

import xlwings as xw
from datetime import datetime


def scan(wb, worksheet, peserta):
    data = qrcode_scanner(key="qrcode_scanner")
    if data:
        try:
            id = peserta.index(data)

            now = datetime.now()
            dt_string = now.strftime("%d/%m/%Y %H:%M:%S")

            if worksheet["B" + str(2 + id)].value == "HADIR":
                st.write(
                    data
                    + " -- SUDAH HADIR PADA -- "
                    + str(worksheet["C" + str(2 + id)].value)
                )
            else:
                worksheet["B" + str(2 + id)].value = "HADIR"
                worksheet["C" + str(2 + id)].value = dt_string
                st.write(data + " -- HADIR --" + dt_string)

        except:
            st.write(data + " -- TIDAK TERDAFTAR --")