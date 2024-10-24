import streamlit as st
import pandas as pd
from spire.xls import *
from spire.xls.common import *
import os
import tempfile




st.set_page_config(page_title="Excel File Image Extractor", layout="wide")


def process_excel(file):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xls') as temp_file:
        temp_file.write(file.getvalue())
        temp_file_path = temp_file.name

    workbook = Workbook()
    workbook.LoadFromFile(temp_file_path)

    sheet = workbook.Worksheets[0]

    save_dir = "extracted_images"
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    desired_size_in_bytes = 2740
    arrNum = []
    count = 0

    for i in range(sheet.Pictures.Count):
        try:
            pic = sheet.Pictures[i]
            
            if isinstance(pic, XlsBitmapShape):
                image_stream = pic.Picture
                image_data = image_stream.ToArray()
                image_size_in_bytes = len(image_data)

                if image_size_in_bytes == desired_size_in_bytes:
                    arrNum.append(count)
                    if count < 3:
                        count += 1
                    else:
                        count = 0
                    
                    image_filename = f"image_{i}_size_{image_size_in_bytes}.png"
                    image_path = os.path.join(save_dir, image_filename)
                    
                    with open(image_path, 'wb') as f:
                        f.write(image_data)
                else:
                    if count < 3:
                        count += 1
                    else:
                        count = 0
            else:
                st.warning(f"Picture {i} is not of type XlsBitmapShape")
        except Exception as e:
            st.error(f"Error processing picture {i}: {e}")

    workbook.Dispose()

    df = pd.read_excel(temp_file_path, skiprows=8)  
    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')
    df = df.drop(0)  
    if 'Unnamed: 6' in df.columns:
        df.drop(columns=['Unnamed: 6'], inplace=True)
    df = df.dropna(axis=1, how='all')
    df = df.dropna(axis=0, how='all')

    df['Response'] = arrNum

    return df

st.title("Excel File Image Extractor")

uploaded_file = st.file_uploader("Choose an Excel file", type=["xls"])

if uploaded_file:
    df = process_excel(uploaded_file)
    st.dataframe(df, use_container_width=True)
 