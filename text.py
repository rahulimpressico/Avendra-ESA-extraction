import streamlit as st
import pandas as pd
from spire.xls import Workbook, XlsBitmapShape
import os
import tempfile
import numpy as np
import io


st.set_page_config(page_title="Excel File Image Extractor", layout="wide")

dd = ""


def process_excel(file):
    global dd
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xls") as temp_file:
        temp_file.write(file.getvalue())
        temp_file_path = temp_file.name
        dd = temp_file_path

    workbook = Workbook()
    workbook.LoadFromFile(temp_file_path)

    sheet = workbook.Worksheets[0]

    save_dir = "extracted_images"
    if not os.path.exists(save_dir):
        os.makedirs(save_dir)

    desired_size_in_bytes = 2740
    image_list = []

    for i in range(sheet.Pictures.Count):
        try:
            pic = sheet.Pictures[i]
            if isinstance(pic, XlsBitmapShape):
                image_stream = pic.Picture
                image_data = image_stream.ToArray()
                image_size_in_bytes = len(image_data)
                image_list.append(image_size_in_bytes)
                if image_size_in_bytes == desired_size_in_bytes:

                    image_filename = f"image_{i}_size_{image_size_in_bytes}.png"
                    image_path = os.path.join(save_dir, image_filename)

                    with open(image_path, "wb") as f:
                        f.write(image_data)
            else:
                st.warning(f"Picture {i} is not of type XlsBitmapShape")
        except Exception as e:
            st.error(f"Error processing picture {i}: {e}")

    workbook.Dispose()

    df = pd.read_excel(temp_file_path, skiprows=8)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")

    first_row_df = pd.read_excel(temp_file_path, nrows=7)
    site_id = "Site ID Not Found"
    building_name = "Building Name Not Found"

    for row in first_row_df.itertuples():
        row_data = list(row)
        if "School District" in str(row_data[3]):
            name = row_data[3]
            try:
                parts = name.split(";")
                if len(parts) > 1:
                    site_id = parts[0].split("\\")[-1].strip()
                    building_name = parts[1].strip()
            except IndexError:
                pass

    site_ids = [site_id] * len(df)
    building_names = [building_name] * len(df)

    unnamed_6_groups = []
    current_group = []

    for index, row in df.iterrows():
        if not pd.isna(row["No."]):
            if current_group:
                unnamed_6_groups.append(current_group)
            current_group = [np.nan]
        if "Unnamed: 6" in df.columns and not pd.isna(row["Unnamed: 6"]):
            current_group.append(row["Unnamed: 6"])

    if current_group:
        unnamed_6_groups.append(current_group)

   


    Number = df["No."]
    arrHeading = []
    headingCount = 0

    for i in range(1, len(Number)):
        if pd.isna(Number.iloc[i]):
            headingCount += 1
        else:
            if headingCount > 0:
                arrHeading.append(headingCount + 1)
            if headingCount == 0:
                arrHeading.append(None)
            headingCount = 0
    if headingCount > 0:
        arrHeading.append(headingCount)

    arrSubType = []
    j = 0
    for i in range(0, len(arrHeading)):
        if arrHeading[i] is None:
            if j < len(df["Question Text"]):
                if j == 0:
                    arrSubType.append(df["Question Text"].iloc[j])
                    # j+=1
                else:
                    print(df["Question Text"].values)
                    arrSubType.append(df["Question Text"].iloc[j + 1])
                    j += 1
        if arrHeading[i] is not None:

            j += arrHeading[i]

    j = 0

    print()
    print()

    subType = [arrSubType[0]]

    for i in range(1, len(arrHeading)):
        if arrHeading[i] is None:
            j += 1
        subType.append(arrSubType[j])

    df = df.drop(0)
    if "Unnamed: 6" in df.columns:
        df.drop(columns=["Unnamed: 6"], inplace=True)
    df = df.dropna(axis=1, how="all")
    df = df.dropna(axis=0, how="all")

    getlist = []
    index = 0

    for group in unnamed_6_groups[1:]:
        length_of_group = len(group)
        if group == [np.nan]:
            getlist.append([])
        else:
            sliced_data = image_list[index : index + length_of_group]
            getlist.append(sliced_data)
            index += length_of_group
    positions_of_2740 = []
    for sublist in getlist:
        try:
            position1 = max(sublist)
            position = sublist.index(position1)
            positions_of_2740.append(position)
        except ValueError:
            positions_of_2740.append(None)
    df["Response"] = positions_of_2740

    columns = [
        "Site ID",
        "Building Name",
        "No.",
        "Space Type",
        "Support Space Focus",
        "Response",
        "Comment",
    ]
    df_final = pd.DataFrame(columns=columns)
    num_rows = max(len(subType), len(df))
    df_final = df_final.reindex(range(num_rows))

    num_rows = len(df)
    df_final = df_final.reindex(range(num_rows))

    df_final["Site ID"] = site_ids[:num_rows]
    df_final["Building Name"] = building_names[:num_rows]
    df_final["No."] = df["No."].values
    df_final["Space Type"] = subType[:num_rows]
    df_final["Support Space Focus"] = df['Question Text'].values
    df_final["Response"] = df["Response"].values
    df_final["Comment"] = df["Comment"].values

    df_final = df_final.dropna(subset=['Response', 'Comment'], how='all')

    df_final = df_final[~((df_final['Response'].isna()) | (df_final['Response'] == '') &
                        (df_final['Comment'].isna()) | (df_final['Comment'] == ''))]

    return df_final



def convert_excel_to_image(file_path):
    workbook = Workbook()
    workbook.LoadFromFile(file_path)

    sheet = workbook.Worksheets[0]

    sheet.PageSetup.LeftMargin = 0
    sheet.PageSetup.BottomMargin = 0
    sheet.PageSetup.TopMargin = 0
    sheet.PageSetup.RightMargin = 0

    image = sheet.ToImage(
        sheet.FirstRow, sheet.FirstColumn, sheet.LastRow, sheet.LastColumn
    )

    image_path = tempfile.mktemp(suffix=".png")
    image.Save(image_path)

    workbook.Dispose()

    return image_path



col1, col2, col3 = st.columns(3)
with col1:
    st.title("Aramark ESA Data Extraction")
    uploaded_file = st.file_uploader("Choose an Excel file", type=["xls"])
    if uploaded_file: 
        df = process_excel(uploaded_file)
        file_name = uploaded_file.name
        name_of_the_file = file_name.split(".")[0]

        buffer = io.BytesIO()
        with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Processed Data")
        buffer.seek(0)
        st.download_button(
            label="Download Processed Excel File",
            data=buffer,
            file_name=f"{name_of_the_file}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
        st.dataframe(df, use_container_width=True)

with col2:
        st.title("Uploaded Document View")
        if uploaded_file:
            print(dd)
            snapshot_image_path = convert_excel_to_image(dd)
            st.image(
                snapshot_image_path,
                use_column_width=True,
                caption="Image of Uploaded Excel Sheet",
            )



with col3:
    st.title("Combine all the Document into One")
    uploaded_files2 = st.file_uploader("Choose Excel files for Additional Analysis", type=["xlsx"], accept_multiple_files=True, key="file3")
    if uploaded_files2:
        finalexcelsheet = pd.DataFrame()

        for file in uploaded_files2:
            df = pd.concat(pd.read_excel(file, sheet_name=None), ignore_index=True, sort=False)
            finalexcelsheet = pd.concat([finalexcelsheet, df], ignore_index=True)

        st.subheader("Combined Excel Sheet:")
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            finalexcelsheet.to_excel(writer, index=False, sheet_name="Combined Data")
        output.seek(0)
        st.download_button(
            label="Download Combined Excel (.xlsx)",
            data=output,
            file_name="Final_Combined.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
        st.write(finalexcelsheet)



