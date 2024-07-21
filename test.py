import streamlit as st
import pandas as pd
import openpyxl
import os
import zipfile
import shutil
from zipfile import ZipFile
from PIL import Image

def extract_all_images_sequential(file_path, output_dir):
    with zipfile.ZipFile(file_path, 'r') as archive:
        # Ensure the output directory exists
        os.makedirs(output_dir, exist_ok=True)
        
        image_files = [f for f in archive.namelist() if f.startswith('xl/media/')]
        images_info = []
        
        for i, image_file in enumerate(image_files, start=1):
            image_name = f"{i}.jpeg"
            image_path = os.path.join(output_dir, image_name)
            with open(image_path, 'wb') as img_file:
                img_file.write(archive.read(image_file))
            images_info.append({'image_name': image_name, 'image_path': image_path})
            print(f"Extracted {image_name}")
        
        return images_info

def rename_images_based_on_sheet(file_path, output_dir):
    # Load the Excel sheet to get the names
    excel_data = pd.read_excel(file_path, sheet_name=0)
    
    # Extract all images sequentially from the provided Excel file
    extracted_images = extract_all_images_sequential(file_path, output_dir)
    
    # Rename images based on the names in the Excel sheet
    for idx, row in excel_data.iterrows():
        name = row['Name']
        if pd.notna(name):
            old_image_path = os.path.join(output_dir, f"{idx + 1}.jpeg")
            new_image_path = os.path.join(output_dir, f"{name}.jpeg")
            if os.path.exists(old_image_path):
                os.rename(old_image_path, new_image_path)
                print(f"Renamed {old_image_path} to {new_image_path}")

# Streamlit app
def main():
    st.title("Excel Image Extractor")
    
    uploaded_file = st.file_uploader("Upload an Excel file", type=["xlsx", "xls","csv"])
    
    if uploaded_file is not None:
        # Create a temporary directory for the uploaded file
        temp_dir = "temp"
        if not os.path.exists(temp_dir):
            os.makedirs(temp_dir)
        
        # Save the uploaded file to disk
        file_path = os.path.join(temp_dir, uploaded_file.name)
        with open(file_path, "wb") as f:
            f.write(uploaded_file.getbuffer())
        
        # Create a temporary directory for extracted images
        output_dir = os.path.join(temp_dir, "extracted_images")
        if os.path.exists(output_dir):
            shutil.rmtree(output_dir)
        
        os.makedirs(output_dir, exist_ok=True)
        
        # Process the uploaded file
        rename_images_based_on_sheet(file_path, output_dir)
        
        # Display the extracted images in a grid with two images per row
        st.subheader("Extracted Images")
        images = os.listdir(output_dir)
        col1, col2 = st.columns(2)
        for i, image in enumerate(images):
            image_path = os.path.join(output_dir, image)
            img = Image.open(image_path)
            if i % 2 == 0:
                with col1:
                    st.image(img, caption=image, use_column_width=True)
            else:
                with col2:
                    st.image(img, caption=image, use_column_width=True)
        
        # Create a zip file of the extracted images
        zip_file_path = os.path.join(temp_dir, "images.zip")
        with ZipFile(zip_file_path, 'w') as zipf:
            for root, dirs, files in os.walk(output_dir):
                for file in files:
                    zipf.write(os.path.join(root, file), file)
        
        # Allow the user to download the zip file
        with open(zip_file_path, "rb") as f:
            st.download_button("Download Images as Zip", f, file_name="images.zip")

if __name__ == "__main__":
    main()
