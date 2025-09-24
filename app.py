# @title Universal File-to-Text Converter
# This script is designed to run in a Streamlit environment.
# It provides functions to upload files, convert them to plain text or Markdown,
# and then provides a preview and a download link.

import os
import io
import zipfile
import subprocess
import streamlit as st
from pathlib import Path

# The following imports are removed as they are not needed in the Streamlit app.
# from google.colab import files
# from IPython.display import display, HTML

# The subprocess calls for installation are also removed, as these are handled
# by the requirements.txt and packages.txt files in Streamlit Cloud.

def upload_file():
  """
  Provides a file upload widget for the user in a Streamlit environment.
  Returns the file path of the uploaded file.
  """
  st.title("Universal File-to-Text Converter")
  uploaded_file = st.file_uploader("Upload your file(s) (Word, Excel, PPTX, HTML, ZIP, etc.)", type=None)
  if uploaded_file is not None:
    # Save the uploaded file to a temporary location
    file_path = uploaded_file.name
    with open(file_path, "wb") as f:
        f.write(uploaded_file.getbuffer())
    return file_path
  else:
    return None

def universal_file_converter(file_path):
  """
  Converts a wide variety of file types into a single Markdown/text string.

  Args:
    file_path (str): The path to the file to be converted.

  Returns:
    tuple: A tuple containing the converted text (str) and a message (str).
  """
  # Determine the file type based on its extension
  file_extension = Path(file_path).suffix.lower()

  # Output placeholder
  output_text = ""
  message = f"Successfully converted {file_path} to Markdown."

  try:
    if file_extension in ['.docx', '.pptx', '.odt', '.rtf']:
      # Use pypandoc for direct document conversion
      import pypandoc
      output_text = pypandoc.convert_file(file_path, to='markdown_strict')
    elif file_extension in ['.xlsx']:
      # Handle Excel files using openpyxl
      from openpyxl import load_workbook
      workbook = load_workbook(file_path)
      for sheet_name in workbook.sheetnames:
        sheet = workbook[sheet_name]
        output_text += f"\n\n# Excel Sheet: {sheet_name}\n"
        for row in sheet.iter_rows():
          row_text = ' | '.join([str(cell.value) if cell.value is not None else '' for cell in row])
          output_text += row_text + '\n'
    elif file_extension in ['.html', '.htm']:
      # Use BeautifulSoup for HTML files
      from bs4 import BeautifulSoup
      with open(file_path, 'r', encoding='utf-8') as f:
        html_content = f.read()
      soup = BeautifulSoup(html_content, 'html.parser')
      # Extract all text, clean up extra spaces and newlines
      text_content = soup.get_text()
      output_text = os.linesep.join([s for s in text_content.splitlines() if s])
    elif file_extension == '.zip':
      # Handle zip files by extracting and converting each file
      with zipfile.ZipFile(file_path, 'r') as zip_ref:
        for member in zip_ref.infolist():
          if not member.is_dir():
            extracted_path = zip_ref.extract(member)
            converted_content, _ = universal_file_converter(extracted_path)
            output_text += f"\n\n-- ZIP Archive: {member.filename} --\n"
            output_text += converted_content
            os.remove(extracted_path) # Clean up extracted file
    elif file_extension == '.pdf':
      message = "PDF conversion requires additional, complex packages. This function does not support it."
      output_text = ""
    else:
      # For all other text-based files, read them directly
      with open(file_path, 'r', encoding='utf-8') as f:
        output_text = f.read()

  except Exception as e:
    message = f"An error occurred during conversion: {e}"
    output_text = ""

  return output_text, message

def display_and_download_output(text, file_name):
  """
  Displays a preview of the converted text and provides a download link.

  Args:
    text (str): The converted text to display.
    file_name (str): The original file name used for the download link.
  """
  if not text:
    st.write("No content to display or download.")
    return

  # Display the first 1000 characters as a preview
  preview_text = text[:1000] + ('...' if len(text) > 1000 else '')
  st.subheader("Preview (first 1000 characters)")
  st.code(preview_text)

  # Create a download link for the full text
  download_file_name = f"{Path(file_name).stem}_converted.md"
  st.download_button(
      label="Download Full Text",
      data=text,
      file_name=download_file_name,
      mime="text/markdown"
  )

# --- Main execution block ---
if __name__ == "__main__":
  uploaded_file_path = upload_file()
  if uploaded_file_path:
    converted_text, status_message = universal_file_converter(uploaded_file_path)
    st.write(status_message)
    display_and_download_output(converted_text, uploaded_file_path)

