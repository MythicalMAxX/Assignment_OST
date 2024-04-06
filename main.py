# Importing necessary libraries
import os
import pdfplumber
import docx
import re
import pandas as pd

# Function to extract information from a PDF file
def Extract_Info_From_Pdf(File_Path):
    # Open the PDF file
    with pdfplumber.open(File_Path) as Pdf:
        Text = ''
        # Loop through each page in the PDF
        for Page in Pdf.pages:
            # Extract the text from the page and add it to the Text variable
            Text += Page.extract_text()

        # Use regular expressions to find the email and phone number in the text
        Email = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', Text)
        Phone = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', Text)

        # Return a dictionary with the extracted email, phone number, and text
        return {
            'email': Email.group() if Email else None,
            'phone': Phone.group() if Phone else None,
            'text': Text
        }

# Function to extract information from a Word file
def Extract_Info_From_Docx(File_Path):
    # Open the Word file
    Doc = docx.Document(File_Path)
    # Extract the text from the document
    Text = ' '.join([Paragraph.text for Paragraph in Doc.paragraphs])

    # Use regular expressions to find the email and phone number in the text
    Email = re.search(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b', Text)
    Phone = re.search(r'\(?\d{3}\)?[-.\s]?\d{3}[-.\s]?\d{4}', Text)

    # Return a dictionary with the extracted email, phone number, and text
    return {
        'email': Email.group() if Email else None,
        'phone': Phone.group() if Phone else None,
        'text': Text
    }

# Function to write the extracted information to an Excel file
def Write_To_Excel(Data, Output_Path):
    # Convert the data to a pandas DataFrame
    Df = pd.DataFrame(Data)
    # Write the DataFrame to an Excel file
    Df.to_excel(Output_Path, index=False)

# Path to the folder containing the CVs
Folder_Path = 'CV/Sample2/'

# Initialize an empty list to store the extracted information
Data = []
# Loop through each file in the folder
for File_Name in os.listdir(Folder_Path):
    # If the file is a PDF, use the PDF extraction function
    if File_Name.endswith('.pdf'):
        Data.append(Extract_Info_From_Pdf(os.path.join(Folder_Path, File_Name)))
    # If the file is a Word document, use the Word extraction function
    elif File_Name.endswith('.docx'):
        Data.append(Extract_Info_From_Docx(os.path.join(Folder_Path, File_Name)))

# Write the extracted information to an Excel file
Write_To_Excel(Data, 'output.xlsx')