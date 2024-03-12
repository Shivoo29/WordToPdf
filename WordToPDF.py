import os  # Import the os module for operating system functions
import comtypes.client  # Import the comtypes.client module for working with COM objects
import docx  # Import the docx module for working with Word documents

word_path = "word.docx"  # Define the path to the input Word document
pdf_path = "pdf.pdf"  # Define the path for the output PDF document

doc = docx.Document(word_path)  # Open the Word document using the docx module

# Create a new instance of the Word application using COM
word = comtypes.client.CreateObject("Word.Application")

# Get the absolute path of the input Word document
docx_path = os.path.abspath(word_path)

# Get the absolute path of the output PDF document
pdf_path = os.path.abspath(pdf_path)

# Define the PDF file format (17 corresponds to PDF format)
pdf_format = 17

# Set Word application visibility to False (run in the background)
word.Visible = False

# Open the input Word document using the Word application
in_file = word.Documents.Open(docx_path)

# Save the opened Word document as a PDF file with the specified format
in_file.SaveAs(pdf_path, FileFormat=pdf_format)

# Close the input Word document
in_file.Close()

# Quit the Word application
word.Quit()
