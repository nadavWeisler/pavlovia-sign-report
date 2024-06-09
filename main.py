import pandas as pd
from docx import Document
from docx.shared import Inches
import base64
from io import BytesIO
from PIL import Image
import os
import tqdm

# Define a class to represent a column in the data
class Column:
    def __init__(self, name, display_name):
        self.name = name
        self.display_name = display_name

# Define the column that contains the unique identifier for each document
ID_COL = "תעודת זהות"

# Define the columns to be included in the document as text
STRING_COLUMNS = [
    Column("שם", "שם"),
    Column("תעודת זהות", "תעודת זהות"),
    Column("סכום", "סכום"),
    Column("responseDate", "תאריך")
]

# Define the columns that contain images
IMAGE_COLUMNS = [
    Column("חתימה", "חתימה")
]

TITLE = "חתימת נבדק" # Title of the document

# Define the size of the image to be inserted into the document
IMAGE_WIDTH = 4.0 # Inches
IMAGE_HEIGHT = 4.0 # Inches

# Define the folder where the results will be saved
RESULTS_FOLDER = "results"

# Create the results folder if it does not exist
if not os.path.exists(RESULTS_FOLDER):
    os.makedirs(RESULTS_FOLDER)

# Function to create a document for a row in the data frame and save it 
def create_doc(row):
    # Create a new document
    doc = Document()

    # Add the title to the document
    doc.add_heading(TITLE, 0)

    # Add the text columns to the document
    for column in STRING_COLUMNS:
        doc.add_paragraph(f"{column.display_name}: {row[column.name]}")

    # Add the image columns to the document 
    for column in IMAGE_COLUMNS:
        try:
            base64_signature = row[column.name]
            base64_signature = base64_signature.replace("data:image/png;base64,", "")
            signature_bytes = base64.b64decode(base64_signature)
            img = Image.open(BytesIO(signature_bytes))

            temp_signature_path = "temp_signature.png"
            img.save(temp_signature_path)

            # Add the image to the document
            doc.add_picture(temp_signature_path, width=Inches(IMAGE_WIDTH), height=Inches(IMAGE_HEIGHT))

            # Clean up: remove the temporary signature image file
            os.remove(temp_signature_path)
        except Exception as e:
            print(e)
            doc.add_paragraph(f"{column.display_name}: {row[column.name]}")
    
    # Save the document to a file with the ID of the row as the filename in the results folder 
    doc.save(f"{RESULTS_FOLDER}/{row[ID_COL]}.docx")

if __name__ == "__main__":
    # Load the data from the CSV file into a pandas DataFrame 
    df = pd.read_csv('recieps.csv')  # Load the data

    # Iterate over the rows in the DataFrame and create a document for each row 
    for index, row in tqdm.tqdm(df.iterrows(), total=df.shape[0]):
        create_doc(row)

    # Open the results folder in the file explorer
    os.startfile(RESULTS_FOLDER)