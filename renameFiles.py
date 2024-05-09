import os
from tkinter import *
from tkinter import filedialog
import pandas as pd


root = Tk()
root.withdraw()

# Select the folder containing the files to be renamed
filenames = filedialog.askopenfilenames()

# Get a list of all the files in the folder
file_list = [f for f in filenames if f.endswith('.vtk')]

# Sort the files alphabetically
file_list.sort()

# Initialize the new names list
new_names = []

# Rename the files to a series of numbers starting at 1
for i, old_name in enumerate(file_list):
    _, file_extension = os.path.splitext(old_name)
    new_name = f"{i + 1}" + file_extension  # Change the extension as per file type
    os.rename(os.path.join(os.path.dirname(filenames[i]), old_name), os.path.join(os.path.dirname(filenames[i]), new_name))
    new_names.append(new_name)

# Create a dictionary containing the old and new names
data = {'Old Name': file_list, 'New Name': new_names}

# Convert the dictionary to a Pandas DataFrame
df = pd.DataFrame(data)

# Create a new Excel file and write the DataFrame to it
excel_path = os.path.join(os.path.dirname(filenames[0]), 'file_names.xlsx')
df.to_excel(excel_path, index=False)

print('Files renamed and excel sheet created successfully!')
