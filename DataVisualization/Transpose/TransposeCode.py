import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog

def flip_xlsx(input_file, output_file):
    data = pd.read_excel(input_file, skiprows=27)  # Read input .xlsx file and skip 27 rows

    # Check for empty cells to stop
    empty_rows = data.index[data.iloc[:, 0].isnull()]
    if len(empty_rows) > 0:
        empty_row = empty_rows[0]
        data = data.iloc[:empty_row]

    transposed_data = data.transpose()  # Flips y and x axis
    transposed_data.columns = [None] * len(transposed_data.columns)  # Remove unnecessary numbering
    transposed_data.to_excel(output_file, index=False, header=False)  # Write to output, keeping index

    # Add average and standard deviation columns
    data2 = pd.read_excel(output_file)
    cols = data2.columns.tolist()

    data3 = pd.DataFrame()

    for i in range(0, len(cols)):
        data3[cols[i]] = data2[cols[i]]
        if i % 3 == 0 and i >= 3:  # Calculate average and std deviation only when there are at least three columns
            data3[f'Average {i//3}'] = data2.iloc[:, i-2:i+1].mean(axis=1)
            data3[f'Std Dev {i//3}'] = data2.iloc[:, i-2:i+1].std(axis=1)

    data3.to_excel(output_file, index=False)


# Ask user for input and output file names
root = tk.Tk()
root.attributes('-alpha', 0.0)  # Makes the extra window fully transparent
input_file = filedialog.askopenfilename()  # Opens the file dialog
output_file = filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])  # Asks user for output file
root.destroy()  # Gets rid of the main window

flip_xlsx(input_file, output_file)  # Call flip function which flips the file & adds avg & std. dev.

print("Flipped Excel file has been saved as", output_file)