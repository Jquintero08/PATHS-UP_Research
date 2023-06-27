import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog




def DI_Water():
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes extra window fully transparent
    input_file = filedialog.askopenfilename()  # Opens the file dialog
    output_file = filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])  #Asks user for output file
    root.destroy()  #Destroys window
    
    
    
    
    data = pd.read_excel(input_file, skiprows=27)  #Read input .xlsx file & skip 27 rows like in flip_xlsx

    #Check - Empty cells to stop
    empty_rows = data.index[data.iloc[:, 0].isnull()]
    if len(empty_rows) > 0:
        empty_row = empty_rows[0]
        data = data.iloc[:empty_row]


    di_water_row = data[data.iloc[:, 0] == "DI water"] #Find "DI water" row

    # Transpose and skip the first row which is the header
    di_water_row = di_water_row.transpose()[1:]

    di_water_row.to_excel(output_file, index=False, header=False) #Writing DI water row to new excel file
    
    print("DI Water Excel file has been saved as", output_file)
    
    

def flip_xlsx():
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes extra window fully transparent
    input_file = filedialog.askopenfilename()  #Opens file dialog
    DIWater_file = filedialog.askopenfilename()
    output_file = filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])  #Asks user for output file
    root.destroy()  #Destroys window
    
    
    data = pd.read_excel(input_file, skiprows=27)  # Read input .xlsx file and skip 27 rows

    # Check for empty cells to stop
    empty_rows = data.index[data.iloc[:, 0].isnull()]
    if len(empty_rows) > 0:
        empty_row = empty_rows[0]
        data = data.iloc[:empty_row]

    transposed_data = data.transpose()  #FLips x and y
    transposed_data.columns = [None] * len(transposed_data.columns)  #Remove unnecessary numbering
    transposed_data.to_excel(output_file, index=False, header=False)  #Write to output, keeping index

    # Add average and standard deviation columns
    data2 = pd.read_excel(output_file)
    cols = data2.columns.tolist()

    data3 = pd.DataFrame()

    for i in range(0, len(cols)):
        data3[cols[i]] = data2[cols[i]]
        if i % 3 == 0 and i >= 3:  #Calculate average and std deviation only when there are at least three columns
            data3[f'Average - DI water {i//3}'] = data2.iloc[:, i-2:i+1].mean(axis=1)
            data3[f'Std Dev {i//3}'] = data2.iloc[:, i-2:i+1].std(axis=1)


    DIWater_data = pd.read_excel(DIWater_file, header=None) #Read DI-Water
    
    
    if len(data3.columns) > 4:  # make sure there is a 5th column
        print("Shape of data3:", data3.iloc[1:133, 4].shape)
        print("Shape of di_water_data:", DIWater_data.iloc[1:133, 0].shape)
        data3.iloc[0:133, 4] = data3.iloc[0:133, 4].values - DIWater_data.iloc[0:133, 0].values  #Subtracts row by row

    data3.to_excel(output_file, index=False)
        


    data3.to_excel(output_file, index=False)
    
    print("Flipped Excel file has been saved as", output_file)
    


    
    
    
flip_xlsx();
#DI_Water();




