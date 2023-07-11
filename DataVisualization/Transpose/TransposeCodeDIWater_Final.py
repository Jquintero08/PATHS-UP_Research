import pandas as pd
import numpy as np
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox




def DI_Water():
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes extra window fully transparent
    
    try:
        inputFile = filedialog.askopenfilename()  #Opens compiled ethanol file
        if not inputFile:  #If no file selected
            raise ValueError('No input file selected.')

        outputFile = filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])  #Asks user for output file
        if not outputFile:  #If no file given to be created
            raise ValueError('No output file given.')
    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Destroys window
    
    
    
    data = pd.read_excel(inputFile, skiprows=27)  #Read input .xlsx file & skip 27 rows like in flip_xlsx

    #Check - Empty cells to stop
    empty_rows = data.index[data.iloc[:, 0].isnull()]
    if len(empty_rows) > 0:
        empty_row = empty_rows[0]
        data = data.iloc[:empty_row]


    di_water_row = data[data.iloc[:, 0] == "DI water"] #Find "DI water" row

    
    di_water_row = di_water_row.transpose()[1:] #Transpose and skip first row which is the header

    di_water_row.to_excel(outputFile, index=False, header=False) #Writing DI water row to new excel file
    
    print("DI Water Excel file has been saved as", outputFile)
    
    

def flip_xlsx():
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes extra window fully transparent
    
    
    try:
        inputFile = filedialog.askopenfilename()  #Opens compiled ethanol file
        if not inputFile:  #If no file selected
            raise ValueError('No input file selected.')

        DIWater_file = filedialog.askopenfilename()  #Opens file to shift
        if not DIWater_file:  #If no file selected
            raise ValueError('No DI-Water file selected.')

        outputFile = filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])  #Asks user for output file
        if not outputFile:  #If no file given to be created
            raise ValueError('No output file given.')
    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Destroys window
    
    
    data = pd.read_excel(inputFile, skiprows=27)  # Read input .xlsx file and skip 27 rows

    #Check for empty cells to stop
    empty_rows = data.index[data.iloc[:, 0].isnull()]
    if len(empty_rows) > 0:
        empty_row = empty_rows[0]
        data = data.iloc[:empty_row]

    transposedData = data.transpose()  #FLips x and y
    transposedData.columns = [None] * len(transposedData.columns)  #Remove unnecessary numbering
    transposedData.to_excel(outputFile, index=False, header=False)  #Write to output, keeping index

    #Add average & standard deviation columns
    data2 = pd.read_excel(outputFile)
    cols = data2.columns.tolist()

    data3 = pd.DataFrame()

    for i in range(0, len(cols)):
        data3[cols[i]] = data2[cols[i]]
        if i % 3 == 0 and i >= 3:  #Calculate average and std deviation only when there are at least three columns
            data3[f'Average {i//3}'] = data2.iloc[:, i-2:i+1].mean(axis=1)
            data3[f'Std Dev {i//3}'] = data2.iloc[:, i-2:i+1].std(axis=1)


    DIWater_data = pd.read_excel(DIWater_file, header=None) #Read DI-Water
    
    
    if len(data3.columns) >= 5:  #Make sure there are 5 columns, if not, just print out averages
        for col in np.arange(4, len(data3.columns), 5):
            data3.iloc[0:133, col] = data3.iloc[0:133, col].values - DIWater_data.iloc[0:133, 0].values  #Subtracts row by row

    data3.to_excel(outputFile, index=False)
        

    
    print("Flipped Excel file has been saved as", outputFile)
    


    
def main():
    import PySimpleGUI as sg
    event, values = sg.Window('Choose an option', [[sg.Text('Select one->'), sg.Listbox(['DI-Water Transpose', 'Sample Transpose'], size=(20, 3), key='LB')],
    [sg.Button('Ok'), sg.Button('Cancel')]]).read(close=True)

    if event == 'Ok':
        if 'DI-Water Transpose' in values['LB']:
            
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Info", "Select a File Containing a DI Water Row")  #Show Info
            root.destroy() #Close
            

            DI_Water();

            
        elif 'Sample Transpose' in values['LB']:
            
            root = tk.Tk()
            root.withdraw()
            messagebox.showinfo("Info", "1st File - Sample file to Transpose \n2nd File - Transposed DI Water File \n3rd - Where to output the file")  #Show Info
            root.destroy() #Close
            
            flip_xlsx(); #Run Tranpose Sample code

    else:
        sg.popup_cancel('User aborted')
    



main()



