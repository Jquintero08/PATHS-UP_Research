import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import argrelextrema
import tkinter as tk
from tkinter import filedialog
import PySimpleGUI as sg
from tkinter import messagebox


def Combo():
    
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes the extra window fully transparent
    
    try:
        filename = filedialog.askopenfilenames()  #User picks which files
        if not filename:  #If no file selected
            raise ValueError('No files selected.')

    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Destroys window
        filename = list(filename) #Convers to list from Tuple
    

    df = pd.read_csv(filename[0], skiprows=17, usecols=[0, 1], header=None) #Process first file skipping 17 rows from machine


    for file in filename[1:]: #Keep processing the remaining files
        new_df = pd.read_csv(file, skiprows=17, usecols=[1],header=None)
        df = pd.concat([df, new_df], axis=1)
        
        
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes the extra window transparent
    
    
    try:
        outputFile = filedialog.asksaveasfilename(filetypes=[('Excel file', '*.xlsx')]) #Asks user for output file
        if not outputFile:  #If no file given to be created
            raise ValueError('No output file given.')
    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Gets rid of main window
    

    df.to_excel(outputFile, index=False, header=False)
    
    print("Success!")
    

def percentChange():
    
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #No extra window will show
    try:
        ethanolInput = filedialog.askopenfilename()  #Opens compiled ethanol file
        if not ethanolInput:  #If no file selected
            raise ValueError('No ethanol data file selected.')

        sampleInput = filedialog.askopenfilename()  #Opens file to shift
        if not sampleInput:  #If no file selected
            raise ValueError('No sample data file selected.')

        outputFile = filedialog.asksaveasfilename(filetypes=[('Excel files', '*.xlsx'), ('CSV files', '*.csv')])  #Asks user for output file
        if not outputFile:  #If no file given to be created
            raise ValueError('No output file given.')
    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Destroys window
    

    df_ethanol = pd.read_excel(ethanolInput, header=None) #Inputted excel files read
    df_sample = pd.read_excel(sampleInput, header=None)

    x_axis = df_sample.iloc[:, 0] #Save wavelength column by itself so no data analysis done

    df_ethanol = df_ethanol.iloc[:, 1:] #First column taken out of data comparison because it's the first day
    df_sample = df_sample.iloc[:, 1:]

    first_day_max = df_ethanol.iloc[:, 0].max() #Finds the max of the first day

    for day in range(1, df_ethanol.shape[1]): #Go through leftover columns

        current_day_max = df_ethanol.iloc[:, day].max() #Max of the current day
        percent_change = (current_day_max / first_day_max) #Calculate percent change from day 1 to current day


        df_sample.iloc[:, day] *= percent_change #Multiply sample data by the percent change to compensate for external factors on sample

    df_sample = pd.concat([x_axis, df_sample], axis=1) #Concatenate the wavelength column and data columns to the file

    df_sample.to_excel(outputFile, index=False, header=False) #Export to outputFile
    
    print("Success!")



def create_popup(): #Uses PySimpleGUI
    layout = [[sg.Button('Continue', key='-CONTINUE-', size=(10,2))],[sg.Button('Exit', key='-EXIT-', size=(10,2))]]

    window = sg.Window('Info', layout)

    while True:
        event, values = window.read()
        if event == sg.WINDOW_CLOSED or event == '-EXIT-':
            window.close()
            print("Exiting...")
            return False #Return False = Exit code
        elif event == '-CONTINUE-':
            window.close()
            return True #Return True = Continue doing more 
        else:
            sg.popup_cancel('User aborted')
            return False
        
        
        
#Add Data Analysis Functionality Later
        
        
        

def main(): 

    loop = True
    while loop:
        layout = [[sg.Text('Select one->'), sg.Listbox(['Combine Files', 'Data Shift'], size=(20, 3), key='LB')],[sg.Button('Ok'), sg.Button('Cancel')]]

        window = sg.Window('Choose an option', layout)
        event, values = window.read(close=True)

        if event == 'Ok':
            if 'Combine Files' in values['LB']:
                
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo("Info", "Select as many files as you'd like in the same folder to combine") #Show Info
                #Comment out above messagebox if you know what to do
                root.destroy() #Close
                
                print("Combining files...")
                Combo() #Runs the Combo function
                loop = create_popup()

            elif 'Data Shift' in values['LB']:
                
                root = tk.Tk()
                root.withdraw()
                messagebox.showinfo("Info", "1st - Compiled Ethanol File\n2nd - Compiled Sample Data File\n3rd - Name the output file with .xlsx extension") #Show Info
                #Comment out above messagebox if you know what to do
                root.destroy() #Close
                
                print("Performing data shift...")
                percentChange() #Runs the Analysis function
                loop = create_popup()
        else:
            sg.popup_cancel('User aborted')
            loop = False




if __name__ == "__main__":
    main()







