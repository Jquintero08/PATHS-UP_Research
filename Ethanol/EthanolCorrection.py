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
    filename = filedialog.askopenfilenames()  #Opens the file dialogue
    root.destroy()  #Gets rid of the main window
    filename = list(filename)

    if not filename:  # If there are no selected files, end the function
        print("No files selected.")
        return

    # Process the first file
    df = pd.read_csv(filename[0], skiprows=17, usecols=[0, 1], header=None)

    # Process the remaining files
    for file in filename[1:]:
        new_df = pd.read_csv(file, skiprows=17, usecols=[1],header=None)
        df = pd.concat([df, new_df], axis=1)
        
        
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes the extra window transparent
    #output_dir = filedialog.askdirectory()  #Opens the file dialog (For input in terminal)
    outputFile = filedialog.asksaveasfilename(filetypes=[('Excel file', '*.xlsx')])
    root.destroy()  #Gets rid of the main window

    df.to_excel(outputFile, index=False)
    
    print(df.shape)
    
Combo()