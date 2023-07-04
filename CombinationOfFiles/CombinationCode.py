import pandas as pd

import os
import tkinter as tk
from tkinter import filedialog

########################## Change the Master Folder Destination ##########################
##########################################################################################

#desired_location = (r'C:\Users\Jakey\Desktop\Summer_2023_Research\Combination\2021-05-09')


root = tk.Tk()
root.attributes('-alpha', 0.0)  #Makes the extra window fully transparent
desired_location = filedialog.askdirectory()  #Opens the file dialog
root.destroy()  #Gets rid of the main window

##########################################################################################

def main():

   ## Series of code to summarized code from Sergio
   
   SampleRepetition = input('Which repeitition?     ') + '\/'
   SampleRepetition1 = input('Which repeitition?     ') + '\/'

   Spectrum = ["Spectrum"]
   SampleDF = pd.read_csv('Wavenumber-testfile-donotdelete.csv')
   ## Sample file in general location to pull x axis from ##
   SampleDF.columns = ["Wavenumber", "Spectrum", "Standard Deviation"]
   Wavenumber = SampleDF.loc[0:,'Wavenumber']
   print(Wavenumber)
   NewDF = pd.DataFrame()

   OS_finalloc = desired_location + SampleRepetition
   print(OS_finalloc)
   sorted_OS_listdir = sorted(os.listdir(OS_finalloc))

   for file1 in sorted_OS_listdir :
       df_file = pd.read_csv(OS_finalloc + file1)
       df_file.columns = ["Integration ", "S1",]
       print(df_file.columns)
       df_fileloc = df_file.loc[15:,['S1']]
       df_fileloc.reset_index(drop=True, inplace=True)
       print(df_fileloc)
       NewDF = pd.concat([df_fileloc, NewDF], axis=1)
       NewDF.reset_index(drop=True, inplace=True)
       #print(NewDF)
   NewDF2 = pd.concat([Wavenumber, NewDF], axis=1)
   #print(NewDF2)
   NewDF2.reset_index(drop=True, inplace=True)


   Outfile_pathq = input('What folder name would you like?       ')
   ParentDir = (r'C:\Users\Angela\PycharmProjects\CompilingSERS\/')
   Outfile_path = (r'C:\Users\Angela\PycharmProjects\CompilingSERS\/') + Outfile_pathq
   #print(Outfile_path)
   Outdir = os.mkdir(Outfile_path)
   #print(Outdir)

   OutFileDF = input('What will you name your fle for your set of samples -MatlabProcessing  ')
   OutFileDF2 = input('What will you name your fle for your set of samples + with header/wavenumber   ')
   NewDFOut = (OutFileDF + '.csv')
   NewDFOut2 = (OutFileDF2 + '.csv')


   NewDF.to_csv(os.path.join(Outfile_path, NewDFOut), header = False, index = False)
   NewDF2.to_csv(os.path.join(Outfile_path, NewDFOut2), header = True)
   restart = input('Would you like to compile more data?      ')
   YesList = ["Yes", "y", "Y", "yes"]
   if restart in YesList :
       main()
   else:
       quit()
       
       
     

main()
