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

    dayOne_max_list = df_ethanol.iloc[:, 0].nlargest(10).tolist() #Finds top 10 maxes of the first day & stores as list
    
    #print(dayOne_max_list)
    
    total_one = sum(dayOne_max_list) #Total of day 0's top 10 maxes
    count_one = len(dayOne_max_list) #Count of them
    
    dayOne_10Max = total_one/count_one #Find Average
    
    #print(dayOne_10Max) #Remove
    #print('') #Remove
    #print('') #Remove

    for day in range(1, df_ethanol.shape[1]): #Go through leftover columns


        top10_maxValues = df_ethanol.iloc[:, day].nlargest(10).tolist() #List of top 10 maxes
        #print(day)
        
        total = sum(top10_maxValues)
        count = len(top10_maxValues)
        
            
        #print(total) #Remove
        #print(count) #Remove
        currentDayMaxAvgs = total/count
        #print(currentDayMaxAvgs) #Remove
        
        #print(top10_maxValues) #Remove
        

        percent_change = ((currentDayMaxAvgs / dayOne_10Max)-1) #Calculate percent change from day 1 to current day
        
        
        print("{:.2f}% change from Day 0 to Day {}".format((percent_change*100), day))
        
        
        #print('') #Remove
        #print('') #Remove

        df_sample.iloc[:, day] -= (df_sample.iloc[:, day] * percent_change) #Subtract data amount with the amount the data changed (could be negative) from ethanol file
        #If ethanol change is 13% from day 0 -> day 1, then the sample data should be shifted down 13% to 87% of what it was before to compensate for external factors

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

def Analysis():
    
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes the extra window fully transparent
    
    try:
        filename = filedialog.askopenfilename()  #User picks which files
        if not filename:  #If no file selected
            raise ValueError('No files selected.')

    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Destroys window
    
    
    
    df = pd.read_excel(filename, header=None)
    
    #Create a sample dataframe with header names
    header_names = sg.popup_get_text('Enter the header names separated by commas').split(',')
    
    df.columns = header_names
    

    print(df) #Shows Dataframe
    
    x_col = df.columns[0] #Header name of the first col.
    

    df = df[(df[x_col] >= 450) & (df[x_col] <= 1600)] #Filter data b/w 450 - 1600
    
    for i in range(1, df.shape[1]): #Plot all spectra with 1st column = x-axis

        z = np.polyfit(df[x_col], df.iloc[:, i], 30) #Data -> Polynomial
        p = np.poly1d(z)
    
    
        baseline_corrected_data = df.iloc[:, i] - p(df[x_col]) #Subtract min value of polynomial fit from data to get baseline-corrected data
        min_value = np.min(baseline_corrected_data)
        baseline_corrected_data += abs(min_value)
        

        if np.min(baseline_corrected_data) < 0: #All negative values -> position one
            negative_values = baseline_corrected_data[baseline_corrected_data < 0]
            baseline_corrected_data[baseline_corrected_data < 0] = 1
            baseline_corrected_data[0:len(negative_values)] = negative_values
    

        plt.plot(df[x_col], baseline_corrected_data, label=df.columns[i]) #Plot baseline-corrected data
    

        maxima_indices = argrelextrema(baseline_corrected_data.values, np.greater) #Find indices of local max in data
    

        maxima_indices = maxima_indices[0][np.argsort(-baseline_corrected_data.values[maxima_indices])] #Sort indices based on value of the data at the indices
    
        #Mark the top three highest peaks with markers on the plot
        # for j in range(min(3, len(maxima_indices))):
        #     plt.plot(df[x_col].values[maxima_indices[j]], baseline_corrected_data.values[maxima_indices[j]], 'ro')
    
    #Adds a straight line
    plt.axvspan(1186, 1226, color='blue', alpha=0.3)
    plt.axvspan(1250, 1290, color='black', alpha=0.3)
    plt.axvspan(1447, 1487, color='yellow', alpha=0.3)
    plt.axvspan(1495, 1535, color='red', alpha=0.3)
    
    
    
    # plt.text(1206 -50, 18000, '[1206] \nC-H \nrocking of \ncyclohexene ring', ha='right')
    # plt.text(1270 + 10, 15000, '[1270] \nC-H \nrocking of \nphenyl ring', ha='left')
    # plt.text(1467 -10, 5000, '[1467] \nC=C \nbending of \nphenyl', ha='right')
    
    # plt.text(1515 + 50, 4500, '[1515] \nC-H \nstretching \nof phenyl', ha='left')
    
    ### Fix this later... Causes plot size to be irregular ###
    
    
    #Add title to label axis
    #plt.title('SERS measurement of silica coated, IR780 encapsulated GNS')
    plt.xlabel('Wavenumber (cm$^{-1}$)', fontsize=12)
    plt.ylabel('Raman Intensity (counts)', fontsize=12)
    

    plt.ylim(bottom=min(df.iloc[:, 1:].min().min(), 0)) #Set min value of y-axis -> lower than min value of data
    plt.legend(loc='upper left') #Legend
    plt.show() #Output the plot
    
    
    
    #Plot all the spectra with the first column as x-axis
    fig, ax = plt.subplots()
    fig.set_size_inches(6, 4)
    x = df[x_col].values
    y_offset = 1
    
    for i in range(1, df.shape[1]):

        z = np.polyfit(x, df.iloc[:, i], 3) #Data -> Polynomial
        p = np.poly1d(z)
    
        #Subtract min value of polynomial fit from data to get baseline-corrected data
        baseline_corrected_data = df.iloc[:, i] - p(x) + y_offset 
        min_value = np.min(baseline_corrected_data)
        baseline_corrected_data += abs(min_value)
    
    
        ax.plot(x, baseline_corrected_data, label=df.columns[i]) #Plot baseline-corrected data
    

        maxima_indices = argrelextrema(baseline_corrected_data.values, np.greater) #Indicies of local maxima of data
    
        #Sort indices based on value @ that index
        maxima_indices = maxima_indices[0][np.argsort(-baseline_corrected_data.values[maxima_indices])]
    
        for j in range(min(3, len(maxima_indices))): #Mark top 3 highest peaks on the plot
            ax.plot(x[maxima_indices[j]], baseline_corrected_data.values[maxima_indices[j]], 'ro')
    
        y_offset += 25000 #y-offset for next plot
    
    #Add labels -> axes
    #ax.set_title('Baseline-corrected spectra', fontsize=14, loc="center")
    ax.set_xlabel('Wavenumber (cm$^{-1}$)', fontsize=12)
    ax.set_ylabel('Raman intensity', fontsize=12)
    
    ax.legend(loc = 'upper right') #Legend
    

    plt.show() #Show plot
    
    
    

    highest_peaks = [] #Empty list to store highest peaks
    second_highest_peaks = [] #Empty list to store second highest peak
    third_highest_peaks = []  #Empty list to store third highest peaks
    
    
    sample_names = [] #Empty list to store sample names
    
    highest_wavenumbers = [] #Empty list to store the highest wavenumbers
    second_highest_wavenumbers = [] #Empty list to store the second highest wavenumbers
    third_highest_wavenumbers = [] #Empty list to store the third highest wavenumbers
    
    
    #Loop through spectrum inside the Dataframe
    for i in range(1, df.shape[1]):

        z = np.polyfit(df[x_col], df.iloc[:, i], 3) #Data -> Polynomial
        p = np.poly1d(z)
    
    
        baseline_corrected_data = df.iloc[:, i] - p(df[x_col]) #Subtract min-value from polynomial fit to get baseline-corrected data
        min_value = np.min(baseline_corrected_data)
        baseline_corrected_data += abs(min_value)
    

        maxima_indices = argrelextrema(baseline_corrected_data.values, np.greater) #Find indices of local maxima
    
        #Sort indices based on value of data @ respective index
        sorted_maxima_indices = maxima_indices[0][np.argsort(-baseline_corrected_data.values[maxima_indices])]
        

        highest_peaks.append(baseline_corrected_data.values[sorted_maxima_indices[0]]) #Append highest peak
        second_highest_peaks.append(baseline_corrected_data.values[sorted_maxima_indices[1]]) #Append second highest peak
        third_highest_peaks.append(baseline_corrected_data.values[sorted_maxima_indices[2]]) #Append third highest peak
        
        highest_wavenumbers.append(df[x_col].values[sorted_maxima_indices[0]])
        second_highest_wavenumbers.append(df[x_col].values[sorted_maxima_indices[1]])
        third_highest_wavenumbers.append(df[x_col].values[sorted_maxima_indices[2]])
        
        sample_names.append(df.columns[i])
    

    meanHighPeak = np.mean(highest_peaks) #Find mean of highest peaks
    std_highest_peak = np.std(highest_peaks) #Find standard deviation of highest peaks
    
    mean2ndHighPeak = np.mean(second_highest_peaks)
    std_second_highest_peak = np.std(second_highest_peaks)
    
    mean3rdHighPeak = np.mean(third_highest_peaks)
    std_third_highest_peak = np.std(third_highest_peaks)
    

    rsd_highest_peak = (std_highest_peak / meanHighPeak) * 100 #RSD of highest peaks
    rsd_second_highest_peak = (std_second_highest_peak / mean2ndHighPeak) * 100 #RSD of second highest peaks
    rsd_third_highest_peak = (std_third_highest_peak / mean3rdHighPeak) * 100 #RSD of third highest peaks
    
    #Labels for mean and standard deviaton
    mean_label_highest = f"Mean: {meanHighPeak:.2f}\n+/- {std_highest_peak:.2f}\nRSD: {rsd_highest_peak:.2f}%"
    mean_label_second_highest = f"Mean: {mean2ndHighPeak:.2f}\n+/- {std_second_highest_peak:.2f}\nRSD: {rsd_second_highest_peak:.2f}%"
    mean_label_third_highest = f"Mean: {mean3rdHighPeak:.2f}\n+/- {std_third_highest_peak:.2f}\nRSD: {rsd_third_highest_peak:.2f}%"
    

    fig, ax = plt.subplots(figsize=(10, 5)) #Bar chart of highest peak for each after baseline correction
    x = np.arange(len(sample_names))
    
    ax.bar(x, highest_peaks, width=0.8, label='Highest Peak', tick_label=sample_names)
    #ax.errorbar(x, highest_peaks, yerr=std_highest_peak, fmt='none', ecolor='black', capsize=5)
    ax.text(3,-5000, mean_label_highest, ha='right') 
    
    ax.bar(x + 20, second_highest_peaks, width=0.8, label='Second Highest Peak', tick_label=sample_names)
    #ax.errorbar(x + 20, second_highest_peaks, yerr=std_second_highest_peak, fmt='none', ecolor='black', capsize=5)
    ax.text(23,-5000, mean_label_second_highest, ha='right')
    
    ax.bar(x + 10, third_highest_peaks, width=0.8, label='Third Highest Peak', tick_label=sample_names)
    #ax.errorbar(x + 10, third_highest_peaks, yerr=std_third_highest_peak, fmt='none', ecolor='black', capsize=5)
    ax.text(13,-5000, mean_label_third_highest, ha='right')
    
    
    
    width = 10
    xticks = np.tile(x, 3) + np.repeat([0, width, 2*width], len(sample_names))
    ax.set_xticks(xticks)
    ax.set_xticklabels(np.tile(sample_names, 3), rotation=90)
    ax.tick_params(axis='x', pad=10, which='minor')
    ax.set_ylabel('Raman Intensity', fontsize=12)
    ax.set_xlabel('Samples', fontsize=12)
    
    
    ax2 = ax.twiny()
    ax2.set_xticks([4, 15, 26, 30])
    ax2.set_xticklabels([round(highest_wavenumbers[0],0), round(second_highest_wavenumbers[0],0), round(third_highest_wavenumbers[0],0), ""])
    ax2.set_xlabel('Wavenumber (cm$^{-1}$)', fontsize=12)
    
    fig.tight_layout()
    plt.show()
    
    print("Success!") #Show the function was a success to the terminal
        
        
        

def main(): 

    loop = True
    while loop:
        layout = [[sg.Text('Select one->'), sg.Listbox(['Combine Files', 'Data Shift', 'Data Analysis'], size=(20, 3), key='LB')],[sg.Button('Ok'), sg.Button('Cancel')]]

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
                
            elif 'Data Analysis' in values['LB']:
                
                root = tk.Tk()
                root.withdraw()
                #messagebox.showinfo("Info", "Display Message") #Show Info
                #Comment out above messagebox if you know what to do
                root.destroy() #Close
                
                print("Performing data analysis...")
                Analysis() #Runs the Analysis function
                loop = create_popup()
                
        else:
            sg.popup_cancel('User aborted')
            loop = False




if __name__ == "__main__":
    main()







