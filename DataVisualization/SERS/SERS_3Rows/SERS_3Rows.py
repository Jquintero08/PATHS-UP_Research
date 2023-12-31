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
    
    
    if not filename:  # If there are no selected files, end the function
        print("No files selected.")
        return

    # Process the first file
    df = pd.read_csv(filename[0], skiprows=17, usecols=[0, 1], header = None)

    # Process the remaining files
    for file in filename[1:]:
        new_df = pd.read_csv(file, skiprows=17, usecols=[1], header = None)
        df = pd.concat([df, new_df], axis=1)
        
        
    root = tk.Tk()
    root.attributes('-alpha', 0.0)  #Makes the extra window transparent
    #output_dir = filedialog.askdirectory()  #Opens the file dialog (For input in terminal)
    
    try:
        outputFile = filedialog.asksaveasfilename(filetypes=[('Excel file', '*.xlsx')])
        if not outputFile:  #If no file selected
            raise ValueError('No files selected.')

    except ValueError as e:
        print(e)
        return  #Exit

    finally:
        root.destroy()  #Removes window
    

    df.to_excel(outputFile, index=False, header=False)
    
    print('Success!')
    
    
    
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
        
        
        
        
    
    df = pd.read_excel(filename)
    
    header_names = [df.columns[0]] ### Comment out if you want your own header names ###

    # Append every 5th header, starting from the 5th column
    header_names += [df.columns[i] for i in np.arange(4, df.shape[1], 5)] ### Comment out if you want your own header names ###
    
    
    df = df.iloc[:, [0] + list(np.arange(4, df.shape[1], 5))]
    
    # Create a sample dataframe with header names
    #header_names = sg.popup_get_text('Enter the header names separated by commas').split(',') ### Uncomment if you want your own header names ###

    
    
    
    df.columns = header_names
    
    # display the dataframeg
    print(df)
    
    # Get the header name of the first column
    x_col = df.columns[0]
    
    

    
    # Filter the data between 450 and 1600
    df = df[(df[x_col] >= 450) & (df[x_col] <= 1600)]
    

    
    # Plot all the spectra with the first column as the x-axis
    for i in range(1, df.shape[1]):
        # Fit a polynomial to the data
        z = np.polyfit(df[x_col], df.iloc[:, i], 30)
        p = np.poly1d(z)
    
        # Subtract the minimum value of the polynomial fit from the data to get the baseline-corrected data
        baseline_corrected_data = df.iloc[:, i] - p(df[x_col])
        min_value = np.min(baseline_corrected_data)
        baseline_corrected_data += abs(min_value)
        
        # Bring all the negative values on the curve to position one
        if np.min(baseline_corrected_data) < 0:
            negative_values = baseline_corrected_data[baseline_corrected_data < 0]
            baseline_corrected_data[baseline_corrected_data < 0] = 1
            baseline_corrected_data[0:len(negative_values)] = negative_values
    
        # Plot the baseline-corrected data
        plt.plot(df[x_col], baseline_corrected_data, label=df.columns[i])
    
        # Find the indices of the local maxima of the data
        maxima_indices = argrelextrema(baseline_corrected_data.values, np.greater)
    
        # Sort the indices based on the value of the data at the indices
        maxima_indices = maxima_indices[0][np.argsort(-baseline_corrected_data.values[maxima_indices])]
    
        # Mark the top three highest peaks with markers on the plot
        #for j in range(min(3, len(maxima_indices))):
         #   plt.plot(df[x_col].values[maxima_indices[j]], baseline_corrected_data.values[maxima_indices[j]], 'ro')
    
    # Add a straight line
    plt.axvspan(1186, 1226, color='blue', alpha=0.3)
    plt.axvspan(1250, 1290, color='black', alpha=0.3)
    plt.axvspan(1447, 1487, color='yellow', alpha=0.3)
    plt.axvspan(1495, 1535, color='red', alpha=0.3)
    
    
    
   # plt.text(1206 -50, 18000, '[1206] \nC-H \nrocking of \ncyclohexene ring', ha='right')
   # plt.text(1270 + 10, 15000, '[1270] \nC-H \nrocking of \nphenyl ring', ha='left')
   # plt.text(1467 -10, 5000, '[1467] \nC=C \nbending of \nphenyl', ha='right')
    
    #plt.text(1515 + 50, 4500, '[1515] \nC-H \nstretching \nof phenyl', ha='left')
    
    
    # add a title and labels to the axes
    #plt.title('SERS measurement of silica coated, IR780 encapsulated GNS')
    plt.xlabel('Wavenumber (cm$^{-1}$)', fontsize=12)
    plt.ylabel('Raman Intensity (counts)', fontsize=12)
    
    # Set the minimum value of the y-axis to be lower than the minimum value of the data
    plt.ylim(bottom=min(df.iloc[:, 1:].min().min(), 0))
    
    # Add a legend
    plt.legend(loc='upper left')
    
    # display the plot
    plt.show()
    
    
    
    # Plot all the spectra with the first column as the x-axis
    fig, ax = plt.subplots()
    fig.set_size_inches(6, 4)
    x = df[x_col].values
    y_offset = 1
    
    for i in range(1, df.shape[1]):
        # Fit a polynomial to the data
        z = np.polyfit(x, df.iloc[:, i], 3)
        p = np.poly1d(z)
    
        # Subtract the polynomial fit from the data to get the baseline-corrected data
        baseline_corrected_data = df.iloc[:, i] - p(x) + y_offset
        min_value = np.min(baseline_corrected_data)
        baseline_corrected_data += abs(min_value)
    
        # Plot the baseline-corrected data with an offset in the y-axis
        ax.plot(x, baseline_corrected_data, label=df.columns[i])
    
        # Find the indices of the local maxima of the data
        maxima_indices = argrelextrema(baseline_corrected_data.values, np.greater)
    
        # Sort the indices based on the value of the data at the indices
        maxima_indices = maxima_indices[0][np.argsort(-baseline_corrected_data.values[maxima_indices])]
    
        # Mark the top three highest peaks with markers on the plot
        for j in range(min(3, len(maxima_indices))):
            ax.plot(x[maxima_indices[j]], baseline_corrected_data.values[maxima_indices[j]], 'ro')
    
        # Increment the y offset for the next plot
        y_offset += 25000
    
    # Add labels to the axes
    #ax.set_title('Baseline-corrected spectra', fontsize=14, loc="center")
    ax.set_xlabel('Wavenumber (cm$^{-1}$)', fontsize=12)
    ax.set_ylabel('Raman intensity', fontsize=12)
    
    # Display the legend
    ax.legend(loc = 'upper right')
    
    # Display the plot
    plt.show()
    
    
    
    # Create empty lists to store the highest peak, second highest peak, and third highest peak of each spectrum and its corresponding sample name
    highest_peaks = []
    second_highest_peaks = []
    third_highest_peaks = []
    
    
    sample_names = []
    
    highest_wavenumbers = []
    second_highest_wavenumbers = []
    third_highest_wavenumbers = []
    
    
    # Loop through each spectrum in the DataFrame
    for i in range(1, df.shape[1]):
        # Fit a polynomial to the data
        z = np.polyfit(df[x_col], df.iloc[:, i], 3)
        p = np.poly1d(z)
    
        # Subtract the minimum value of the polynomial fit from the data to get the baseline-corrected data
        baseline_corrected_data = df.iloc[:, i] - p(df[x_col])
        min_value = np.min(baseline_corrected_data)
        baseline_corrected_data += abs(min_value)
    
        # Find the indices of the local maxima of the data
        maxima_indices = argrelextrema(baseline_corrected_data.values, np.greater)
    
        # Sort the indices based on the value of the data at the indices
        sorted_maxima_indices = maxima_indices[0][np.argsort(-baseline_corrected_data.values[maxima_indices])]
        
        # Add the highest peak, second highest peak, third highest peak value and sample name to the corresponding lists
        highest_peaks.append(baseline_corrected_data.values[sorted_maxima_indices[0]])
        second_highest_peaks.append(baseline_corrected_data.values[sorted_maxima_indices[1]])
        third_highest_peaks.append(baseline_corrected_data.values[sorted_maxima_indices[2]])
        
        highest_wavenumbers.append(df[x_col].values[sorted_maxima_indices[0]])
        second_highest_wavenumbers.append(df[x_col].values[sorted_maxima_indices[1]])
        third_highest_wavenumbers.append(df[x_col].values[sorted_maxima_indices[2]])
        
        sample_names.append(df.columns[i])
    
    # Compute the mean and standard deviation of the highest peaks
    mean_highest_peak = np.mean(highest_peaks)
    std_highest_peak = np.std(highest_peaks)
    
    mean_second_highest_peak = np.mean(second_highest_peaks)
    std_second_highest_peak = np.std(second_highest_peaks)
    
    mean_third_highest_peak = np.mean(third_highest_peaks)
    std_third_highest_peak = np.std(third_highest_peaks)
    
    # Compute the RSD of the highest peaks
    rsd_highest_peak = (std_highest_peak / mean_highest_peak) * 100
    rsd_second_highest_peak = (std_second_highest_peak / mean_second_highest_peak) * 100
    rsd_third_highest_peak = (std_third_highest_peak / mean_third_highest_peak) * 100
    
    # Create labels for the mean and standard deviation
    mean_label_highest = f"Mean: {mean_highest_peak:.2f}\n+/- {std_highest_peak:.2f}\nRSD: {rsd_highest_peak:.2f}%"
    mean_label_second_highest = f"Mean: {mean_second_highest_peak:.2f}\n+/- {std_second_highest_peak:.2f}\nRSD: {rsd_second_highest_peak:.2f}%"
    mean_label_third_highest = f"Mean: {mean_third_highest_peak:.2f}\n+/- {std_third_highest_peak:.2f}\nRSD: {rsd_third_highest_peak:.2f}%"
    
    # Create a bar chart of the highest peak of each spectrum after baseline correction
    fig, ax = plt.subplots(figsize=(10, 5))
    x = np.arange(len(sample_names))
    
    ax.bar(x, highest_peaks, width=0.8, label='Highest Peak', tick_label=sample_names)
    #ax.errorbar(x, highest_peaks, yerr=std_highest_peak, fmt='none', ecolor='black', capsize=5)
    ax.annotate(mean_label_highest, xy=(0.19, -0.12), xycoords='axes fraction', xytext=(-5, -5), textcoords='offset points', ha='right', va='top')
    ### CHANGE ABOVE [INSIDE xy=(...,...)] X-AXIS TO MOVE LEFT AND RIGHT. Y DOESN'T NEED TO BE CHANGED ###
    
    
    ax.bar(x + 20, second_highest_peaks, width=0.8, label='Second Highest Peak', tick_label=sample_names)
    #ax.errorbar(x + 20, second_highest_peaks, yerr=std_second_highest_peak, fmt='none', ecolor='black', capsize=5)
    ax.annotate(mean_label_second_highest, xy=(0.57, -0.12), xycoords='axes fraction', xytext=(-5, -5), textcoords='offset points', ha='right', va='top')
    ### CHANGE ABOVE [INSIDE xy=(...,...)] X-AXIS TO MOVE LEFT AND RIGHT. Y DOESN'T NEED TO BE CHANGED ###
    
    
    ax.bar(x + 10, third_highest_peaks, width=0.8, label='Third Highest Peak', tick_label=sample_names)
    #ax.errorbar(x + 10, third_highest_peaks, yerr=std_third_highest_peak, fmt='none', ecolor='black', capsize=5)
    ax.annotate(mean_label_third_highest, xy=(0.95, -0.12), xycoords='axes fraction', xytext=(-5, -5), textcoords='offset points', ha='right', va='top')
    ### CHANGE ABOVE [INSIDE xy=(...,...)] X-AXIS TO MOVE LEFT AND RIGHT. Y DOESN'T NEED TO BE CHANGED ###
    
    
    
    
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
    
    print('Success!')
    
    
def AvgandStdDev():
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
    
    data2 = pd.read_excel(inputFile, header=None)

    cols = data2.columns.tolist()
    
    data3 = pd.DataFrame()
    data3['Wavelength'] = data2.iloc[:, 0]

    for i in range(1, len(cols)):
        data3[cols[i]] = data2[cols[i]]
        if i % 3 == 0 and i >= 3:  #Calculate average and std deviation only when there are at least three columns
            data3[f'Average {i//3}'] = data2.iloc[:, i-2:i+1].mean(axis=1)
            data3[f'Std Dev {i//3}'] = data2.iloc[:, i-2:i+1].std(axis=1)
            
    data3.to_excel(outputFile, index=False)
        

    
    print('Success!')


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

def main(): 

    loop = True
    while loop:
        layout = [[sg.Text('Select one->'), sg.Listbox(['Combine Files', 'Averages and Std Devs', 'Data Analysis'], size=(20, 3), key='LB')],[sg.Button('Ok'), sg.Button('Cancel')]]

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

            elif 'Averages and Std Devs' in values['LB']:
                
                root = tk.Tk()
                root.withdraw()
            
                root.destroy() #Close
                
                print("Performing calculations...")
                AvgandStdDev() #Runs the Analysis function
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