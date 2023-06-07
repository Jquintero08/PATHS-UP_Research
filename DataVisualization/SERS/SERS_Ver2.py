import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import argrelextrema
from pandas import ExcelWriter
import os
import openpyxl





choice = input("Would you like to just do computations on one file (1) or combine multiple .csv files into an .xlsx file (2): ")

if choice == '2':
    csv_files = [] # Initialize an empty list

    while True: # Continue prompting the user until they choose to stop
        csv_file = input("Enter the full path of a .csv file (or 'quit' to finish): ")
        
        if csv_file.lower() == 'quit':
            break
        
        # Checks to see if file exists
        if not os.path.isfile(csv_file): 
            print("The file does not exist. Please check the path and try again.")
        else:
            # Append entered file to the list
            csv_files.append(csv_file)
            
    # Prompts user to enter the output file name with .xlsx
    outputfileName = input("What would you like the output file to be named (with .xlsx): ")
    # Create a Pandas Excel writer using XlsxWriter as the engine
    with ExcelWriter(outputfileName) as writer: 
        # Iterate through all the .csv files and add each one to a separate worksheet
        for csv_file in csv_files: 
            df = pd.read_csv(csv_file)
            # Use the .csv file name (without extension) as the name of the worksheet
            df.to_excel(writer, sheet_name=os.path.splitext(os.path.basename(csv_file))[0], index=False)
        writer.save()
        
        # Input which worksheet to work with
    looper = 1
    while looper == 1:
        selectSheet = input("Enter which worksheet you would like to work with (for list, enter 'GiveList'): ")
        if selectSheet == 'GiveList':
            #Lists all the worksheets inside .xlsx file
            xls = pd.ExcelFile(outputfileName)
            sheet_list = xls.sheet_names
            print(sheet_list)
        else:
            # Reads the new Excel file skipping the first 17 lines inside the respective worksheet
            df = pd.read_excel(outputfileName, header=None, skiprows = 17, sheet_name = selectSheet)
            looper = 0
    
elif choice == '1':
    filename = input("Please input filename with respective path and type: ")
    #filename = r"C:\Users\Jakey\Desktop\Summer_2023_Research\Python\780-AuNS-3_0.5-S1.1" + '.xlsx'
    #filename = r"C:\Users\movaq\OneDrive\Spring 2023\Gold Nano-star\data\python_data" + '.xlsx'
    
    # If the filetype is .xlsx which includes multiple worksheets
    if filename.endswith('.xlsx'): 
        # Input which worksheet to work with
        looper = 1
        while looper == 1:
            selectSheet = input("Enter which worksheet you would like to work with (for list, enter 'GiveList'): ")
            if selectSheet == 'GiveList':
                #Lists all the worksheets inside .xlsx file
                xls = pd.ExcelFile(filename)
                sheet_list = xls.sheet_names
                print(sheet_list)
            else:
                # Reads the new Excel file skipping the first 17 lines inside the respective worksheet
                df = pd.read_excel(filename, header=None, skiprows = 17, sheet_name = selectSheet)
                looper = 0
            
 
    # If the filetype is .csv which only has 1 worksheet
    elif filename.endswith('.csv'): 
        # Reads the new .csv file skipping the first 17 lines
        df = pd.read_csv(filename, header=None, skiprows = 17) 
    else:
        raise ValueError('Invalid file type. Please use .xlsx or .csv file')   

else:
    raise ValueError('Invalid choice')

# Create a sample dataframe with header names
header_names = input("Enter the header names separated by commas: ").split(',')

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



plt.text(1206 -50, 18000, '[1206] \nC-H \nrocking of \ncyclohexene ring', ha='right')
plt.text(1270 + 10, 15000, '[1270] \nC-H \nrocking of \nphenyl ring', ha='left')
plt.text(1467 -10, 5000, '[1467] \nC=C \nbending of \nphenyl', ha='right')

plt.text(1515 + 50, 4500, '[1515] \nC-H \nstretching \nof phenyl', ha='left')


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
    for j in range(min(4, len(maxima_indices))):
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






























# ##########################################################
# # Inside these comments is the combining files code
# # If you just want the skip 17 lines and select the worksheet code, comment this out and uncomment all above


# # Function to handle different files
# def process_file(filename, selectSheet=None):
#     if filename.endswith('.xlsx'):
#         df = pd.read_excel(filename, header=None, skiprows = 17, sheet_name = selectSheet)
#     elif filename.endswith('.csv'):
#         df = pd.read_csv(filename, header=None, skiprows = 17)
#     else:
#         raise ValueError('Invalid file type. Please use .xlsx or .csv file')
#     return df

# # Function to combine files
# def combine_files(filenames):
#     combined_data = None
#     for filename in filenames:
#         filename = filename.strip()  # Remove any leading/trailing spaces
#         if filename.endswith('.xlsx'):
#             selectSheet = input(f"Enter which worksheet you would like to work with for file {filename}: ")
#             df = process_file(filename, selectSheet)
#         else:
#             df = process_file(filename)

#         if combined_data is None:
#             combined_data = df
#         else:
#             combined_data = pd.concat([combined_data, df], axis=1)

#     # Write the combined data to a new Excel file
#     combined_filename = "combined.xlsx"
#     write_to_excel(combined_data, combined_filename)
#     print(f"All files combined into {combined_filename}. Now proceeding with the data analysis.")

#     return combined_filename

# def write_to_excel(df, filename):
#     writer = ExcelWriter(filename)
#     df.to_excel(writer, 'Sheet1', index=False)
#     writer.save()

# choice = input("Do you want to (1) process a single file or (2) combine multiple files? ")

# if choice == '1':
#     filename = input("Please input the full file path and filename: ")
#     if filename.endswith('.xlsx'):
#         selectSheet = input("Enter which worksheet you would like to work with: ")
#         df = process_file(filename, selectSheet)
#     else:
#         df = process_file(filename)
# elif choice == '2':
#     filenames = input("Please input the full file paths and filenames, separated by commas: ").split(',')
#     filename = combine_files(filenames)
#     df = process_file(filename)
# else:
#     raise ValueError('Invalid choice. Please enter 1 or 2.')

# # Read the file (either the single file or the combined file)
# df = pd.read_excel(filename, header=None, skiprows=17)



# ##################################################################################

