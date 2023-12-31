import pandas as pd
import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import argrelextrema


filename = r"C:\Users\Jakey\OneDrive\Desktop\Summer_2023_Research\Python\python_data_UV-vis" + '.xlsx'

df = pd.read_excel(filename, header=None)

# Create a sample dataframe with header names
header_names = input("Enter the header names separated by commas: ").split(',')

df.columns = header_names

# display the dataframe
print(df)

# Get the header name of the first column
x_col = df.columns[0]

# Filter the data between 350 and 1000
df = df[(df[x_col] >= 350) & (df[x_col] <= 1000)]



# Plot all the spectra with the first column as the x-axis
for i in range(1, df.shape[1]):

    plt.plot(df[x_col], df.iloc[:, i], label=df.columns[i])

# Add a straight line
plt.axvline(x=785, color='black', linestyle='--')
#plt.axvline(x=850, color='blue', linestyle=':')
#plt.text(760, 0.3, '760', ha='right')
#plt.text(850, 0.2, '850', ha='left')

plt.text(785, 0.8, '785 nm', ha='left')

# add a title and labels to the axes
plt.xlabel('Wavelength (nm)', fontsize=12)
plt.ylabel('Optical Density', fontsize=12)

# Set the minimum value of the y-axis to be lower than the minimum value of the data
plt.ylim(bottom=min(df.iloc[:, 1:].min().min(), 0.3))

# Add a legend
plt.legend(loc='upper right')

# display the plot
plt.show()


# Filter the data between 450 and 1600
df = df[(df[x_col] >= 350) & (df[x_col] <= 1000)]

# Plot all the spectra with the first column as the x-axis
for i in range(1, df.shape[1]):
    # Fit a polynomial to the data
    z = np.polyfit(df[x_col], df.iloc[:, i], 1)
    p = np.poly1d(z)

    # Subtract the minimum value of the polynomial fit from the data to get the baseline-corrected data
    baseline_corrected_data = df.iloc[:, i] - p(df[x_col])
    min_value = np.min(baseline_corrected_data)
    baseline_corrected_data += abs(min_value)

    plt.plot(df[x_col], baseline_corrected_data, label=df.columns[i])


# Add a straight line
plt.axvline(x=760, color='black', linestyle='--')
plt.axvline(x=850, color='blue', linestyle=':')
plt.text(760, 0.2, '760', ha='right')
plt.text(850, 0.2, '850', ha='left')


# add a title and labels to the axes
plt.xlabel('Wavelength (nm)', fontsize=12)
plt.ylabel('Optical Density', fontsize=12)

# Set the minimum value of the y-axis to be lower than the minimum value of the data
plt.ylim(bottom=min(df.iloc[:, 1:].min().min(), 0))

# Add a legend
plt.legend(loc='upper right')

# display the plot
plt.show()


# Plot all the normalized spectra with the first column as the x-axis
for i in range(1, df.shape[1]):
    # Fit a polynomial to the data
    z = np.polyfit(df[x_col], df.iloc[:, i], 1)
    p = np.poly1d(z)

    # Subtract the minimum value of the polynomial fit from the data to get the baseline-corrected data
    baseline_corrected_data = df.iloc[:, i] - p(df[x_col])
    min_value = np.min(baseline_corrected_data)
    baseline_corrected_data += abs(min_value)

    # Normalize the data by dividing by the maximum value
    normalized_data = baseline_corrected_data / np.max(baseline_corrected_data)

    plt.plot(df[x_col], normalized_data, label=df.columns[i])

# Add a straight line
#plt.axvline(x=760, color='black', linestyle='--')
#plt.axvline(x=850, color='blue', linestyle=':')
#plt.text(760, 0.2, '760', ha='right')
#plt.text(850, 0.2, '850', ha='left')

# Add a title and labels to the axes
plt.xlabel('Wavelength (nm)', fontsize=12)
plt.ylabel('Normalized Optical Density', fontsize=12)

# Set the minimum value of the y-axis to be lower than the minimum value of the data
plt.ylim(bottom=min(normalized_data.min(), 0.0))

# Add a legend
plt.legend(loc='upper right')

# Display the plot
plt.show()