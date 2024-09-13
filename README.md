# PATHS-UP_Research
Progress of programs throughout my time apart of PATHS-UP program.


# CombinationOfFiles

This folder includes a program and sample files for testing. The program prompts the user to choose a directory for the master folder, then scours as deep as possible within it for similar .csv files. If the master folder contains subfolders, each becomes its own worksheet within a single .xlsx document. The program reads .csv files from each subfolder, appends necessary data to the respective worksheet, and repeats this process for all subfolders or until no more .csv files are available. The result is a singule .xlsx file with each subfolder represented as a worksheet. Each worksheet contains multiple datasets, with the first two columns representing one dataset, and each column thereafter representing an additional dataset



# DataVisualization

This folder contains different ways of transferring data as well as visualizing it in graphical form.

  ### SERS ###

  This folder contains a program that has two functionalities. The first functionality is combining many different .csv files into one .xlsx file, makings data analysis much easier. The second functionality is the data visualization of the SERS-analyzed data. It will plot a graph showing the raw data, the baseline corrected data, and the normalized data. The program also has a simple, but friendly, user interface to make it easier for someone with a lack of computer science knowledge to use. It also has sample data to perform these calculations.

  ### Tranpose ###

  This folder houses a peer-requested program with a user-friendly interface to facilitate data processing. Users are prompted to either select a "DI Water" data file or a file for transposition along with a pre-compiled DI Water file. If the first option is selected, the program reads a selected file with the "DI water" row, discards the header, and transposes the remaining data into an output file for future use. If the second option is chosen, users must select a file to transpose with at least one sample data set (three iterations of the same sample) and a pre-transposed DI Water file. An output filename with the .xlsx extension is then requested. The output file contains the transposed sample data, an additional average minus DI Water column (to adjust for any water content in the sample), and a standard deviation column. The output file headers are arranged as follows: "Wavelength, A1, A2, A3, Average 1 - DI Water, Standard Deviation 1, A5, A6, A7, Average 2 - DI Water, Standard Deviation 2", and so on.

  ### UV_vis ###

  This folder contains a program that was given to me to do data visualization, but I modified it to take the output file of the "Transpose" folder as an input and do the same data analysis. I also added another simple UI to make the program easier to use.

  


# EthanolAnalysis

This folder includes two functions for ethanol data analysis. The first merges various ethanol .csv files from different days into a unified .xlsx file. The second compares actual sample data with the consolidated ethanol data. After the user selects the ethanol and sample data files, the program notes any changes in the ethanol data due to external factors. It calculates the percentage change between the first day's top 10 maximum values and all the days following. This percentage change is then used to adjust each data point in the equivalent day column of the sample data file to reflect the external influences. For instance, if there was a 14% increase in the average of the top 10 values from day 0 to day 1 in the ethanol data, every data value in the day 2 column of the sample data file would be shifted down to 86%. The program repeats this process for each day, and the final output is saved to an .xlsx file. The user then has the option to visualize this data.
