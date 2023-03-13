import os
import matplotlib.pyplot as plt
import pandas as pd
import xlrd

# Set directory to where the Excel files are stored
dir_path = 'C:\\Users\\offic\\OneDrive\\Desktop\\DevCodeCamp\\CAPSTONE\\Data'

# Get a list of all files in the directory
files = os.listdir(dir_path)

# Initialize a list to store dataframes
dfs = []

# Loop through each file
for file in files:
    # Check if file is an Excel file
    if not file.endswith(('.xls', '.xlsx', '.xlsm', '.xlsb')):
        print(f"{file} is not an Excel file. Skipping...")
        continue
        
    # Load the Excel file into a dataframe
    try:
        xls = pd.ExcelFile(os.path.join(dir_path, file))
        df = pd.concat([pd.read_excel(xls, sheet_name=tab) for tab in xls.sheet_names])
    except xlrd.biffh.XLRDError:
        print(f"Error: Could not load {file}")
        continue
    
    # Remove rows with null data
    df = df.dropna()
    
    # Find columns with 'x' and 'y' in their names
    x_col = [col for col in df.columns if 'x' in col.lower()]
    y_col = [col for col in df.columns if 'y' in col.lower()]
    
    # Check if there are any 'x' and 'y' columns
    if not x_col or not y_col:
        print(f"No 'x' or 'y' column found in {file}. Skipping...")
        print(f"x_col: {x_col}")
        print(f"y_col: {y_col}")
        continue
    
    # Create a dataframe with only 'x' and 'y' columns
    df = df[x_col + y_col]
    
    # Append the dataframe to the list of dataframes
    dfs.append(df)

# Create a scatter plot of the dataframes
fig, ax = plt.subplots()
for i, df in enumerate(dfs):
    ax.scatter(df[x_col[0]], df[y_col[0]], label=f"Dataset {i+1}")

# Add legend to the plot
ax.legend()

# Show the plot
plt.show()