import pandas as pd
import sys

args = sys.argv[1::]

if len(args) != 2:
    raise Exception("2 files needed, check syntax: \n "
                    "python3 main.py \"filename1\" \"filename2\"")

# Load the xlsx as variables in the program
print("Loading " + args[0] + " ...")
first_dataframe = pd.read_excel(args[0], "Amplitude sweep - 1", skiprows=[0])

print("Loading " + args[1] + " ...")
second_dataframe = pd.read_excel(args[1], "Amplitude sweep - 1", skiprows=[0])
print("Successfully read")

# Merge the two dataframes in one
print("Merging xls...")
output = first_dataframe.join(second_dataframe, lsuffix='_A', rsuffix='_B')
print("Merged ")

# Remove the header with units. It's mandatory, so we can automate the calculations
output = output.drop([0])

# Adding the mean and the deviation of the Storage modulus columns
output["G'"] = output[['Storage modulus_A', 'Storage modulus_B']].mean(axis=1)
output["G' devest"] = output[['Storage modulus_A', 'Storage modulus_B']].std(axis=1)

# Adding the mean and the deviation of the Loss modulus columns
output["G''"] = output[['Loss modulus_A', 'Loss modulus_B']].mean(axis=1)
output["G'' devest"] = output[['Loss modulus_A', 'Loss modulus_B']].std(axis=1)

print("Creating additional columns xls...")

# Create a Pandas Excel writer using XlsxWriter as the engine.
excel_file = "output.xlsx"
sheet_name = 'Sheet1'

writer = pd.ExcelWriter(excel_file, engine='xlsxwriter')
output.to_excel(writer, sheet_name=sheet_name)

# Access the XlsxWriter workbook and worksheet objects from the dataframe.
workbook = writer.book
worksheet = writer.sheets[sheet_name]

# Create a chart object.
chart = workbook.add_chart({'type': 'scatter'})

# Configure the series of the chart from the dataframe data.
# Using a list of values instead of category/value formulas:
# #     [sheetname, first_row, first_col, last_row, last_col]
max_row = 44
col_number_g_prime = 29
chart.add_series({
    'name': ['Sheet1', 0, col_number_g_prime],
    'categories': ['Sheet1', 0, col_number_g_prime, 0, col_number_g_prime],
    'values': ['Sheet1', 1, col_number_g_prime, max_row, col_number_g_prime],
    'marker': {'type': 'circle', 'size': 7}})

col_number_g_double_prime = 31
chart.add_series({
    'name': ['Sheet1', 0, col_number_g_double_prime],
    'categories': ['Sheet1', 0, col_number_g_double_prime, 0, col_number_g_double_prime],
    'values': ['Sheet1', 1, col_number_g_double_prime, max_row, col_number_g_double_prime],
    'marker': {'type': 'circle', 'size': 7, 'fill': {'color': 'red'}}})

# Configure the chart axes.
chart.set_x_axis({'name': "X Axis"})
chart.set_y_axis({'name': "Y Axis",
                  'major_gridlines': {'visible': False}})

# Insert the chart into the worksheet.
worksheet.insert_chart('A60', chart)
print("Inserting scatter chart ...")

# Export the output in a xlsx file
print("Exporting file ...")
output.to_excel("output.xlsx", index=False)
writer.close()
print("Exported successfully as \"output.xlsx\" ")
