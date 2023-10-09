# backend-repo1
import openpyxl
import requests

# Open the input Excel file
input_file = "input.xlsx"
workbook = openpyxl.load_workbook(input_file)
sheet = workbook.active

# Create a new Excel sheet for results
output_file = "output.xlsx"
output_workbook = openpyxl.Workbook()
output_sheet = output_workbook.active
output_sheet.append(["URL", "Status Code", "Validation Result"])

# Iterate through rows in the input sheet
for row in sheet.iter_rows(min_row=2, values_only=True):
    url = row[0]
    try:
        response = requests.get(url)
        status_code = response.status_code
        if status_code == 200:
            validation_result = "Valid"
        else:
            validation_result = "Invalid"
    except requests.exceptions.RequestException:
        status_code = "N/A"
        validation_result = "Error"

    # Append the results to the output sheet
    output_sheet.append([url, status_code, validation_result])

# Save the results to the output Excel file
output_workbook.save(output_file)
