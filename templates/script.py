from openpyxl import load_workbook
from openpyxl.chart import Reference

# Load the workbook
wb = load_workbook(filename='Avg_Template.xlsx')

# Loop over the worksheets
for sheet in wb.worksheets:

    # Check if the worksheet has any charts
    if sheet._charts:

        # Loop over the charts
        for idx, chart in enumerate(sheet._charts, start=1):
            # Change the chart title to the chart ID number
            chart.title = str(idx)

# Save the workbook
wb.save('modified_file.xlsx')