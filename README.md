# openpyxl_calc
Summation and visualisation of multiple sheets excel of data (with or without conditions) aiutomated using the openpyxl module in Python.

This project takes an input of an excel workbook with corresponding sheets containing expenditure data for that particular month. Goal is to provide a summary of ONE category of expenditure, and provide a visualisation of the same.

The module openpyxl from pypi is used here.

* Sample Input file - SampleData.xlsx

# Installation
With pip

```pip install openpyxl```

# Output
[![Output-Screenshot.png](https://i.postimg.cc/85v23TwW/Output-Screenshot.png)](https://postimg.cc/cgsktyp4)

Here is the final result, which shows the summation of Gym expenditures for the months of January and February. As more sheets are added, the chart will add more bars.

The January worksheet has various types of expenditures, but in our final result we are taking a use-case where only gym expenditures are required.
Use cases can change as per convenience. Will soon add a dynamic input for a type of expense(s).

### Calculating expenditure of a category for each month (in each sheet)
```    
for sheet_name in sheet_names:
    total_exp = 0
    sheet = wb[sheet_name]
    for row in range(1, sheet.max_row+1):
        if str(sheet.cell(row, 1).value).upper() == "GYM":  # Explicitly searching for Gym
            total_exp += sheet.cell(row, 3).value
```
### Creating a new sheet with titles
```
sheet_names = wb.sheetnames
ws = wb.create_sheet('Summary')
rows = [("Month", "Total Expenditure")] 
```
### Adding the data into the new sheet
```
rows = [(sheet_name, total_exp)]
for row in rows:
    ws.append(row)
wb.save('SampleData.xlsx') #  Saving changes to the workbook
```
### Creating a Bar chart for the data
```
wb = xl.load_workbook('SampleData.xlsx')
ws = wb['Summary']

data = Reference(ws, min_row=1, max_row=ws.max_row, min_col=2, max_col=2)
titles = Reference(ws, min_row=2, max_row=ws.max_row, min_col=1, max_col=1)
chart = BarChart()
chart.title = "Gym Expense Chart"
chart.add_data(data=data, titles_from_data=True)
chart.set_categories(titles)

ws.add_chart(chart, "H2")
wb.save("SampleData.xlsx")
```