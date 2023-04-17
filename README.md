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