
# Create a temporary worksheet to clean data


```python
from openpyxl import Workbook
wb = Workbook()
#Use the next line to get the default worksheet
#ws = wb.active
#OR create a new worksheet at 0 position, by default it creates at the end
ws = wb.create_sheet("CleanWater_VS_Country", 0) # 0 means insert at first position
print(wb.sheetnames)
```


```python
from openpyxl import load_workbook
wb_raw = load_workbook('TableRaw.xlsx')
print(wb_raw.get_sheet_names())
```


```python
ws_read = wb_raw['Table4.9Worksheet - Sheet1']
```


```python
values = ws_read['B1'].value
values = values.split(" ")
print(values)
```


```python
for i, col in enumerate(ws_read.iter_cols()):
    if i == 1:
        print(type(col))
        print([len(x.value.split(" ")) for x in col])
        for cel in col:
            arr = cel.value.split(" ")
            if len(arr) == 17:
                print(cel.row)
                print(arr)
```


```python
countries = list()
for i, col in enumerate(ws_read.iter_cols()):
    if i == 0:
        countries = [x.value for x in col]
        print(countries)
```

    ['Afghanistan', 'Algeria', 'Argentina', 'Australia', 'Austria', 'Bangladesh', 'Belgium-Lux.', 'Brazil', 'Bulgaria', 'Canada', 'China', 'Colombia', 'Congo, DR', "CÃ´te d'Ivoire", 'Denmark', 'Egypt', 'Ethiopia', 'France', 'Germany', 'Ghana', 'Greece', 'India', 'Indonesia', 'Iran', 'Iraq', 'Israel', 'Italy', 'Japan', 'Jordan', 'Kazakhstan', 'Kenya', 'Korea, DPR ', 'Korea, Rep ', 'Malaysia', 'Mexico', 'Morocco', 'Myanmar', 'Nepal', 'Netherlands', 'Nigeria', 'Norway', 'Pakistan', 'Peru', 'Philippines', 'Poland', 'Portugal', 'Romania', 'Russia', 'Saudi Arabia', 'South Africa', 'Spain', 'Sri Lanka', 'Sudan', 'Sweden', 'Switzerland', 'Tanzania', 'Thailand', 'Turkey', 'Turkmenistan', 'Ukraine', 'United Kingdom', 'USA', 'Uzbekistan', 'Venezuela', 'Viet Nam', 'Global total/average ']
    


```python
for i, col in enumerate(ws_read.iter_cols()):
    if i == 1:
        print(type(col))
        #print([len(x.value.split(" ")) for x in col])
        for cel in col:
            arr = cel.value.split(" ")
            if len(arr) < 16:
                print(cel.row)
                print(arr)
```


```python
#I want to insert rows in the sheet ws
# Step 1: Insert values that are 17 in number
for i, col in enumerate(ws_read.iter_cols()):
    if i == 1:
        print(type(col))
        #print([len(x.value.split(" ")) for x in col])
        for cel in col:
            arr = cel.value.split(" ")
            if len(arr) == 17:
                [ws.cell(row=cel.row, column=i, value=arr[i-1]) for i,_ in enumerate(arr) if i > 0]
                ws.cell(row=cel.row, column=len(arr), value=arr[len(arr) - 1])
            if len(arr) == 16:
                [ws.cell(row=cel.row, column=i+1, value=arr[i]) for i,_ in enumerate(arr)]
wb.save(filename = "temp.xlsx")
```


```python
# I added column names using google sheets and downloaded the file after a little bit of cleaning
#Filename : Water Consumption Vs Country.xlsx
# Fill country names
wb_final = load_workbook('Water Consumption Vs Country.xlsx')
print(wb_raw.get_sheet_names())
```

    ['CleanWater_VS_Country', 'Sheet']
    


```python
final_sheet = wb_final['CleanWater_VS_Country']
```


```python
for i in range(2, 67):
    final_sheet.cell(row=i, column=1, value=countries[i-2])
wb_final.save('Output.xlsx')
```


```python
#print(ws['B1'].value)
```
