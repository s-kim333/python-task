# python-task

Basic program that processes data from excel file and update the result

## Approaches
### Method 1
Used library: pandas dataframe, openpyxl

1. Open excel file using pandas library
2. Calculate result and save the result in results array. 
3. Pass the array from step 2 to writing function (using openpyxl)
4. Save the updated excel sheet to the same path where the original excel file was stored. 

### Method 2
Used library: openpyxl
1. Open excel workbook using openpyxl
2. Get cell values using .cell(row,col) method 
3. Calculate the result and update result cell's value.
4. Save the updated excel sheet to the same path where the original excel file was stored.

### Method 3
Used library: pandas
1. Open excel workbook using pandas dataframe
2. Copy the dataframe -> drop result column
3. Calculate result and store it in an array -> concatenate to the dataframe created at step 2.
4. Save the updated excel sheet to the same path where the original excel file was stored.
