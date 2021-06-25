# python-task

Basic program that processes data from excel file and update the result

## Approaches
Tried 3 methods wit simple flow as below:
![basic flow chart](https://user-images.githubusercontent.com/64014524/123373620-00977200-d5c9-11eb-87b9-7d97f5f8b39b.PNG)

### Method 1
Used library: pandas dataframe, openpyxl

1. Open excel file using pandas library
2. Calculate result and save the result in results array. 
3. Pass the array from step 2 to writing function (using openpyxl)
4. Save the updated excel sheet to the same path where the original excel file was stored. 

![Capture](https://user-images.githubusercontent.com/64014524/123374571-a7304280-d5ca-11eb-80fc-42b928ae2261.PNG)


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

## Test plan 
For this program, simple functional test is possible by following steps

1. Create test input
2. Compute expected outcomes with the selected input
3. Execute the test input
4. Comparison of actual and computed expected result

At this stage, simple test code (test_1.py) using pytest is created to check correctness of the program, barrowing the code from given Sample.py code. 

