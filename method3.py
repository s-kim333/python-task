import pandas as pd
#import xlrd
import openpyxl
import sys
import os

# define Python user-defined exceptions
class Error(Exception):
    """Base class for other exceptions"""
    pass

class NoPathError(Error):
    """Raised when no path input"""
    pass

try:
    if len(sys.argv) >= 2:
        path = sys.argv[1]
        #r'C:\Users\kimmy\Downloads\SampleInput.xlsx'
        print(path)

        df = pd.read_excel(path)
        print(df)
        df1 = df.drop(['result'],axis=1)
        print(df)
        df = df.to_numpy()

        print (df)
        results = []
        for i,element in enumerate(df):
            x, y, z = element[0], element[1], element[2]
            if z.startswith("add"):
                result = x + y
            if z.startswith("mul"):
                result = x * y
            if z.startswith("exp"):
                result = x ** y
            if z.startswith("sub"):
                result = x - y
            if z.startswith("div"):
                if y != 0:
                    result = x/y
                else:
                    result = "Error"
            if result == -0.0:
                result = 0.0
            results.append(result)
        print(results)

        result_df = pd.DataFrame(results, columns =['result'])
        df = pd.concat([df1, result_df], axis=1)
        print(df)
        df.to_excel(path, index=False)
    else:
        raise NoPathError
except NoPathError:
    print("no file path given to parse. please try again!")

