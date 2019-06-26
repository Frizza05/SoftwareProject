import pandas as pd
import csv
data = pd.read_csv("EmployeeIN.csv")

employeedata = []
print (data)
print(employeedata)

if len(data.loc[:0]["ID":"Total"]) == 0:
    print("None")

elif len(data["ID"]) != 0:
    counter = 0
    employeedata = []
    
    employeedata.append(data["ID"][1])
    print(employeedata)
                            
