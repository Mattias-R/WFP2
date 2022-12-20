from sklearn import linear_model
import numpy as np
import matplotlib.pyplot as plt
import openpyxl
Z = []
#X=np.array([1,2,3,4,5,6,7,8,9,10 ]).reshape(-1, 1)
#Y=[2,4,3,6,8,9,9,10,11,13]
Y=[]
lm = linear_model.LinearRegression()

inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\verbrauchsdaten_Neuhaus.xlsx'
#inputExcelFile =r'C:\Users\matti\PycharmProjects\csvIntoJson\verbrauchsdaten_schwabbeg.xlsx'
# creating or loading an excel workbook
newWorkbook = openpyxl.load_workbook(inputExcelFile)
sheet_obj = newWorkbook.active
# Accessing specific sheet of the workbook.
firstWorksheet = newWorkbook["Tabelle1"]
maxLines = sheet_obj.max_row
test = ["C"]
tableName = firstWorksheet["D"]
helper = []
value = []
help = 0
for i in range(maxLines):
    if (i == 0):
        continue
    # if value is unrealistic, make it NA
    if (tableName[i].value != "NA" and float(tableName[i].value) > 3000):
        tableName[i].value = "NA"
    # skip loop if value is NA
    if (tableName[i].value == "NA"):
        continue
    # skip loop if value is WM
    if (str(tableName[i - 1].value).startswith("WM")):
        continue
    helper.append(float(tableName[i].value))
    if (tableName[i - 1].value == "NA"):
        difference = round(float(tableName[i].value) - float(helper[len(helper) - 2]), 2)
        Y.append(difference)
        print(tableName[i].value)
        help = help + 1
        Z.append(i)
        continue
    else:
        difference = round(float(tableName[i].value) - float(tableName[i - 1].value), 2)
        Y.append(difference)
        print(tableName[i].value)
    help = help + 1
    Z.append(help)
X=np.array(Z).reshape(-1, 1)
lm.fit(X, Y)
plt.scatter(X, Y, color = "r",marker = "o", s = 30)
y_pred = lm.predict(X)
plt.plot(X, y_pred, color = "k")
plt.xlabel('x')
plt.ylabel('y')
plt.title("Simple Linear Regression")
plt.show()