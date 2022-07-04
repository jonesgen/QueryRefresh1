import win32com.client as win32

xlApp = win32.Dispatch('Excel.Application')

Q1 = "C:/Users/Gene/Documents/New folder/querytest.xlsx"
Q2 = "C:/Users/Gene/Documents/New folder/querytest1.xlsx"
Q3 = "C:/Users/Gene/Documents/New folder/querytest2.xlsx"
names = [Q1,Q2,Q3]

for x in names:
    xlApp.Visible = True

    Workbook = xlApp.Workbooks.open (x)

    Workbook.RefreshAll()
    xlApp.CalculateUntilAsyncQueriesDone()
    xlApp.DisplayAlerts = False

    Workbook.Save()
    Workbook.Close()
    xlApp.Quit()







