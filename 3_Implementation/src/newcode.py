import openpyxl  # import openpyxl module
from openpyxl.chart import BarChart, Reference  # importing openpyxl for Bar Graph

# To open the workbook,theFile is created
theFile = openpyxl.load_workbook('crime.xlsx')
allSheetNames = theFile.sheetnames
# print("All sheet names {} " .format(theFile.sheetnames))
master = theFile['mastersheet']  # to open mastersheet , master is created

print('How Many Input u Want: ')  # Taking how much input user want
b = int(input())
row_data = 1

for i in range(0, b):
    cityname = (input('Enter the Area name: '))  # Taking input from user
    for sheet in allSheetNames:  # itreating through all sheet avilable in my Excel file

        currentSheet = theFile[sheet]  # assiging avilable sheets to variable accoding to the loop

        for row in range(1, currentSheet.max_row + 1):  # itreating through all rows in currentsheet
            # print(row)
            for column in 'ABCDEFGHIJ':  # Here you can add or reduce the columns
                cell_name = "{}{}".format(column, row)
                # print(cell_name)

                if currentSheet[cell_name].value == cityname:  # checking if user input is equal to cell value or not
                    r = row  # where the input will be equal to cell value it will store that row number.

        for j in range(1, currentSheet.max_column + 1):
            master.cell(row=row_data, column=j).value = currentSheet.cell(row=r,
                                                                          column=j).value  # using this loop to print the data in master sheet
            # print(row_data)
        row_data = row_data + 1

theFile.save('crime.xlsx')

mastersheet = theFile["mastersheet"]  # to open mastersheet, mastersheet object is crteated
values = Reference(mastersheet, min_col=1, min_row=3, max_col=6, max_row=mastersheet.max_row)
chart = BarChart()
chart.add_data(values)
chart.title = " BAR-CHART "  # title of the Bar chart
chart.x_axis.title = " Area "  # Giving name for X axis
chart.y_axis.title = " Crime "  # Giving name for Y axis
mastersheet.add_chart(chart, "E2")
theFile.save("crime.xlsx")  # saving the bar graph in mastersheet
