import openpyxl

from openpyxl import load_workbook

while True:
    print('Please enter the data you would like to add to your Excel document.')
    data = input()
    if (data == ''):
        print('Please enter some data.')
    else:
        while True:
            print('How is your data separated?')
            print('Please enter , (a comma) if your data is separated by commas. \n' +
            'Please enter ; (a semicolon) if your data is separated by semicolons. \n' +
            'Please enter  (a space) if your data is separated by spaces. \n')
            separationOptions = [',', ';', ' ']
            separationType = input()
            if separationType not in separationOptions:
                print('Please enter a valid separation method.')
            else:
                dataList = data.split(separationType)
                dataLength = len(dataList)
                while True:
                    print('Please enter the absolute path to an Excel document (e.g. file://Users/...)')
                    filePath = input()
                    if not (filePath.endswith('.xls') or filePath.endswith('.xlsx')):
                        print('Please input a valid Excel file.')
                    elif (load_workbook(filePath) == None):
                        print('This file could not be found. Please confirm your file path.')
                    else:
                        wb = load_workbook(filePath)
                        print('Please enter the name of the sheet where your data is located. \n' +
                        'Otherwise, the active sheet will be used by default.')
                        sheetName = input()
                        if ((sheetName == '') or (wb[sheetName] == None)):
                            ws = wb.active
                            print('You are using the active sheet.')
                        else:
                            ws = wb[sheetName]
                            print(f'You are using the sheet "{sheetName}"')
                        while True:
                            print('Enter the letter of the column you would like to check for duplicates.')
                            columnLetter = input()
                            if (columnLetter == ''):
                                print('Please enter a column letter.')
                            else:
                                print('If you are counting these items in a separate column,' +
                                'please enter the corresponding letter for that column.' +
                                'Otherwise, press ENTER to continue.')
                                countColumn = input()
                                sortedColumn = []
                                rows = ws[columnLetter].max_row
                                for i in range(rows):
                                    i += 1
                                    unsortedCell = ws[f'{columnLetter}{i}'].value
                                    sortedColumn.append(unsortedCell)
                                currentRow = int(rows) + 1
                                dataAdded = 0
                                for i in range(dataLength):
                                    if dataList[i] in sortedColumn:
                                        print(f'{dataList[i]} already exists.')
                                    else:
                                        ws[f'{columnLetter}{currentRow}'] = dataList[i]
                                        if not ((countColumn == '') or (countColumn == columnLetter)):
                                            continue
                                        else:
                                            previousRow = currentRow - 1
                                            previousCount = int(ws[f'{countColumn}{previousRow}'].value)
                                            ws[f'{countColumn}{currentRow}'] = previousCount
                                        currentRow += 1
                                        dataAdded += 1
                                print(f'Finished searching for duplicates. {dataAdded} value(s) added.')
                        break
                break
        break
    print('Program finished.')