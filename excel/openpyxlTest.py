import openpyxl

'''
openpyxl 使用例子
'''

def handle_excel_by_openpyxl(excelPath):
    file_path = excelPath
    workbook = openpyxl.load_workbook(file_path)

    # print(workbook.worksheets)

    sheet1 = workbook.worksheets

    for sheet in workbook.worksheets:
        for row in sheet.rows:
            for cell in row:
                print('%s |' % cell.value, end="")
            print('\n')


if __name__ == '__main__':
    handle_excel_by_openpyxl('/Users/longzhiwu/Downloads/test.xlsx')
