from openpyxl import Workbook

def create_workbook(path):
    workbook = Workbook()
    workbook.save(path)

create_workbook("student.xlsx")
