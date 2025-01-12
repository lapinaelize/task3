import openpyxl
wb = openpyxl.load_workbook("Task3.xlsx")
sheet = wb.active

texts = [cell.value.upper() for cell in sheet["A"][1:] if isinstance(cell.value, str)]

with open("uppercase_task3.txt", "w") as f:
    f.writelines("\n".join(texts))
