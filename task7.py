import openpyxl
wb = openpyxl.load_workbook("Task7.xlsx")
sheet = wb.active

old_value = input("Ievadiet vērtību, ko aizvietot: ")
new_value = input("Ievadiet jauno vērtību: ")

for cell in sheet["E"]:
    if cell.value == old_value:
        cell.value = new_value

wb.save("Task7.xlsx")

with open("updated_task7.txt", "w") as f:
    for row in sheet.iter_rows(values_only=True):
        f.write(f"{row}\n")
