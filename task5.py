import openpyxl
wb = openpyxl.load_workbook("Task5.xlsx")
sheet = wb.active

name = input("Ievadiet vārdu: ")
age = int(input("Ievadiet vecumu: "))
score = float(input("Ievadiet rezultātu: "))

sheet.append([name, age, score])
wb.save("Task5.xlsx")

with open("task5_data.txt", "a") as f:
    f.write(f"{name}, {age}, {score}\n")
