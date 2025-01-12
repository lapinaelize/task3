import openpyxl

# 1. uzdevums: Datu summēšana un vidējā vērtība no kolonnas B
wb = openpyxl.load_workbook("Task1_Task2.xlsx")
sheet = wb.active

# Nolasām datus no kolonnas B (Value)
values = [cell.value for cell in sheet["B"][1:] if isinstance(cell.value, (int, float))]

# Aprēķinām summu un vidējo vērtību
total = sum(values)
average = total / len(values)

# Saglabājam rezultātu TXT failā
with open("results_task1.txt", "w", encoding="utf-8") as f:
    f.write(f"Kopējā summa: {total}\n")
    f.write(f"Vidējā vērtība: {average:.2f}\n")

# 2. uzdevums: Filtrēšana pēc noteikta sliekšņa
# Filtrējam rindas, kur kolonnas C vērtība ir lielāka par 50
filtered_rows = [
    row for row in sheet.iter_rows(min_row=2, values_only=True) if row[2] and row[2] > 50
]

# Saglabājam filtrētās rindas TXT failā
with open("filtered_task2.txt", "w", encoding="utf-8") as f:
    for row in filtered_rows:
        f.write(", ".join(map(str, row)) + "\n")


