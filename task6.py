import openpyxl
# Nolasām Excel failu
wb = openpyxl.load_workbook("Task6.xlsx")
sheet = wb.active

# Apvienojam datus no kolonnām A un B
merged_data = [f"{row[0]}-{row[1]}" for row in sheet.iter_rows(min_row=2, values_only=True) if row[0] and row[1]]

# Saglabājam TXT failā
with open("merged_task6.txt", "w") as f:
    f.writelines("\n".join(merged_data))