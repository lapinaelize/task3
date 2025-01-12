import openpyxl
import matplotlib.pyplot as plt

# Nolasām Excel failu
wb = openpyxl.load_workbook("Task4.xlsx")
sheet = wb.active

# Nolasām datus no kolonnas C
values = [cell.value for cell in sheet["C"][1:] if isinstance(cell.value, (int, float))]

# Saglabājam TXT failā
with open("values_task4.txt", "w") as f:
    f.writelines("\n".join(map(str, values)))

# Izveidojam histogrammu
plt.hist(values, bins=10, color='blue', edgecolor='black')
plt.title("Histogramma")
plt.xlabel("Vērtības")
plt.ylabel("Biežums")
plt.savefig("histogram_task4.png")
plt.show()