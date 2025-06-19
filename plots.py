import pandas as pd
import matplotlib.pyplot as plt

# Define file path and load Excel File
file_path = "Data.xlsx"
xl = pd.ExcelFile(file_path)

# Read 'Export Gen kWh' Sheet
df = pd.read_excel(xl, "Export Gen kWh")
df.columns = [col.strip() for col in df.columns]  # Clean column names

# Extract kWh where kWh > 3000 and Date is April 10, 2023
target_mpan = 1099999999999
target_date = '2023-04-02'
condition = (df['Meter Point Administration Number'] == target_mpan) & (df['Date'] == pd.to_datetime(target_date))
specific_rows = df['kWh'][condition]
print(specific_rows)

list_mpan_day = specific_rows.tolist()

print(list_mpan_day)


x = []
for i in range(len(list_mpan_day)):
    x.append(i)

fig, ax = plt.subplots()
ax.plot(x,list_mpan_day)
ax.set_title('24-Hour Export Generation for Customer A')
ax.set_xlabel('Settlement Period (1/2 Hour)')
ax.set_ylabel('Generation (kWh)')
plt.show()
