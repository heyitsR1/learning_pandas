from openpyxl import Workbook

wb = Workbook()
ws = wb.active
ws.title = "Users"

data = [
    ["id","name","email","age","city","country","signup_date","is_active"],
    [1,"Alice Johnson","alice.johnson@example.com",24,"New York","USA","2024-01-12",True],
    [2,"Bob Smith","bob.smith@example.com",31,"London","UK","2023-11-03",True],
]

for row in data:
    ws.append(row)

wb.save("users.xlsx")
