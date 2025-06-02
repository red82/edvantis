import os
from openpyxl import load_workbook
import pandas as pd

# Ensure the output directory exists
os.makedirs("outputs", exist_ok=True)

file_path = "MMA_Club_Stats_20_Students.xlsx"
workbook = load_workbook(file_path)
sheet = workbook.active

data = sheet.values
columns = next(data)
df = pd.DataFrame(data, columns=columns)

# Filtering by belt level
black_belts = df[df["Belt Level"] == "Black"]
print(f"Number of Black Belts: {len(black_belts)}")

# Calculation
df["Efficiency Score"] = (df["Sparring Wins"] - df["Sparring Losses"]) * df["Endurance (rounds)"]

# Top-5 by attendance
top_attendance = df.sort_values(by="Total Attendance", ascending=False).head(5)
print("Top 5 Members by Attendance:")
print(top_attendance[["Name", "Total Attendance"]])

# Save intermediate file
df.to_excel("outputs/processed_stats.xlsx", index=False)
print("Saved successfully in 'outputs/processed_stats.xlsx'")