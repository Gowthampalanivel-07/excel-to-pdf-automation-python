import pandas as pd
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
import os
from datetime import datetime

excel_file = "student_marks.xlsx"
df = pd.read_excel(excel_file)

print("Excel Data:")
print(df)

total_students = len(df)
average_marks = df["Marks"].mean()

output_dir = os.path.join(os.getcwd(), "output")
os.makedirs(output_dir, exist_ok=True)

date_str = datetime.now().strftime("%Y-%m-%d_%H-%M-%S")
pdf_file = os.path.join(output_dir, f"report_{date_str}.pdf")

c = canvas.Canvas(pdf_file, pagesize=A4)
width, height = A4

c.setFont("Helvetica-Bold", 14)
c.drawString(50, height - 50, "Student Marks Report")

c.setFont("Helvetica", 11)
c.drawString(50, height - 80, f"Total Students: {total_students}")
c.drawString(50, height - 100, f"Average Marks: {average_marks:.2f}")

y = height - 140
c.setFont("Helvetica", 10)

for _, row in df.iterrows():
    line = f"{row['Name']} | {row['Department']} | Marks: {row['Marks']}"
    c.drawString(50, y, line)
    y -= 18

c.save()

print("PDF created successfully:", pdf_file)
