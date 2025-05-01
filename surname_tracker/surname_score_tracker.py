import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import openpyxl
import os

EXCEL_FILE = "student_scores.xlsx"

# Ensure the Excel file exists with correct headers
def create_or_load_workbook():
    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Scores"
        sheet.append(["Student Name", "Score", "Status"])
        workbook.save(EXCEL_FILE)
    return openpyxl.load_workbook(EXCEL_FILE)

# Add student score, determine pass/fail, and calculate average
def add_student_score(name, score):
    workbook = create_or_load_workbook()
    sheet = workbook["Scores"]

    # Remove previous "Average Score" row if it exists
    for row in sheet.iter_rows():
        if row[0].value == "Average Score":
            sheet.delete_rows(row[0].row, 1)
            break

    # Add student entry
    status = "Pass" if score >= 75 else "Fail"
    sheet.append([name, score, status])
    
    # Recalculate and append average score
    scores = [
        row[1].value for row in sheet.iter_rows(min_row=2, max_col=2)
        if isinstance(row[1].value, (int, float))
    ]
    if scores:
        avg_score = sum(scores) / len(scores)
        sheet.append(["Average Score", round(avg_score, 2), ""])
    
    workbook.save(EXCEL_FILE)   

# Get all student records (exclude "Average Score" row)
def get_all_records():
    workbook = create_or_load_workbook()
    sheet = workbook["Scores"]
    records = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if row[0] != "Average Score":
            records.append(row)
    return records

# Get average score (read from Excel)
def get_average_score():
    workbook = create_or_load_workbook()
    sheet = workbook["Scores"]
    for row in sheet.iter_rows(values_only=True):
        if row[0] == "Average Score":
            return row[1]
    return None

# Submit new score via GUI
def submit_score():
    name = entry_name.get()
    try:
        score = int(entry_score.get())
        if not name:
            messagebox.showwarning("Input Error", "Please enter a student name.")
            return
        if 0 <= score <= 100:
            add_student_score(name, score)
            messagebox.showinfo("Success", f"{name}'s score saved.")
            entry_name.delete(0, tk.END)
            entry_score.delete(0, tk.END)
            refresh_records()
        else:
            messagebox.showwarning("Input Error", "Score must be between 0 and 100.")
    except ValueError:
        messagebox.showwarning("Input Error", "Please enter a valid integer for the score.")

# Refresh display table and average
def refresh_records():
    for row in tree.get_children():
        tree.delete(row)
    for record in get_all_records():
        tree.insert("", tk.END, values=record)

    avg = get_average_score()
    if avg is not None:
        label_avg.config(text=f"Average Score: {avg}")
    else:
        label_avg.config(text="Average Score: N/A")

# GUI Layout
root = tk.Tk()
root.title("Student Score Tracker")

# Input Frame
frame_input = tk.Frame(root, padx=10, pady=10)
frame_input.pack()

tk.Label(frame_input, text="Student Name:").grid(row=0, column=0, padx=5, pady=5)
entry_name = tk.Entry(frame_input)
entry_name.grid(row=0, column=1, padx=5, pady=5)

tk.Label(frame_input, text="Score:").grid(row=1, column=0, padx=5, pady=5)
entry_score = tk.Entry(frame_input)
entry_score.grid(row=1, column=1, padx=5, pady=5)

btn_submit = tk.Button(frame_input, text="Submit Score", command=submit_score)
btn_submit.grid(row=2, column=0, columnspan=2, pady=10)

# Record Display Frame
frame_display = tk.Frame(root, padx=10, pady=10)
frame_display.pack()

tree = ttk.Treeview(frame_display, columns=("Name", "Score", "Status"), show="headings")
tree.heading("Name", text="Student Name")
tree.heading("Score", text="Score")
tree.heading("Status", text="Status")
tree.pack()

# Average Score Label
label_avg = tk.Label(root, text="Average Score: N/A", font=("Arial", 12, "bold"))
label_avg.pack(pady=10)

# Initialize record view
refresh_records()

root.mainloop()
