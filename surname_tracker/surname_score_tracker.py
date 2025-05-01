import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
import openpyxl
import os

EXCEL_FILE = "student_scores.xlsx"

def create_or_load_workbook():
    if not os.path.exists(EXCEL_FILE):
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = "Scores"
        sheet.append(["Student Name", "Score", "Status"])
        workbook.save(EXCEL_FILE)
    return openpyxl.load_workbook(EXCEL_FILE)

def add_student_score(name, score):
    workbook = create_or_load_workbook()
    sheet = workbook["Scores"]
    status = "Pass" if score >= 75 else "Fail"
    sheet.append([name, score, status])
    workbook.save(EXCEL_FILE)

def get_all_records():
    workbook = create_or_load_workbook()
    sheet = workbook["Scores"]
    return list(sheet.iter_rows(min_row=2, values_only=True))

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
            messagebox.showwarning("Input Error", "Score must be betwee 0 and 100.")
    except ValueError:
        messagebox.showwarning("Input Error", "Please eter a valid integer for the score.")

def refresh_records():
    for row in tree.get_children():
        tree.delete(row)
    for record in get_all_records():
        tree.insert("", tk.END, values=record)

root = tk.Tk()
root.title("Student Score Tracker")

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

frame_display = tk.Frame(root, padx=10, pady=10)
frame_display.pack()

tree = ttk.Treeview(frame_display, columns=("Name", "Score", "Status"), show="headings")
tree.heading("Name", text="Student Name")
tree.heading("Score", text="Score")
tree.heading("Status", text="Status")
tree.pack()

refresh_records()

root.mainloop()
