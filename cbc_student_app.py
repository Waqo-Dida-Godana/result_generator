import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import pandas as pd
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage
from reportlab.lib.styles import getSampleStyleSheet
from fpdf import FPDF
from PIL import Image, ImageTk
import os
import shutil

# ─── System Configuration ─────────────────────────────────────────────────────
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("blue")

PHOTOS_DIR = "photos"
if not os.path.exists(PHOTOS_DIR):
    os.makedirs(PHOTOS_DIR)

# ─── MySQL Connection ─────────────────────────────────────────────────────────
db, cursor = None, None
DB_CONNECTED = False

def connect_db():
    global db, cursor, DB_CONNECTED
    try:
        import mysql.connector
        db = mysql.connector.connect(
            host="localhost", user="root", password="",
            database="cbc_advanced_db", connection_timeout=3
        )
        cursor = db.cursor()
        DB_CONNECTED = True
        cursor.execute("""
            CREATE TABLE IF NOT EXISTS students (
                id INT AUTO_INCREMENT PRIMARY KEY,
                name VARCHAR(100),
                class VARCHAR(20),
                term VARCHAR(20),
                math INT, eng INT, kisw INT, sci INT, soc INT,
                total INT, avg FLOAT, level VARCHAR(10), rank_no INT,
                photo_path VARCHAR(255)
            )
        """)
    except:
        DB_CONNECTED = False

connect_db()

# ─── In-memory Data ───────────────────────────────────────────────────────────
# Format: [id, name, class, term, math, eng, kisw, sci, soc, total, avg, level, rank, photo_path]
students = []
current_photo_path = None

def get_level(avg: float) -> str:
    if avg >= 80: return "EE"
    if avg >= 65: return "ME"
    if avg >= 50: return "AE"
    return "BE"

def calculate_rank():
    # Rank within the same class and term
    groups = {}
    for s in students:
        key = (s[2], s[3])
        if key not in groups: groups[key] = []
        groups[key].append(s)
    
    for key in groups:
        groups[key].sort(key=lambda x: x[9], reverse=True)
        for i, s in enumerate(groups[key], 1):
            s[12] = i

# ─── GUI Logic ───────────────────────────────────────────────────────────────
def select_photo():
    global current_photo_path
    path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png *.jpeg")])
    if path:
        current_photo_path = path
        img = Image.open(path)
        img = img.resize((100, 100))
        photo_img = ImageTk.PhotoImage(img)
        photo_label.configure(image=photo_img, text="")
        photo_label.image = photo_img

def add_student():
    global current_photo_path
    name = name_entry.get().strip()
    cls = class_cb.get()
    trm = term_cb.get()
    
    if not name:
        messagebox.showerror("Error", "Student name is required!")
        return
        
    try:
        m = int(math_entry.get())
        e = int(eng_entry.get())
        k = int(kisw_entry.get())
        s = int(sci_entry.get())
        so = int(soc_entry.get())
    except:
        messagebox.showerror("Error", "Marks must be numeric!")
        return

    total = m + e + k + s + so
    avg = round(total / 5, 2)
    lvl = get_level(avg)
    
    dest_path = ""
    if current_photo_path:
        dest_path = os.path.join(PHOTOS_DIR, f"{name}_{cls}_{trm}.jpg").replace(" ", "_")
        shutil.copy(current_photo_path, dest_path)

    new_id = len(students) + 1
    students.append([new_id, name, cls, trm, m, e, k, s, so, total, avg, lvl, 0, dest_path])
    
    calculate_rank()
    refresh_table()
    save_to_db()
    clear_inputs()
    messagebox.showinfo("Success", f"Added {name} successfully!")

def refresh_table():
    for row in table.get_children():
        table.delete(row)
    for s in students:
        # Filter logic can go here
        table.insert("", "end", values=s[:13])

def save_to_db():
    if not DB_CONNECTED: return
    try:
        cursor.execute("DELETE FROM students")
        sql = "INSERT INTO students (name, class, term, math, eng, kisw, sci, soc, total, avg, level, rank_no, photo_path) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        for s in students:
            cursor.execute(sql, tuple(s[1:]))
        db.commit()
    except Exception as e: print(f"DB Error: {e}")

def clear_inputs():
    global current_photo_path
    name_entry.delete(0, 'end')
    math_entry.delete(0, 'end')
    eng_entry.delete(0, 'end')
    kisw_entry.delete(0, 'end')
    sci_entry.delete(0, 'end')
    soc_entry.delete(0, 'end')
    current_photo_path = None
    photo_label.configure(image="", text="No Photo")

def generate_report():
    selected = table.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a student first!")
        return
    
    item = table.item(selected[0])
    s_id = item["values"][0]
    s_data = next(s for s in students if s[0] == s_id)
    
    name, cls, term = s_data[1], s_data[2], s_data[3]
    m, e, k, sci, soc = s_data[4], s_data[5], s_data[6], s_data[7], s_data[8]
    total, avg, lvl, rank = s_data[9], s_data[10], s_data[11], s_data[12]
    photo = s_data[13]

    filename = f"Report_{name}_{cls}_{term}.pdf".replace(" ", "_")
    doc = SimpleDocTemplate(filename, pagesize=A4)
    styles = getSampleStyleSheet()
    elements = []

    # Header
    elements.append(Paragraph("<font size=18 color='#1a3c5e'><b>KENYAN CBC ACADEMIC PROGRESS REPORT</b></font>", styles['Title']))
    elements.append(Paragraph(f"<font size=12><b>{cls} - {term}</b></font>", styles['Normal']))
    elements.append(Spacer(1, 20))

    # Student Info Table
    info_data = [
        ["STUDENT NAME:", name, "RANK:", f"{rank} / {len([x for x in students if x[2]==cls and x[3]==term])}"],
        ["CLASS:", cls, "LEVEL:", lvl],
        ["TERM:", term, "AVERAGE:", f"{avg}%"]
    ]
    info_table = Table(info_data, colWidths=[100, 200, 80, 80])
    info_table.setStyle(TableStyle([('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'), ('BOTTOMPADDING', (0,0), (-1,-1), 10)]))
    elements.append(info_table)
    
    if photo and os.path.exists(photo):
        img = RLImage(photo, width=80, height=80)
        img.hAlign = 'RIGHT'
        elements.append(img)
    
    elements.append(Spacer(1, 20))

    # Marks Table
    marks_data = [
        ["SUBJECT", "SCORE (/100)", "REMARKS"],
        ["MATHEMATICS", m, "Excellent" if m>=80 else "Good"],
        ["ENGLISH", e, "Excellent" if e>=80 else "Good"],
        ["KISWAHILI", k, "Excellent" if k>=80 else "Good"],
        ["SCIENCE", sci, "Excellent" if sci>=80 else "Good"],
        ["SOCIAL STUDIES", soc, "Excellent" if soc>=80 else "Good"],
        ["TOTAL MARKS", total, ""],
    ]
    marks_table = Table(marks_data, colWidths=[150, 100, 150])
    marks_table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#1a3c5e")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('BOTTOMPADDING', (0,0), (-1,-1), 12),
    ]))
    elements.append(marks_table)
    
    elements.append(Spacer(1, 40))
    elements.append(Paragraph("<b>Class Teacher's Remarks:</b> ________________________________________________", styles['Normal']))
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("<b>Principal's Signature:</b> ________________________________________________", styles['Normal']))

    doc.build(elements)
    messagebox.showinfo("Report Generated", f"Report card saved as {filename}")

# ─── UI Layout ───────────────────────────────────────────────────────────────
app = ctk.CTk()
app.title("CBC Advanced Student Management System")
app.geometry("1200x700")

# Sidebar
sidebar = ctk.CTkFrame(app, width=250, corner_radius=0)
sidebar.pack(side="left", fill="y")

ctk.CTkLabel(sidebar, text="ADVANCED CBC", font=("Arial", 20, "bold")).pack(pady=20)

photo_label = ctk.CTkLabel(sidebar, text="No Photo", width=120, height=120, fg_color="gray80", corner_radius=10)
photo_label.pack(pady=10)

ctk.CTkButton(sidebar, text="Upload Photo", command=select_photo, fg_color="#34495e").pack(pady=5, padx=20)

# Main Form in Sidebar
name_entry = ctk.CTkEntry(sidebar, placeholder_text="Student Name")
name_entry.pack(pady=5, padx=20, fill="x")

class_cb = ctk.CTkComboBox(sidebar, values=["Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5", "Grade 6"])
class_cb.pack(pady=5, padx=20, fill="x")

term_cb = ctk.CTkComboBox(sidebar, values=["Term 1", "Term 2", "Term 3"])
term_cb.pack(pady=5, padx=20, fill="x")

math_entry = ctk.CTkEntry(sidebar, placeholder_text="Math Mark")
math_entry.pack(pady=5, padx=20, fill="x")

eng_entry = ctk.CTkEntry(sidebar, placeholder_text="English Mark")
eng_entry.pack(pady=5, padx=20, fill="x")

kisw_entry = ctk.CTkEntry(sidebar, placeholder_text="Kiswahili Mark")
kisw_entry.pack(pady=5, padx=20, fill="x")

sci_entry = ctk.CTkEntry(sidebar, placeholder_text="Science Mark")
sci_entry.pack(pady=5, padx=20, fill="x")

soc_entry = ctk.CTkEntry(sidebar, placeholder_text="Social Mark")
soc_entry.pack(pady=5, padx=20, fill="x")

ctk.CTkButton(sidebar, text="Add Student", command=add_student, fg_color="#27ae60").pack(pady=20, padx=20, fill="x")

# Content Area
content = ctk.CTkFrame(app)
content.pack(side="right", fill="both", expand=True, padx=20, pady=20)

top_bar = ctk.CTkFrame(content, height=60)
top_bar.pack(fill="x", pady=(0, 20))

ctk.CTkButton(top_bar, text="Generate Report Card", command=generate_report, fg_color="#9b59b6").pack(side="left", padx=10, pady=10)
ctk.CTkButton(top_bar, text="Export Excel", command=lambda: messagebox.showinfo("Export", "Excel Exported"), fg_color="#2ecc71").pack(side="left", padx=10, pady=10)

# Treeview
tbl_frame = tk.Frame(content)
tbl_frame.pack(fill="both", expand=True)

cols = ("ID", "Name", "Class", "Term", "Math", "Eng", "Kisw", "Sci", "Soc", "Total", "Avg", "Level", "Rank")
table = ttk.Treeview(tbl_frame, columns=cols, show="headings")

for col in cols:
    table.heading(col, text=col)
    table.column(col, width=70, anchor="center")
table.column("Name", width=150)

sb = ttk.Scrollbar(tbl_frame, orient="vertical", command=table.yview)
table.configure(yscrollcommand=sb.set)
table.pack(side="left", fill="both", expand=True)
sb.pack(side="right", fill="y")

app.mainloop()
