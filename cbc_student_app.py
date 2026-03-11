import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import customtkinter as ctk
import pandas as pd
import openpyxl
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.lib.units import inch
from fpdf import FPDF
from PIL import Image, ImageTk
import os
import shutil

# ─── System Configuration ─────────────────────────────────────────────────────
# Theme colors inspired by the Mt Olives Adventist School logo
# Using green (olives/nature) and blue (trust/education) theme
ctk.set_appearance_mode("Light")
ctk.set_default_color_theme("green")

# Primary Theme Colors - School Logo Inspired
PRIMARY_GREEN = "#2E7D32"       # Forest green - main brand color (olives)
PRIMARY_GREEN_DARK = "#1B5E20"  # Darker green for hover states
PRIMARY_BLUE = "#1565C0"        # Blue - trust and education
PRIMARY_BG = "#F5F7FA"          # Light gray-blue background

# Sidebar Colors
SIDEBAR_BG = "#1B3A2B"          # Dark green - sidebar background
SIDEBAR_HOVER = "#2D5A45"       # Lighter green for hover
SIDEBAR_ACTIVE = "#3D7A5F"      # Active nav highlight
SIDEBAR_TEXT = "#E8F5E9"        # Light green text
SIDEBAR_TEXT_MUTED = "#A5D6A7" # Muted green text

# Card Colors
CARD_BG = "#FFFFFF"
CARD_BORDER = "#E0E7FF"
CARD_SHADOW = "#E2E8F0"

# Accent Colors
ACCENT = "#1565C0"              # Blue accent
ACCENT_LIGHT = "#42A5F5"        # Light blue
ACCENT_GREEN = "#4CAF50"       # Green accent
GOLD = "#FFB300"               # Gold for highlights/trophies

# Button Colors
BUTTON_BG = "#2E7D32"           # Green button
BUTTON_HOVER = "#1B5E20"       # Darker green hover
BUTTON_SECONDARY = "#1565C0"   # Blue secondary button

# Text Colors
LABEL_COLOR = "#1E3A5F"        # Dark blue-gray
TEXT_PRIMARY = "#1A1A2E"       # Near black
TEXT_SECONDARY = "#64748B"     # Gray
TEXT_LIGHT = "#FFFFFF"         # White text

# Dashboard Stats Colors
STAT_TOTAL = "#2E7D32"         # Green
STAT_AVG = "#1565C0"            # Blue
STAT_TOP = "#FFB300"           # Gold
STAT_SUBJECTS = "#7B1FA2"      # Purple

# Footer Colors
FOOTER_BG = "#1B3A2B"          # Match sidebar
FOOTER_TEXT = "#A5D6A7"        # Muted green

# Grade Colors
GRADE_EE = "#4CAF50"          # Green
GRADE_ME = "#8BC34A"           # Light green
GRADE_AE = "#FFC107"          # Amber
GRADE_BE = "#FF9800"          # Orange
GRADE_IE = "#F44336"           # Red

# Load logo (place a file named logo.png/jpg in the project folder)
# Also try loading from Downloads folder if not in project
LOGO_IMAGE = None
LOGO_SMALL = None
for logo_path in ("logo.png", "logo.jpg", "logo.jpeg", "logo.gif", 
                  "C:/Users/WDG/Downloads/mos.jpg", "C:\\Users\\WDG\\Downloads\\mos.jpg"):
    if os.path.exists(logo_path):
        try:
            _img = Image.open(logo_path).convert("RGBA")
            _img_resized = _img.resize((100, 100), Image.LANCZOS)
            LOGO_IMAGE = ImageTk.PhotoImage(_img_resized)
            # Smaller version for dashboard
            _img_small = _img.resize((60, 60), Image.LANCZOS)
            LOGO_SMALL = ImageTk.PhotoImage(_img_small)
            break
        except Exception:
            pass

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
                admission_no VARCHAR(50),
                name VARCHAR(100),
                class VARCHAR(20),
                term VARCHAR(20),
                math INT, eng INT, kisw INT, int_sci INT, agri INT,
                sst INT, cre INT, cia INT, pre_tech INT,
                total INT, avg FLOAT, level VARCHAR(10), rank_no INT,
                photo_path VARCHAR(255),
                gender VARCHAR(20)
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
        groups[key].sort(key=lambda x: x[13], reverse=True)
        for i, s in enumerate(groups[key], 1):
            s[16] = i

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

def add_student_from_form(admission_no: str, name: str, cls: str, term: str, gender: str, math: str="0", eng: str="0", kisw: str="0", int_sci: str="0", agri: str="0", sst: str="0", cre: str="0", cia: str="0", pre_tech: str="0", photo_path: str=""):
    global current_photo_path

    admission_no = admission_no.strip()
    name = name.strip()
    if not name or not admission_no:
        messagebox.showerror("Error", "Admission No and Student Name are required!")
        return

    try:
        m = int(math)
        e = int(eng)
        k = int(kisw)
        isc = int(int_sci)
        agr = int(agri)
        ss = int(sst)
        cr = int(cre)
        ci = int(cia)
        pt = int(pre_tech)
    except:
        messagebox.showerror("Error", "Marks must be numeric!")
        return

    total = m + e + k + isc + agr + ss + cr + ci + pt
    avg = round(total / 9, 2)
    lvl = get_level(avg)

    dest_path = ""
    if photo_path:
        dest_path = os.path.join(PHOTOS_DIR, f"{name}_{cls}_{term}.jpg").replace(" ", "_")
        try:
            shutil.copy(photo_path, dest_path)
        except Exception:
            dest_path = ""

    students.append([admission_no, name, cls, term, m, e, k, isc, agr, ss, cr, ci, pt, total, avg, lvl, 0, dest_path, gender])

    calculate_rank()
    refresh_table()
    save_to_db()
    messagebox.showinfo("Success", f"Added {name} successfully!")

def import_from_excel():
    """Import students from an Excel file"""
    file_path = filedialog.askopenfilename(
        title="Select Excel File",
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    if not file_path:
        return
    
    try:
        df = pd.read_excel(file_path)
        
        # Expected columns: admission_no, name, class, term, gender, math, eng, kisw, sci, soc, photo_path
        required_cols = ['admission_no', 'name', 'class']
        missing_cols = [col for col in required_cols if col not in df.columns]
        
        if missing_cols:
            messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_cols)}")
            return
        
        imported_count = 0
        for _, row in df.iterrows():
            admission_no = str(row.get('admission_no', '')).strip()
            name = str(row.get('name', '')).strip()
            cls = str(row.get('class', row.get('Class', 'Grade 1'))).strip()
            term = str(row.get('term', row.get('Term', 'Term 1'))).strip()
            gender = str(row.get('gender', row.get('Gender', 'Male'))).strip()
            photo_path = str(row.get('photo_path', row.get('photo', ''))).strip()
            
            # Get marks (default to 0 if not present)
            math = int(row.get('math', row.get('Math', 0)))
            eng = int(row.get('eng', row.get('Eng', 0)))
            kis = int(row.get('kis', row.get('Kis', 0)))
            isc = int(row.get('int_sci', row.get('Int Sci', 0)))
            agr = int(row.get('agri', row.get('Agri', 0)))
            sst = int(row.get('sst', row.get('SST', 0)))
            cre = int(row.get('cre', row.get('CRE', 0)))
            cia = int(row.get('cia', row.get('CIA', 0)))
            pt = int(row.get('pre_tech', row.get('Pre-Tech', 0)))
            
            if not name or not admission_no:
                continue
            
            # Copy photo if path exists
            dest_path = ""
            if photo_path and os.path.exists(photo_path):
                dest_path = os.path.join(PHOTOS_DIR, f"{name}_{cls}_{term}.jpg").replace(" ", "_")
                try:
                    shutil.copy(photo_path, dest_path)
                except Exception:
                    dest_path = ""
            
            total = math + eng + kis + isc + agr + sst + cre + cia + pt
            avg = round(total / 9, 2) if total > 0 else 0
            lvl = get_level(avg)
            
            students.append([admission_no, name, cls, term, math, eng, kis, isc, agr, sst, cre, cia, pt, total, avg, lvl, 0, dest_path, gender])
            imported_count += 1
        
        calculate_rank()
        refresh_table()
        save_to_db()
        messagebox.showinfo("Import Complete", f"Successfully imported {imported_count} students!")
        
    except Exception as e:
        messagebox.showerror("Import Error", f"Failed to import Excel file: {str(e)}")

def download_template():
    """Download an Excel template for adding students"""
    file_path = filedialog.asksaveasfilename(
        title="Save Template",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="student_template.xlsx"
    )
    if not file_path:
        return
    
    try:
        # Create template DataFrame
        template_data = {
            'admission_no': ['001', '002', '003'],
            'name': ['John Doe', 'Jane Smith', 'Michael Johnson'],
            'class': ['Grade 1', 'Grade 2', 'Grade 3'],
            'term': ['Term 1', 'Term 1', 'Term 1'],
            'gender': ['Male', 'Female', 'Male'],
            'math': [0, 0, 0],
            'eng': [0, 0, 0],
            'kis': [0, 0, 0],
            'int_sci': [0, 0, 0],
            'agri': [0, 0, 0],
            'sst': [0, 0, 0],
            'cre': [0, 0, 0],
            'cia': [0, 0, 0],
            'pre_tech': [0, 0, 0],
            'photo_path': ['', '', '']
        }
        df = pd.DataFrame(template_data)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Template Created", f"Template saved to: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to create template: {str(e)}")

def export_student_list():
    """Export all students to an Excel file"""
    file_path = filedialog.asksaveasfilename(
        title="Export Students",
        defaultextension=".xlsx",
        filetypes=[("Excel files", "*.xlsx")],
        initialfile="students_list.xlsx"
    )
    if not file_path:
        return
    
    try:
        export_data = []
        for s in students:
            while len(s) < 19: s.append("")
            export_data.append({
                'admission_no': s[0],
                'name': s[1],
                'class': s[2],
                'term': s[3],
                'math': s[4],
                'eng': s[5],
                'kis': s[6],
                'int_sci': s[7],
                'agri': s[8],
                'sst': s[9],
                'cre': s[10],
                'cia': s[11],
                'pre_tech': s[12],
                'total': s[13],
                'average': s[14],
                'level': s[15],
                'rank': s[16],
                'photo_path': s[17],
                'gender': s[18]
            })
        
        df = pd.DataFrame(export_data)
        df.to_excel(file_path, index=False)
        messagebox.showinfo("Export Complete", f"Exported {len(students)} students to: {file_path}")
    except Exception as e:
        messagebox.showerror("Error", f"Failed to export: {str(e)}")

def refresh_table():
    for row in table.get_children():
        table.delete(row)
    for s in students:
        # Display only the columns we need for the students view
        student_id = s[0]
        name = s[1]
        cls = s[2]
        gender = s[18] if len(s) > 18 else "-"
        table.insert("", "end", values=(student_id, name, cls, gender, "Delete"))


def delete_selected_student():
    selected = table.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a student to delete.")
        return

    item = table.item(selected[0])
    s_id = item["values"][0]
    for i, s in enumerate(students):
        if s[0] == s_id:
            del students[i]
            break

    calculate_rank()
    refresh_table()


def init_add_student_page():
    for child in add_student_frame.winfo_children():
        child.destroy()
    
    add_student_frame.configure(fg_color=PRIMARY_BG)
    
    # Header
    header_frame = ctk.CTkFrame(add_student_frame, fg_color="transparent")
    header_frame.pack(fill="x", padx=24, pady=(20, 10))
    
    ctk.CTkLabel(header_frame, text="Add New Student", font=("Arial", 28, "bold"), 
                 text_color=LABEL_COLOR).pack(side="left")
    
    form_container = ctk.CTkFrame(add_student_frame, fg_color="transparent")
    form_container.pack(fill="both", expand=True, padx=24, pady=10)

    def create_field(parent, label_text, placeholder):
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=8)
        
        ctk.CTkLabel(field_frame, text=label_text, font=("Arial", 14, "bold"), text_color=LABEL_COLOR).pack(anchor="w", pady=(0, 4))
        entry = ctk.CTkEntry(
            field_frame, 
            placeholder_text=placeholder, 
            height=45, 
            fg_color="#ffffff", 
            border_color=CARD_BORDER, 
            text_color=TEXT_PRIMARY,
            corner_radius=10
        )
        entry.pack(fill="x")
        return entry

    def create_dropdown(parent, label_text, values):
        field_frame = ctk.CTkFrame(parent, fg_color="transparent")
        field_frame.pack(fill="x", pady=8)
        
        ctk.CTkLabel(field_frame, text=label_text, font=("Arial", 14, "bold"), text_color=LABEL_COLOR).pack(anchor="w", pady=(0, 4))
        combo = ctk.CTkComboBox(
            field_frame, 
            values=values, 
            height=45, 
            fg_color="#ffffff", 
            border_color=CARD_BORDER, 
            text_color=TEXT_PRIMARY,
            button_color="#ffffff",
            button_hover_color="#f3f4f6",
            dropdown_fg_color="#ffffff",
            corner_radius=10
        )
        combo.pack(fill="x")
        return combo

    adm_no_entry = create_field(form_container, "Admission No", "e.g. 001")
    name_entry = create_field(form_container, "Full Name", "e.g. Desmond Kipchumba")
    class_cb = create_dropdown(form_container, "Class", ["Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5", "Grade 6", "Grade 7"])
    gender_cb = create_dropdown(form_container, "Gender", ["Male", "Female", "Other"])

    # Action Buttons
    btn_frame = ctk.CTkFrame(add_student_frame, fg_color="transparent")
    btn_frame.pack(fill="x", side="bottom", padx=24, pady=24)

    # Import from Excel button
    ctk.CTkButton(
        btn_frame, 
        text="📥 Import Excel", 
        command=lambda: [import_from_excel(), show_frame(students_frame)], 
        fg_color=PRIMARY_BLUE, 
        hover_color="#0D47A1", 
        text_color="#ffffff",
        height=45,
        width=140,
        font=("Arial", 13, "bold"),
        corner_radius=10
    ).pack(side="left")

    def _save():
        add_student_from_form(
            adm_no_entry.get(),
            name_entry.get(),
            class_cb.get(),
            "Term 1", # Default term
            gender_cb.get()
        )
        show_frame(students_frame)

    ctk.CTkButton(
        btn_frame, 
        text="Add", 
        command=_save, 
        fg_color=PRIMARY_GREEN, 
        hover_color=PRIMARY_GREEN_DARK, 
        text_color="#ffffff",
        height=45,
        width=120,
        font=("Arial", 14, "bold"),
        corner_radius=10
    ).pack(side="right", padx=(10, 0))

    ctk.CTkButton(
        btn_frame, 
        text="Cancel", 
        command=lambda: show_frame(students_frame), 
        fg_color="#ffffff", 
        hover_color="#f3f4f6", 
        text_color=LABEL_COLOR,
        border_width=1,
        border_color=CARD_BORDER,
        height=45,
        width=120,
        font=("Arial", 14, "bold"),
        corner_radius=10
    ).pack(side="right")


def save_to_db():
    if not DB_CONNECTED: return
    try:
        cursor.execute("DELETE FROM students")
        sql = "INSERT INTO students (admission_no, name, class, term, math, eng, kisw, int_sci, agri, sst, cre, cia, pre_tech, total, avg, level, rank_no, photo_path, gender) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"
        for s in students:
            # Ensure s has 19 elements
            while len(s) < 19: s.append("")
            cursor.execute(sql, tuple(s[:19]))
        db.commit()
    except Exception as e: print(f"DB Error: {e}")

def clear_inputs():
    global current_photo_path
    try:
        name_entry.delete(0, 'end')
        math_entry.delete(0, 'end')
        eng_entry.delete(0, 'end')
        kisw_entry.delete(0, 'end')
        sci_entry.delete(0, 'end')
        soc_entry.delete(0, 'end')
    except NameError:
        pass
    current_photo_path = None
    try:
        photo_label.configure(image="", text="No Photo")
    except NameError:
        pass

def generate_report():
    selected = table.selection()
    if not selected:
        messagebox.showwarning("Warning", "Please select a student first!")
        return
    
    item = table.item(selected[0])
    s_id = item["values"][0]
    s_data = next(s for s in students if s[0] == s_id)
    
    # Ensure s_data has enough elements
    while len(s_data) < 19: s_data.append("")
    
    adm, name, cls, term = s_data[0], s_data[1], s_data[2], s_data[3]
    m, e, k, isc, agr, ss, cr, ci, pt = s_data[4], s_data[5], s_data[6], s_data[7], s_data[8], s_data[9], s_data[10], s_data[11], s_data[12]
    total, avg, lvl, rank = s_data[13], s_data[14], s_data[15], s_data[16]
    photo = s_data[17]
    gender = s_data[18]

    filename = f"Report_{name}_{cls}_{term}.pdf".replace(" ", "_")
    doc = SimpleDocTemplate(filename, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    styles = getSampleStyleSheet()
    
    # Custom styles
    styles.add(ParagraphStyle(name='SchoolName', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, textColor=colors.HexColor("#2159A6"), alignment=1))
    styles.add(ParagraphStyle(name='SchoolInfo', parent=styles['Normal'], fontSize=9, textColor=colors.gray, alignment=1))
    styles.add(ParagraphStyle(name='ReportTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor("#D9534F"), alignment=1, spaceBefore=10, spaceAfter=10))
    styles.add(ParagraphStyle(name='SectionHeader', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=12, textColor=colors.white))
    styles.add(ParagraphStyle(name='CommentTitle', parent=styles['Normal'], fontName='Helvetica-Bold', fontSize=10, textColor=colors.HexColor("#2159A6")))
    styles.add(ParagraphStyle(name='CommentText', parent=styles['Normal'], fontName='Helvetica-Oblique', fontSize=10, textColor=colors.gray))

    elements = []

    # Header section
    header_data = [
        [Paragraph("VISION PRIMARY AND JUNIOR SCHOOLS", styles['SchoolName'])],
        [Paragraph("P.O. BOX 54 KENYA", styles['SchoolInfo'])],
        [Paragraph("visionprimaryschool@gmail.com", styles['SchoolInfo'])],
        [Paragraph("+254718481515/+254718481515", styles['SchoolInfo'])],
        [Paragraph("<i>Education For Excellence</i>", styles['SchoolInfo'])],
        [Spacer(1, 10)],
        [Paragraph("<hr color='#2159A6' width='100%'/>", styles['Normal'])],
        [Paragraph("LEARNER ASSESSMENT REPORT CARD", styles['ReportTitle'])]
    ]
    header_table = Table(header_data, colWidths=[520])
    elements.append(header_table)

    # Student Info Grid
    info_data = [
        ["NAME", name.upper(), "GRADE", cls],
        ["STREAM", "001", "YEAR", "2026"],
        ["TERM", term, "GENDER", gender]
    ]
    info_table = Table(info_data, colWidths=[80, 200, 80, 160])
    info_table.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), 'Helvetica-Bold'),
        ('FONTSIZE', (0,0), (-1,-1), 10),
        ('LINEBELOW', (1,0), (1,0), 0.5, colors.gray),
        ('LINEBELOW', (3,0), (3,0), 0.5, colors.gray),
        ('LINEBELOW', (1,1), (1,1), 0.5, colors.gray),
        ('LINEBELOW', (3,1), (3,1), 0.5, colors.gray),
        ('LINEBELOW', (1,2), (1,2), 0.5, colors.gray),
        ('LINEBELOW', (3,2), (3,2), 0.5, colors.gray),
        ('BOTTOMPADDING', (0,0), (-1,-1), 10),
        ('TOPPADDING', (0,0), (-1,-1), 10),
    ]))
    elements.append(info_table)
    elements.append(Spacer(1, 10))

    # Marks Table
    def get_perf_level(score):
        if score >= 80: return "Exceeding Expectations"
        if score >= 65: return "Meeting Expectations"
        if score >= 50: return "Approaching Expectations"
        return "Below Expectations"

    subjects = [
        ("Math", m), ("Eng", e), ("Kis", k), ("Int Sci", isc),
        ("Agri", agr), ("SST", ss), ("CRE", cr), ("CIA", ci), ("Pre-Tech", pt)
    ]
    
    marks_header = ["LEARNING AREA", "MARKS", "AVG", "PERFORMANCE LEVEL"]
    marks_data = [marks_header]
    for subj, score in subjects:
        marks_data.append([subj, score, score, get_perf_level(score)])
    
    # Summary Rows
    total_row = ["Total Scores", f"{total}/900", f"{avg}/100", f"Termly Performance Level: {lvl}"]
    avg_row = ["Average Scores", f"{avg}/100", "", f"Position: {rank} of {len([x for x in students if x[2]==cls and x[3]==term])}"]
    marks_data.append(total_row)
    marks_data.append(avg_row)

    marks_table = Table(marks_data, colWidths=[150, 100, 100, 170])
    marks_style = TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2EFDA")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#2159A6")),
        ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
        ('ALIGN', (0,0), (-1,-1), 'LEFT'),
        ('GRID', (0,0), (-1,-3), 0.5, colors.gray),
        ('FONTSIZE', (0,0), (-1,-1), 9),
        ('BOTTOMPADDING', (0,0), (-1,-1), 8),
        ('TOPPADDING', (0,0), (-1,-1), 8),
        # Summary row styling
        ('BACKGROUND', (0,-2), (-1,-1), colors.HexColor("#D0F0E0")),
        ('FONTNAME', (0,-2), (-1,-1), 'Helvetica-Bold'),
        ('SPAN', (2,-1), (2,-1)), # Optional spans
    ])
    marks_table.setStyle(marks_style)
    elements.append(marks_table)
    elements.append(Spacer(1, 10))

    # Performance Trend Chart
    elements.append(Paragraph("<u>Performance Trend</u>", styles['SchoolInfo']))
    elements.append(Spacer(1, 5))
    
    drawing = Drawing(400, 150)
    bc = VerticalBarChart()
    bc.x = 50
    bc.y = 20
    bc.height = 100
    bc.width = 350
    bc.data = [[m, e, k, isc, agr, ss, cr, ci, pt]]
    bc.strokeColor = colors.black
    bc.valueAxis.valueMin = 0
    bc.valueAxis.valueMax = 100
    bc.valueAxis.valueStep = 25
    bc.categoryAxis.labels.boxAnchor = 'ne'
    bc.categoryAxis.labels.dx = 0
    bc.categoryAxis.labels.dy = -2
    bc.categoryAxis.labels.angle = 30
    bc.categoryAxis.categoryNames = [s[0] for s in subjects]
    
    # Bar colors based on performance
    for i, (_, score) in enumerate(subjects):
        if score >= 80: bc.bars[0, i].fillColor = colors.HexColor("#2159A6")
        elif score >= 50: bc.bars[0, i].fillColor = colors.HexColor("#4CAF50")
        else: bc.bars[0, i].fillColor = colors.HexColor("#D9534F")

    drawing.add(bc)
    elements.append(drawing)
    elements.append(Spacer(1, 20))

    # Comments Section
    comments_data = [
        [Paragraph("Class Teacher's Comment: <font color='gray'><i>Fair performance. Keep working harder.</i></font>", styles['CommentTitle'])],
        [Spacer(1, 10)],
        [Paragraph("Head Teacher's Comment: <font color='gray'><i>Encourage the learner to improve.</i></font>", styles['CommentTitle'])]
    ]
    comments_table = Table(comments_data, colWidths=[520])
    comments_table.setStyle(TableStyle([
        ('BOX', (0,0), (-1,0), 0.5, colors.gray),
        ('BOX', (0,2), (-1,2), 0.5, colors.gray),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('TOPPADDING', (0,0), (-1,-1), 5),
    ]))
    elements.append(comments_table)
    elements.append(Spacer(1, 10))

    # Footer
    footer_text = f"This term closed on: 3/11/2026 | Next term opens on: ___________"
    elements.append(Paragraph(footer_text, styles['SchoolInfo']))
    elements.append(Spacer(1, 5))
    elements.append(Paragraph("<font size=7 color='gray' backColor='#EEEEEE'>This Exam Report Card has Been Issued Without Any Alterations Whatsoever. Any Alterations Will Invalidate Its Authenticity.</font>", styles['SchoolInfo']))

    # Border - using Page Canvas for border
    def add_border(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#2159A6"))
        canvas.setLineWidth(2)
        canvas.rect(10, 10, A4[0]-20, A4[1]-20)
        canvas.restoreState()

    doc.build(elements, onFirstPage=add_border, onLaterPages=add_border)
    messagebox.showinfo("Report Generated", f"Report card saved as {filename}")

# ─── UI Layout ───────────────────────────────────────────────────────────────
app = ctk.CTk()
app.title("MOAS School Management System")
app.geometry("1200x700")
app.configure(fg_color=PRIMARY_BG)

if LOGO_IMAGE:
    try:
        app.iconphoto(False, LOGO_IMAGE)
    except Exception:
        pass

# --- Login Screen ---
login_frame = ctk.CTkFrame(app, fg_color=PRIMARY_BG)
login_frame.pack(fill="both", expand=True)

# Login content centered
login_content = ctk.CTkFrame(login_frame, fg_color="transparent")
login_content.place(relx=0.5, rely=0.5, anchor="center")

# Logo at top of login
if LOGO_IMAGE:
    login_logo = ctk.CTkLabel(login_content, image=LOGO_IMAGE, text="")
    login_logo.pack(pady=(0, 20))

ctk.CTkLabel(login_content, text="MT OLIVES", font=("Arial", 28, "bold"), 
             text_color=PRIMARY_GREEN).pack()
ctk.CTkLabel(login_content, text="Adventist School", font=("Arial", 14), 
             text_color=TEXT_SECONDARY).pack(pady=(0, 30))

ctk.CTkLabel(login_content, text="MOAS School Management System", font=("Arial", 16, "bold"), 
             text_color=LABEL_COLOR).pack(pady=(0, 30))

# Login button - hide main app frames initially
def show_main_app():
    login_frame.pack_forget()
    # Show sidebar and content
    sidebar.pack(side="left", fill="y")
    content.pack(side="right", fill="both", expand=True, padx=20, pady=20)

ctk.CTkButton(login_content, text="Enter System", command=show_main_app,
              fg_color=PRIMARY_GREEN, hover_color=PRIMARY_GREEN_DARK,
              text_color="#ffffff", height=50, width=200,
              font=("Arial", 16, "bold"), corner_radius=12).pack()

# --- Main App Layout (hidden initially) ---
# Sidebar and content will be shown after login
sidebar = ctk.CTkFrame(app, width=260, corner_radius=0, fg_color=SIDEBAR_BG)

# Sidebar Header with Logo Area
sidebar_header = ctk.CTkFrame(sidebar, fg_color="transparent")
sidebar_header.pack(pady=(20, 10), fill="x")

if LOGO_IMAGE:
    logo_label = ctk.CTkLabel(sidebar_header, image=LOGO_IMAGE, text="")
    logo_label.pack(pady=(0, 8))

# School Name - styled prominently
ctk.CTkLabel(sidebar_header, text="MT OLIVES", font=("Arial", 22, "bold"), 
             text_color="#FFFFFF", fg_color="transparent").pack(pady=(5, 0))
ctk.CTkLabel(sidebar_header, text="Adventist School", font=("Arial", 13, "bold"), 
             text_color=SIDEBAR_TEXT_MUTED, fg_color="transparent").pack(pady=(2, 0))

# Divider
ctk.CTkFrame(sidebar, height=2, fg_color=SIDEBAR_HOVER).pack(fill="x", padx=20, pady=(10, 20))

# Navigation section label
ctk.CTkLabel(sidebar, text="MAIN MENU", font=("Arial", 10, "bold"), 
             text_color=SIDEBAR_TEXT_MUTED, fg_color="transparent").pack(anchor="w", padx=24, pady=(0, 8))

nav_btns = []

def add_nav_button(text, command, icon=""):
    btn = ctk.CTkButton(
        sidebar,
        text=f"  {icon} {text}" if icon else f"  {text}",
        command=command,
        width=220,
        anchor="w",
        fg_color="transparent",
        hover_color=SIDEBAR_HOVER,
        text_color=SIDEBAR_TEXT,
        border_width=0,
        corner_radius=8,
        height=42,
        font=("Arial", 13, "bold"),
    )
    btn.pack(pady=3, padx=16, fill="x")
    nav_btns.append(btn)
    return btn

# Pages
content = ctk.CTkFrame(app, fg_color=PRIMARY_BG)
content.grid_rowconfigure(0, weight=1)
content.grid_columnconfigure(0, weight=1)

dashboard_frame = ctk.CTkFrame(content, fg_color=PRIMARY_BG)
students_frame = ctk.CTkFrame(content, fg_color=PRIMARY_BG)
add_student_frame = ctk.CTkFrame(content, fg_color=PRIMARY_BG)
marks_frame = ctk.CTkFrame(content, fg_color=PRIMARY_BG)
reports_frame = ctk.CTkFrame(content, fg_color=PRIMARY_BG)

for frame in (dashboard_frame, students_frame, add_student_frame, marks_frame, reports_frame):
    frame.grid(row=0, column=0, sticky="nsew")


def show_frame(frame):
    frame.tkraise()
    if frame is dashboard_frame:
        refresh_dashboard()
    elif frame is marks_frame:
        refresh_marks_table()


# --- Dashboard ---
stats_vars = {
    "total_students": tk.StringVar(value="0"),
    "class_avg": tk.StringVar(value="-"),
    "top_student": tk.StringVar(value="-"),
    "subjects": tk.StringVar(value="5"),
}

def make_stat_card(parent, title, value_var, icon_text=None, accent_color=STAT_TOTAL):
    card = ctk.CTkFrame(parent, corner_radius=16, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
    card.grid_propagate(False)
    card.grid_rowconfigure(0, weight=1)
    card.grid_columnconfigure(0, weight=1)
    card.grid_columnconfigure(1, weight=0)

    # Colored accent strip at top
    accent_strip = ctk.CTkFrame(card, height=4, fg_color=accent_color, corner_radius=0)
    accent_strip.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0, pady=(0, 0))

    card_title = ctk.CTkLabel(card, text=title, font=("Arial", 12, "bold"), text_color="#64748B")
    card_title.grid(row=0, column=0, sticky="w", padx=(16, 8), pady=(20, 0))

    card_value = ctk.CTkLabel(card, textvariable=value_var, font=("Arial", 32, "bold"), text_color=accent_color)
    card_value.grid(row=1, column=0, sticky="w", padx=(16, 8), pady=(0, 12))

    if icon_text:
        icon = ctk.CTkLabel(card, text=icon_text, font=("Arial", 24))
        icon.grid(row=0, column=1, rowspan=2, sticky="e", padx=(0, 16), pady=16)

    return card


def refresh_dashboard():
    total = len(students)
    stats_vars["total_students"].set(str(total))

    if total > 0:
        avg = round(sum(s[10] for s in students) / total, 2)
        stats_vars["class_avg"].set(f"{avg}")

        top = max(students, key=lambda s: s[10])
        stats_vars["top_student"].set(top[1])
    else:
        stats_vars["class_avg"].set("-")
        stats_vars["top_student"].set("-")

    stats_vars["subjects"].set("5")

    # Keep table in sync
    refresh_table()

# Dashboard layout
# Header section with welcome message and logo avatar
dashboard_header = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
dashboard_header.pack(fill="x")

# Left side - Welcome text
welcome_text = ctk.CTkFrame(dashboard_header, fg_color="transparent")
welcome_text.pack(side="left", fill="both", expand=True)

ctk.CTkLabel(welcome_text, text="Welcome Back!", font=("Arial", 32, "bold"), 
             text_color=PRIMARY_GREEN, anchor="w").pack(anchor="w")
ctk.CTkLabel(welcome_text, text="School Report Management System overview", 
             font=("Arial", 14), text_color=TEXT_SECONDARY, anchor="w").pack(anchor="w", pady=(4, 0))

# Right side - Logo Avatar
if LOGO_SMALL:
    avatar_frame = ctk.CTkFrame(dashboard_header, fg_color="transparent")
    avatar_frame.pack(side="right", padx=(0, 10))
    
    avatar_container = ctk.CTkFrame(avatar_frame, corner_radius=35, fg_color=CARD_BG, 
                                     border_width=2, border_color=PRIMARY_GREEN)
    avatar_container.pack(pady=5)
    
    avatar_label = ctk.CTkLabel(avatar_container, image=LOGO_SMALL, text="")
    avatar_label.pack(padx=5, pady=5)

cards_container = ctk.CTkFrame(dashboard_frame, fg_color="transparent")
cards_container.pack(fill="x", pady=(0, 20))

cards_container.grid_columnconfigure((0, 1, 2, 3), weight=1, uniform="cards")

make_stat_card(cards_container, "Total Students", stats_vars["total_students"], icon_text="👥", accent_color=STAT_TOTAL).grid(row=0, column=0, padx=8, ipadx=8, ipady=8, sticky="nsew")
make_stat_card(cards_container, "Class Average", stats_vars["class_avg"], icon_text="📈", accent_color=STAT_AVG).grid(row=0, column=1, padx=8, ipadx=8, ipady=8, sticky="nsew")
make_stat_card(cards_container, "Top Student", stats_vars["top_student"], icon_text="🏆", accent_color=STAT_TOP).grid(row=0, column=2, padx=8, ipadx=8, ipady=8, sticky="nsew")
make_stat_card(cards_container, "Subjects", stats_vars["subjects"], icon_text="📚", accent_color=STAT_SUBJECTS).grid(row=0, column=3, padx=8, ipadx=8, ipady=8, sticky="nsew")

# Grading Scale
grade_frame = ctk.CTkFrame(dashboard_frame, corner_radius=16, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
grade_frame.pack(fill="x", pady=(20, 0))

ctk.CTkLabel(grade_frame, text="CBC Grading Scale", font=("Arial", 18, "bold"), anchor="w", text_color=LABEL_COLOR).pack(anchor="w", padx=20, pady=(20, 12))

scale_row = ctk.CTkFrame(grade_frame, fg_color="transparent")
scale_row.pack(fill="x", padx=16, pady=(0, 20))

GRADE_STYLES = [
    ("EE", "80-100", "Exceeding Expectations", GRADE_EE),
    ("ME", "65-79", "Meeting Expectations", GRADE_ME),
    ("AE", "50-64", "Approaching Expectations", GRADE_AE),
    ("BE", "0-49", "Below Expectations", GRADE_BE),
    ("IE", "0-49", "Inadequate", GRADE_IE),
]

for i, (grade, rng, desc, color) in enumerate(GRADE_STYLES):
    box = ctk.CTkFrame(scale_row, corner_radius=14, fg_color=color)
    box.grid(row=0, column=i, padx=8, sticky="nsew")
    scale_row.grid_columnconfigure(i, weight=1)
    ctk.CTkLabel(box, text=grade, font=("Arial", 18, "bold"), text_color="#ffffff").pack(pady=(12, 4))
    ctk.CTkLabel(box, text=rng, font=("Arial", 10), text_color="#ffffff").pack()
    ctk.CTkLabel(box, text=desc, font=("Arial", 9), text_color="#ffffff").pack(pady=(4, 12))

# --- Students Page ---
header_bar = ctk.CTkFrame(students_frame, fg_color="transparent")
header_bar.pack(fill="x", pady=(0, 16))

left_header = ctk.CTkFrame(header_bar, fg_color="transparent")
left_header.pack(side="left", fill="both", expand=True)

ctk.CTkLabel(left_header, text="Students", font=("Arial", 28, "bold"), text_color=LABEL_COLOR, anchor="w").pack(anchor="w")
ctk.CTkLabel(left_header, text="Manage student registrations", font=("Arial", 12), text_color=TEXT_SECONDARY, anchor="w").pack(anchor="w")

def show_add_student_page():
    init_add_student_page()
    show_frame(add_student_frame)

add_student_btn = ctk.CTkButton(header_bar, text="+ Add Student", command=show_add_student_page, 
                                fg_color=PRIMARY_GREEN, hover_color=PRIMARY_GREEN_DARK, 
                                text_color="#ffffff", corner_radius=12, font=("Arial", 12, "bold"))
add_student_btn.pack(side="right", pady=4, padx=(0, 8))

# Template button
ctk.CTkButton(header_bar, text="📥 Template", command=download_template,
              fg_color=PRIMARY_BLUE, hover_color="#0D47A1",
              text_color="#ffffff", corner_radius=12, font=("Arial", 12)).pack(side="right", pady=4, padx=(0, 8))

# Export button
ctk.CTkButton(header_bar, text="📤 Export", command=export_student_list,
              fg_color="#FF9800", hover_color="#F57C00",
              text_color="#ffffff", corner_radius=12, font=("Arial", 12)).pack(side="right", pady=4)

# Table section
table_card = ctk.CTkFrame(students_frame, corner_radius=16, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
table_card.pack(fill="both", expand=True)

table_card.grid_rowconfigure(0, weight=1)
table_card.grid_columnconfigure(0, weight=1)

# Treeview
tbl_frame = tk.Frame(table_card, bg=TABLE_BG if 'TABLE_BG' in globals() else CARD_BG)
tbl_frame.grid(row=0, column=0, sticky="nsew", padx=8, pady=12)

cols = ("Adm No", "Name", "Class", "Gender", "Actions")
table = ttk.Treeview(tbl_frame, columns=cols, show="headings")

for col in cols:
    table.heading(col, text=col)
    table.column(col, width=100, anchor="center")

table.column("Name", width=180)

def _on_select(event):
    # Optional: future hooks for enabling edit/delete buttons
    pass

table.bind("<<TreeviewSelect>>", _on_select)

sb = ttk.Scrollbar(tbl_frame, orient="vertical", command=table.yview)
table.configure(yscrollcommand=sb.set)
table.pack(side="left", fill="both", expand=True)
sb.pack(side="right", fill="y")

# Actions below table
actions_bar = ctk.CTkFrame(table_card, fg_color="transparent")
actions_bar.grid(row=1, column=0, sticky="ew", padx=16, pady=(0, 16))

ctk.CTkButton(actions_bar, text="Delete Selected", command=delete_selected_student, 
                    fg_color="#E53935", hover_color="#C62828", text_color="#ffffff").pack(side="left")

# --- Marks Entry Page ---
marks_header = ctk.CTkFrame(marks_frame, fg_color="transparent")
marks_header.pack(fill="x", pady=(0, 16))

left_marks_header = ctk.CTkFrame(marks_header, fg_color="transparent")
left_marks_header.pack(side="left", fill="both", expand=True)

ctk.CTkLabel(left_marks_header, text="Enter Marks", font=("Arial", 28, "bold"), text_color=LABEL_COLOR, anchor="w").pack(anchor="w")
ctk.CTkLabel(left_marks_header, text="Enter subject marks for each student", font=("Arial", 12), text_color=TEXT_SECONDARY, anchor="w").pack(anchor="w")

marks_controls = ctk.CTkFrame(marks_header, fg_color="transparent")
marks_controls.pack(side="right", pady=4)

grade_filter_var = tk.StringVar(value="Grade 1")
ctk.CTkComboBox(marks_controls, values=["Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5", "Grade 6"], variable=grade_filter_var).pack(side="left", padx=(0, 12))
ctk.CTkButton(marks_controls, text="💾 Save All", command=lambda: save_marks_for_grade(grade_filter_var.get()), 
              fg_color=PRIMARY_GREEN, hover_color=PRIMARY_GREEN_DARK, 
              text_color="#ffffff", corner_radius=12).pack(side="left")

# Marks table container
marks_table_card = ctk.CTkFrame(marks_frame, corner_radius=16, fg_color=CARD_BG, border_width=1, border_color=CARD_BORDER)
marks_table_card.pack(fill="both", expand=True)

marks_table_card.grid_rowconfigure(0, weight=1)
marks_table_card.grid_columnconfigure(0, weight=1)

marks_scroll = ctk.CTkScrollableFrame(marks_table_card)
marks_scroll.grid(row=0, column=0, sticky="nsew", padx=16, pady=12)

marks_rows_container = ctk.CTkFrame(marks_scroll, fg_color="transparent")
marks_rows_container.pack(fill="both", expand=True)

marks_row_widgets = []


def refresh_marks_table():
    for w in marks_row_widgets:
        if hasattr(w, 'destroy'):
            w.destroy()
    marks_row_widgets.clear()

    grade = grade_filter_var.get()
    filtered = [s for s in students if s[2] == grade]

    if not filtered:
        empty_label = ctk.CTkLabel(marks_rows_container, text=f"No students in {grade}", font=("Arial", 12), text_color="#5b5b5b")
        empty_label.pack(pady=32)
        marks_row_widgets.append(empty_label)
        return

    header = ctk.CTkFrame(marks_rows_container, fg_color="transparent")
    header.pack(fill="x", pady=(0, 8))
    columns = ["Student", "Math", "Eng", "Kis", "Int Sci", "Agri", "SST", "CRE", "CIA", "Pre-Tech"]
    widths = [180, 70, 70, 70, 70, 70, 70, 70, 70, 80]

    for c, w in zip(columns, widths):
        lbl = ctk.CTkLabel(header, text=c, font=("Arial", 10, "bold"), text_color="#4a4a4a")
        lbl.pack(side="left", padx=(0, 4), ipady=6)

    marks_row_widgets.append(header)

    for s in filtered:
        row = ctk.CTkFrame(marks_rows_container, fg_color="transparent")
        row.pack(fill="x", pady=2)

        # make sure we have enough slots in student record for the extra subjects
        while len(s) < 19:
            s.append(0)

        name_lbl = ctk.CTkLabel(row, text=s[1], text_color="#2c2c2c")
        name_lbl.pack(side="left", padx=(0, 4), ipady=6, ipadx=4, fill="x", expand=True)

        # Subjects from index 4 to 12
        entry_values = [s[4], s[5], s[6], s[7], s[8], s[9], s[10], s[11], s[12]]
        entry_vars = []
        for val in entry_values:
            v = tk.StringVar(value=str(val))
            entry = ctk.CTkEntry(row, width=60, textvariable=v)
            entry.pack(side="left", padx=(0, 4))
            entry_vars.append(v)

        marks_row_widgets.append(row)

        # store for saving later
        row._entry_vars = entry_vars
        row._student = s


def save_marks_for_grade(grade: str):
    # loop through row widgets and save into student records
    for w in marks_row_widgets:
        if not w or not hasattr(w, "_student"):
            continue
        s = w._student
        vars = w._entry_vars
        try:
            # Update subjects 4-12
            for i in range(9):
                s[4+i] = int(vars[i].get())
            
            # Recalculate total, avg, level
            s[13] = sum(s[4:13])
            s[14] = round(s[13] / 9, 2)
            s[15] = get_level(s[14])
        except ValueError:
            # ignore invalid input
            pass

    calculate_rank()
    refresh_table()
    save_to_db()

# --- Reports Page ---
SUBJECT_SCORES = [
    ("Math", 4),
    ("Eng", 5),
    ("Kis", 6),
    ("Int Sci", 7),
    ("Agri", 8),
    ("SST", 9),
    ("CRE", 10),
    ("CIA", 11),
    ("Pre-Tech", 12),
]

report_filter_var = tk.StringVar(value="All")


def get_subject_value(student, idx):
    return student[idx] if len(student) > idx else 0


def calculate_report_summary(filtered_students):
    summary = {}
    total_students = len(filtered_students)
    for subject, idx in SUBJECT_SCORES:
        if total_students == 0:
            summary[subject] = 0
        else:
            summary[subject] = round(sum(get_subject_value(s, idx) for s in filtered_students) / total_students, 2)
    return summary


def get_grade_from_score(score: float) -> str:
    # Reuse the existing get_level logic, but convert to level code
    return get_level(score)


def refresh_reports():
    for child in reports_frame.winfo_children():
        child.destroy()

    reports_frame.configure(fg_color="#f9fafb")

    header_row = ctk.CTkFrame(reports_frame, fg_color="transparent")
    header_row.pack(fill="x", padx=24, pady=(20, 24))

    header_left = ctk.CTkFrame(header_row, fg_color="transparent")
    header_left.pack(side="left", fill="both", expand=True)
    ctk.CTkLabel(header_left, text="Reports & Rankings", font=("Arial", 32, "bold"), text_color="#111827", anchor="w").pack(anchor="w")
    ctk.CTkLabel(header_left, text="View student performance and rankings", font=("Arial", 16), text_color="#6b7280", anchor="w").pack(anchor="w")

    header_right = ctk.CTkFrame(header_row, fg_color="transparent")
    header_right.pack(side="right", pady=10)

    grade_dropdown = ctk.CTkComboBox(
        header_right, 
        values=["All", "Grade 1", "Grade 2", "Grade 3", "Grade 4", "Grade 5", "Grade 6", "Grade 7"], 
        variable=report_filter_var, 
        width=160,
        height=45,
        fg_color="#ffffff",
        border_color="#d1d5db",
        text_color="#111827",
        button_color="#ffffff",
        button_hover_color="#f3f4f6",
        dropdown_fg_color="#ffffff",
        corner_radius=10
    )
    grade_dropdown.pack(side="left", padx=(0, 16))
    grade_dropdown.bind("<<ComboboxSelected>>", lambda e: refresh_reports())

    def export_csv():
        filename = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if not filename:
            return
        filtered = [s for s in students if report_filter_var.get() in ("All", s[2])]
        with open(filename, "w", encoding="utf-8") as f:
            headers = ["Pos", "Name", "Class"] + [s[0] for s in SUBJECT_SCORES] + ["Total", "Avg", "Grade"]
            f.write(",".join(headers) + "\n")
            for pos, s in enumerate(sorted(filtered, key=lambda x: x[12], reverse=False), start=1):
                values = [str(pos), s[1], s[2]]
                for _, idx in SUBJECT_SCORES:
                    values.append(str(get_subject_value(s, idx)))
                values.append(str(s[9]))
                values.append(str(s[10]))
                values.append(get_grade_from_score(s[10]))
                f.write(",".join(values) + "\n")
        messagebox.showinfo("Export Complete", f"Report exported to {filename}")

    ctk.CTkButton(
        header_right, 
        text="📥 Export CSV", 
        command=export_csv, 
        fg_color="#ffffff", 
        hover_color="#f3f4f6", 
        text_color="#374151",
        border_width=1,
        border_color="#d1d5db",
        height=45,
        width=140,
        font=("Arial", 14, "bold"),
        corner_radius=10
    ).pack(side="left")

    # Subject performance card
    subject_card = ctk.CTkFrame(reports_frame, corner_radius=16, fg_color="#ffffff", border_width=1, border_color="#e5e7eb")
    subject_card.pack(fill="x", padx=24, pady=(0, 24))

    ctk.CTkLabel(subject_card, text="Subject Performance", font=("Arial", 18, "bold"), text_color="#111827", anchor="w").pack(anchor="w", padx=24, pady=(24, 20))

    card_row = ctk.CTkFrame(subject_card, fg_color="transparent")
    card_row.pack(fill="x", padx=16, pady=(0, 24))

    filtered_students = [s for s in students if report_filter_var.get() in ("All", s[2])]
    summary = calculate_report_summary(filtered_students)

    for subj, _ in SUBJECT_SCORES:
        score = summary.get(subj, 0)
        lvl = get_grade_from_score(score)

        card = ctk.CTkFrame(card_row, corner_radius=12, fg_color="#ffffff", border_width=1, border_color="#e5e7eb", width=100, height=130)
        card.pack(side="left", padx=8, fill="y", expand=True)
        card.pack_propagate(False)

        ctk.CTkLabel(card, text=subj, font=("Arial", 12), text_color="#6b7280").pack(pady=(16, 4))
        ctk.CTkLabel(card, text=str(int(score)), font=("Arial", 28, "bold"), text_color="#111827").pack()
        
        # Badge color based on grade
        badge_color = "#ef4444" if lvl in ("BE", "IE") else "#10b981"
        badge = ctk.CTkLabel(card, text=lvl, font=("Arial", 10, "bold"), fg_color=badge_color, text_color="#ffffff", corner_radius=12, width=32, height=22)
        badge.pack(pady=(8, 16))

    # Rankings table card
    ranking_card = ctk.CTkFrame(reports_frame, corner_radius=16, fg_color="#ffffff", border_width=1, border_color="#e5e7eb")
    ranking_card.pack(fill="both", expand=True, padx=24, pady=(0, 24))

    ctk.CTkLabel(ranking_card, text="Student Rankings", font=("Arial", 18, "bold"), text_color="#111827", anchor="w").pack(anchor="w", padx=24, pady=(24, 20))

    ranking_table_frame = tk.Frame(ranking_card, bg="#ffffff")
    ranking_table_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))

    # Styling Treeview to match clean header look
    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview", background="#ffffff", fieldbackground="#ffffff", foreground="#374151", rowheight=45, font=("Arial", 11), borderwidth=0)
    style.configure("Treeview.Heading", background="#f9fafb", foreground="#6b7280", font=("Arial", 11, "bold"), borderwidth=0)
    style.map("Treeview", background=[('selected', '#eff6ff')], foreground=[('selected', '#1e40af')])

    columns = ["Pos", "Name", "Class"] + [s[0] for s in SUBJECT_SCORES] + ["Total", "Avg", "Grade"]
    ranking_table = ttk.Treeview(ranking_table_frame, columns=columns, show="headings", height=10)
    
    for col in columns:
        ranking_table.heading(col, text=col)
        # Adjust width for subjects
        w = 140 if col == "Name" else 60
        if col in ("Pos", "Total", "Avg", "Grade"): w = 70
        ranking_table.column(col, width=w, anchor="center")

    sb2 = ttk.Scrollbar(ranking_table_frame, orient="vertical", command=ranking_table.yview)
    ranking_table.configure(yscrollcommand=sb2.set)
    ranking_table.pack(side="left", fill="both", expand=True)
    sb2.pack(side="right", fill="y")

    if not filtered_students:
        ranking_table.insert("", "end", values=["", "No results to display. Add students and enter marks first."] + ["" for _ in range(len(columns)-2)])
    else:
        # Sort and populate rankings table
        sorted_students = sorted(filtered_students, key=lambda x: x[9], reverse=True)
        for i, s in enumerate(sorted_students, 1):
            row_data = [i, s[1], s[2]]
            for _, idx in SUBJECT_SCORES:
                row_data.append(get_subject_value(s, idx))
            row_data.extend([s[9], s[10], get_grade_from_score(s[10])])
            ranking_table.insert("", "end", values=row_data)

    if not filtered_students:
        ranking_table.insert("", "end", values=["", "No results to display. Add students and enter marks first."] + ["" for _ in range(len(columns) - 2)])
    else:
        sorted_students = sorted(filtered_students, key=lambda s: s[12] if len(s) > 12 else 0)
        for pos, s in enumerate(sorted_students, start=1):
            row = [pos, s[1], s[2]]
            for _, idx in SUBJECT_SCORES:
                row.append(get_subject_value(s, idx))
            row += [s[9], s[10], get_grade_from_score(s[10])]
            ranking_table.insert("", "end", values=row)


# --- Navigation ---
add_nav_button("Dashboard", lambda: show_frame(dashboard_frame), "🏠")
add_nav_button("Students", lambda: show_frame(students_frame), "👥")
add_nav_button("Enter Marks", lambda: show_frame(marks_frame), "📝")
add_nav_button("Reports", lambda: show_frame(reports_frame), "📊")

# Sidebar Footer
sidebar_footer = ctk.CTkFrame(sidebar, fg_color="transparent")
sidebar_footer.pack(side="bottom", fill="x", pady=(20, 15))

# Divider
ctk.CTkFrame(sidebar, height=1, fg_color=SIDEBAR_HOVER).pack(fill="x", padx=16, pady=(0, 15))

ctk.CTkLabel(sidebar_footer, text="MOAS System v1.0", font=("Arial", 10), 
             text_color=SIDEBAR_TEXT_MUTED, fg_color="transparent").pack(pady=(0, 4))
ctk.CTkLabel(sidebar_footer, text="Mt Olives Adventist School", font=("Arial", 9), 
             text_color=SIDEBAR_TEXT_MUTED, fg_color="transparent").pack()

# Show initial page
show_frame(dashboard_frame)

# Run the application with graceful exit handling
try:
    app.mainloop()
except KeyboardInterrupt:
    print("\nApplication closed by user.")
    try:
        app.destroy()
    except:
        pass
