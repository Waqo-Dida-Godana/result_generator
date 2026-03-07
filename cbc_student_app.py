import tkinter as tk
from tkinter import ttk, messagebox
import pandas as pd
from reportlab.pdfgen import canvas
from fpdf import FPDF
import os

# ─── Colours & Font constants ────────────────────────────────────────────────
BG_MAIN      = "#f0f4f8"
BG_HEADER    = "#1a3c5e"
BG_FRAME     = "#ffffff"
BG_BTN_ADD   = "#27ae60"
BG_BTN_DEL   = "#e74c3c"
BG_BTN_CLR   = "#7f8c8d"
FG_WHITE     = "#ffffff"
FG_DARK      = "#2c3e50"
FG_STATUS_OK = "#27ae60"
FG_STATUS_ERR= "#e74c3c"

FONT_TITLE   = ("Arial", 18, "bold")
FONT_SUB     = ("Arial", 10, "italic")
FONT_LABEL   = ("Arial", 10, "bold")
FONT_ENTRY   = ("Arial", 10)
FONT_BTN     = ("Arial", 10, "bold")
FONT_STATUS  = ("Arial", 9, "italic")

LEVEL_COLORS = {
    "EE": "#27ae60",   # green
    "ME": "#2980b9",   # blue
    "AE": "#e67e22",   # orange
    "BE": "#e74c3c",   # red
}

# ─── MySQL Connection (graceful, optional dependency) ────────────────────────
db     = None
cursor = None
DB_CONNECTED = False

def connect_db():
    global db, cursor, DB_CONNECTED
    try:
        import mysql.connector
        from mysql.connector import Error as _Error  # noqa: F401
        db = mysql.connector.connect(
            host="localhost",
            user="root",
            password="",
            database="cbc_student_db",
            connection_timeout=3
        )
        cursor = db.cursor()
        DB_CONNECTED = True
    except Exception:
        DB_CONNECTED = False
        db = None
        cursor = None

connect_db()

# ─── In-memory student list ───────────────────────────────────────────────────
students = []   # each item: [name, math, eng, kisw, sci, soc, total, avg, level, rank]

# ─── CBC logic ───────────────────────────────────────────────────────────────
def get_level(avg: float) -> str:
    if avg >= 80:
        return "EE"
    elif avg >= 65:
        return "ME"
    elif avg >= 50:
        return "AE"
    return "BE"

# ─── Input validation ────────────────────────────────────────────────────────
def validate_mark(value: str, field: str) -> int:
    """Return int mark or raise ValueError with a friendly message."""
    try:
        mark = int(value)
    except ValueError:
        raise ValueError(f"'{field}' must be a whole number.")
    if not (0 <= mark <= 100):
        raise ValueError(f"'{field}' must be between 0 and 100.")
    return mark

# ─── Core actions ────────────────────────────────────────────────────────────
def generate_report(name, math, eng, kisw, sci, soc, total, avg, level):
    file = f"{name}_report.pdf"
    pdf = canvas.Canvas(file)
    pdf.drawString(200,800,"STUDENT REPORT CARD")
    pdf.drawString(100,750,"Name: " + name)
    pdf.drawString(100,700,"Math: " + str(math))
    pdf.drawString(100,680,"English: " + str(eng))
    pdf.drawString(100,660,"Kiswahili: " + str(kisw))
    pdf.drawString(100,640,"Science: " + str(sci))
    pdf.drawString(100,620,"Social: " + str(soc))
    pdf.drawString(100,580,"Total: " + str(total))
    pdf.drawString(100,560,"Average: " + str(avg))
    pdf.drawString(100,540,"Level: " + level)
    pdf.save()
    set_status(f"✔ Report card generated for '{name}'")

def generate_selected_report():
    selected = table.selection()
    if not selected:
        set_status("Select a student to generate report.", error=True)
        return
    item = table.item(selected[0])
    vals = item["values"]
    generate_report(vals[0], vals[1], vals[2], vals[3], vals[4], vals[5], vals[6], vals[7], vals[8])

def export_excel():
    if not students:
        set_status("No data to export.", error=True)
        return
    data = []
    for s in students:
        data.append({
            "Name": s[0], "Math": s[1], "English": s[2], "Kiswahili": s[3],
            "Science": s[4], "Social": s[5], "Total": s[6], "Average": s[7],
            "Level": s[8], "Rank": s[9]
        })
    df = pd.DataFrame(data)
    df.to_excel("cbc_results.xlsx", index=False)
    set_status("✔ Results exported to 'cbc_results.xlsx'")

def export_pdf():
    if not students:
        set_status("No data to export.", error=True)
        return
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, "CBC Student Results", ln=True, align="C")
    pdf.ln(10)
    
    # Headers
    header = ["Name", "Math", "Eng", "Kisw", "Sci", "Soc", "Tot", "Avg", "Lvl", "Rnk"]
    col_widths = [40, 15, 15, 15, 15, 15, 15, 15, 15, 15]
    for i, h in enumerate(header):
        pdf.cell(col_widths[i], 10, h, border=1)
    pdf.ln()
    
    # Data
    for s in students:
        for i, val in enumerate(s):
            pdf.cell(col_widths[i], 10, str(val), border=1)
        pdf.ln()
    
    pdf.output("cbc_results.pdf")
    set_status("✔ Results exported to 'cbc_results.pdf'")

def add_student():
    name = name_entry.get().strip()
    if not name:
        set_status("Name cannot be empty.", error=True)
        return

    try:
        math      = validate_mark(math_entry.get().strip(),      "Math")
        english   = validate_mark(english_entry.get().strip(),   "English")
        kiswahili = validate_mark(kiswahili_entry.get().strip(), "Kiswahili")
        science   = validate_mark(science_entry.get().strip(),   "Science")
        social    = validate_mark(social_entry.get().strip(),    "Social")
    except ValueError as e:
        set_status(str(e), error=True)
        return

    total = math + english + kiswahili + science + social
    avg   = round(total / 5, 2)
    level = get_level(avg)

    students.append([name, math, english, kiswahili, science, social, total, avg, level, 0])
    calculate_rank()
    refresh_table()
    save_to_db()
    clear_entries()
    set_status(f"✔ '{name}' added successfully  |  Total: {total}  |  Avg: {avg}  |  Level: {level}")


def delete_selected():
    global students
    selected = table.selection()
    if not selected:
        set_status("Select a row to delete.", error=True)
        return

    item     = table.item(selected[0])
    row_name = item["values"][0]

    students = [s for s in students if s[0] != row_name]
    calculate_rank()
    refresh_table()
    save_to_db()
    set_status(f"✔ '{row_name}' removed.")


def clear_all():
    global students
    if not students:
        set_status("No data to clear.", error=True)
        return
    if messagebox.askyesno("Clear All", "Remove ALL students from the list and database?"):
        students = []
        refresh_table()
        save_to_db()
        set_status("✔ All records cleared.")


def calculate_rank():
    students.sort(key=lambda x: x[6], reverse=True)
    for rank, s in enumerate(students, start=1):
        s[9] = rank


def refresh_table():
    for row in table.get_children():
        table.delete(row)
    for s in students:
        level = s[8]
        color = LEVEL_COLORS.get(level, FG_DARK)
        tag   = f"lvl_{level}"
        table.insert("", tk.END, values=s, tags=(tag,))
        table.tag_configure(tag, foreground=color)


def save_to_db():
    if not DB_CONNECTED or db is None or cursor is None:
        set_status("⚠ Database not connected — data saved in memory only.", error=True)
        return
    try:
        cursor.execute("DELETE FROM students")
        for s in students:
            sql = """INSERT INTO students
                     (name, math, english, kiswahili, science, social,
                      total, average, level, rank_no)
                     VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s)"""
            cursor.execute(sql, tuple(s))
        db.commit()
    except Exception as e:
        set_status(f"DB error: {e}", error=True)


def clear_entries():
    for entry in entries:
        entry.delete(0, tk.END)
    name_entry.focus()


def set_status(msg: str, error: bool = False):
    status_var.set(msg)
    status_lbl.config(fg=FG_STATUS_ERR if error else FG_STATUS_OK)

# ─── Login System ────────────────────────────────────────────────────────────
def login():
    username = user_entry.get()
    password = pass_entry.get()

    # Admin: admin/1234, Teacher: teacher/1111
    if (username == "admin" and password == "1234") or \
       (username == "teacher" and password == "1111"):
        login_window.destroy()
        open_main_app()
    else:
        messagebox.showerror("Login Failed", "Invalid Username or Password")

def open_main_app():
    global root, table, name_entry, math_entry, english_entry, \
           kiswahili_entry, science_entry, social_entry, status_var, status_lbl, entries

    # ─── UI Construction ──────────────────────────────────────────────────────────
    root = tk.Tk()
    root.title("CBC Student Marks System")
    root.geometry("1100x650")
    root.configure(bg=BG_MAIN)
    root.resizable(True, True)

    # ── Header ──────────────────────────────────────────────────────────────────
    header = tk.Frame(root, bg=BG_HEADER, pady=10)
    header.pack(fill="x")

    tk.Label(header, text="KENYAN CBC STUDENT RESULTS SYSTEM",
             font=FONT_TITLE, bg=BG_HEADER, fg=FG_WHITE).pack()
    tk.Label(header, text="Competency Based Curriculum  |  EE · ME · AE · BE",
             font=FONT_SUB, bg=BG_HEADER, fg="#a8c8e8").pack()

    # ── Input frame ──────────────────────────────────────────────────────────────
    input_frame = tk.Frame(root, bg=BG_FRAME, padx=14, pady=10,
                           relief="ridge", bd=1)
    input_frame.pack(fill="x", padx=10, pady=(8, 0))

    labels = ["Student Name", "Math (/100)", "English (/100)",
              "Kiswahili (/100)", "Science (/100)", "Social (/100)"]
    entries = []

    for i, lbl in enumerate(labels):
        tk.Label(input_frame, text=lbl, font=FONT_LABEL,
                 bg=BG_FRAME, fg=FG_DARK).grid(row=0, column=i, padx=6, sticky="w")
        e = tk.Entry(input_frame, font=FONT_ENTRY, width=13,
                     relief="solid", bd=1)
        e.grid(row=1, column=i, padx=6, pady=4)
        entries.append(e)

    (name_entry, math_entry, english_entry,
     kiswahili_entry, science_entry, social_entry) = entries

    # ── Buttons ──────────────────────────────────────────────────────────────────
    btn_frame = tk.Frame(input_frame, bg=BG_FRAME)
    btn_frame.grid(row=0, column=len(labels), rowspan=2, padx=10, sticky="ns")

    def make_btn(parent, text, cmd, bg):
        return tk.Button(parent, text=text, command=cmd, font=FONT_BTN,
                         bg=bg, fg=FG_WHITE, relief="flat",
                         padx=10, pady=5, cursor="hand2",
                         activebackground=bg, activeforeground=FG_WHITE)

    make_btn(btn_frame, "➕  Add Student", add_student, BG_BTN_ADD).pack(fill="x", pady=2)
    make_btn(btn_frame, "📋  Report Card", generate_selected_report, "#9b59b6").pack(fill="x", pady=2)
    make_btn(btn_frame, "📊  Export Excel", export_excel, "#2ecc71").pack(fill="x", pady=2)
    make_btn(btn_frame, "📄  Export PDF", export_pdf, "#34495e").pack(fill="x", pady=2)
    make_btn(btn_frame, "🗑  Delete Row",  delete_selected, BG_BTN_DEL).pack(fill="x", pady=2)
    make_btn(btn_frame, "✖  Clear All",   clear_all, BG_BTN_CLR).pack(fill="x", pady=2)

    # ── Legend ──────────────────────────────────────────────────────────────────
    legend_frame = tk.Frame(root, bg=BG_MAIN, pady=3)
    legend_frame.pack(anchor="e", padx=14)

    tk.Label(legend_frame, text="Level: ", font=FONT_LABEL,
             bg=BG_MAIN, fg=FG_DARK).pack(side="left")
    for lvl, rng, col in [("EE","80-100","#27ae60"),("ME","65-79","#2980b9"),
                           ("AE","50-64","#e67e22"),("BE","0-49","#e74c3c")]:
        tk.Label(legend_frame, text=f"  {lvl} ({rng})",
                 font=FONT_LABEL, bg=BG_MAIN, fg=col).pack(side="left")

    # ── Table ────────────────────────────────────────────────────────────────────
    tbl_frame = tk.Frame(root, bg=BG_MAIN)
    tbl_frame.pack(fill="both", expand=True, padx=10, pady=(4, 0))

    cols = ("Name","Math","English","Kiswahili","Science","Social",
            "Total","Average","Level","Rank")

    style = ttk.Style()
    style.theme_use("clam")
    style.configure("Treeview",
                    font=("Arial", 10),
                    rowheight=24,
                    background=BG_FRAME,
                    fieldbackground=BG_FRAME,
                    foreground=FG_DARK)
    style.configure("Treeview.Heading",
                    font=("Arial", 10, "bold"),
                    background=BG_HEADER,
                    foreground=FG_WHITE,
                    relief="flat")
    style.map("Treeview.Heading", background=[("active", "#2c5282")])
    style.map("Treeview", background=[("selected", "#d5e8f7")])

    table = ttk.Treeview(tbl_frame, columns=cols, show="headings",
                         selectmode="browse")

    col_widths = {"Name": 160, "Math": 65, "English": 75, "Kiswahili": 85,
                  "Science": 70, "Social": 65, "Total": 65,
                  "Average": 75, "Level": 60, "Rank": 55}
    for c in cols:
        table.heading(c, text=c)
        table.column(c, width=col_widths.get(c, 80), anchor="center")

    table.column("Name", anchor="w")

    vsb = ttk.Scrollbar(tbl_frame, orient="vertical",   command=table.yview)
    hsb = ttk.Scrollbar(tbl_frame, orient="horizontal", command=table.xview)
    table.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

    table.grid(row=0, column=0, sticky="nsew")
    vsb.grid(row=0, column=1, sticky="ns")
    hsb.grid(row=1, column=0, sticky="ew")
    tbl_frame.rowconfigure(0, weight=1)
    tbl_frame.columnconfigure(0, weight=1)

    # ── Status bar ───────────────────────────────────────────────────────────────
    status_var = tk.StringVar()
    db_badge   = "🟢 MySQL Connected" if DB_CONNECTED else "🔴 MySQL Not Connected (offline mode)"
    status_var.set(db_badge)

    status_bar = tk.Frame(root, bg="#dde3ea", pady=3)
    status_bar.pack(fill="x", side="bottom")

    tk.Label(status_bar, text=db_badge,
             font=FONT_STATUS, bg="#dde3ea",
             fg=FG_STATUS_OK if DB_CONNECTED else FG_STATUS_ERR).pack(side="right", padx=10)

    status_lbl = tk.Label(status_bar, textvariable=status_var,
                          font=FONT_STATUS, bg="#dde3ea", fg=FG_STATUS_OK)
    status_lbl.pack(side="left", padx=10)

    # ── Key bindings ─────────────────────────────────────────────────────────────
    root.bind("<Return>", lambda e: add_student())
    root.bind("<Delete>", lambda e: delete_selected())

    name_entry.focus()
    root.mainloop()

# ─── Login Window ────────────────────────────────────────────────────────────
login_window = tk.Tk()
login_window.title("Login System")
login_window.geometry("350x250")
login_window.configure(bg=BG_MAIN)

tk.Label(login_window, text="CBC ADMIN LOGIN", font=FONT_TITLE, bg=BG_MAIN, fg=BG_HEADER).pack(pady=10)

tk.Label(login_window, text="Username", font=FONT_LABEL, bg=BG_MAIN).pack()
user_entry = tk.Entry(login_window, font=FONT_ENTRY)
user_entry.pack(pady=5)

tk.Label(login_window, text="Password", font=FONT_LABEL, bg=BG_MAIN).pack()
pass_entry = tk.Entry(login_window, show="*", font=FONT_ENTRY)
pass_entry.pack(pady=5)

tk.Button(login_window, text="Login", command=login, font=FONT_BTN,
          bg=BG_HEADER, fg=FG_WHITE, width=15, pady=5).pack(pady=20)

login_window.bind("<Return>", lambda e: login())
login_window.mainloop()
