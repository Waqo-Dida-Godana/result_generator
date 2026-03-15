# -*- coding: utf-8 -*-
"""
MOAS School Management System - Complete Merged App
Pure Tkinter • Modern UI • Full Features • Legacy DB Compatible
"""

import os
import shutil
import sys
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from PIL import Image, ImageTk
from database import db
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import csv
import pandas as pd
from datetime import datetime
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib import colors
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image as RLImage, PageBreak
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.graphics.shapes import Drawing, Rect
from reportlab.graphics.charts.barcharts import VerticalBarChart
from reportlab.lib.units import inch
from fpdf import FPDF

# ====================== DESIGN TOKENS ======================
# Logo Color Theme: Green-based
SIDEBAR_BG        = '#1b5e20'       # Dark Olive Green (darker shade)
SIDEBAR_ACTIVE    = '#4CAF50'       # Light Green
SIDEBAR_HOVER     = '#2E7D32'       # Dark Olive Green
SIDEBAR_TEXT      = '#a5d6a7'       # Light green text
SIDEBAR_TEXT_ACT  = '#ffffff'       # White

CONTENT_BG  = '#f8f9ff'              # Light background
CARD_BG     = '#ffffff'               # White
BORDER_CLR  = '#e0e0e0'               # Light gray border

TEXT_PRIMARY   = '#333333'            # Dark Gray / Black
TEXT_SECONDARY = '#666666'            # Medium Gray

BLUE   = '#4CAF50'                    # Light Green (logo theme)
GREEN  = '#2E7D32'                    # Dark Olive Green
ORANGE = '#f59e0b'                    # Orange (kept)
PURPLE = '#6366f1'                    # Purple (kept)

GRADE_COLORS = {
    'EE': '#2ecc71',
    'ME': '#3498db',
    'AE': '#f39c12',
    'BE': '#e67e22',
    'IE': '#e74c3c',
}
GRADE_LABELS = {
    'EE': 'Exceeding Expectations',
    'ME': 'Meeting Expectations',
    'AE': 'Approaching Expectations',
    'BE': 'Below Expectations',
    'IE': 'Inadequate',
}

FF = 'Segoe UI'   # font family

# ====================== TOPBAR TOKENS ======================
TOPBAR_H        = 56                # height in px
TOPBAR_RIGHT_BG = '#1b5e20'         # same dark green as sidebar/logo theme
TOPBAR_BTN_BG   = '#2E7D32'         # button bg (slightly lighter green)
TOPBAR_BTN_HOV  = '#388E3C'         # button hover
TOPBAR_ICON_FG  = '#ffffff'         # button icon/text color (white on green)
TOPBAR_USER_BG  = '#7b1fa2'         # purple user section
TOPBAR_YR_BG    = '#4CAF50'         # green year badge

# ====================== CBC LEVELS & SUBJECTS ======================
# Kenyan Competency Based Curriculum (CBC) Structure

# Education Levels
LEVELS = [
    'Lower Primary (Grade 1-3)',
    'Upper Primary (Grade 4-6)',
    'Junior School (Grade 7-9)'
]

# Subjects by Level
SUBJECTS_BY_LEVEL = {
    'Lower Primary (Grade 1-3)': [
        'Literacy Activities (Reading & Writing)',
        'Kiswahili Language Activities',
        'English Language Activities',
        'Mathematical Activities',
        'Environmental Activities',
        'Hygiene & Nutrition Activities',
        'Religious Education Activities',
        'Movement & Creative Activities'
    ],
    'Upper Primary (Grade 4-6)': [
        'English',
        'Kiswahili / Kenyan Sign Language',
        'Mathematics',
        'Science & Technology',
        'Agriculture',
        'Social Studies',
        'Religious Education',
        'Christian Religious Education (CRE)',
        'Islamic Religious Education (IRE)',
        'Hindu Religious Education (HRE)',
        'Creative Arts',
        'Physical & Health Education',
        'Home Science'
    ],
    'Junior School (Grade 7-9)': {
        'core': [
            'English',
            'Kiswahili / Kenyan Sign Language',
            'Mathematics',
            'Integrated Science',
            'Health Education',
            'Pre-Technical Studies',
            'Social Studies',
            'Religious Education (CRE/IRE/HRE)',
            'Agriculture',
            'Life Skills Education',
            'Sports & Physical Education',
            'Visual Arts',
            'Performing Arts'
        ],
        'optional': [
            'Computer Science / ICT',
            'Foreign Languages (French, German, Arabic)',
            'Kenyan Sign Language'
        ]
    }
}

# Classes by Level
CLASSES_BY_LEVEL = {
    'Lower Primary (Grade 1-3)': ['Grade 1', 'Grade 2', 'Grade 3'],
    'Upper Primary (Grade 4-6)': ['Grade 4', 'Grade 5', 'Grade 6'],
    'Junior School (Grade 7-9)': ['Grade 7', 'Grade 8', 'Grade 9']
}

# Grading System by Level (CBC Competency Levels)
GRADING_BY_LEVEL = {
    'Lower Primary (Grade 1-3)': {
        'levels': {
            'EE': {'label': 'Exceeding Expectation', 'description': 'Learner performs above the required level'},
            'ME': {'label': 'Meeting Expectation', 'description': 'Learner understands and performs well'},
            'AE': {'label': 'Approaching Expectation', 'description': 'Learner is improving but needs support'},
            'BE': {'label': 'Below Expectation', 'description': 'Learner needs more help'}
        },
        'assessment_methods': 'Class activities, Oral work, Practical activities, Teacher observation'
    },
    'Upper Primary (Grade 4-6)': {
        'levels': {
            'EE': {'label': 'Exceeding Expectation', 'description': 'Learner performs above the required level'},
            'ME': {'label': 'Meeting Expectation', 'description': 'Learner understands and performs well'},
            'AE': {'label': 'Approaching Expectation', 'description': 'Learner is improving but needs support'},
            'BE': {'label': 'Below Expectation', 'description': 'Learner needs more help'}
        },
        'assessment_components': '60% School Based Assessment (SBA) + 40% National Assessment (KNEC)'
    },
    'Junior School (Grade 7-9)': {
        'levels': {
            'EE': {'label': 'Exceeding Expectation', 'description': 'Learner performs above the required level'},
            'ME': {'label': 'Meeting Expectation', 'description': 'Learner understands and performs well'},
            'AE': {'label': 'Approaching Expectation', 'description': 'Learner is improving but needs support'},
            'BE': {'label': 'Below Expectation', 'description': 'Learner needs more help'}
        },
        'uses_percentage': True,
        'description': 'Competency levels with percentage scores'
    }
}

# Legacy compatibility - Default to Junior School
SUBJECTS = SUBJECTS_BY_LEVEL['Junior School (Grade 7-9)']['core']
CLASSES  = CLASSES_BY_LEVEL['Junior School (Grade 7-9)']

# ====================== CONSTANTS ==========================
TERMS    = ['One', 'Two', 'Three']
COLORS = {
    'primary': '#4CAF50', 'sidebar': '#1b5e20', 'card': '#ffffff',
    'text': '#333333', 'text_sec': '#666666', 'border': '#e0e0e0'
}

# ====================== CBC GRADE SUB-LEVELS ===============
def get_cbc_grade_sublevel(mark):
    """Return the CBC grade sub-level string (EE1, EE2, ME1, ME2, AE1, AE2, BE1, BE2)."""
    try:
        mark = int(mark)
    except (TypeError, ValueError):
        return ''
    if mark >= 90: return 'EE1'
    if mark >= 75: return 'EE2'
    if mark >= 60: return 'ME1'
    if mark >= 50: return 'ME2'
    if mark >= 35: return 'AE1'
    if mark >= 25: return 'AE2'
    if mark >= 12: return 'BE1'
    return 'BE2'

# ====================== HELPERS ============================
def _rr(canvas, x0, y0, x1, y1, r, fill, outline=None):
    """Draw a rounded rectangle on a Canvas."""
    oc = outline or fill
    canvas.create_arc(x0,      y0,      x0+2*r, y0+2*r, start=90,  extent=90, fill=fill, outline=oc)
    canvas.create_arc(x1-2*r,  y0,      x1,     y0+2*r, start=0,   extent=90, fill=fill, outline=oc)
    canvas.create_arc(x0,      y1-2*r,  x0+2*r, y1,     start=180, extent=90, fill=fill, outline=oc)
    canvas.create_arc(x1-2*r,  y1-2*r,  x1,     y1,     start=270, extent=90, fill=fill, outline=oc)
    canvas.create_rectangle(x0+r, y0,   x1-r, y1,   fill=fill, outline=fill)
    canvas.create_rectangle(x0,   y0+r, x1,   y1-r, fill=fill, outline=fill)


def rounded_badge(parent, icon_text, color, size=40):
    """Draw a coloured rounded-square icon badge on a Canvas."""
    c = tk.Canvas(parent, width=size, height=size, bg=CARD_BG, highlightthickness=0)
    r = 9
    x0, y0, x1, y1 = 2, 2, size - 2, size - 2
    # fill rounded rect via overlapping shapes
    c.create_arc(x0,      y0,      x0+2*r, y0+2*r, start=90,  extent=90, fill=color, outline=color)
    c.create_arc(x1-2*r, y0,      x1,     y0+2*r, start=0,   extent=90, fill=color, outline=color)
    c.create_arc(x0,      y1-2*r, x0+2*r, y1,     start=180, extent=90, fill=color, outline=color)
    c.create_arc(x1-2*r, y1-2*r, x1,     y1,     start=270, extent=90, fill=color, outline=color)
    c.create_rectangle(x0+r, y0,   x1-r, y1,   fill=color, outline=color)
    c.create_rectangle(x0,   y0+r, x1,   y1-r, fill=color, outline=color)
    c.create_text(size // 2, size // 2, text=icon_text, fill='white',
                  font=(FF, size // 3, 'bold'))
    return c


def make_card(parent, padx=20, pady=16, **grid_or_pack):
    """Return a white card frame with a 1-px border."""
    outer = tk.Frame(parent, bg=BORDER_CLR)
    inner = tk.Frame(outer, bg=CARD_BG, padx=padx, pady=pady)
    inner.pack(fill='both', expand=True, padx=1, pady=1)
    return outer, inner


def scrollable_frame(parent, bg=CONTENT_BG):
    """Return (canvas, scrollbar, inner_frame) for a vertically scrollable area."""
    canvas = tk.Canvas(parent, bg=bg, highlightthickness=0)
    sb = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
    canvas.configure(yscrollcommand=sb.set)
    sb.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    inner = tk.Frame(canvas, bg=bg)
    win = canvas.create_window((0, 0), window=inner, anchor='nw')

    def _on_resize(e):
        canvas.itemconfig(win, width=e.width)
    canvas.bind('<Configure>', _on_resize)

    def _on_frame_resize(e):
        canvas.configure(scrollregion=canvas.bbox('all'))
    inner.bind('<Configure>', _on_frame_resize)

    def _on_mousewheel(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')
    canvas.bind_all('<MouseWheel>', _on_mousewheel)

    return canvas, sb, inner


# ====================== TREEVIEW STYLE =====================
def setup_treeview_style():
    style = ttk.Style()
    style.theme_use('clam')
    style.configure('App.Treeview',
        background=CARD_BG,
        foreground=TEXT_PRIMARY,
        rowheight=34,
        fieldbackground=CARD_BG,
        borderwidth=0,
        font=(FF, 10),
    )
    style.configure('App.Treeview.Heading',
        background='#f8fafc',
        foreground=TEXT_SECONDARY,
        relief='flat',
        font=(FF, 10, 'bold'),
        padding=8,
    )
    style.map('App.Treeview',
        background=[('selected', '#eff6ff')],
        foreground=[('selected', TEXT_PRIMARY)],
    )
    style.configure('App.Vertical.TScrollbar',
        background=BORDER_CLR,
        troughcolor='#f8fafc',
        arrowcolor=TEXT_SECONDARY,
        borderwidth=0,
    )
    style.configure('App.TCombobox',
        fieldbackground=CARD_BG,
        background=CARD_BG,
        foreground=TEXT_PRIMARY,
        arrowcolor=TEXT_SECONDARY,
        bordercolor=BORDER_CLR,
        lightcolor=BORDER_CLR,
        darkcolor=BORDER_CLR,
        padding=6,
        font=(FF, 10),
    )
    style.configure('App.TEntry',
        fieldbackground=CARD_BG,
        bordercolor=BORDER_CLR,
        lightcolor=BORDER_CLR,
        padding=8,
        font=(FF, 10),
    )


# ====================== MAIN APP ===========================
class SchoolReportApp:
    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title('MT OLIVES ADVENTIST SCHOOL,NGONG')
        self.root.geometry('1280x760')
        self.root.configure(bg=CONTENT_BG)
        self.root.minsize(960, 620)

        # CBC Level - Default to Junior School
        self.current_level = 'Junior School (Grade 7-9)'
        
        self.current_user = None
        self.user_role = 'admin'  # Default role
        self.nav_frames: dict = {}
        self.active_nav: str = ''
        self.logo_img = self.load_logo()
        
        # Load logo image for use in dashboard
        self.logo_image = None
        try:
            # Get the directory where the script is located
            script_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(script_dir, 'moas.jpg')
            img = Image.open(logo_path)
            # Resize logo to fit dashboard nicely
            img = img.resize((80, 80), Image.Resampling.LANCZOS)
            self.logo_image = ImageTk.PhotoImage(img)
        except Exception as e:
            print(f'Could not load logo image: {e}')

        # ── Window / taskbar icon (Windows-compatible .ico) ──────────────────
        try:
            script_dir = os.path.dirname(os.path.abspath(__file__))
            jpg_path = os.path.join(script_dir, 'moas.jpg')
            ico_path = os.path.join(script_dir, 'moas.ico')
            # Generate .ico from jpg if not already present
            if not os.path.exists(ico_path):
                icon_img = Image.open(jpg_path).convert('RGBA')
                icon_img.save(ico_path, format='ICO',
                              sizes=[(16,16),(32,32),(48,48),(64,64)])
            self.root.iconbitmap(ico_path)
        except Exception as e:
            print(f'Could not set window icon: {e}')

        setup_treeview_style()
        self.show_login()

    # ------------------- utilities -------------------
    def get_current_subjects(self):
        """Get subjects for the current user based on role"""
        # Check if user is a teacher with assigned subjects
        if self.current_user and hasattr(self, 'user_role'):
            role = getattr(self, 'user_role', 'admin')
            teacher_id = self.current_user.get('id')
            
            # If teacher, return only assigned subjects
            if role == 'teacher' and teacher_id:
                assignments = db.get_teacher_subjects(teacher_id)
                if assignments:
                    return [a['subject'] for a in assignments]
        
        # Default to CBC level subjects
        return SUBJECTS_BY_LEVEL.get(self.current_level, SUBJECTS)
    
    def get_teacher_assigned_classes(self):
        """Get classes assigned to the current teacher"""
        if self.current_user and hasattr(self, 'user_role'):
            role = getattr(self, 'user_role', 'admin')
            teacher_id = self.current_user.get('id')
            
            if role == 'teacher' and teacher_id:
                # Get unique classes from subject assignments
                assignments = db.get_teacher_subjects(teacher_id)
                classes = list(set([a['class_name'] for a in assignments]))
                if classes:
                    return classes
            elif role == 'class_teacher' and teacher_id:
                return db.get_teacher_classes(teacher_id)
        
        return CLASSES_BY_LEVEL.get(self.current_level, CLASSES)
    
    def get_current_classes(self):
        """Get classes for the current CBC level"""
        return CLASSES_BY_LEVEL.get(self.current_level, CLASSES)
    
    def get_current_grading(self):
        """Get grading system for the current CBC level"""
        return GRADING_BY_LEVEL.get(self.current_level, GRADING_BY_LEVEL['Junior School (Grade 7-9)'])
    
    def set_level(self, level):
        """Set the current CBC level and update subjects/classes"""
        if level in LEVELS:
            self.current_level = level
            # Update legacy compatibility variables
            global SUBJECTS, CLASSES
            # Handle both list and dict structures (Junior School has core/optional)
            level_subjects = SUBJECTS_BY_LEVEL[level]
            if isinstance(level_subjects, dict):
                SUBJECTS = level_subjects.get('core', [])
            else:
                SUBJECTS = level_subjects
            CLASSES = CLASSES_BY_LEVEL[level]
            return True
        return False
    
    def load_logo(self):
        try:
            img = Image.open('moas.jpg').resize((60, 60))
            return ImageTk.PhotoImage(img)
        except:
            return None

    def _clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    def clear_frame(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

    def clear_root(self):
        self._clear()

    def _page_header(self, title: str, subtitle: str):
        hdr = tk.Frame(self.content_frame, bg=CONTENT_BG)
        hdr.pack(fill='x', pady=(0, 22))
        tk.Label(hdr, text=title, bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 26, 'bold')).pack(anchor='w')
        tk.Label(hdr, text=subtitle, bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 11)).pack(anchor='w')

    def _toolbar_btn(self, parent, text, command, bg=BLUE, fg='white'):
        b = tk.Label(parent, text=text, bg=bg, fg=fg,
                     font=(FF, 10, 'bold'), padx=14, pady=7, cursor='hand2')
        b.bind('<Button-1>', lambda e: command())
        b.bind('<Enter>', lambda e: b.config(bg='#388E3C' if bg == '#4CAF50' else '#1b5e20'))
        b.bind('<Leave>', lambda e: b.config(bg=bg))
        return b

    # ─────────────────────── Auth screen ───────────────────────────────────
    _AUTH_BG   = '#f5f5f5'
    _CARD_W    = 400
    _BLUE      = '#4CAF50'      # Light Green (logo theme)
    _DARK_BLUE = '#2E7D32'      # Dark Olive Green
    _FIELD_BG  = '#f0f0f0'
    _TEXT      = '#333333'
    _GRAY_T    = '#666666'

    def show_login(self):
        self._clear()
        self._auth_tab = 'login'
        self._auth_outer = tk.Frame(self.root, bg=self._AUTH_BG)
        self._auth_outer.pack(fill='both', expand=True)
        self._build_auth_card()

    def _build_auth_card(self):
        for w in self._auth_outer.winfo_children():
            w.destroy()

        tab    = self._auth_tab
        CW     = self._CARD_W
        CARD_H = 492 if tab == 'login' else 565
        PAD    = 6

        # ── Canvas: shadow + rounded white card ─────────────────────────────
        canvas = tk.Canvas(self._auth_outer, width=CW + PAD*2,
                           height=CARD_H + PAD*2, bg=self._AUTH_BG,
                           highlightthickness=0)
        canvas.place(relx=0.5, rely=0.5, anchor='center')

        # shadow
        _rr(canvas, PAD+3, PAD+5, CW+PAD+3, CARD_H+PAD+5, 16, '#cdd1dc')
        # card
        _rr(canvas, PAD, PAD, CW+PAD, CARD_H+PAD, 16, 'white')

        # content frame embedded in canvas
        frame = tk.Frame(canvas, bg='white')
        canvas.create_window(CW // 2 + PAD, PAD + 1, window=frame,
                             anchor='n', width=CW - 40)
        self._auth_frame = frame

        # ── Logo ────────────────────────────────────────────────────────────
        if self.logo_img:
            tk.Label(frame, image=self.logo_img, bg='white').pack(pady=20)
        else:
            logo = tk.Canvas(frame, width=56, height=56, bg='white', highlightthickness=0)
            logo.pack(pady=(24, 12))
            _rr(logo, 0, 0, 56, 56, 12, self._BLUE)
            logo.create_text(28, 28, text='\U0001f393', font=(FF, 22), fill='white')

        # ── Title ───────────────────────────────────────────────────────────
        tk.Label(frame, text='MOAS', bg='white', fg=self._TEXT,
                 font=(FF, 16, 'bold')).pack(pady=(5, 3))
        tk.Label(frame, text='School Management System', bg='white',
                 fg=self._GRAY_T, font=(FF, 11)).pack(pady=(0, 18))

        # ── Tab switcher ────────────────────────────────────────────────────
        TH = 42
        tab_c = tk.Canvas(frame, height=TH, bg='white', highlightthickness=0)
        tab_c.pack(fill='x', pady=(0, 14))

        def draw_tabs(w):
            tab_c.delete('all')
            _rr(tab_c, 0, 0, w, TH, 10, self._FIELD_BG)
            half = w // 2
            # active indicator
            if tab == 'login':
                _rr(tab_c, 4, 4, half - 2, TH - 4, 7, 'white')
            else:
                _rr(tab_c, half + 2, 4, w - 4, TH - 4, 7, 'white')
            # Login label
            lo = tk.Label(tab_c,
                          text='Login',
                          bg='white' if tab == 'login' else self._FIELD_BG,
                          fg=self._TEXT if tab == 'login' else self._GRAY_T,
                          font=(FF, 11, 'bold' if tab == 'login' else 'normal'),
                          cursor='hand2')
            tab_c.create_window(half // 2, TH // 2, window=lo, width=half - 10)
            lo.bind('<Button-1>', lambda e: self._switch_tab('login'))
            # Sign Up label
            su = tk.Label(tab_c,
                          text='Sign Up',
                          bg='white' if tab == 'signup' else self._FIELD_BG,
                          fg=self._TEXT if tab == 'signup' else self._GRAY_T,
                          font=(FF, 11, 'bold' if tab == 'signup' else 'normal'),
                          cursor='hand2')
            tab_c.create_window(half + half // 2, TH // 2, window=su, width=half - 10)
            su.bind('<Button-1>', lambda e: self._switch_tab('signup'))

        tab_c.bind('<Configure>', lambda e: draw_tabs(e.width))

        # ── Form fields ─────────────────────────────────────────────────────
        self._auth_entries = {}

        def mk_field(label, key, show=''):
            tk.Label(frame, text=label, bg='white', fg=self._TEXT,
                     font=(FF, 11, 'bold'), anchor='w').pack(fill='x', pady=(10, 4))
            wrap = tk.Frame(frame, bg=self._FIELD_BG, padx=14)
            wrap.pack(fill='x')
            e = tk.Entry(wrap, bg=self._FIELD_BG, fg=self._TEXT, relief='flat',
                         font=(FF, 12), show=show, bd=0,
                         highlightthickness=0, insertbackground=self._TEXT)
            e.pack(fill='x', ipady=12)
            e.bind('<Return>', lambda ev: (self._do_login() if tab == 'login'
                                           else self._do_register()))
            self._auth_entries[key] = e
            return e

        if tab == 'signup':
            mk_field('Full Name', 'name')
        first = mk_field('Email', 'email')
        mk_field('Password', 'password', show='\u25cf')
        first.focus_set()

        # ── Action button ────────────────────────────────────────────────────
        btn_text = 'Sign In' if tab == 'login' else 'Create Account'
        btn_cmd  = self._do_login if tab == 'login' else self._do_register

        btn_c = tk.Canvas(frame, height=46, bg='white', highlightthickness=0,
                          cursor='hand2')
        btn_c.pack(fill='x', pady=(20, 0))

        def draw_btn(color=None):
            color = color or self._BLUE
            w = btn_c.winfo_width() or (CW - 40)
            btn_c.delete('all')
            _rr(btn_c, 0, 0, w, 46, 8, color)
            btn_c.create_text(w // 2, 23, text=btn_text, fill='white',
                              font=(FF, 12, 'bold'))

        btn_c.bind('<Configure>', lambda e: draw_btn())
        btn_c.bind('<Button-1>',  lambda e: btn_cmd())
        btn_c.bind('<Enter>',     lambda e: draw_btn('#388E3C' if self._BLUE == '#4CAF50' else self._DARK_BLUE))
        btn_c.bind('<Leave>',     lambda e: draw_btn(self._BLUE))

        # ── Forgot password / bottom spacer ─────────────────────────────────
        if tab == 'login':
            fp = tk.Label(frame, text='Forgot password?', bg='white',
                          fg=self._BLUE, font=(FF, 10), cursor='hand2')
            fp.pack(pady=(12, 24))
            fp.bind('<Button-1>', lambda e: messagebox.showinfo(
                'Forgot Password',
                'Please contact your administrator to reset your password.'))
            
            # Demo Mode button
            tk.Button(frame, text='Demo Mode (Skip Login)', command=self.show_main, 
                     bg='#e0e0e0', width=20, pady=5).pack(pady=5)
        else:
            tk.Frame(frame, bg='white', height=24).pack()

    def _switch_tab(self, tab):
        if self._auth_tab != tab:
            self._auth_tab = tab
            self._build_auth_card()

    def _do_login(self):
        email = self._auth_entries.get('email', tk.Entry()).get().strip()
        pwd   = self._auth_entries.get('password', tk.Entry()).get().strip()
        if not email or not pwd:
            messagebox.showerror('Error', 'Please fill in all fields')
            return
        user = db.authenticate(email, pwd)
        if user:
            self.current_user = user
            # Get user's role
            user_role = user.get('role', 'admin')
            self.user_role = user_role
            self.show_main()
        else:
            messagebox.showerror('Login Failed', 'Invalid email or password.\n\n'
                                 'Default: admin / admin123')

    def _do_register(self):
        name  = self._auth_entries.get('name',  tk.Entry()).get().strip()
        email = self._auth_entries.get('email', tk.Entry()).get().strip()
        pwd   = self._auth_entries.get('password', tk.Entry()).get().strip()
        if not all([name, email, pwd]):
            messagebox.showerror('Error', 'All fields are required')
            return
        if len(pwd) < 6:
            messagebox.showerror('Error', 'Password must be at least 6 characters')
            return
        if db.register_user(name, email, pwd):
            messagebox.showinfo('Success', 'Account created! Please sign in.')
            self._switch_tab('login')
        else:
            messagebox.showerror('Error', 'This email is already registered.')

    def do_login(self):
        """Legacy – kept for compatibility"""
        email = self.email_entry.get()
        pwd = self.pwd_entry.get()
        if db.authenticate(email, pwd):
            self.show_main()
        else:
            messagebox.showerror('Login Failed', 'Invalid credentials. Try admin/admin123')

    # ------------------- main layout -------------------
    def show_main(self):
        self._clear()

        wrapper = tk.Frame(self.root, bg=CONTENT_BG)
        wrapper.pack(fill='both', expand=True)

        # ── Top navbar ──────────────────────────────────────
        self._build_topbar(wrapper)

        # ── Bottom: sidebar + content ────────────────────────
        bottom = tk.Frame(wrapper, bg=CONTENT_BG)
        bottom.pack(fill='both', expand=True)

        self._build_sidebar(bottom)

        # 1-px divider between sidebar and content
        tk.Frame(bottom, bg=BORDER_CLR, width=1).pack(side='left', fill='y')

        # content scroller
        content_wrapper = tk.Frame(bottom, bg=CONTENT_BG)
        content_wrapper.pack(side='left', fill='both', expand=True)

        self.content = tk.Frame(content_wrapper, bg=CONTENT_BG)
        self.content.pack(fill='both', expand=True, padx=28, pady=24)
        self.content_frame = self.content

        self.show_dashboard()

    # ─────────────────────── Top Navbar ────────────────────────────────────
    def _draw_topbar_icon(self, c, icon_type, bg):
        """Draw crisp vector-style icons on a Canvas widget."""
        c.delete('all')
        c.config(bg=bg)
        FG = TOPBAR_ICON_FG
        if icon_type == 'logout':
            # Door frame (left rectangle)
            c.create_rectangle(2, 2, 13, 20, outline=FG, width=1.5)
            # Door knob
            c.create_oval(9, 10, 11.5, 12.5, fill=FG, outline='')
            # Arrow shaft pointing right
            c.create_line(13, 11, 22, 11, fill=FG, width=2)
            # Arrow head
            c.create_line(18, 7,  22, 11, fill=FG, width=2)
            c.create_line(18, 15, 22, 11, fill=FG, width=2)
        elif icon_type == 'about':
            # Circle outline
            c.create_oval(1, 1, 21, 21, outline=FG, width=1.5)
            # Question mark glyph
            c.create_text(11, 12, text='?', fill=FG,
                          font=(FF, 11, 'bold'))

    def _topbar_btn(self, parent, icon_type, label, bg, command=None):
        """Icon-over-text button with crisp Canvas-drawn icons."""
        hover = TOPBAR_BTN_HOV

        f = tk.Frame(parent, bg=bg, padx=16, cursor='hand2')
        f.pack(side='right', fill='y')

        # Canvas icon (22×22 drawing area)
        ic = tk.Canvas(f, width=22, height=22, bg=bg, highlightthickness=0)
        ic.pack(pady=(10, 2))
        self._draw_topbar_icon(ic, icon_type, bg)

        txt = tk.Label(f, text=label, bg=bg, fg=TOPBAR_ICON_FG,
                       font=(FF, 8))
        txt.pack(pady=(0, 10))

        all_w = [f, ic, txt]
        if command:
            for w in all_w:
                w.bind('<Button-1>', lambda e: command())

        def _on_enter(e):
            for w in [f, txt]: w.config(bg=hover)
            ic.config(bg=hover)
            self._draw_topbar_icon(ic, icon_type, hover)

        def _on_leave(e):
            for w in [f, txt]: w.config(bg=bg)
            ic.config(bg=bg)
            self._draw_topbar_icon(ic, icon_type, bg)

        for w in all_w:
            w.bind('<Enter>', _on_enter)
            w.bind('<Leave>', _on_leave)
        return f

    def _build_topbar(self, parent):
        """Build the Gestio-RPS-style top navigation bar."""
        now = datetime.now()
        acad_year = f'{now.year}/{now.year + 1}'
        uname = self.current_user.get('username', 'Admin') if self.current_user else 'Admin'
        role  = uname.title()

        # ── Outer bar – right-area bg fills the whole bar ─────────────────
        bar = tk.Frame(parent, bg=TOPBAR_RIGHT_BG, height=TOPBAR_H)
        bar.pack(fill='x', side='top')
        bar.pack_propagate(False)

        # ── Right buttons – packed right→left so visual order = left→right ─
        # About
        self._topbar_btn(bar, 'about', 'About', TOPBAR_BTN_BG,
                         command=lambda: messagebox.showinfo(
                             'About', 'School MIS v1.0\nMT Olives Adventist School'))
        # Log Out  (License removed)
        self._topbar_btn(bar, 'logout', 'Log Out', TOPBAR_BTN_BG,
                         command=self.logout)

        # ── User section (purple) ──────────────────────────────────────────
        usr_f = tk.Frame(bar, bg=TOPBAR_USER_BG, padx=14, cursor='hand2')
        usr_f.pack(side='right', fill='y')

        # Crisp Canvas person icon
        av = tk.Canvas(usr_f, width=28, height=28,
                       bg=TOPBAR_USER_BG, highlightthickness=0)
        av.pack(side='left', pady=14)
        av.create_oval(7, 1, 21, 15, fill='white', outline='')       # head
        av.create_arc(1, 13, 27, 31, start=0, extent=180,
                      fill='white', outline='')                        # shoulders

        tk.Label(usr_f, text=f'  admin | {role}',
                 bg=TOPBAR_USER_BG, fg='white',
                 font=(FF, 10, 'bold')).pack(side='left')

        # ── Academic year badge (green) ───────────────────────────────────
        yr_f = tk.Frame(bar, bg=TOPBAR_YR_BG, padx=12, cursor='hand2')
        yr_f.pack(side='right', fill='y', padx=(0, 2))

        # Canvas checkmark
        ck = tk.Canvas(yr_f, width=22, height=18,
                       bg=TOPBAR_YR_BG, highlightthickness=0)
        ck.pack(pady=(10, 0))
        ck.create_line(2, 9, 8, 15, fill='white', width=2.5,
                       capstyle='round', joinstyle='round')
        ck.create_line(8, 15, 20, 4, fill='white', width=2.5,
                       capstyle='round', joinstyle='round')

        tk.Label(yr_f, text=acad_year, bg=TOPBAR_YR_BG, fg='white',
                 font=(FF, 7, 'bold')).pack(pady=(0, 10))

        # ── CBC Level Selector ────────────────────────────────────────────────
        level_frame = tk.Frame(bar, bg='#1565c0', padx=10, cursor='hand2')
        level_frame.pack(side='right', fill='y', padx=(8, 2))
        
        tk.Label(level_frame, text='📚 CBC Level:', bg='#1565c0', fg='white',
                 font=(FF, 9)).pack(side='left', pady=14)
        
        # Level dropdown
        self.level_var = tk.StringVar(value=self.current_level)
        level_cb = ttk.Combobox(level_frame, textvariable=self.level_var, 
                                values=LEVELS, state='readonly', width=22)
        level_cb.pack(side='left', padx=(5, 0), pady=10)
        level_cb.bind('<<ComboboxSelected>>', lambda e: self._on_level_change())
        
        # ── Live clock ────────────────────────────────────────────────────
        dt_lbl = tk.Label(bar, bg=TOPBAR_RIGHT_BG, fg='#c8e6c9',
                          font=(FF, 10), padx=16)
        dt_lbl.pack(side='right', fill='y')

        def _tick():
            try:
                dt_lbl.config(
                    text=datetime.now().strftime('%a, %d %b %Y  %H:%M:%S'))
                bar.after(1000, _tick)
            except tk.TclError:
                pass
        _tick()

    def _on_level_change(self):
        """Handle CBC level change - update subjects and classes"""
        new_level = self.level_var.get()
        if new_level and new_level != self.current_level:
            self.set_level(new_level)
            # Refresh current view if logged in
            if self.active_nav:
                # Navigate to dashboard to refresh
                self.show_dashboard()
            messagebox.showinfo('CBC Level Changed', 
                f'Switched to: {new_level}\n\nSubjects updated to:\n' + 
                '\n'.join(f'• {s}' for s in self.get_current_subjects()[:5]) +
                ('...' if len(self.get_current_subjects()) > 5 else ''))

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=SIDEBAR_BG, width=238)
        sb.pack(side='left', fill='y')
        sb.pack_propagate(False)

        # --- logo ---
        logo_row = tk.Frame(sb, bg=SIDEBAR_BG)
        logo_row.pack(fill='x', padx=18, pady=(10, 10))

        if self.logo_img:
            tk.Label(logo_row, image=self.logo_img, bg=SIDEBAR_BG).pack(side='left')
        else:
            icon_c = tk.Canvas(logo_row, width=40, height=40, bg=SIDEBAR_BG, highlightthickness=0)
            icon_c.pack(side='left')
            icon_c.create_oval(0, 0, 40, 40, fill=SIDEBAR_ACTIVE, outline=SIDEBAR_ACTIVE)
            icon_c.create_text(20, 20, text='\U0001f393', font=(FF, 17))

        title_box = tk.Frame(logo_row, bg=SIDEBAR_BG)
        title_box.pack(side='left', padx=10)
        tk.Label(title_box, text='MT OLIVES', bg=SIDEBAR_BG, fg='white',
                 font=(FF, 11, 'bold')).pack(anchor='w')
        tk.Label(title_box, text='ADVENTIST SCHOOL', bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                 font=(FF, 8)).pack(anchor='w')

        tk.Frame(sb, bg='#2E7D32', height=1).pack(fill='x', padx=18, pady=10)

        # --- nav items ---
        nav_cfg = self._get_role_based_nav()
        nav_box = tk.Frame(sb, bg=SIDEBAR_BG)
        nav_box.pack(fill='x', padx=10)
        self.nav_frames = {}

        for icon, label, cmd in nav_cfg:
            self._nav_item(nav_box, icon, label, cmd)

        # --- bottom ---
        bot = tk.Frame(sb, bg=SIDEBAR_BG)
        bot.pack(side='bottom', fill='x', padx=18, pady=18)
        tk.Frame(bot, bg='#2E7D32', height=1).pack(fill='x', pady=(0, 12))

        uname = self.current_user.get('username', 'admin') if self.current_user else 'admin'
        email = uname if '@' in uname else uname + '@school.ac'
        tk.Label(bot, text=email, bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                 font=(FF, 10), anchor='w').pack(fill='x', pady=(0, 8))

        so = tk.Frame(bot, bg=SIDEBAR_BG, cursor='hand2')
        so.pack(fill='x')
        lbl_so = tk.Label(so, text='\u2192  Sign Out', bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                          font=(FF, 11), anchor='w', padx=4, pady=6)
        lbl_so.pack(fill='x')

        for w in (so, lbl_so):
            w.bind('<Button-1>', lambda e: self.logout())
            w.bind('<Enter>', lambda e: [so.config(bg=SIDEBAR_HOVER), lbl_so.config(bg=SIDEBAR_HOVER)])
            w.bind('<Leave>', lambda e: [so.config(bg=SIDEBAR_BG), lbl_so.config(bg=SIDEBAR_BG)])

    def _get_role_based_nav(self):
        """Get navigation items based on user role"""
        role = getattr(self, 'user_role', 'admin')
        
        # Base navigation for all users
        nav = [
            ('🏠',  'Dashboard',   self.show_dashboard),
        ]
        
        # Admin gets full access
        if role == 'admin':
            nav.extend([
                ('⚙️',  'Settings',    self.show_settings),
                ('👥',  'Students',    self.show_students),
                ('👨‍🏫', 'Teachers',    self.show_teachers),
                ('📝',  'Enter Marks', self.show_marks_entry),
                ('📊',  'Reports',     self.show_reports),
                ('📈',  'Charts',      self.show_charts),
                ('📄',  'Report Cards', self.show_report_cards),
                ('📚',  'CBC Info',    self.show_cbc_info),
            ])
        # Subject teacher
        elif role == 'teacher':
            nav.extend([
                ('📝',  'Enter Marks', self.show_marks_entry),
                ('📚',  'CBC Info',    self.show_cbc_info),
            ])
        # Class teacher
        elif role == 'class_teacher':
            nav.extend([
                ('👥',  'My Students', self.show_class_students),
                ('📝',  'Enter Marks', self.show_marks_entry),
                ('💬',  'Add Comments', self.show_add_comments),
                ('📄',  'Report Cards', self.show_report_cards),
                ('📚',  'CBC Info',    self.show_cbc_info),
            ])
        
        return nav

    def _nav_item(self, parent, icon: str, label: str, cmd):
        frame = tk.Frame(parent, bg=SIDEBAR_BG, cursor='hand2')
        frame.pack(fill='x', pady=2)

        row = tk.Frame(frame, bg=SIDEBAR_BG, padx=12, pady=9)
        row.pack(fill='x')

        ico = tk.Label(row, text=icon, bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                       font=(FF, 13), width=2, anchor='center')
        ico.pack(side='left')
        txt = tk.Label(row, text=f'  {label}', bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                       font=(FF, 11), anchor='w')
        txt.pack(side='left', fill='x')

        widgets = [frame, row, ico, txt]
        self.nav_frames[label] = widgets

        def activate(e=None):
            self._set_nav(label)
            cmd()

        def hover_on(e=None):
            if self.active_nav != label:
                for w in widgets: w.config(bg=SIDEBAR_HOVER)
                ico.config(fg='white')
                txt.config(fg='white')

        def hover_off(e=None):
            if self.active_nav != label:
                for w in widgets: w.config(bg=SIDEBAR_BG)
                ico.config(fg=SIDEBAR_TEXT)
                txt.config(fg=SIDEBAR_TEXT)

        for w in widgets:
            w.bind('<Button-1>', activate)
            w.bind('<Enter>',    hover_on)
            w.bind('<Leave>',    hover_off)

    def _set_nav(self, label: str):
        # deactivate old
        if self.active_nav and self.active_nav in self.nav_frames:
            for w in self.nav_frames[self.active_nav]:
                w.config(bg=SIDEBAR_BG)
            self.nav_frames[self.active_nav][2].config(fg=SIDEBAR_TEXT)  # icon
            self.nav_frames[self.active_nav][3].config(fg=SIDEBAR_TEXT)  # text
        # activate new
        self.active_nav = label
        if label in self.nav_frames:
            for w in self.nav_frames[label]:
                w.config(bg=SIDEBAR_ACTIVE)
            self.nav_frames[label][2].config(fg='white')
            self.nav_frames[label][3].config(fg='white')

    def logout(self):
        self.current_user = None
        self.user_role = 'admin'
        self.show_login()

    def change_password(self):
        old = simpledialog.askstring('Change Password', 'Current password:', show='*', parent=self.root)
        if not old: return
        new = simpledialog.askstring('Change Password', 'New password:', show='*', parent=self.root)
        if not new: return
        if new != simpledialog.askstring('Change Password', 'Confirm new password:', show='*', parent=self.root):
            messagebox.showerror('Error', 'Passwords do not match')
            return
        if db.change_password(self.current_user['username'], old, new):
            messagebox.showinfo('Success', 'Password changed successfully')
        else:
            messagebox.showerror('Error', 'Current password is incorrect')

    # ==================== SETTINGS ====================
    def show_settings(self):
        """Show settings/admin page for managing classes, streams, subjects"""
        self.clear_frame()
        self._set_nav('Settings')
        self._page_header('Settings', 'Manage school configuration')
        
        # Create notebook for different settings
        notebook = ttk.Notebook(self.content_frame)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Style tabs
        style = ttk.Style()
        style.theme_use('default')
        style.configure('Settings.TNotebook', background=CONTENT_BG)
        style.configure('Settings.TNotebook.Tab', background=CARD_BG, padding=[10, 5])
        notebook.configure(style='Settings.TNotebook')
        
        # Classes Tab
        classes_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(classes_frame, text='  Classes  ')
        self._build_classes_tab(classes_frame)
        
        # Streams Tab
        streams_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(streams_frame, text='  Streams  ')
        self._build_streams_tab(streams_frame)
        
        # Subjects Tab
        subjects_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(subjects_frame, text='  Subjects  ')
        self._build_subjects_tab(subjects_frame)
        
        # Teacher Assignments Tab
        assignments_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(assignments_frame, text='  Teacher Assignments  ')
        self._build_assignments_tab(assignments_frame)
    
    def _build_classes_tab(self, parent):
        """Build classes management tab"""
        # Toolbar
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)
        
        tk.Button(toolbar, text='+ Add Class', bg=GREEN, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=self._add_class_dialog).pack(side='left', padx=5)
        
        tk.Button(toolbar, text='Refresh', bg='#666', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._refresh_classes(parent)).pack(side='left', padx=5)
        
        # Classes list
        list_frame = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        cols = ('id', 'name', 'level', 'stream')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        tree.heading('id', text='ID')
        tree.heading('name', text='Class Name')
        tree.heading('level', text='Level')
        tree.heading('stream', text='Stream')
        
        tree.column('id', width=50)
        tree.column('name', width=150)
        tree.column('level', width=150)
        tree.column('stream', width=150)
        
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Load classes
        self._load_classes(tree)
        
        # Delete button
        tk.Button(list_frame, text='Delete Selected', bg='#e74c3c', fg='white',
                 font=(FF, 10), padx=10, pady=5, command=lambda: self._delete_class(tree)).pack(pady=5)
    
    def _load_classes(self, tree):
        """Load classes into tree"""
        for item in tree.get_children():
            tree.delete(item)
        
        classes = db.get_all_classes()
        for cls in classes:
            tree.insert('', 'end', values=(
                cls.get('id', '')[:8],
                cls.get('name', ''),
                cls.get('level', ''),
                cls.get('stream', '') or '-'
            ))
    
    def _refresh_classes(self, parent):
        """Refresh classes tab"""
        for widget in parent.winfo_children():
            widget.destroy()
        self._build_classes_tab(parent)
    
    def _add_class_dialog(self):
        """Dialog to add a new class"""
        dialog = tk.Toplevel(self.root)
        dialog.title('Add Class')
        dialog.geometry('400x300')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Class Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Level:', font=(FF, 11)).pack(pady=(15, 5))
        level_var = tk.StringVar()
        level_cb = ttk.Combobox(dialog, textvariable=level_var, values=LEVELS, state='readonly', font=(FF, 10))
        level_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Stream (optional):', font=(FF, 11)).pack(pady=(15, 5))
        stream_entry = tk.Entry(dialog, font=(FF, 11))
        stream_entry.pack(fill='x', padx=20)
        
        def save():
            name = name_entry.get().strip()
            level = level_var.get()
            stream = stream_entry.get().strip()
            
            if not name or not level:
                messagebox.showerror('Error', 'Class name and level are required')
                return
            
            success, msg = db.add_class(name, level, stream or None)
            if success:
                messagebox.showinfo('Success', msg)
                dialog.destroy()
                self.show_settings()
            else:
                messagebox.showerror('Error', msg)
        
        tk.Button(dialog, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(pady=20)
    
    def _delete_class(self, tree):
        """Delete selected class"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select', 'Please select a class to delete')
            return
        
        if not messagebox.askyesno('Confirm', 'Delete this class?'):
            return
        
        item = tree.item(selected)
        class_id = item['values'][0]
        
        if db.delete_class(class_id):
            messagebox.showinfo('Success', 'Class deleted')
            self._load_classes(tree)
        else:
            messagebox.showerror('Error', 'Failed to delete class')
    
    def _build_streams_tab(self, parent):
        """Build streams management tab"""
        # Toolbar
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)
        
        tk.Button(toolbar, text='+ Add Stream', bg=GREEN, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=self._add_stream_dialog).pack(side='left', padx=5)
        
        # Streams list
        list_frame = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        cols = ('id', 'name', 'class_name')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        tree.heading('id', text='ID')
        tree.heading('name', text='Stream Name')
        tree.heading('class_name', text='Class')
        
        tree.column('id', width=50)
        tree.column('name', width=200)
        tree.column('class_name', width=200)
        
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Load streams
        classes = db.get_all_classes()
        for cls in classes:
            streams = db.get_streams_for_class(cls['id'])
            for stream in streams:
                tree.insert('', 'end', values=(
                    stream.get('id', '')[:8],
                    stream.get('name', ''),
                    cls.get('name', '')
                ))
    
    def _add_stream_dialog(self):
        """Dialog to add a new stream"""
        classes = db.get_all_classes()
        if not classes:
            messagebox.showwarning('No Classes', 'Please add classes first')
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title('Add Stream')
        dialog.geometry('400x200')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Stream Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Class:', font=(FF, 11)).pack(pady=(15, 5))
        class_var = tk.StringVar()
        class_cb = ttk.Combobox(dialog, textvariable=class_var, 
                               values=[c['name'] for c in classes], state='readonly', font=(FF, 10))
        class_cb.pack(fill='x', padx=20)
        
        def save():
            name = name_entry.get().strip()
            class_name = class_var.get()
            
            if not name or not class_name:
                messagebox.showerror('Error', 'All fields are required')
                return
            
            # Get class ID
            class_id = next((c['id'] for c in classes if c['name'] == class_name), None)
            if not class_id:
                messagebox.showerror('Error', 'Invalid class')
                return
            
            success, msg = db.add_stream(name, class_id)
            if success:
                messagebox.showinfo('Success', msg)
                dialog.destroy()
                self.show_settings()
            else:
                messagebox.showerror('Error', msg)
        
        tk.Button(dialog, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(pady=20)
    
    def _build_subjects_tab(self, parent):
        """Build subjects management tab"""
        # Toolbar
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)
        
        tk.Button(toolbar, text='+ Add Subject', bg=GREEN, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=self._add_subject_dialog).pack(side='left', padx=5)
        
        # Subjects list
        list_frame = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        cols = ('id', 'name', 'level', 'category', 'optional')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        tree.heading('id', text='ID')
        tree.heading('name', text='Subject Name')
        tree.heading('level', text='Level')
        tree.heading('category', text='Category')
        tree.heading('optional', text='Optional')
        
        tree.column('id', width=50)
        tree.column('name', width=200)
        tree.column('level', width=150)
        tree.column('category', width=150)
        tree.column('optional', width=80)
        
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Load subjects
        subjects = db.get_subjects_by_level()
        for subj in subjects:
            tree.insert('', 'end', values=(
                subj.get('id', '')[:8],
                subj.get('name', ''),
                subj.get('level', ''),
                subj.get('category', ''),
                'Yes' if subj.get('is_optional') else 'No'
            ))
    
    def _add_subject_dialog(self):
        """Dialog to add a new subject"""
        dialog = tk.Toplevel(self.root)
        dialog.title('Add Subject')
        dialog.geometry('400x300')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Subject Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Level:', font=(FF, 11)).pack(pady=(15, 5))
        level_var = tk.StringVar()
        level_cb = ttk.Combobox(dialog, textvariable=level_var, values=LEVELS, state='readonly', font=(FF, 10))
        level_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Category:', font=(FF, 11)).pack(pady=(15, 5))
        category_entry = tk.Entry(dialog, font=(FF, 11))
        category_entry.pack(fill='x', padx=20)
        
        optional_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dialog, text='Optional Subject', variable=optional_var,
                      font=(FF, 11)).pack(pady=15)
        
        def save():
            name = name_entry.get().strip()
            level = level_var.get()
            category = category_entry.get().strip()
            
            if not name or not level or not category:
                messagebox.showerror('Error', 'All fields are required')
                return
            
            success, msg = db.add_subject(name, level, category, optional_var.get())
            if success:
                messagebox.showinfo('Success', msg)
                dialog.destroy()
                self.show_settings()
            else:
                messagebox.showerror('Error', msg)
        
        tk.Button(dialog, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(pady=20)
    
    def _build_assignments_tab(self, parent):
        """Build teacher assignments tab"""
        # This reuses the teachers page functionality
        container = tk.Frame(parent, bg=CONTENT_BG)
        container.pack(fill='both', expand=True)
        
        # Subject Teacher Assignments
        tk.Label(container, text='Subject Teacher Assignments', font=(FF, 14, 'bold'),
                bg=CONTENT_BG, fg=TEXT_PRIMARY).pack(pady=10)
        
        subj_frame = tk.Frame(container, bg=CARD_BG, relief='flat', bd=1)
        subj_frame.pack(fill='x', padx=10, pady=5)
        
        cols = ('teacher', 'subject', 'class')
        subj_tree = ttk.Treeview(subj_frame, columns=cols, show='headings', height=8)
        subj_tree.heading('teacher', text='Teacher')
        subj_tree.heading('subject', text='Subject')
        subj_tree.heading('class', text='Class')
        subj_tree.column('teacher', width=150)
        subj_tree.column('subject', width=150)
        subj_tree.column('class', width=150)
        subj_tree.pack(fill='x', padx=10, pady=10)
        
        assignments = db.get_subject_teacher_assignments()
        for a in assignments:
            subj_tree.insert('', 'end', values=(
                a.get('full_name', ''),
                a.get('subject', ''),
                a.get('class_name', '')
            ))
        
        # Class Teacher Assignments
        tk.Label(container, text='Class Teacher Assignments', font=(FF, 14, 'bold'),
                bg=CONTENT_BG, fg=TEXT_PRIMARY).pack(pady=10)
        
        class_frame = tk.Frame(container, bg=CARD_BG, relief='flat', bd=1)
        class_frame.pack(fill='x', padx=10, pady=5)
        
        cols = ('teacher', 'class')
        class_tree = ttk.Treeview(class_frame, columns=cols, show='headings', height=5)
        class_tree.heading('teacher', text='Teacher')
        class_tree.heading('class', text='Class')
        class_tree.column('teacher', width=200)
        class_tree.column('class', width=200)
        class_tree.pack(fill='x', padx=10, pady=10)
        
        class_assignments = db.get_class_teacher_assignments()
        for a in class_assignments:
            class_tree.insert('', 'end', values=(
                a.get('full_name', ''),
                a.get('class_name', '')
            ))
        
        # Buttons
        btn_frame = tk.Frame(container, bg=CONTENT_BG)
        btn_frame.pack(pady=10)
        
        tk.Button(btn_frame, text='Assign Subject Teacher', bg=BLUE, fg='white',
                 font=(FF, 10), padx=10, pady=5, command=self._assign_subject_dialog).pack(side='left', padx=5)
        
        tk.Button(btn_frame, text='Assign Class Teacher', bg=PURPLE, fg='white',
                 font=(FF, 10), padx=10, pady=5, command=self._assign_class_teacher_dialog).pack(side='left', padx=5)

    # ==================== DASHBOARD ====================
    def show_dashboard(self):
        self.clear_frame()
        self._set_nav('Dashboard')
        
        # Header with logo
        header_frame = tk.Frame(self.content_frame, bg=CONTENT_BG)
        header_frame.pack(fill='x', pady=(0, 22))
        
        if self.logo_image:
            logo_label = tk.Label(header_frame, image=self.logo_image, bg=CONTENT_BG)
            logo_label.pack(side='left', padx=(0, 15))
        
        title_frame = tk.Frame(header_frame, bg=CONTENT_BG)
        title_frame.pack(side='left')
        tk.Label(title_frame, text='Dashboard', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 26, 'bold')).pack(anchor='w')
        tk.Label(title_frame, text='Overview', bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 11)).pack(anchor='w')

        stats = db.get_statistics('One')

        stat_configs = [
            ('Total Students', str(stats['students']),  '⊙', BLUE),
            ('Class Average',  str(stats['avg_score']), '\u2197', GREEN),
            ('Top Student',    stats['top_student'],    '\u2605', ORANGE),
            ('Subjects',       str(stats['subjects']),  '\u25a6', PURPLE),
        ]

        # ---- stat cards row ----
        cards_row = tk.Frame(self.content_frame, bg=CONTENT_BG)
        cards_row.pack(fill='x', pady=(0, 20))

        for col, (title, value, icon, color) in enumerate(stat_configs):
            outer = tk.Frame(cards_row, bg=BORDER_CLR)
            outer.grid(row=0, column=col, padx=8, pady=4, sticky='nsew')
            card = tk.Frame(outer, bg=CARD_BG, padx=20, pady=18)
            card.pack(fill='both', expand=True, padx=1, pady=1)

            # title row
            tr = tk.Frame(card, bg=CARD_BG)
            tr.pack(fill='x')
            tk.Label(tr, text=title, bg=CARD_BG, fg=TEXT_SECONDARY,
                     font=(FF, 11)).pack(side='left')
            badge = rounded_badge(tr, icon, color, size=40)
            badge.pack(side='right')

            # value
            tk.Label(card, text=value, bg=CARD_BG, fg=TEXT_PRIMARY,
                     font=(FF, 24, 'bold'), anchor='w').pack(fill='x', pady=(10, 0))

        for c in range(4):
            cards_row.columnconfigure(c, weight=1)

        # ---- merged: CBC Info (left) + Grade Scale (right) ----
        bottom_row = tk.Frame(self.content_frame, bg=CONTENT_BG)
        bottom_row.pack(fill='both', expand=True, pady=(8, 4))
        bottom_row.columnconfigure(0, weight=3)
        bottom_row.columnconfigure(1, weight=2)

        # ── CBC Info card (left) ──────────────────────────────────────────
        cbc_outer = tk.Frame(bottom_row, bg=BORDER_CLR)
        cbc_outer.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        cbc_card = tk.Frame(cbc_outer, bg=CARD_BG, padx=20, pady=18)
        cbc_card.pack(fill='both', expand=True, padx=1, pady=1)

        level = self.current_level
        _short = {
            'Lower Primary (Grade 1-3)': 'Lower Primary',
            'Upper Primary (Grade 4-6)': 'Upper Primary',
            'Junior School (Grade 7-9)': 'Junior School',
        }
        level_short = _short.get(level, level)

        # Title row with level badge
        cbc_hdr = tk.Frame(cbc_card, bg=CARD_BG)
        cbc_hdr.pack(fill='x', pady=(0, 12))
        tk.Label(cbc_hdr, text='📚  CBC Information', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 13, 'bold')).pack(side='left')
        tk.Label(cbc_hdr, text=level_short, bg=GREEN, fg='white',
                 font=(FF, 9, 'bold'), padx=10, pady=3).pack(side='left', padx=(10, 0))

        # Subjects grid (up to 8 chips, 2 columns)
        subjects_data = SUBJECTS_BY_LEVEL[level]
        subjects = subjects_data['core'] if isinstance(subjects_data, dict) else subjects_data
        subj_grid = tk.Frame(cbc_card, bg=CARD_BG)
        subj_grid.pack(fill='x', pady=(0, 10))
        for i, subj in enumerate(subjects[:8]):
            chip = tk.Label(subj_grid, text=f'▸  {subj}', bg='#e8f5e9', fg=GREEN,
                            font=(FF, 9), padx=8, pady=4)
            chip.grid(row=i // 2, column=i % 2, padx=4, pady=3, sticky='ew')
        subj_grid.columnconfigure(0, weight=1)
        subj_grid.columnconfigure(1, weight=1)
        if len(subjects) > 8:
            tk.Label(cbc_card, text=f'+ {len(subjects) - 8} more subjects — see CBC Info page',
                     bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 9, 'italic')).pack(anchor='w', pady=(0, 6))

        # Divider
        tk.Frame(cbc_card, bg=BORDER_CLR, height=1).pack(fill='x', pady=(4, 8))

        # Assessment note
        grading_info = GRADING_BY_LEVEL[level]
        if 'assessment_components' in grading_info:
            assess_text = grading_info['assessment_components']
        elif 'assessment_methods' in grading_info:
            assess_text = grading_info['assessment_methods']
        else:
            assess_text = 'Competency levels with percentage scores'
        tk.Label(cbc_card, text=f'📝  Assessment:  {assess_text}', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 9), wraplength=400, justify='left').pack(anchor='w')

        # ── Grade Scale card (right) ──────────────────────────────────────
        gs_outer = tk.Frame(bottom_row, bg=BORDER_CLR)
        gs_outer.grid(row=0, column=1, sticky='nsew', padx=(6, 0))
        gs_card = tk.Frame(gs_outer, bg=CARD_BG, padx=20, pady=18)
        gs_card.pack(fill='both', expand=True, padx=1, pady=1)

        tk.Label(gs_card, text='📊  Grade Scale', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 13, 'bold')).pack(anchor='w', pady=(0, 10))

        grade_data = [
            ('EE', '80 – 100', 'Exceeding Expectations',   GRADE_COLORS['EE']),
            ('ME', '70 – 79',  'Meeting Expectations',     GRADE_COLORS['ME']),
            ('AE', '60 – 69',  'Approaching Expectations', GRADE_COLORS['AE']),
            ('BE', '50 – 59',  'Below Expectations',       GRADE_COLORS['BE']),
            ('IE', '0 – 49',   'Inadequate',               GRADE_COLORS['IE']),
        ]
        for code, rng, desc, clr in grade_data:
            tile = tk.Frame(gs_card, bg=clr, padx=14, pady=9)
            tile.pack(fill='x', pady=2)
            tk.Label(tile, text=code, bg=clr, fg='white',
                     font=(FF, 13, 'bold'), width=4, anchor='w').pack(side='left')
            tk.Label(tile, text=rng, bg=clr, fg='white',
                     font=(FF, 9), width=10, anchor='w').pack(side='left')
            tk.Label(tile, text=desc, bg=clr, fg='white',
                     font=(FF, 9), anchor='w').pack(side='left', padx=(6, 0))

    # ==================== TEACHERS ====================
    def show_teachers(self):
        """Show teacher management page"""
        self.clear_frame()
        self._set_nav('Teachers')
        self._page_header('Teachers Management', 'Manage teachers and their subject/class assignments')
        
        # Toolbar
        toolbar = tk.Frame(self.content_frame, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=(0, 10))
        
        tk.Button(toolbar, text='+ Add Teacher', bg=GREEN, fg='white', 
                 font=(FF, 10), padx=15, pady=5, command=self._add_teacher_dialog).pack(side='left', padx=5)
        
        tk.Button(toolbar, text='Refresh', bg='#666', fg='white',
                 font=(FF, 10), padx=15, pady=5, command=self.show_teachers).pack(side='left', padx=5)
        
        # Teachers list
        teachers_frame = tk.Frame(self.content_frame, bg=CARD_BG, relief='flat', bd=1)
        teachers_frame.pack(fill='both', expand=True)
        
        # Treeview
        cols = ('id', 'name', 'username', 'role', 'assignments')
        tree = ttk.Treeview(teachers_frame, columns=cols, show='headings', style='App.Treeview')
        
        tree.heading('id', text='ID')
        tree.heading('name', text='Full Name')
        tree.heading('username', text='Username')
        tree.heading('role', text='Role')
        tree.heading('assignments', text='Assignments')
        
        tree.column('id', width=50)
        tree.column('name', width=150)
        tree.column('username', width=120)
        tree.column('role', width=100)
        tree.column('assignments', width=300)
        
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Load teachers
        teachers = db.get_all_teachers()
        for teacher in teachers:
            # Get assignments
            assignments = []
            subject_assignments = db.get_subject_teacher_assignments()
            class_assignments = db.get_class_teacher_assignments()
            
            for sa in subject_assignments:
                if sa['teacher_id'] == teacher['id']:
                    assignments.append(f"{sa['subject']} ({sa['class_name']})")
            
            for ca in class_assignments:
                if ca['teacher_id'] == teacher['id']:
                    assignments.append(f"Class Teacher: {ca['class_name']}")
            
            role_label = 'Subject Teacher' if teacher.get('role') == 'teacher' else 'Class Teacher'
            assignment_text = ', '.join(assignments) if assignments else 'No assignments'
            
            tree.insert('', 'end', values=(
                teacher.get('id', '')[:8],
                teacher.get('full_name', ''),
                teacher.get('username', ''),
                role_label,
                assignment_text
            ))
        
        # Action buttons
        action_frame = tk.Frame(teachers_frame, bg=CARD_BG)
        action_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        tk.Button(action_frame, text='Assign Subject', bg=BLUE, fg='white',
                 font=(FF, 10), padx=10, pady=5, command=self._assign_subject_dialog).pack(side='left', padx=5)
        
        tk.Button(action_frame, text='Assign Class Teacher', bg=PURPLE, fg='white',
                 font=(FF, 10), padx=10, pady=5, command=self._assign_class_teacher_dialog).pack(side='left', padx=5)
        
        tk.Button(action_frame, text='Delete', bg='#e74c3c', fg='white',
                 font=(FF, 10), padx=10, pady=5, command=lambda: self._delete_teacher(tree)).pack(side='left', padx=5)
    
    def _add_teacher_dialog(self):
        """Dialog to add a new teacher"""
        dialog = tk.Toplevel(self.root)
        dialog.title('Add Teacher')
        dialog.geometry('400x350')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Full Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Username:', font=(FF, 11)).pack(pady=(15, 5))
        username_entry = tk.Entry(dialog, font=(FF, 11))
        username_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Password:', font=(FF, 11)).pack(pady=(15, 5))
        password_entry = tk.Entry(dialog, font=(FF, 11), show='*')
        password_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Role:', font=(FF, 11)).pack(pady=(15, 5))
        role_var = tk.StringVar(value='teacher')
        role_frame = tk.Frame(dialog)
        role_frame.pack()
        tk.Radiobutton(role_frame, text='Subject Teacher', variable=role_var, 
                      value='teacher', font=(FF, 10)).pack(side='left', padx=10)
        tk.Radiobutton(role_frame, text='Class Teacher', variable=role_var, 
                      value='class_teacher', font=(FF, 10)).pack(side='left', padx=10)
        
        def save_teacher():
            name = name_entry.get().strip()
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            role = role_var.get()
            
            if not all([name, username, password]):
                messagebox.showerror('Error', 'All fields are required')
                return
            
            if len(password) < 4:
                messagebox.showerror('Error', 'Password must be at least 4 characters')
                return
            
            success, msg = db.add_teacher(name, username, password, role)
            if success:
                messagebox.showinfo('Success', msg)
                dialog.destroy()
                self.show_teachers()
            else:
                messagebox.showerror('Error', msg)
        
        tk.Button(dialog, text='Save Teacher', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save_teacher).pack(pady=25)
    
    def _assign_subject_dialog(self):
        """Dialog to assign a subject to a teacher"""
        teachers = db.get_all_teachers()
        if not teachers:
            messagebox.showwarning('No Teachers', 'Please add teachers first')
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title('Assign Subject to Teacher')
        dialog.geometry('400x300')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Select Teacher:', font=(FF, 11)).pack(pady=(20, 5))
        teacher_names = [f"{t['full_name']} ({t['username']})" for t in teachers]
        teacher_var = tk.StringVar()
        teacher_cb = ttk.Combobox(dialog, textvariable=teacher_var, values=teacher_names, 
                                  state='readonly', font=(FF, 10))
        teacher_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Select Class:', font=(FF, 11)).pack(pady=(15, 5))
        class_var = tk.StringVar()
        class_cb = ttk.Combobox(dialog, textvariable=class_var, values=CLASSES,
                               state='readonly', font=(FF, 10))
        class_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Select Subject:', font=(FF, 11)).pack(pady=(15, 5))
        subject_var = tk.StringVar()
        subject_cb = ttk.Combobox(dialog, textvariable=subject_var, values=SUBJECTS,
                                 state='readonly', font=(FF, 10))
        subject_cb.pack(fill='x', padx=20)
        
        def save_assignment():
            if not all([teacher_var.get(), class_var.get(), subject_var.get()]):
                messagebox.showerror('Error', 'All fields are required')
                return
            
            # Get selected teacher ID
            selected_idx = teacher_cb.current()
            teacher_id = teachers[selected_idx]['id']
            
            success = db.assign_subject_teacher(teacher_id, class_var.get(), subject_var.get())
            if success:
                messagebox.showinfo('Success', 'Subject assigned successfully!')
                dialog.destroy()
                self.show_teachers()
            else:
                messagebox.showerror('Error', 'Failed to assign subject')
        
        tk.Button(dialog, text='Assign Subject', bg=BLUE, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save_assignment).pack(pady=25)
    
    def _assign_class_teacher_dialog(self):
        """Dialog to assign a class teacher"""
        teachers = db.get_all_teachers()
        if not teachers:
            messagebox.showwarning('No Teachers', 'Please add teachers first')
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title('Assign Class Teacher')
        dialog.geometry('400x250')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Select Teacher:', font=(FF, 11)).pack(pady=(20, 5))
        teacher_names = [f"{t['full_name']} ({t['username']})" for t in teachers]
        teacher_var = tk.StringVar()
        teacher_cb = ttk.Combobox(dialog, textvariable=teacher_var, values=teacher_names,
                                  state='readonly', font=(FF, 10))
        teacher_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Select Class:', font=(FF, 11)).pack(pady=(15, 5))
        class_var = tk.StringVar()
        class_cb = ttk.Combobox(dialog, textvariable=class_var, values=CLASSES,
                               state='readonly', font=(FF, 10))
        class_cb.pack(fill='x', padx=20)
        
        def save_assignment():
            if not all([teacher_var.get(), class_var.get()]):
                messagebox.showerror('Error', 'All fields are required')
                return
            
            selected_idx = teacher_cb.current()
            teacher_id = teachers[selected_idx]['id']
            
            success = db.assign_class_teacher(teacher_id, class_var.get())
            if success:
                messagebox.showinfo('Success', 'Class teacher assigned successfully!')
                dialog.destroy()
                self.show_teachers()
            else:
                messagebox.showerror('Error', 'Failed to assign class teacher')
        
        tk.Button(dialog, text='Assign Class Teacher', bg=PURPLE, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save_assignment).pack(pady=25)
    
    def _delete_teacher(self, tree):
        """Delete selected teacher"""
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select Teacher', 'Please select a teacher to delete')
            return
        
        if not messagebox.askyesno('Confirm Delete', 'Are you sure you want to delete this teacher?'):
            return
        
        item = tree.item(selected)
        teacher_id = item['values'][0]
        
        if db.delete_user(teacher_id):
            messagebox.showinfo('Success', 'Teacher deleted successfully')
            self.show_teachers()
        else:
            messagebox.showerror('Error', 'Failed to delete teacher')

    # ==================== CLASS TEACHER VIEWS ====================
    def show_class_students(self):
        """Show class teacher their assigned students"""
        if not self.current_user:
            messagebox.showerror('Error', 'Please login first')
            return
        
        teacher_id = self.current_user.get('id')
        classes = db.get_teacher_classes(teacher_id)
        
        if not classes:
            messagebox.showwarning('No Class', 'You are not assigned as a class teacher for any class')
            return
        
        self.clear_frame()
        self._set_nav('My Students')
        self._page_header('My Students', f'Class Teacher for: {classes[0]}')
        
        # Class selector
        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 10))
        
        tk.Label(ctrl, text='Class:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=5)
        class_var = tk.StringVar(value=classes[0])
        class_cb = ttk.Combobox(ctrl, textvariable=class_var, values=classes, state='readonly', font=(FF, 10))
        class_cb.pack(side='left', padx=5)
        class_cb.bind('<<ComboboxSelected>>', lambda e: self._load_class_students(class_var.get()))
        
        # Students list
        list_frame = tk.Frame(self.content_frame, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True)
        
        cols = ('adm', 'name', 'gender')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        tree.heading('adm', text='Admission No')
        tree.heading('name', text='Name')
        tree.heading('gender', text='Gender')
        tree.column('adm', width=120)
        tree.column('name', width=200)
        tree.column('gender', width=80)
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.class_students_tree = tree
        self._load_class_students(classes[0])
    
    def _load_class_students(self, class_name):
        """Load students for a specific class"""
        tree = self.class_students_tree
        for item in tree.get_children():
            tree.delete(item)
        
        students = db.get_students_by_class(class_name)
        for s in students:
            tree.insert('', 'end', values=(
                s.get('admission_no', ''),
                s.get('name', ''),
                s.get('gender', '')
            ))
    
    def show_add_comments(self):
        """Show page for adding student comments"""
        if not self.current_user:
            messagebox.showerror('Error', 'Please login first')
            return
        
        teacher_id = self.current_user.get('id')
        classes = db.get_teacher_classes(teacher_id)
        
        if not classes:
            messagebox.showwarning('No Class', 'You are not assigned as a class teacher for any class')
            return
        
        self.clear_frame()
        self._set_nav('Add Comments')
        self._page_header('Add Comments', 'Add comments for students in your class')
        
        # Controls
        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 10))
        
        tk.Label(ctrl, text='Class:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=5)
        class_var = tk.StringVar(value=classes[0])
        class_cb = ttk.Combobox(ctrl, textvariable=class_var, values=classes, state='readonly', font=(FF, 10))
        class_cb.pack(side='left', padx=5)
        
        tk.Label(ctrl, text='Term:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=15)
        term_var = tk.StringVar(value='One')
        term_cb = ttk.Combobox(ctrl, textvariable=term_var, values=TERMS, state='readonly', font=(FF, 10))
        term_cb.pack(side='left', padx=5)
        
        tk.Button(ctrl, text='Load Students', bg=BLUE, fg='white', font=(FF, 10),
                 command=lambda: self._load_students_for_comments(class_var.get(), term_var.get())).pack(side='left', padx=15)
        
        # Students with comments
        list_frame = tk.Frame(self.content_frame, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True)
        
        cols = ('adm', 'name', 'comment')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings')
        tree.heading('adm', text='Admission No')
        tree.heading('name', text='Name')
        tree.heading('comment', text='Current Comment')
        tree.column('adm', width=120)
        tree.column('name', width=180)
        tree.column('comment', width=350)
        tree.pack(fill='both', expand=True, padx=10, pady=10)
        
        self.comments_tree = tree
        self.comments_class_var = class_var
        self.comments_term_var = term_var
    
    def _load_students_for_comments(self, class_name, term):
        """Load students for adding comments"""
        tree = self.comments_tree
        for item in tree.get_children():
            tree.delete(item)
        
        students = db.get_students_by_class(class_name)
        for s in students:
            # Get existing comment
            comment_data = db.get_student_comment(s['id'], term)
            comment = comment_data.get('comment_text', '') if comment_data else ''
            
            tree.insert('', 'end', values=(
                s.get('admission_no', ''),
                s.get('name', ''),
                comment
            ), tags=(s['id'],))
        
        # Add double-click to edit
        tree.bind('<Double-1>', lambda e: self._edit_comment(tree, class_name, term))
    
    def _edit_comment(self, tree, class_name, term):
        """Edit comment for selected student"""
        selected = tree.selection()
        if not selected:
            return
        
        item = tree.item(selected)
        student_id = item['tags'][0]
        current_comment = item['values'][2]
        
        # Dialog to edit comment
        dialog = tk.Toplevel(self.root)
        dialog.title('Add/Edit Comment')
        dialog.geometry('500x250')
        dialog.transient(self.root)
        dialog.grab_set()
        
        tk.Label(dialog, text='Comment:', font=(FF, 11)).pack(pady=(20, 5))
        comment_text = tk.Text(dialog, font=(FF, 11), height=8)
        comment_text.pack(fill='both', expand=True, padx=20, pady=10)
        comment_text.insert('1.0', current_comment)
        
        def save():
            text = comment_text.get('1.0', 'end').strip()
            if db.save_comment(student_id, self.current_user['id'], term, text):
                messagebox.showinfo('Success', 'Comment saved!')
                dialog.destroy()
                self._load_students_for_comments(class_name, term)
            else:
                messagebox.showerror('Error', 'Failed to save comment')
        
        tk.Button(dialog, text='Save Comment', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(pady=10)

    # ==================== STUDENTS ====================
    def show_students(self):
        self.clear_frame()
        self._set_nav('Students')
        self._page_header('Students', 'Manage student registrations')

        # toolbar
        tb = tk.Frame(self.content_frame, bg=CONTENT_BG)
        tb.pack(fill='x', pady=(0, 12))

        self._toolbar_btn(tb, '+ Add Student', self.add_student).pack(side='left')
        self._toolbar_btn(tb, '📥 Template', self.download_template, bg=ORANGE).pack(side='left', padx=10)
        self._toolbar_btn(tb, '📥 Import Excel', self.import_excel, bg=BLUE).pack(side='left')
        self._toolbar_btn(tb, '📤 Export Excel', self.export_excel, bg=PURPLE).pack(side='left', padx=10)

        tk.Label(tb, text='Search:', bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(side='left', padx=(20, 6))
        self.student_search = ttk.Entry(tb, style='App.TEntry', width=28)
        self.student_search.pack(side='left', ipady=4)
        self.student_search.bind('<KeyRelease>', lambda e: self.filter_students())

        # table card
        tc_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        tc_outer.pack(fill='both', expand=True, pady=4)
        tc = tk.Frame(tc_outer, bg=CARD_BG)
        tc.pack(fill='both', expand=True, padx=1, pady=1)

        cols = ('adm_no', 'name', 'class', 'gender')
        self.students_tree = ttk.Treeview(tc, columns=cols, show='headings',
                                          style='App.Treeview')
        for col, text, w in [('adm_no','Admission No',130), ('name','Name',220),
                              ('class','Class',110), ('gender','Gender',90)]:
            self.students_tree.heading(col, text=text)
            self.students_tree.column(col, width=w)

        sb = ttk.Scrollbar(tc, orient='vertical', command=self.students_tree.yview,
                           style='App.Vertical.TScrollbar')
        self.students_tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.students_tree.pack(fill='both', expand=True)

        self.students_tree.bind('<Double-Button-1>', lambda e: self.edit_student())
        self.students_tree.bind('<Button-3>', self._student_ctx)
        self.load_students()

    def load_students(self):
        for i in self.students_tree.get_children():
            self.students_tree.delete(i)
        for s in db.get_all_students():
            self.students_tree.insert('', 'end',
                values=(s['admission_no'], s['name'], s['class'], s['gender']),
                tags=(s['id'], s.get('photo_path', '')))

    def filter_students(self):
        q = self.student_search.get()
        for i in self.students_tree.get_children():
            self.students_tree.delete(i)
        rows = db.search_students(q) if q else db.get_all_students()
        for s in rows:
            self.students_tree.insert('', 'end',
                values=(s['admission_no'], s['name'], s['class'], s['gender']),
                tags=(s['id'], s.get('photo_path', '')))

    def _student_ctx(self, event):
        item = self.students_tree.identify('item', event.x, event.y)
        if item:
            self.students_tree.selection_set(item)
            m = tk.Menu(self.root, tearoff=0)
            m.add_command(label='Edit',   command=self.edit_student)
            m.add_command(label='Delete', command=self.delete_student)
            m.add_command(label='Generate PDF Report', command=self.generate_pdf_report)
            m.post(event.x_root, event.y_root)

    def _student_dialog(self, title, adm='', name='', cls=CLASSES[0], gender='Male', photo_path='',
                        on_save=None):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry('500x520')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        outer = tk.Frame(dlg, bg=BORDER_CLR)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=CARD_BG, padx=30, pady=24)
        card.pack(padx=1, pady=1)

        tk.Label(card, text=title, bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 14, 'bold')).pack(anchor='w', pady=(0, 20))

        # Photo selection
        photo_frame = tk.Frame(card, bg=CARD_BG)
        photo_frame.pack(fill='x', pady=(0, 15))
        
        self.temp_photo_path = photo_path
        photo_label = tk.Label(photo_frame, text="No Photo", bg='#f1f5f9', width=12, height=6)
        photo_label.pack(side='left')
        
        def update_photo_display(p):
            if p and os.path.exists(p):
                try:
                    img = Image.open(p)
                    img.thumbnail((100, 100))
                    photo_img = ImageTk.PhotoImage(img)
                    photo_label.config(image=photo_img, text="")
                    photo_label.image = photo_img
                except:
                    photo_label.config(image="", text="Error")
        
        update_photo_display(photo_path)
        
        def pick_photo():
            p = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.png *.jpeg")])
            if p:
                self.temp_photo_path = p
                update_photo_display(p)

        tk.Button(photo_frame, text="Select Photo", command=pick_photo, bg=BLUE, fg='white', relief='flat', padx=10).pack(side='left', padx=10)

        def lbl_entry(label_text, default=''):
            tk.Label(card, text=label_text, bg=CARD_BG, fg=TEXT_SECONDARY,
                     font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
            e = ttk.Entry(card, style='App.TEntry')
            e.insert(0, default)
            e.pack(fill='x', ipady=4, pady=(0, 10))
            return e

        adm_e  = lbl_entry('Admission No', adm)
        name_e = lbl_entry('Full Name', name)

        tk.Label(card, text='Class', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
        cls_cb = ttk.Combobox(card, values=self.get_current_classes(), state='readonly', style='App.TCombobox')
        cls_cb.set(cls)
        cls_cb.pack(fill='x', ipady=4, pady=(0, 10))

        tk.Label(card, text='Gender', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
        gen_cb = ttk.Combobox(card, values=['Male', 'Female'], state='readonly',
                              style='App.TCombobox')
        gen_cb.set(gender)
        gen_cb.pack(fill='x', ipady=4, pady=(0, 18))

        btn_row = tk.Frame(card, bg=CARD_BG)
        btn_row.pack(fill='x')
        cancel = tk.Label(btn_row, text='Cancel', bg='#e8f5e9', fg=TEXT_PRIMARY,
                          font=(FF, 10, 'bold'), padx=20, pady=8, cursor='hand2')
        cancel.pack(side='left', padx=(0, 8))
        cancel.bind('<Button-1>', lambda e: dlg.destroy())

        def do_save():
            if not adm_e.get().strip():
                messagebox.showerror('Error', 'Admission No is required', parent=dlg)
                return
            if not name_e.get().strip():
                messagebox.showerror('Error', 'Name is required', parent=dlg)
                return
            on_save(adm_e.get().strip(), name_e.get().strip(), cls_cb.get(), gen_cb.get(), self.temp_photo_path)
            dlg.destroy()

        save = tk.Label(btn_row, text='Save', bg=BLUE, fg='white',
                        font=(FF, 10, 'bold'), padx=20, pady=8, cursor='hand2')
        save.pack(side='left')
        save.bind('<Button-1>', lambda e: do_save())

    def add_student(self):
        def on_save(adm, name, cls, gender, photo):
            db.add_student(name, cls, gender, adm, photo)
            self.load_students()
        self._student_dialog('Add Student', on_save=on_save)

    def edit_student(self):
        sel = self.students_tree.selection()
        if not sel: return
        item = self.students_tree.item(sel[0])
        adm, name, cls, gender = item['values']
        sid = item['tags'][0]
        photo = item['tags'][1] if len(item['tags']) > 1 else ""
        def on_save(a, n, c, g, p):
            db.update_student(sid, n, c, g, a, p)
            self.load_students()
        self._student_dialog('Edit Student', adm, name, cls, gender, photo, on_save=on_save)

    def delete_student(self):
        sel = self.students_tree.selection()
        if not sel: return
        sid = self.students_tree.item(sel[0])['tags'][0]
        if messagebox.askyesno('Confirm Delete', 'Delete this student?', parent=self.root):
            db.delete_student(sid)
            self.load_students()

    def import_excel(self):
        """Import students from an Excel file"""
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        
        try:
            df = pd.read_excel(file_path)
            
            # Expected columns: admission_no, name, class, gender
            required_cols = ['admission_no', 'name', 'class']
            missing_cols = [col for col in required_cols if col not in df.columns]
            
            if missing_cols:
                messagebox.showerror("Error", f"Missing required columns: {', '.join(missing_cols)}")
                return
            
            imported_count = 0
            updated_count  = 0
            for _, row in df.iterrows():
                admission_no = str(row.get('admission_no', '')).strip()
                name         = str(row.get('name', '')).strip()
                cls          = str(row.get('class', 'Grade 7')).strip()
                gender       = str(row.get('gender', 'Male')).strip()
                photo_path   = str(row.get('photo_path', '')).strip()
                # pandas reads missing cells as 'nan'
                if photo_path.lower() == 'nan':
                    photo_path = ''

                if not name or not admission_no:
                    continue

                # Upsert: update existing student (preserving photo if not provided),
                # or add new student.
                existing = db.get_student_by_admission_no(admission_no)
                if existing:
                    db.update_student(existing['id'], name, cls, gender, admission_no, photo_path)
                    updated_count += 1
                else:
                    db.add_student(name, cls, gender, admission_no, photo_path)
                    imported_count += 1

            self.load_students()
            messagebox.showinfo("Import Complete",
                                f"Done!  {imported_count} new student(s) added, "
                                f"{updated_count} updated.\n"
                                f"Existing photos were preserved where no new photo was supplied.")
            
        except Exception as e:
            messagebox.showerror("Import Error", f"Failed to import Excel file: {str(e)}")

    def export_excel(self):
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
            students = db.get_all_students()
            df = pd.DataFrame(students)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Export Complete", f"Exported {len(students)} students to: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export: {str(e)}")

    def download_template(self):
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
            template_data = {
                'admission_no': ['001', '002', '003'],
                'name': ['John Doe', 'Jane Smith', 'Michael Johnson'],
                'class': ['Grade 7', 'Grade 8', 'Grade 9'],
                'gender': ['Male', 'Female', 'Male'],
                'photo_path': ['', '', '']
            }
            df = pd.DataFrame(template_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Template Created", f"Template saved to: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create template: {str(e)}")

    # ==================== ENTER MARKS ====================
    def show_marks_entry(self):
        self.clear_frame()
        self._set_nav('Enter Marks')
        self._page_header('Enter Marks', 'Enter subject marks for each student')

        # controls
        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 12))

        def lbl(text): tk.Label(ctrl, text=text, bg=CONTENT_BG, fg=TEXT_SECONDARY,
                                 font=(FF, 10)).pack(side='left', padx=(10, 4))

        lbl('Class:')
        self.marks_class_cb = ttk.Combobox(ctrl, values=self.get_current_classes(), state='readonly',
                                           style='App.TCombobox', width=12)
        self.marks_class_cb.set(CLASSES[0])
        self.marks_class_cb.pack(side='left', ipady=4)

        lbl('Term:')
        self.marks_term_cb = ttk.Combobox(ctrl, values=TERMS, state='readonly',
                                          style='App.TCombobox', width=10)
        self.marks_term_cb.set(TERMS[0])
        self.marks_term_cb.pack(side='left', ipady=4)

        self.marks_class_cb.bind('<<ComboboxSelected>>', lambda e: self._load_marks_table())
        self.marks_term_cb.bind('<<ComboboxSelected>>', lambda e: self._load_marks_table())

        self._toolbar_btn(ctrl, '\U0001f4be  Save All Marks', self.save_marks).pack(
            side='left', padx=16)
        self._toolbar_btn(ctrl, '\U0001f4cb  Template', self.download_marks_template,
                          bg=ORANGE).pack(side='left', padx=4)
        self._toolbar_btn(ctrl, '\U0001f4e5  Import Marks', self.import_marks_excel,
                          bg=PURPLE).pack(side='left', padx=4)
        
        def validate_mark(event, sid, sub):
            val = event.widget.get().strip()
            if val == '': return True
            try:
                num = int(val)
                if num < 0 or num > 100:
                    event.widget.delete(0, tk.END)
                    event.widget.insert(0, '0')
                return True
            except ValueError:
                event.widget.delete(0, tk.END)
                return False
        
        self._validate_mark = validate_mark

        # editable marks grid card
        tc_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        tc_outer.pack(fill='both', expand=True, pady=4)
        self.marks_card = tk.Frame(tc_outer, bg=CARD_BG, padx=1, pady=1)
        self.marks_card.pack(fill='both', expand=True)

        # scrollable marks grid
        self.marks_scroll_canvas, marks_scroll_sb, self.marks_inner = scrollable_frame(self.marks_card, CARD_BG)
        
        self.marks_entries = {}  # sid -> {subject: Entry}
        self.student_widgets = []  # for double-click
        
        self._load_marks_table()

    def _load_marks_table(self):
        # clear existing
        for w in self.marks_inner.winfo_children():
            w.destroy()
        self.marks_entries.clear()
        self.student_widgets.clear()
        
        cls = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        students = db.get_students_by_class(cls)
        
        if not students:
            tk.Label(self.marks_inner, text='No students in this class', 
                     bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 12)).pack(pady=40)
            return
        
        # header row
        hdr = tk.Frame(self.marks_inner, bg=CARD_BG)
        hdr.pack(fill='x', pady=(0, 8))
        tk.Label(hdr, text='Student', bg=CARD_BG, fg=TEXT_PRIMARY, 
                 font=(FF, 11, 'bold'), width=20, anchor='w').pack(side='left', padx=12)
        for sub in self.get_current_subjects():
            tk.Label(hdr, text=sub, bg=CARD_BG, fg=TEXT_PRIMARY, 
                     font=(FF, 10, 'bold'), width=9, anchor='center').pack(side='left', padx=2)
        
        # student rows
        for s in students:
            sid = s['id']
            row_frame = tk.Frame(self.marks_inner, bg=CARD_BG)
            row_frame.pack(fill='x', pady=2)
            
            # student name (clickable)
            name_btn = tk.Label(row_frame, text=s['name'], bg=CARD_BG, fg=TEXT_PRIMARY,
                               font=(FF, 10, 'bold'), width=20, anchor='w', cursor='hand2',
                               padx=12, pady=8)
            name_btn.pack(side='left')
            name_btn.bind('<Button-1>', lambda e, sid=sid: self._edit_student_marks(sid))
            self.student_widgets.append(name_btn)
            
            # mark entries
            self.marks_entries[sid] = {}
            m = db.get_student_marks(sid, term)
            for sub in self.get_current_subjects():
                e_frame = tk.Frame(row_frame, bg=CARD_BG, width=70)
                e_frame.pack(side='left', padx=1)
                e_frame.pack_propagate(False)
                
                e = ttk.Entry(e_frame, style='App.TEntry', width=6, justify='center')
                e.pack(pady=4, padx=4)
                val = m.get(sub, '')
                e.insert(0, str(val) if val else '')
                e.bind('<KeyRelease>', lambda ev, s=sub, sid=sid: self._validate_mark(ev.widget, sid, s))
                
                self.marks_entries[sid][sub] = e
            
            # alternating row color
            if len(self.student_widgets) % 2 == 0:
                for child in row_frame.winfo_children():
                    child.config(bg='#f8fafc')

    def _edit_student_marks(self, sid=None):
        if sid is None:
            # Fallback for old Treeview binding
            sel = getattr(self, 'marks_tree', None) and self.marks_tree.selection()
            if sel:
                item = self.marks_tree.item(sel[0])
                sid = item['tags'][0]
        if not sid: return
        
        students = db.get_students_by_class(self.marks_class_cb.get())
        name = next((s['name'] for s in students if s['id'] == sid), 'Unknown')
        term = self.marks_term_cb.get()
        cur = db.get_student_marks(sid, term)

        dlg = tk.Toplevel(self.root)
        dlg.title(f'Marks – {name}')
        dlg.geometry('460x420')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        outer = tk.Frame(dlg, bg=BORDER_CLR)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=CARD_BG, padx=30, pady=22)
        card.pack(padx=1, pady=1)

        tk.Label(card, text=f'Marks for {name}  (Term {term})',
                 bg=CARD_BG, fg=TEXT_PRIMARY, font=(FF, 13, 'bold')).pack(anchor='w', pady=(0, 16))

        entries: dict = {}
        grid = tk.Frame(card, bg=CARD_BG)
        grid.pack(fill='x')
        for i, sub in enumerate(self.get_current_subjects()):
            r, c = divmod(i, 3)
            tk.Label(grid, text=sub, bg=CARD_BG, fg=TEXT_SECONDARY,
                     font=(FF, 10, 'bold'), anchor='w', width=9).grid(
                         row=r*2, column=c, padx=6, pady=(6, 2), sticky='w')
            e = ttk.Entry(grid, style='App.TEntry', width=8)
            e.insert(0, cur.get(sub, ''))
            e.grid(row=r*2+1, column=c, padx=6, pady=(0, 6), ipady=4, sticky='ew')
            entries[sub] = e
        for c in range(3): grid.columnconfigure(c, weight=1)

        btn_row = tk.Frame(card, bg=CARD_BG)
        btn_row.pack(fill='x', pady=(14, 0))
        ca = tk.Label(btn_row, text='Cancel', bg='#e8f5e9', fg=TEXT_PRIMARY,
                      font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        ca.pack(side='left', padx=(0, 8))
        ca.bind('<Button-1>', lambda e: dlg.destroy())

        def do_save():
            marks = {}
            for sub, ent in entries.items():
                val = ent.get().strip()
                if val:
                    try: marks[sub] = min(100, max(0, int(val)))
                    except ValueError: pass
            db.save_student_marks(sid, marks, term)
            self._load_marks_table()
            dlg.destroy()

        sv = tk.Label(btn_row, text='Save Marks', bg=BLUE, fg='white',
                      font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        sv.pack(side='left')
        sv.bind('<Button-1>', lambda e: do_save())

    def save_marks(self):
        cls = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        
        saved_count = 0
        for sid, subject_entries in self.marks_entries.items():
            marks = {}
            for sub, entry in subject_entries.items():
                val = entry.get().strip()
                if val:
                    try:
                        mark = min(100, max(0, int(val)))
                        marks[sub] = mark
                        saved_count += 1
                    except ValueError:
                        pass
            if marks:
                db.save_student_marks(sid, marks, term)
        
        messagebox.showinfo('Success', f'Marks saved successfully! ({saved_count} values updated)')

    # ── Marks import / template helpers ──────────────────────────────────────

    def download_marks_template(self):
        """Export an Excel template for the current class & term pre-filled with
        student names / admission numbers and blank subject columns."""
        cls  = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        subjects = self.get_current_subjects()
        students = db.get_students_by_class(cls)

        file_path = filedialog.asksaveasfilename(
            title='Save Marks Template',
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f"marks_template_{cls.replace(' ', '_')}_T{term}.xlsx"
        )
        if not file_path:
            return

        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Marks'

            HDR_FILL = PatternFill('solid', fgColor='2E7D32')
            HDR_FONT = Font(bold=True, color='FFFFFF', name='Calibri', size=10)
            CTR = Alignment(horizontal='center', vertical='center')

            # Header row: admission_no | name | subject1 | subject2 | ...
            headers = ['admission_no', 'name'] + subjects
            for ci, h in enumerate(headers, 1):
                cell = ws.cell(row=1, column=ci, value=h)
                cell.fill = HDR_FILL
                cell.font = HDR_FONT
                cell.alignment = CTR

            # Column widths
            ws.column_dimensions['A'].width = 15
            ws.column_dimensions['B'].width = 26
            for i in range(len(subjects)):
                from openpyxl.utils import get_column_letter
                ws.column_dimensions[get_column_letter(3 + i)].width = 11

            # Pre-fill student rows with existing marks where available
            for ri, s in enumerate(students, 2):
                existing = db.get_student_marks(s['id'], term)
                ws.cell(row=ri, column=1, value=s['admission_no'])
                ws.cell(row=ri, column=2, value=s['name'])
                for ci, subj in enumerate(subjects, 3):
                    v = existing.get(subj)
                    ws.cell(row=ri, column=ci, value=v if v else '')

            wb.save(file_path)
            messagebox.showinfo(
                'Template Saved',
                f'Template saved to:\n{file_path}\n\n'
                f'Fill in the mark columns then use "Import Marks" to load them.'
            )
        except Exception as exc:
            messagebox.showerror('Error', f'Could not create template:\n{exc}')

    def import_marks_excel(self):
        """Import marks from an Excel file into the current class & term.

        Both 'admission_no' and 'name' columns are optional — at least one
        identifier is enough. Column names are matched case-insensitively and
        common aliases are accepted.
        """
        cls      = self.marks_class_cb.get()
        term     = self.marks_term_cb.get()
        subjects = self.get_current_subjects()

        file_path = filedialog.askopenfilename(
            title='Select Marks File',
            filetypes=[('Excel files', '*.xlsx *.xls')]
        )
        if not file_path:
            return

        try:
            df = pd.read_excel(file_path)

            # ── Normalise all column names to lowercase-stripped for matching ──
            orig_cols = list(df.columns)
            df.columns = [str(c).strip().lower() for c in df.columns]

            # Accepted aliases for the two identifier fields
            ADM_ALIASES  = {'admission_no', 'adm_no', 'admission no', 'adm no',
                            'admission', 'adm', 'reg_no', 'reg no', 'regno'}
            NAME_ALIASES = {'name', 'student_name', 'student name', 'learner',
                            'learner name', 'full_name', 'fullname', 'pupil'}

            adm_col  = next((c for c in df.columns if c in ADM_ALIASES),  None)
            name_col = next((c for c in df.columns if c in NAME_ALIASES), None)

            if not adm_col and not name_col:
                messagebox.showerror(
                    'No Identifier Column Found',
                    'The file needs at least one column to identify students.\n\n'
                    'Accepted column names for admission number:\n'
                    '  admission_no, adm_no, admission, reg_no …\n\n'
                    'Accepted column names for student name:\n'
                    '  name, learner, student_name, full_name …\n\n'
                    'Use the Template button to get a ready-made file.'
                )
                return

            # ── Build subject-column lookup (case-insensitive) ─────────────────
            # Map lowercase subject name → actual df column name
            subj_col_map = {}
            for subj in subjects:
                subj_lower = subj.strip().lower()
                if subj_lower in df.columns:
                    subj_col_map[subj] = subj_lower

            if not subj_col_map:
                messagebox.showerror(
                    'No Subject Columns Found',
                    f'No subject columns were found in the file.\n\n'
                    f'Expected columns (any capitalisation):\n'
                    + ', '.join(subjects)
                )
                return

            # ── Student lookup tables ──────────────────────────────────────────
            students = db.get_students_by_class(cls)
            adm_to_sid  = {s['admission_no'].strip(): s['id'] for s in students}
            name_to_sid = {s['name'].strip().lower():  s['id'] for s in students}

            updated   = 0
            skipped   = 0
            not_found = []

            for _, row in df.iterrows():
                # Extract identifiers (treat pandas 'nan' as empty)
                def _clean(val):
                    v = str(val).strip()
                    return '' if v.lower() == 'nan' else v

                adm      = _clean(row[adm_col])  if adm_col  else ''
                name_key = _clean(row[name_col]).lower() if name_col else ''

                # Resolve: admission_no first, name as fallback
                sid = adm_to_sid.get(adm) or name_to_sid.get(name_key)
                if not sid:
                    label = adm or name_key
                    if label:
                        not_found.append(label)
                    continue

                # Collect valid marks from subject columns
                marks = {}
                for subj, col in subj_col_map.items():
                    raw = row.get(col)
                    if raw is not None and str(raw).strip() not in ('', 'nan'):
                        try:
                            marks[subj] = min(100, max(0, int(float(raw))))
                        except (ValueError, TypeError):
                            pass

                if marks:
                    db.save_student_marks(sid, marks, term)
                    updated += 1
                else:
                    skipped += 1

            # Refresh on-screen grid
            self._load_marks_table()

            used_id = f'"{adm_col}"' if adm_col else f'"{name_col}"'
            msg = f'Import complete!  (matched by {used_id})\n\n✅ {updated} student(s) updated.'
            if skipped:
                msg += f'\n⚠️  {skipped} row(s) had no valid mark values (skipped).'
            if not_found:
                msg += (f'\n❌ {len(not_found)} identifier(s) not found in {cls}:\n   '
                        + ', '.join(not_found[:10]))
                if len(not_found) > 10:
                    msg += f'  … and {len(not_found) - 10} more'
            messagebox.showinfo('Import Complete', msg)

        except Exception as exc:
            messagebox.showerror('Import Error', f'Failed to import marks:\n{exc}')

    # ==================== REPORTS ====================
    def show_reports(self):
        self.clear_frame()
        self._set_nav('Reports')
        self._page_header('Reports & Rankings', 'View student performance and rankings')

        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 12))

        def lbl(t): tk.Label(ctrl, text=t, bg=CONTENT_BG, fg=TEXT_SECONDARY,
                              font=(FF, 10)).pack(side='left', padx=(10, 4))

        lbl('Class:')
        self.rep_cls_cb = ttk.Combobox(ctrl, values=['All'] + self.get_current_classes(), state='readonly',
                                       style='App.TCombobox', width=12)
        self.rep_cls_cb.set('All')
        self.rep_cls_cb.pack(side='left', ipady=4)

        lbl('Term:')
        self.rep_term_cb = ttk.Combobox(ctrl, values=TERMS, state='readonly',
                                        style='App.TCombobox', width=10)
        self.rep_term_cb.set(TERMS[0])
        self.rep_term_cb.pack(side='left', ipady=4)

        self.rep_cls_cb.bind('<<ComboboxSelected>>',  lambda e: self.load_reports())
        self.rep_term_cb.bind('<<ComboboxSelected>>', lambda e: self.load_reports())

        self._toolbar_btn(ctrl, '\u2193  Export CSV', self.export_csv,
                          bg='#475569').pack(side='left', padx=16)
        self._toolbar_btn(ctrl, '\U0001F4CA  Spotlight Excel', self.export_spotlight_excel,
                          bg='#1B5E20').pack(side='left', padx=4)

        # subject averages strip
        subj_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        subj_outer.pack(fill='x', pady=(0, 10))
        subj_card = tk.Frame(subj_outer, bg=CARD_BG, padx=16, pady=14)
        subj_card.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Label(subj_card, text='Subject Averages', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', pady=(0, 8))
        self.subj_row = tk.Frame(subj_card, bg=CARD_BG)
        self.subj_row.pack(fill='x')

        # rankings table
        tc_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        tc_outer.pack(fill='both', expand=True, pady=4)
        tc = tk.Frame(tc_outer, bg=CARD_BG)
        tc.pack(fill='both', expand=True, padx=1, pady=1)

        cols = ['pos', 'name', 'class'] + self.get_current_subjects() + ['total', 'avg', 'grade']
        self.rep_tree = ttk.Treeview(tc, columns=cols, show='headings', style='App.Treeview')

        spec = [('pos','Pos',45,'center'), ('name','Name',170,'w'), ('class','Class',90,'center')]
        for col, txt, w, anchor in spec:
            self.rep_tree.heading(col, text=txt)
            self.rep_tree.column(col, width=w, anchor=anchor)
        for s in self.get_current_subjects():
            self.rep_tree.heading(s, text=s)
            self.rep_tree.column(s, width=54, anchor='center')
        for col, txt, w in [('total','Total',65), ('avg','Avg',65), ('grade','Grade',60)]:
            self.rep_tree.heading(col, text=txt)
            self.rep_tree.column(col, width=w, anchor='center')

        sb = ttk.Scrollbar(tc, orient='vertical', command=self.rep_tree.yview,
                           style='App.Vertical.TScrollbar')
        self.rep_tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.rep_tree.pack(fill='both', expand=True)

        self.load_reports()

    def load_reports(self):
        for i in self.rep_tree.get_children():
            self.rep_tree.delete(i)
        for w in self.subj_row.winfo_children():
            w.destroy()

        cls  = self.rep_cls_cb.get()
        term = self.rep_term_cb.get()
        results = db.calculate_results(cls, term)

        # subject averages
        subj_totals = {s: [] for s in self.get_current_subjects()}
        for r in results:
            for s in self.get_current_subjects():
                if r['marks'].get(s): subj_totals[s].append(r['marks'][s])

        for s in self.get_current_subjects():
            vals = subj_totals[s]
            avg  = round(sum(vals) / len(vals), 1) if vals else 0
            grade = 'EE' if avg >= 80 else 'ME' if avg >= 70 else 'AE' if avg >= 60 else 'BE' if avg >= 50 else 'IE'
            clr   = GRADE_COLORS[grade]
            tile  = tk.Frame(self.subj_row, bg=clr, padx=10, pady=8)
            tile.pack(side='left', padx=3, expand=True, fill='both')
            tk.Label(tile, text=s,    bg=clr, fg='white', font=(FF, 9, 'bold')).pack()
            tk.Label(tile, text=str(avg), bg=clr, fg='white', font=(FF, 13, 'bold')).pack()
            tk.Label(tile, text=grade, bg=clr, fg='white', font=(FF, 8)).pack()

        # rows
        for r in results:
            vals = [r['position'], r['student']['name'], r['student']['class']]
            for s in self.get_current_subjects(): vals.append(r['marks'].get(s, '—'))
            vals += [r['total'], r['average'], r['grade']]
            self.rep_tree.insert('', 'end', values=vals)

    def export_csv(self):
        cls  = self.rep_cls_cb.get()
        term = self.rep_term_cb.get()
        results = db.calculate_results(cls, term)
        fn = f"report_{cls.replace(' ', '_')}_term_{term}.csv"
        with open(fn, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['Position', 'Adm No', 'Name', 'Class'] + self.get_current_subjects() + ['Total', 'Average', 'Grade'])
            for r in results:
                row = [r['position'], r['student']['admission_no'],
                       r['student']['name'], r['student']['class']]
                row += [r['marks'].get(s, 0) for s in self.get_current_subjects()]
                row += [r['total'], r['average'], r['grade']]
                w.writerow(row)
        messagebox.showinfo('Exported', f'Report saved to {fn}')

    # ==================== WESTERN SPOTLIGHT EXPORT ====================
    def export_spotlight_excel(self):
        """Open a dialog to configure and export the Western Spotlight class report."""
        # Pre-populate from Reports page if available
        try:
            initial_cls  = self.rep_cls_cb.get()
            if initial_cls == 'All':
                initial_cls = CLASSES[0]
            initial_term = self.rep_term_cb.get()
        except AttributeError:
            initial_cls  = CLASSES[0]
            initial_term = TERMS[0]

        dlg = tk.Toplevel(self.root)
        dlg.title('Export – Western Spotlight')
        dlg.geometry('400x340')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        outer = tk.Frame(dlg, bg=BORDER_CLR)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=CARD_BG, padx=30, pady=24)
        card.pack(padx=1, pady=1)
        card.columnconfigure(1, weight=1)

        tk.Label(card, text='Western Spotlight Export', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 13, 'bold')).grid(row=0, column=0, columnspan=2,
                                              sticky='w', pady=(0, 18))

        def row_field(r, label, widget):
            tk.Label(card, text=label, bg=CARD_BG, fg=TEXT_SECONDARY,
                     font=(FF, 10, 'bold'), anchor='w').grid(
                         row=r, column=0, sticky='w', padx=(0, 12), pady=5)
            widget.grid(row=r, column=1, sticky='ew', pady=5, ipady=4)

        cls_cb = ttk.Combobox(card, values=self.get_current_classes(),
                               state='readonly', style='App.TCombobox', width=20)
        cls_cb.set(initial_cls)
        row_field(1, 'Class:', cls_cb)

        term_cb = ttk.Combobox(card, values=TERMS, state='readonly',
                                style='App.TCombobox', width=20)
        term_cb.set(initial_term)
        row_field(2, 'Term:', term_cb)

        stream_var = tk.StringVar(value='GREEN')
        stream_e = ttk.Entry(card, textvariable=stream_var,
                              style='App.TEntry', width=20)
        row_field(3, 'Stream / Group:', stream_e)

        assess_cb = ttk.Combobox(card, values=['MID-TERM', 'END-TERM'],
                                  state='readonly', style='App.TCombobox', width=20)
        assess_cb.set('MID-TERM')
        row_field(4, 'Assessment:', assess_cb)

        year_var = tk.StringVar(value=str(datetime.now().year))
        year_e = ttk.Entry(card, textvariable=year_var,
                            style='App.TEntry', width=20)
        row_field(5, 'Year:', year_e)

        btn_row = tk.Frame(card, bg=CARD_BG)
        btn_row.grid(row=6, column=0, columnspan=2, pady=(20, 0))

        cancel = tk.Label(btn_row, text='Cancel', bg='#e8f5e9', fg=TEXT_PRIMARY,
                          font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        cancel.pack(side='left', padx=(0, 8))
        cancel.bind('<Button-1>', lambda e: dlg.destroy())

        def do_export():
            selected_cls    = cls_cb.get()
            selected_term   = term_cb.get()
            selected_stream = stream_var.get().strip() or 'GREEN'
            selected_assess = assess_cb.get()
            selected_year   = year_var.get().strip() or str(datetime.now().year)
            dlg.destroy()
            self._do_spotlight_export(selected_cls, selected_term,
                                      selected_stream, selected_assess, selected_year)

        export_btn = tk.Label(btn_row, text='Export Excel', bg='#1B5E20', fg='white',
                              font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        export_btn.pack(side='left')
        export_btn.bind('<Button-1>', lambda e: do_export())

    def _do_spotlight_export(self, cls, term, stream, assess, year):
        """Generate the Western Spotlight Excel workbook and save it."""
        results = db.calculate_results(cls, term)
        if not results:
            messagebox.showwarning('No Data',
                                   f'No results found for {cls}, Term {term}.\n'
                                   'Please enter marks first.')
            return

        file_path = filedialog.asksaveasfilename(
            title='Save – Western Spotlight',
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f"spotlight_{cls.replace(' ', '_')}_T{term}_{year}.xlsx"
        )
        if not file_path:
            return

        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = 'Spotlight'
            ws.sheet_view.showGridLines = False

            # ── Palette ──────────────────────────────────────────────────────
            BG_TITLE  = PatternFill('solid', fgColor='1B5E20')   # dark green – titles
            BG_HDR    = PatternFill('solid', fgColor='2E7D32')   # medium green – headers
            BG_DATA   = PatternFill('solid', fgColor='4CAF50')   # bright green – data rows
            FNT_YLW   = Font(bold=True, color='FFFF00', size=12, name='Calibri')
            FNT_YLW_S = Font(bold=True, color='FFFF00', size=10, name='Calibri')
            FNT_WHT   = Font(bold=True, color='FFFFFF', size=11, name='Calibri')
            FNT_WHT_S = Font(bold=True, color='FFFFFF', size=10, name='Calibri')
            ALIGN_CTR = Alignment(horizontal='center', vertical='center', wrap_text=True)
            ALIGN_LFT = Alignment(horizontal='left',   vertical='center')
            _thin     = Side(style='thin', color='000000')
            BORDER    = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

            subjects   = self.get_current_subjects()
            SUBJ_LABEL = {
                'Math': 'MATH', 'Eng': 'ENG', 'Kis': 'KIS',
                'Int Sci': 'INT. SCI', 'Agri': 'AGRI', 'SST': 'SST',
                'CRE': 'CRE', 'CIA': 'C/A', 'Pre-Tech': 'PRE-TECH'
            }
            n_subj    = len(subjects)
            total_max = n_subj * 100

            # Column numbers (1-based)
            C_NO    = 1
            C_NAME  = 2
            C_S0    = 3                    # first subject score col
            C_TOTAL = C_S0 + n_subj * 2
            C_AVG   = C_TOTAL + 1
            C_PSN   = C_AVG + 1
            C_LAST  = C_PSN

            def cl(c):
                return get_column_letter(c)

            def apply_style(r, c, fill=None, font=None, align=ALIGN_CTR, border=None):
                cell = ws.cell(row=r, column=c)
                if fill:   cell.fill      = fill
                if font:   cell.font      = font
                if align:  cell.alignment = align
                if border: cell.border    = border
                return cell

            def fill_entire_row(r, fill, font, height=None):
                for c in range(1, C_LAST + 1):
                    apply_style(r, c, fill=fill, font=font)
                if height:
                    ws.row_dimensions[r].height = height

            def merged_cell(r1, c1, r2, c2, value='', fill=None, font=None, align=ALIGN_CTR):
                ws.merge_cells(start_row=r1, start_column=c1,
                               end_row=r2,   end_column=c2)
                cell = ws.cell(row=r1, column=c1)
                cell.value     = value
                if fill:  cell.fill      = fill
                if font:  cell.font      = font
                if align: cell.alignment = align
                return cell

            # ── Title section (rows 1-5) ─────────────────────────────────────
            grade_num   = cls.replace('Grade ', '')
            grade_words = {'1':'ONE','2':'TWO','3':'THREE','4':'FOUR','5':'FIVE',
                           '6':'SIX','7':'SEVEN','8':'EIGHT','9':'NINE'}.get(grade_num, grade_num)

            fill_entire_row(1, BG_TITLE, FNT_YLW, height=26)
            merged_cell(1, C_NO, 1, C_LAST,
                        'MT.  OLIVES ADVENTIST SCHOOL,  NGONG',
                        BG_TITLE, FNT_YLW, ALIGN_CTR)

            fill_entire_row(2, BG_TITLE, FNT_YLW, height=8)
            merged_cell(2, C_NO, 2, C_LAST, '', BG_TITLE, FNT_YLW)

            title3 = (f'GRADE {grade_words} ({grade_num}) {stream} '
                      f'TERM {term.upper()} {assess} ASSESSMENT REPORT {year}')
            fill_entire_row(3, BG_TITLE, FNT_YLW, height=22)
            merged_cell(3, C_NO, 3, C_LAST, title3, BG_TITLE, FNT_YLW, ALIGN_CTR)

            fill_entire_row(4, BG_TITLE, FNT_YLW, height=8)
            merged_cell(4, C_NO, 4, C_LAST, '', BG_TITLE, FNT_YLW)

            fill_entire_row(5, BG_TITLE, FNT_YLW, height=22)
            merged_cell(5, C_NO, 5, C_LAST,
                        'THE WESTERN SPOTLIGHT', BG_TITLE, FNT_YLW, ALIGN_CTR)

            # ── Column header rows (6 & 7) ───────────────────────────────────
            ws.row_dimensions[6].height = 20
            ws.row_dimensions[7].height = 18
            for c in range(1, C_LAST + 1):
                for r in (6, 7):
                    apply_style(r, c, fill=BG_HDR, font=FNT_WHT,
                                align=ALIGN_CTR, border=BORDER)

            # NO.  – merged rows 6-7
            merged_cell(6, C_NO, 7, C_NO, 'NO.', BG_HDR, FNT_WHT, ALIGN_CTR)
            for r in (6, 7):
                ws.cell(r, C_NO).border = BORDER

            # LEARNER – merged rows 6-7
            merged_cell(6, C_NAME, 7, C_NAME, 'LEARNER', BG_HDR, FNT_WHT, ALIGN_CTR)
            for r in (6, 7):
                ws.cell(r, C_NAME).border = BORDER

            # Subject names row 6 (merged score+grade cols), row 7 sub-labels
            for i, subj in enumerate(subjects):
                sc  = C_S0 + i * 2      # score col
                gc  = sc + 1             # grade col
                lbl = SUBJ_LABEL.get(subj, subj.upper())
                merged_cell(6, sc, 6, gc, lbl, BG_HDR, FNT_WHT, ALIGN_CTR)
                for c in (sc, gc):
                    ws.cell(6, c).border = BORDER
                # Row 7: '100' under score, blank under grade
                c7 = ws.cell(7, sc)
                c7.value = '100'; c7.fill = BG_HDR; c7.font = FNT_WHT_S
                c7.alignment = ALIGN_CTR; c7.border = BORDER
                ws.cell(7, gc).border = BORDER

            # TOTAL
            ws.cell(6, C_TOTAL).value = 'TOTAL'
            ws.cell(7, C_TOTAL).value = str(total_max)
            for r in (6, 7):
                apply_style(r, C_TOTAL, fill=BG_HDR, font=FNT_WHT,
                            align=ALIGN_CTR, border=BORDER)

            # AVERAGE
            ws.cell(6, C_AVG).value = 'AVERAGE'
            ws.cell(7, C_AVG).value = '100%'
            for r in (6, 7):
                apply_style(r, C_AVG, fill=BG_HDR, font=FNT_WHT,
                            align=ALIGN_CTR, border=BORDER)

            # PSN – merged rows 6-7
            merged_cell(6, C_PSN, 7, C_PSN, 'PSN', BG_HDR, FNT_WHT, ALIGN_CTR)
            for r in (6, 7):
                ws.cell(r, C_PSN).border = BORDER

            # ── Data rows ────────────────────────────────────────────────────
            for idx, result in enumerate(results):
                r  = 8 + idx
                ws.row_dimensions[r].height = 16
                s  = result['student']
                mk = result['marks']

                def dc(col, value, align=ALIGN_CTR):
                    cell = ws.cell(row=r, column=col)
                    cell.value     = value
                    cell.fill      = BG_DATA
                    cell.font      = FNT_YLW_S
                    cell.alignment = align
                    cell.border    = BORDER

                dc(C_NO,   result['position'])
                dc(C_NAME, s['name'].upper(), ALIGN_LFT)

                for i, subj in enumerate(subjects):
                    sc       = C_S0 + i * 2
                    gc       = sc + 1
                    raw      = mk.get(subj)
                    mark_val = int(raw) if raw else 0
                    dc(sc, mark_val if raw else '')
                    dc(gc, get_cbc_grade_sublevel(mark_val) if raw else '')

                dc(C_TOTAL, result['total'])
                dc(C_AVG,   result['average'])
                dc(C_PSN,   result['position'])

            # ── Column widths ─────────────────────────────────────────────────
            ws.column_dimensions[cl(C_NO)].width   = 5
            ws.column_dimensions[cl(C_NAME)].width = 24
            for i in range(n_subj):
                sc = C_S0 + i * 2
                ws.column_dimensions[cl(sc)].width     = 5.5
                ws.column_dimensions[cl(sc + 1)].width = 5.5
            ws.column_dimensions[cl(C_TOTAL)].width = 7
            ws.column_dimensions[cl(C_AVG)].width   = 9
            ws.column_dimensions[cl(C_PSN)].width   = 5

            # ── Print settings (A4 landscape) ─────────────────────────────────
            ws.page_setup.orientation = 'landscape'
            ws.page_setup.paperSize   = 9     # A4
            ws.page_setup.fitToPage   = True
            ws.page_setup.fitToWidth  = 1
            ws.page_setup.fitToHeight = 0

            wb.save(file_path)
            messagebox.showinfo('Export Complete',
                                f'Western Spotlight report saved:\n{file_path}')

        except ImportError:
            messagebox.showerror('Missing Library',
                                 'openpyxl is required. Run:\n  pip install openpyxl')
        except Exception as exc:
            messagebox.showerror('Export Error', f'Failed to generate file:\n{exc}')

    def generate_pdf_report(self, student_id=None):
        """Generate a PDF report card for a student"""
        if not student_id:
            sel = self.students_tree.selection()
            if not sel:
                messagebox.showwarning("Warning", "Please select a student first!")
                return
            item = self.students_tree.item(sel[0])
            student_id = item["tags"][0]
        
        # Get student and their marks
        cls = "Grade 7" # Default or get from current selection
        term = "One"    # Default or get from current selection
        
        results = db.calculate_results(cls, term)
        result = next((r for r in results if r['student']['id'] == student_id), None)
        
        if not result:
            # Try other classes if not found
            for c in CLASSES:
                if c == cls: continue
                results = db.calculate_results(c, term)
                result = next((r for r in results if r['student']['id'] == student_id), None)
                if result: 
                    cls = c
                    break
        
        if not result:
            messagebox.showerror("Error", "Could not find results for this student in the current term.")
            return

        s = result['student']
        name = s['name']
        filename = f"Report_{name.replace(' ', '_')}_{cls.replace(' ', '_')}_{term}.pdf"
        
        file_path = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=filename
        )
        if not file_path:
            return

        try:
            doc = SimpleDocTemplate(file_path, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
            styles = getSampleStyleSheet()
            
            # Custom styles
            styles.add(ParagraphStyle(name='SchoolName', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, textColor=colors.HexColor("#1b5e20"), alignment=1))
            styles.add(ParagraphStyle(name='SchoolInfo', parent=styles['Normal'], fontSize=9, textColor=colors.gray, alignment=1))
            styles.add(ParagraphStyle(name='ReportTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor("#4CAF50"), alignment=1, spaceBefore=10, spaceAfter=10))
            
            elements = []

            # Header section
            header_data = [
                [Paragraph("MT OLIVES ADVENTIST SCHOOL", styles['SchoolName'])],
                [Paragraph("Sajin Close, Along Ngong-Matasia Road,", styles['SchoolInfo'])],
                [Paragraph("Next to Oryx Petrol Station,", styles['SchoolInfo'])],
                [Paragraph("Ngong", styles['SchoolInfo'])],
                [Spacer(1, 5)],
                [Paragraph("Website", styles['Normal'])],
                [Paragraph("https://mountolivessda.org/", styles['SchoolInfo'])],
                [Spacer(1, 5)],
                [Paragraph("Phone", styles['Normal'])],
                [Paragraph("+254 788 700073", styles['SchoolInfo'])],
                [Spacer(1, 5)],
                [Paragraph("Email", styles['Normal'])],
                [Paragraph("school@mountolivessda.org", styles['SchoolInfo'])],
                [Paragraph("<i>In God We Excel</i>", styles['SchoolInfo'])],
                [Spacer(1, 10)],
                [Paragraph("<hr color='#1b5e20' width='100%'/>", styles['Normal'])],
                [Paragraph("LEARNER ASSESSMENT REPORT CARD", styles['ReportTitle'])]
            ]
            header_table = Table(header_data, colWidths=[520])
            elements.append(header_table)

            # Student Info Grid
            info_data = [
                ["NAME", s['name'].upper(), "GRADE", s['class']],
                ["ADM NO", s['admission_no'], "YEAR", "2026"],
                ["TERM", term, "GENDER", s['gender']]
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
            marks_header = ["LEARNING AREA", "MARKS", "AVG", "PERFORMANCE LEVEL"]
            marks_data = [marks_header]
            
            for subj in self.get_current_subjects():
                mk = result['marks'].get(subj, 0)
                # Subject grade
                subj_grade = 'EE' if mk >= 80 else 'ME' if mk >= 70 else 'AE' if mk >= 60 else 'BE' if mk >= 50 else 'IE'
                marks_data.append([subj, mk, mk, GRADE_LABELS[subj_grade]])
            
            # Summary Rows
            possible = len(self.get_current_subjects()) * 100
            marks_data.append(["Total Scores", f"{result['total']}/{possible}", f"{result['average']}/100", f"Level: {result['grade']}"])
            marks_data.append(["Average Scores", f"{result['average']}/100", "", f"Position: {result['position']}"])

            marks_table = Table(marks_data, colWidths=[150, 100, 100, 170])
            marks_table.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E2EFDA")),
                ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1b5e20")),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('GRID', (0,0), (-1,-3), 0.5, colors.gray),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BACKGROUND', (0,-2), (-1,-1), colors.HexColor("#D0F0E0")),
                ('FONTNAME', (0,-2), (-1,-1), 'Helvetica-Bold'),
            ]))
            elements.append(marks_table)
            elements.append(Spacer(1, 10))

            # Performance Trend Chart
            drawing = Drawing(400, 150)
            bc = VerticalBarChart()
            bc.x = 50
            bc.y = 20
            bc.height = 100
            bc.width = 350
            bc.data = [[result['marks'].get(sub, 0) for sub in self.get_current_subjects()]]
            bc.strokeColor = colors.black
            bc.valueAxis.valueMin = 0
            bc.valueAxis.valueMax = 100
            bc.valueAxis.valueStep = 25
            bc.categoryAxis.labels.boxAnchor = 'ne'
            bc.categoryAxis.labels.angle = 30
            bc.categoryAxis.categoryNames = self.get_current_subjects()
            
            for i, subj in enumerate(self.get_current_subjects()):
                score = result['marks'].get(subj, 0)
                if score >= 80: bc.bars[0, i].fillColor = colors.HexColor("#1b5e20")
                elif score >= 50: bc.bars[0, i].fillColor = colors.HexColor("#4CAF50")
                else: bc.bars[0, i].fillColor = colors.HexColor("#D9534F")

            drawing.add(bc)
            elements.append(drawing)
            
            # Comments Section
            comments_data = [
                [Paragraph("Class Teacher's Comment: <font color='gray'><i>Fair performance. Keep working harder.</i></font>", styles['Normal'])],
                [Spacer(1, 10)],
                [Paragraph("Head Teacher's Comment: <font color='gray'><i>Encourage the learner to improve.</i></font>", styles['Normal'])]
            ]
            comments_table = Table(comments_data, colWidths=[520])
            comments_table.setStyle(TableStyle([
                ('BOX', (0,0), (-1,0), 0.5, colors.gray),
                ('BOX', (0,2), (-1,2), 0.5, colors.gray),
            ]))
            elements.append(comments_table)

            def add_border(canvas, doc):
                canvas.saveState()
                canvas.setStrokeColor(colors.HexColor("#1b5e20"))
                canvas.setLineWidth(2)
                canvas.rect(10, 10, A4[0]-20, A4[1]-20)
                canvas.restoreState()

            doc.build(elements, onFirstPage=add_border, onLaterPages=add_border)
            messagebox.showinfo("Report Generated", f"Report card saved as {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {str(e)}")

    # ==================== CBC INFO ====================
    def show_cbc_info(self):
        """Show CBC (Competency Based Curriculum) information page"""
        self.clear_frame()
        self._set_nav('CBC Info')
        self._page_header('CBC Information', 'Competency Based Curriculum Details')
        
        # Create main container with notebook (tabs)
        notebook = ttk.Notebook(self.content_frame)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Style for tabs
        style = ttk.Style()
        style.theme_use('default')
        style.configure('CBC.TNotebook', background=CONTENT_BG)
        style.configure('CBC.TNotebook.Tab', background=CARD_BG, foreground=TEXT_PRIMARY, 
                        padding=[12, 8], font=(FF, 11, 'bold'))
        style.map('CBC.TNotebook.Tab', background=[('selected', GREEN)], 
                  foreground=[('selected', 'white')])
        notebook.configure(style='CBC.TNotebook')
        
        # Create tab frames for each level
        tabs = {}
        for level in LEVELS:
            tab_frame = tk.Frame(notebook, bg=CONTENT_BG)
            notebook.add(tab_frame, text=f"  {level}  ")
            tabs[level] = tab_frame
        
        # Build content for each level
        self._build_lower_primary_tab(tabs['Lower Primary (Grade 1-3)'])
        self._build_upper_primary_tab(tabs['Upper Primary (Grade 4-6)'])
        self._build_junior_school_tab(tabs['Junior School (Grade 7-9)'])
    
    def _build_lower_primary_tab(self, parent):
        """Build Lower Primary (Grade 1-3) tab content"""
        canvas = tk.Canvas(parent, bg=CONTENT_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=CONTENT_BG)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Level Description
        self._add_section_header(scrollable_frame, "About Lower Primary (Grade 1-3)", 
                                  "Learners are mostly taught by one class teacher and focus on basic skills.")
        
        # Subjects Section
        self._add_section_header(scrollable_frame, "📚 Subjects / Learning Areas", "")
        subjects_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        subjects_frame.pack(fill='x', padx=20, pady=5)
        
        subjects = SUBJECTS_BY_LEVEL['Lower Primary (Grade 1-3)']
        for i, subject in enumerate(subjects):
            row = i // 2
            col = i % 2
            self._add_subject_chip(subjects_frame, subject, row, col)
        
        # Grading Scale Section
        self._add_section_header(scrollable_frame, "📊 Grading Scale", 
                                  "Learners are graded using competency levels instead of marks.")
        grading_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        grading_frame.pack(fill='x', padx=20, pady=5)
        
        grading_data = GRADING_BY_LEVEL['Lower Primary (Grade 1-3)']['levels']
        for i, (grade, info) in enumerate(grading_data.items()):
            self._add_grade_card(grading_frame, grade, info['label'], info['description'], i)
        
        # Assessment Methods
        self._add_section_header(scrollable_frame, "📝 Assessment Methods", "")
        assess_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        assess_frame.pack(fill='x', padx=20, pady=5)
        
        assess_methods = [
            "✓ Class activities",
            "✓ Oral work", 
            "✓ Practical activities",
            "✓ Teacher observation"
        ]
        for method in assess_methods:
            tk.Label(assess_frame, text=method, font=(FF, 11), bg=CARD_BG, 
                    fg=TEXT_SECONDARY, padx=15, pady=5, anchor='w').pack(fill='x', padx=15)
    
    def _build_upper_primary_tab(self, parent):
        """Build Upper Primary (Grade 4-6) tab content"""
        canvas = tk.Canvas(parent, bg=CONTENT_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=CONTENT_BG)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Level Description
        self._add_section_header(scrollable_frame, "About Upper Primary (Grade 4-6)", 
                                  "Assessment combines class assessments and national assessments.")
        
        # Subjects Section
        self._add_section_header(scrollable_frame, "📚 Subjects", "")
        subjects_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        subjects_frame.pack(fill='x', padx=20, pady=5)
        
        subjects = SUBJECTS_BY_LEVEL['Upper Primary (Grade 4-6)']
        for i, subject in enumerate(subjects):
            row = i // 2
            col = i % 2
            self._add_subject_chip(subjects_frame, subject, row, col)
        
        # Grading Scale Section
        self._add_section_header(scrollable_frame, "📊 Grading Scale", 
                                  "Performance levels based on competency.")
        grading_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        grading_frame.pack(fill='x', padx=20, pady=5)
        
        grading_data = GRADING_BY_LEVEL['Upper Primary (Grade 4-6)']['levels']
        for i, (grade, info) in enumerate(grading_data.items()):
            self._add_grade_card(grading_frame, grade, info['label'], info['description'], i)
        
        # Assessment Components
        self._add_section_header(scrollable_frame, "📝 Assessment Components", "")
        assess_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        assess_frame.pack(fill='x', padx=20, pady=5)
        
        tk.Label(assess_frame, text="60% School Based Assessment (SBA)", 
                font=(FF, 11, 'bold'), bg=CARD_BG, fg=GREEN, padx=15, pady=8).pack(anchor='w', padx=15, pady=(15, 5))
        tk.Label(assess_frame, text="   Continuous assessment by teachers", 
                font=(FF, 10), bg=CARD_BG, fg=TEXT_SECONDARY, padx=15).pack(anchor='w', padx=15)
        
        tk.Label(assess_frame, text="40% National Assessment (KNEC)", 
                font=(FF, 11, 'bold'), bg=CARD_BG, fg=ORANGE, padx=15, pady=8).pack(anchor='w', padx=15, pady=(15, 5))
        tk.Label(assess_frame, text="   Assessment done by Kenya National Examinations Council", 
                font=(FF, 10), bg=CARD_BG, fg=TEXT_SECONDARY, padx=15).pack(anchor='w', padx=15, pady=(5, 15))
    
    def _build_junior_school_tab(self, parent):
        """Build Junior School (Grade 7-9) tab content"""
        canvas = tk.Canvas(parent, bg=CONTENT_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(parent, orient='vertical', command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=CONTENT_BG)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')
        
        # Level Description
        self._add_section_header(scrollable_frame, "About Junior School (Grade 7-9)", 
                                  "Junior school uses competency levels but also percentage scores.")
        
        # Core Subjects Section
        self._add_section_header(scrollable_frame, "📚 Core Subjects", "")
        subjects_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        subjects_frame.pack(fill='x', padx=20, pady=5)
        
        core_subjects = SUBJECTS_BY_LEVEL['Junior School (Grade 7-9)']['core']
        for i, subject in enumerate(core_subjects):
            row = i // 2
            col = i % 2
            self._add_subject_chip(subjects_frame, subject, row, col)
        
        # Optional Subjects Note
        opt_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        opt_frame.pack(fill='x', padx=20, pady=5)
        tk.Label(opt_frame, text="📌 Optional Subjects (schools may choose some):", 
                font=(FF, 11, 'bold'), bg=CARD_BG, fg=PURPLE, padx=15, pady=8).pack(anchor='w', padx=15)
        optionals = SUBJECTS_BY_LEVEL['Junior School (Grade 7-9)']['optional']
        for opt in optionals:
            tk.Label(opt_frame, text=f"   • {opt}", font=(FF, 10), bg=CARD_BG, 
                    fg=TEXT_SECONDARY, padx=15, pady=3).pack(anchor='w', padx=15)
        
        # Grading Scale Section
        self._add_section_header(scrollable_frame, "📊 Grading Scale", 
                                  "Competency levels with percentage scores.")
        grading_frame = tk.Frame(scrollable_frame, bg=CARD_BG, relief='flat', bd=1)
        grading_frame.pack(fill='x', padx=20, pady=5)
        
        grading_data = GRADING_BY_LEVEL['Junior School (Grade 7-9)']['levels']
        for i, (grade, info) in enumerate(grading_data.items()):
            self._add_grade_card(grading_frame, grade, info['label'], info['description'], i)
    
    def _add_section_header(self, parent, title, subtitle):
        """Add a section header with title and subtitle"""
        header_frame = tk.Frame(parent, bg=CONTENT_BG)
        header_frame.pack(fill='x', padx=20, pady=(15, 5))
        
        tk.Label(header_frame, text=title, font=(FF, 14, 'bold'), 
                bg=CONTENT_BG, fg=TEXT_PRIMARY).pack(anchor='w')
        if subtitle:
            tk.Label(header_frame, text=subtitle, font=(FF, 10), 
                    bg=CONTENT_BG, fg=TEXT_SECONDARY).pack(anchor='w', pady=(2, 0))
    
    def _add_subject_chip(self, parent, subject, row, col):
        """Add a subject chip/badge"""
        chip = tk.Frame(parent, bg='#e8f5e9', relief='flat', bd=1)
        chip.grid(row=row, column=col, padx=10, pady=8, sticky='w')
        
        tk.Label(chip, text=f"▸ {subject}", font=(FF, 10), 
                bg='#e8f5e9', fg=GREEN, padx=10, pady=6).pack()
    
    def _add_grade_card(self, parent, grade, label, description, index):
        """Add a grade card with visual styling"""
        # Determine color based on grade
        grade_colors = {
            'EE': '#2ecc71',  # Green - Exceeding
            'ME': '#3498db',  # Blue - Meeting
            'AE': '#f39c12',  # Orange - Approaching
            'BE': '#e74c3c',  # Red - Below
        }
        color = grade_colors.get(grade, '#95a5a6')
        
        card = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        card.pack(fill='x', padx=10, pady=5)
        
        # Grade badge
        badge = tk.Frame(card, bg=color, relief='flat', bd=1)
        badge.pack(side='left', padx=10, pady=10)
        
        tk.Label(badge, text=grade, font=(FF, 12, 'bold'), 
                bg=color, fg='white', padx=12, pady=6).pack()
        
        # Grade info
        info_frame = tk.Frame(card, bg=CARD_BG)
        info_frame.pack(side='left', fill='both', expand=True, padx=10, pady=10)
        
        tk.Label(info_frame, text=label, font=(FF, 11, 'bold'), 
                bg=CARD_BG, fg=TEXT_PRIMARY).pack(anchor='w')
        tk.Label(info_frame, text=description, font=(FF, 10), 
                bg=CARD_BG, fg=TEXT_SECONDARY).pack(anchor='w', pady=(2, 0))

    # ==================== CHARTS ====================
    def show_charts(self):
        self.clear_frame()
        self._set_nav('Charts')
        self._page_header('Performance Charts', 'Visual analysis of class and subject performance')

        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 12))

        def lbl(t): tk.Label(ctrl, text=t, bg=CONTENT_BG, fg=TEXT_SECONDARY,
                              font=(FF, 10)).pack(side='left', padx=(10, 4))
        lbl('Class:')
        self.ch_cls_cb = ttk.Combobox(ctrl, values=['All'] + self.get_current_classes(), state='readonly',
                                      style='App.TCombobox', width=12)
        self.ch_cls_cb.set('All')
        self.ch_cls_cb.pack(side='left', ipady=4)
        lbl('Term:')
        self.ch_term_cb = ttk.Combobox(ctrl, values=TERMS, state='readonly',
                                       style='App.TCombobox', width=10)
        self.ch_term_cb.set(TERMS[0])
        self.ch_term_cb.pack(side='left', ipady=4)
        self.ch_cls_cb.bind('<<ComboboxSelected>>',  lambda e: self.load_charts())
        self.ch_term_cb.bind('<<ComboboxSelected>>', lambda e: self.load_charts())

        chart_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        chart_outer.pack(fill='both', expand=True, pady=4)
        chart_card = tk.Frame(chart_outer, bg=CARD_BG, padx=10, pady=10)
        chart_card.pack(fill='both', expand=True, padx=1, pady=1)

        self.fig, self.axes = plt.subplots(2, 2, figsize=(12, 8),
                                           constrained_layout=True)
        self.fig.set_facecolor(CARD_BG)
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_card)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
        self.load_charts()

    def load_charts(self):
        cls  = self.ch_cls_cb.get()
        term = self.ch_term_cb.get()
        results = db.calculate_results(cls, term)
        for ax in self.axes.flat:
            ax.clear()
            ax.set_facecolor('#fafafa')

        subj_totals = {s: [] for s in self.get_current_subjects()}
        for r in results:
            for s in self.get_current_subjects():
                if r['marks'].get(s): subj_totals[s].append(r['marks'][s])

        avgs   = [round(sum(subj_totals[s]) / len(subj_totals[s]), 1) if subj_totals[s] else 0 for s in self.get_current_subjects()]
        colors = [GRADE_COLORS['EE'] if a >= 80 else GRADE_COLORS['ME'] if a >= 70 else
                  GRADE_COLORS['AE'] if a >= 60 else GRADE_COLORS['BE'] if a >= 50 else
                  GRADE_COLORS['IE'] for a in avgs]

        bars0 = self.axes[0, 0].bar(self.get_current_subjects(), avgs, color=colors, edgecolor='none', width=0.6)
        self.axes[0, 0].set_title('Subject Averages', fontweight='bold',
                                   color=TEXT_PRIMARY, pad=10, fontsize=11)
        self.axes[0, 0].set_ylim(0, 110)
        self.axes[0, 0].set_ylabel('Average Marks', color=TEXT_SECONDARY, fontsize=9)
        self.axes[0, 0].tick_params(axis='x', labelrotation=40, labelcolor=TEXT_SECONDARY,
                                     labelsize=8)
        self.axes[0, 0].tick_params(axis='y', labelcolor=TEXT_SECONDARY, labelsize=8)
        self.axes[0, 0].spines['top'].set_visible(False)
        self.axes[0, 0].spines['right'].set_visible(False)
        # Align rotated labels to right so they sit under their bar
        for lbl in self.axes[0, 0].get_xticklabels():
            lbl.set_ha('right')
        # Value labels on top of each bar
        for bar, val in zip(bars0, avgs):
            if val > 0:
                self.axes[0, 0].text(bar.get_x() + bar.get_width() / 2,
                                      bar.get_height() + 1, str(val),
                                      ha='center', va='bottom', fontsize=7,
                                      color=TEXT_SECONDARY)

        grade_counts = {g: 0 for g in GRADE_COLORS}
        for r in results: grade_counts[r['grade']] += 1
        grades = [g for g, v in grade_counts.items() if v > 0]
        counts = [grade_counts[g] for g in grades]
        if counts:
            pie_colors = [GRADE_COLORS[g] for g in grades]
            wedges, texts, autotexts = self.axes[0, 1].pie(
                counts, labels=grades, autopct='%1.1f%%',
                colors=pie_colors, startangle=90, pctdistance=0.78,
                labeldistance=1.12)
            for t in texts:
                t.set_fontsize(9)
                t.set_color(TEXT_SECONDARY)
            for t in autotexts:
                t.set_color('white')
                t.set_fontsize(8)
            self.axes[0, 1].set_title('Grade Distribution', fontweight='bold',
                                       color=TEXT_PRIMARY, pad=10, fontsize=11)

        top5 = results[:5]
        if top5:
            names = [r['student']['name'].split()[0] for r in top5]
            avgs5 = [r['average'] for r in top5]
            top_colors = [BLUE, GREEN, ORANGE, PURPLE, GRADE_COLORS['IE']][:len(top5)]
            bars = self.axes[1, 0].barh(names, avgs5, color=top_colors, edgecolor='none', height=0.5)
            self.axes[1, 0].set_title('Top Students', fontweight='bold',
                                       color=TEXT_PRIMARY, pad=10, fontsize=11)
            self.axes[1, 0].set_xlim(0, 110)
            self.axes[1, 0].set_xlabel('Average Marks', color=TEXT_SECONDARY, fontsize=9)
            self.axes[1, 0].tick_params(labelcolor=TEXT_SECONDARY, labelsize=8)
            self.axes[1, 0].spines['top'].set_visible(False)
            self.axes[1, 0].spines['right'].set_visible(False)
            # Value labels at end of each bar
            for bar, val in zip(bars, avgs5):
                self.axes[1, 0].text(val + 1, bar.get_y() + bar.get_height() / 2,
                                      f'{val:.1f}', va='center', fontsize=8,
                                      color=TEXT_SECONDARY)

        if cls == 'All':
            cls_perf = {}
            for c in CLASSES:
                cr = db.calculate_results(c, term)
                cls_perf[c] = round(sum(r['average'] for r in cr) / len(cr), 1) if cr else 0
            cls_bars = self.axes[1, 1].bar(list(cls_perf.keys()), list(cls_perf.values()),
                                           color=BLUE, edgecolor='none', width=0.5)
            self.axes[1, 1].set_title('Class Performance', fontweight='bold',
                                       color=TEXT_PRIMARY, pad=10, fontsize=11)
            self.axes[1, 1].set_ylim(0, 110)
            self.axes[1, 1].set_ylabel('Average Marks', color=TEXT_SECONDARY, fontsize=9)
            self.axes[1, 1].tick_params(labelcolor=TEXT_SECONDARY, labelsize=8)
            self.axes[1, 1].spines['top'].set_visible(False)
            self.axes[1, 1].spines['right'].set_visible(False)
            for bar, val in zip(cls_bars, cls_perf.values()):
                if val > 0:
                    self.axes[1, 1].text(bar.get_x() + bar.get_width() / 2,
                                          val + 1, str(val),
                                          ha='center', va='bottom', fontsize=8,
                                          color=TEXT_SECONDARY)
        else:
            self.axes[1, 1].set_title('Select "All" for class comparison',
                                      color=TEXT_SECONDARY, fontsize=10)
            self.axes[1, 1].axis('off')

        self.canvas.draw()

    # ==================== REPORT CARDS ====================
    def show_report_cards(self):
        self.clear_frame()
        self._set_nav('Report Cards')
        self._page_header('Report Cards', 'Generate and print learner assessment report cards')

        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 12))

        def lbl(t): tk.Label(ctrl, text=t, bg=CONTENT_BG, fg=TEXT_SECONDARY,
                              font=(FF, 10)).pack(side='left', padx=(10, 4))
        lbl('Class:')
        self.rc_cls_cb = ttk.Combobox(ctrl, values=self.get_current_classes(), state='readonly',
                                      style='App.TCombobox', width=12)
        self.rc_cls_cb.set(CLASSES[0])
        self.rc_cls_cb.pack(side='left', ipady=4)
        lbl('Term:')
        self.rc_term_cb = ttk.Combobox(ctrl, values=TERMS, state='readonly',
                                       style='App.TCombobox', width=10)
        self.rc_term_cb.set(TERMS[0])
        self.rc_term_cb.pack(side='left', ipady=4)
        lbl('Student:')
        self.rc_stu_cb = ttk.Combobox(ctrl, state='readonly',
                                      style='App.TCombobox', width=22)
        self.rc_stu_cb.pack(side='left', ipady=4)

        self.rc_cls_cb.bind('<<ComboboxSelected>>',  lambda e: self._load_rc())
        self.rc_term_cb.bind('<<ComboboxSelected>>', lambda e: self._load_rc())
        self.rc_stu_cb.bind('<<ComboboxSelected>>',  lambda e: self._display_rc())

        self._toolbar_btn(ctrl, '\U0001f5a8  Print', self._print_rc).pack(side='left', padx=14)
        self._toolbar_btn(ctrl, 'Print All', self._print_all_rc, bg='#475569').pack(
            side='left', padx=4)

        # Scrollable paper-like preview
        paper_bg = tk.Frame(self.content_frame, bg='#d8dce5')
        paper_bg.pack(fill='both', expand=True, pady=4)

        _cv, _sb, scroll_inner = scrollable_frame(paper_bg, bg='#d8dce5')

        # White "paper" frame with drop shadow
        shadow = tk.Frame(scroll_inner, bg='#aab0bf')
        shadow.pack(pady=(18, 22), padx=40)
        self._rc_paper = tk.Frame(shadow, bg='white', padx=38, pady=26)
        self._rc_paper.pack(padx=3, pady=3)

        self._load_rc()

    def _load_rc(self):
        results = db.calculate_results(self.rc_cls_cb.get(), self.rc_term_cb.get())
        names = [r['student']['name'] for r in results]
        self.rc_stu_cb['values'] = names
        if names:
            self.rc_stu_cb.current(0)
            self._display_rc()

    def _display_rc(self):
        name = self.rc_stu_cb.get()
        if not name: return
        results = db.calculate_results(self.rc_cls_cb.get(), self.rc_term_cb.get())
        result  = next((r for r in results if r['student']['name'] == name), None)
        if not result: return
        for w in self._rc_paper.winfo_children():
            w.destroy()
        self._render_report_card(self._rc_paper, result, len(results),
                                 self.rc_term_cb.get())

    def _render_report_card(self, parent, result, total_students, term):
        """Render the full styled visual report card into parent."""
        s     = result['student']
        marks = result['marks']

        SCH_BLUE = '#1b5e20'
        RED_ACC  = '#c62828'
        TTL_ORG  = '#4CAF50'
        MINT_BG  = '#e8f5e9'
        HDR_BG   = '#e3f2fd'
        PERF_CLR = {'EE': '#1b5e20', 'ME': '#0d47a1',
                    'AE': '#bf360c', 'BE': '#bf360c', 'IE': '#b71c1c'}

        def get_grade(m):
            return ('EE' if m >= 80 else 'ME' if m >= 70 else
                    'AE' if m >= 60 else 'BE' if m >= 50 else 'IE')

        # School Header
        tk.Label(parent, text='MT OLIVES ADVENTIST SCHOOL',
                 bg='white', fg=SCH_BLUE, font=(FF, 15, 'bold')).pack()
        tk.Label(parent, text='Sajin Close, Along Ngong-Matasia Road,\nNext to Oryx Petrol Station, Ngong',
                 bg='white', fg='#666', font=(FF, 9)).pack()
        tk.Label(parent, text='https://mountolivessda.org/',
                 bg='white', fg=SCH_BLUE, font=(FF, 9)).pack()
        tk.Label(parent, text='school@mountolivessda.org',
                 bg='white', fg=SCH_BLUE, font=(FF, 9)).pack()
        tk.Label(parent, text='+254 788 700073',
                 bg='white', fg='#666', font=(FF, 9)).pack()
        tk.Label(parent, text='In God We Excel',
                 bg='white', fg='#999', font=(FF, 9, 'italic')).pack(pady=(0, 8))
        tk.Frame(parent, bg=SCH_BLUE, height=2).pack(fill='x', pady=(0, 12))

        # Section title
        title_border = tk.Frame(parent, bg=RED_ACC)
        title_border.pack(pady=(0, 14))
        title_inner = tk.Frame(title_border, bg='white', padx=30, pady=7)
        title_inner.pack(padx=2, pady=2)
        tk.Label(title_inner, text='LEARNER ASSESSMENT REPORT CARD',
                 bg='white', fg=TTL_ORG, font=(FF, 12, 'bold')).pack()

        # Student Info grid
        grade_num = s.get('class', 'Grade 7').replace('Grade ', '')
        info_rows = [
            ('NAME',   s['name'],         'GRADE',  grade_num),
            ('STREAM', s['admission_no'], 'YEAR',   '2026'),
            ('TERM',   term,              'GENDER', s['gender']),
        ]
        info_f = tk.Frame(parent, bg='white')
        info_f.pack(fill='x', pady=(0, 14))
        info_f.columnconfigure(0, weight=1)
        info_f.columnconfigure(1, weight=1)
        for ri, (l1, v1, l2, v2) in enumerate(info_rows):
            for ci, (label, val) in enumerate([(l1, v1), (l2, str(v2))]):
                cell = tk.Frame(info_f, bg='white')
                cell.grid(row=ri, column=ci, sticky='ew',
                          padx=(0, 10 if ci == 0 else 0), pady=2)
                row_w = tk.Frame(cell, bg='white')
                row_w.pack(fill='x')
                tk.Label(row_w, text=label, bg='white', fg='#222',
                         font=(FF, 10, 'bold'), width=9, anchor='w').pack(side='left')
                tk.Label(row_w, text=val, bg='white', fg='#333',
                         font=(FF, 10), anchor='w').pack(side='left')
                dot_sep = tk.Canvas(cell, height=2, bg='white', highlightthickness=0)
                dot_sep.pack(fill='x', pady=(2, 0))
                def _draw_dots(e, c=dot_sep):
                    c.delete('all')
                    for x in range(0, e.width, 5):
                        c.create_oval(x, 0, x + 2, 2, fill='#bbb', outline='')
                dot_sep.bind('<Configure>', _draw_dots)

        # Marks Table
        tbl = tk.Frame(parent, bg='#cccccc')
        tbl.pack(fill='x', pady=(0, 12))
        for col_i, weight in enumerate([3, 1, 1, 4]):
            tbl.columnconfigure(col_i, weight=weight)
        for text, col, anchor in [
            ('LEARNING AREA', 0, 'w'), ('MARKS', 1, 'center'),
            ('AVG', 2, 'center'), ('PERFORMANCE LEVEL', 3, 'w')
        ]:
            tk.Label(tbl, text=text, bg=HDR_BG, fg=SCH_BLUE,
                     font=(FF, 10, 'bold'), padx=10, pady=8, anchor=anchor
                     ).grid(row=0, column=col, sticky='nsew', padx=1, pady=1)
        for i, subj in enumerate(self.get_current_subjects()):
            mk    = marks.get(subj, 0)
            grade = get_grade(mk)
            row_bg = '#f5f7fa' if i % 2 == 0 else 'white'
            tk.Label(tbl, text=subj, bg=row_bg, fg='#222', font=(FF, 10),
                     padx=10, pady=6, anchor='w'
                     ).grid(row=i + 1, column=0, sticky='nsew', padx=1, pady=1)
            for col in (1, 2):
                tk.Label(tbl, text=str(mk), bg=row_bg, fg='#333', font=(FF, 10),
                         padx=10, pady=6
                         ).grid(row=i + 1, column=col, sticky='nsew', padx=1, pady=1)
            tk.Label(tbl, text=GRADE_LABELS[grade], bg=row_bg,
                     fg=PERF_CLR[grade], font=(FF, 10, 'italic'),
                     padx=10, pady=6, anchor='w'
                     ).grid(row=i + 1, column=3, sticky='nsew', padx=1, pady=1)

        total_marks = result['total']
        avg_score   = result['average']
        grade_ov    = result['grade']
        pos         = result['position']
        possible    = len(self.get_current_subjects()) * 100
        base        = len(self.get_current_subjects()) + 1
        for j, (lab, c1, c2, c3) in enumerate([
            ('Total Scores',   f'{total_marks}/{possible}', f'{avg_score}/100',
             f'Termly Performance Level: {grade_ov}'),
            ('Average Scores', f'{avg_score}/100', '',
             f'Position: {pos} of {total_students}'),
        ]):
            tk.Label(tbl, text=lab, bg=MINT_BG, fg='#222',
                     font=(FF, 10, 'bold'), padx=10, pady=7, anchor='w'
                     ).grid(row=base + j, column=0, sticky='nsew', padx=1, pady=1)
            tk.Label(tbl, text=c1, bg=MINT_BG, fg='#222',
                     font=(FF, 10, 'bold'), padx=10, pady=7
                     ).grid(row=base + j, column=1, sticky='nsew', padx=1, pady=1)
            tk.Label(tbl, text=c2, bg=MINT_BG, fg='#222',
                     font=(FF, 10, 'bold'), padx=10, pady=7
                     ).grid(row=base + j, column=2, sticky='nsew', padx=1, pady=1)
            tk.Label(tbl, text=c3, bg=MINT_BG, fg='#1b5e20',
                     font=(FF, 10, 'bold'), padx=10, pady=7, anchor='w'
                     ).grid(row=base + j, column=3, sticky='nsew', padx=1, pady=1)

        # Performance Trend chart
        tk.Label(parent, text='Performance Trend', bg='white', fg=SCH_BLUE,
                 font=(FF, 12, 'bold', 'underline')).pack(pady=(10, 4))
        fig, ax = plt.subplots(figsize=(6.2, 2.8))
        fig.patch.set_facecolor('white')
        ax.set_facecolor('#fafafa')
        mark_vals  = [marks.get(sub, 0) for sub in self.get_current_subjects()]
        bar_colors = [GRADE_COLORS[get_grade(m)] for m in mark_vals]
        ax.bar(self.get_current_subjects(), mark_vals, color=bar_colors, edgecolor='none', width=0.55)
        ax.set_ylim(0, 105)
        ax.set_yticks([0, 25, 50, 75, 100])
        ax.tick_params(axis='x', rotation=45, labelsize=8, labelcolor='#444')
        ax.tick_params(axis='y', labelsize=8, labelcolor='#444')
        ax.spines['top'].set_visible(False)
        ax.spines['right'].set_visible(False)
        ax.spines['left'].set_color('#ddd')
        ax.spines['bottom'].set_color('#ddd')
        ax.grid(axis='y', color='#eeeeee', linewidth=0.7, zorder=0)
        fig.tight_layout(pad=1.0)
        chart_widget = FigureCanvasTkAgg(fig, master=parent)
        chart_widget.draw()
        chart_widget.get_tk_widget().pack(fill='x', pady=(0, 12))
        plt.close(fig)

        # Comments
        for clabel, bg_c, brd_c in [
            ("Class Teacher's Comment", '#f0f8ff', '#90caf9'),
            ("Head Teacher's Comment",  '#fffff0', '#d4b800'),
        ]:
            outer = tk.Frame(parent, bg=brd_c)
            outer.pack(fill='x', pady=4)
            inner = tk.Frame(outer, bg=bg_c, padx=14, pady=10)
            inner.pack(fill='both', expand=True, padx=1, pady=1)
            row_f = tk.Frame(inner, bg=bg_c)
            row_f.pack(fill='x')
            tk.Label(row_f, text=f'{clabel}:', bg=bg_c, fg='#222',
                     font=(FF, 10, 'bold')).pack(side='left')
            cmnt = ('  Fair performance. Keep working harder.'
                    if 'Class' in clabel else '  Encourage the learner to improve.')
            tk.Label(row_f, text=cmnt, bg=bg_c, fg='#555', font=(FF, 10)).pack(side='left')

        # Footer
        tk.Frame(parent, bg='white', height=8).pack()
        today = datetime.now().strftime('%m/%d/%Y')
        tk.Label(parent,
                 text=f'This term closed on: {today}    |    Next term opens on: ___________',
                 bg='white', fg='#666', font=(FF, 9)).pack()
        tk.Label(parent,
                 text=('This Exam Report Card has been Issued Without Any Alterations '
                       'Whatsoever. Any Alterations Will Invalidate Its Authenticity.'),
                 bg='white', fg='red', font=(FF, 8, 'italic'),
                 wraplength=480, justify='center').pack(pady=(4, 0))

    def _gen_rc_text(self, result, total, term):
        """Plain-text fallback used for printing."""
        s, m = result['student'], result['marks']
        possible = len(self.get_current_subjects()) * 100
        lines = [
            '=' * 62,
            '      MT OLIVES ADVENTIST SCHOOL',
            '          LEARNER ASSESSMENT REPORT CARD',
            '=' * 62,
            f"  Name    : {s['name']:<20}  Grade  : {s.get('class','').replace('Grade ','')}",
            f"  Stream  : {s['admission_no']:<20}  Year   : 2026",
            f"  Term    : {term:<20}  Gender : {s['gender']}",
            '-' * 62,
            f"  {'Subject':<16} {'Marks':>6}  {'Avg':>6}  Performance Level",
            '-' * 62,
        ]
        for sub in self.get_current_subjects():
            mk = m.get(sub, 0)
            g  = ('EE' if mk >= 80 else 'ME' if mk >= 70 else
                  'AE' if mk >= 60 else 'BE' if mk >= 50 else 'IE')
            lines.append(f"  {sub:<16} {mk:>6}  {mk:>6}  {GRADE_LABELS[g]}")
        lines += [
            '-' * 62,
            f"  Total   : {result['total']}/{possible}   Average: {result['average']}/100",
            f"  Grade   : {result['grade']}   Position: {result['position']} of {total}",
            '=' * 62,
        ]
        return '\n'.join(lines)

    def _print_rc(self):
        name = self.rc_stu_cb.get()
        if not name: return
        results = db.calculate_results(self.rc_cls_cb.get(), self.rc_term_cb.get())
        result  = next((r for r in results if r['student']['name'] == name), None)
        if not result: return
        try:
            import os
            tmp = 'temp_report.txt'
            with open(tmp, 'w', encoding='utf-8') as f:
                f.write(self._gen_rc_text(result, len(results), self.rc_term_cb.get()))
            os.startfile(tmp, 'print')
        except Exception as ex:
            messagebox.showerror('Error', str(ex))

    def _print_all_rc(self):
        cls  = self.rc_cls_cb.get()
        term = self.rc_term_cb.get()
        results = db.calculate_results(cls, term)
        if not results:
            messagebox.showwarning('No Data', 'No students found'); return
        fn = f"report_cards_{cls.replace(' ', '_')}_term_{term}.txt"
        with open(fn, 'w', encoding='utf-8') as f:
            for r in results:
                f.write(self._gen_rc_text(r, len(results), term))
                f.write('\n\n' + '=' * 62 + '\n\n')
        messagebox.showinfo('Done', f'All report cards saved to {fn}')


# ====================== ENTRY POINT ========================
def main():
    root = tk.Tk()
    
    # Get the directory where the script is located
    script_dir = os.path.dirname(os.path.abspath(__file__))
    
    # Set the app logo
    try:
        logo_path = os.path.join(script_dir, 'moas.jpg')
        logo_img = Image.open(logo_path)
        # Resize for icon
        logo_img = logo_img.resize((64, 64), Image.Resampling.LANCZOS)
        logo_photo = ImageTk.PhotoImage(logo_img)
        root.iconphoto(True, logo_photo)
    except Exception as e:
        print(f'Could not load logo: {e}')
    
    app = SchoolReportApp(root)
    
    try:
        root.mainloop()
    except KeyboardInterrupt:
        print("\nKeyboardInterrupt received. Shutting down gracefully...")
    finally:
        # Cleanup code - ensure proper shutdown
        print("Cleanup complete. Goodbye!")
        try:
            root.destroy()
        except Exception:
            pass
        sys.exit(0)


if __name__ == '__main__':
    main()