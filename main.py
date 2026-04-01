# -*- coding: utf-8 -*-
"""
MOAS School Management System - Complete Merged App
Pure Tkinter • Modern UI • Full Features • Legacy DB Compatible
"""

import os
import shutil
import sys
import json
import re
import uuid
import smtplib
import tempfile
import threading
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog, filedialog
from email.message import EmailMessage
from PIL import Image, ImageTk
from database import db
from extract_letterhead import extract_letterhead
from promotion import promotion_manager
from exam_analytics import exam_analytics
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
from reportlab.lib.utils import ImageReader
from fpdf import FPDF

# ====================== DESIGN TOKENS ======================
# Lemon green + olive theme
LEMON_ACCENT      = '#D7F171'
LEMON_SOFT        = '#EEF7C7'
OLIVE_PRIMARY     = '#6B764B'
OLIVE_DARK        = '#55603A'
OLIVE_MID         = '#889660'
CREAM_BG          = '#F7F8EE'

SIDEBAR_BG        = OLIVE_PRIMARY
SIDEBAR_ACTIVE    = OLIVE_DARK
SIDEBAR_HOVER     = OLIVE_MID
SIDEBAR_TEXT      = LEMON_ACCENT
SIDEBAR_TEXT_ACT  = '#ffffff'       # White

CONTENT_BG  = CREAM_BG

# Card themes: (border_rgb_hex, inner_surface_hex) — soft tinted panels for variety
CARD_THEMES = {
    'cream':   ('#bec9a0', '#fffef9'),
    'mint':    ('#7fb89e', '#eef8f2'),
    'sky':     ('#8fa0d4', '#f2f4ff'),
    'peach':   ('#d4a088', '#fff3eb'),
    'lilac':   ('#a894c9', '#f4effc'),
    'sand':    ('#c9b87a', '#fff8e8'),
    'azure':   ('#6dadc4', '#edf7fa'),
    'blossom': ('#c97fa0', '#fcedf4'),
}


def _card_colors(theme_key):
    return CARD_THEMES.get(str(theme_key), CARD_THEMES['cream'])


BORDER_CLR, CARD_BG = _card_colors('cream')

TEXT_PRIMARY   = '#2f3521'
TEXT_SECONDARY = '#66704b'

BLUE   = LEMON_ACCENT
GREEN  = OLIVE_PRIMARY
ORANGE = '#C7E36A'
PURPLE = OLIVE_MID

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

SUBJECT_PALETTE = [
    '#2E7D32', '#1565C0', '#EF6C00', '#8E24AA', '#C62828',
    '#00897B', '#6D4C41', '#3949AB', '#7CB342', '#F9A825',
    '#5E35B1', '#D81B60', '#039BE5', '#43A047', '#FB8C00',
    '#546E7A', '#00ACC1', '#9E9D24', '#8D6E63', '#1E88E5',
]


def grade_base_code(grade_code):
    match = re.match(r'[A-Za-z]+', str(grade_code or '').upper())
    return match.group(0) if match else 'IE'


def _hex_to_rgb(value):
    value = str(value or '').strip().lstrip('#')
    if len(value) != 6:
        return (0, 0, 0)
    return tuple(int(value[i:i+2], 16) for i in (0, 2, 4))


def _rgb_to_hex(rgb):
    r, g, b = [max(0, min(255, int(v))) for v in rgb]
    return f'#{r:02x}{g:02x}{b:02x}'


def _mix_hex(color1, color2, ratio):
    ratio = max(0.0, min(1.0, float(ratio)))
    rgb1 = _hex_to_rgb(color1)
    rgb2 = _hex_to_rgb(color2)
    return _rgb_to_hex(tuple(rgb1[i] * (1 - ratio) + rgb2[i] * ratio for i in range(3)))

FF = 'Segoe UI'   # font family

# ====================== TOPBAR TOKENS ======================
TOPBAR_H        = 50                # height in px
TOPBAR_RIGHT_BG = OLIVE_DARK
TOPBAR_BTN_BG   = OLIVE_PRIMARY
TOPBAR_BTN_HOV  = OLIVE_MID
TOPBAR_ICON_FG  = '#ffffff'
TOPBAR_USER_BG  = OLIVE_PRIMARY
TOPBAR_YR_BG    = OLIVE_MID

# ====================== CBC LEVELS & SUBJECTS ======================
# Kenyan Competency Based Curriculum (CBC) Structure

# Education Levels
LEVELS = [
    'Lower Primary (Grade 1-3)',
    'Upper Primary (Grade 4-6)',
    'Junior School (Grade 7-9)'
]
ALL_SCHOOL_LEVEL = 'All School (All Levels)'

# Subjects by Level
SUBJECT_CATALOG = {
    'Lower Primary (Grade 1-3)': [
        ('LIT', 'Literacy Activities', 'Core', False),
        ('ENG', 'English Activities', 'Core', False),
        ('KIS', 'Kiswahili Activities', 'Core', False),
        ('MAT', 'Mathematical Activities', 'Core', False),
        ('ENV', 'Environmental Activities', 'Core', False),
        ('HNA', 'Hygiene & Nutrition Activities', 'Core', False),
        ('CRE', 'Religious Education Activities', 'Core', False),
        ('MCA', 'Movement & Creative Activities', 'Core', False),
    ],
    'Upper Primary (Grade 4-6)': [
        ('ENG', 'English', 'Core', False),
        ('KIS', 'Kiswahili / KSL', 'Core', False),
        ('MAT', 'Mathematics', 'Core', False),
        ('SCI', 'Science & Technology', 'Core', False),
        ('AGR', 'Agriculture', 'Core', False),
        ('SST', 'Social Studies', 'Core', False),
        ('CRE', 'Christian Religious Education', 'Core', False),
        ('IRE', 'Islamic Religious Education', 'Core', False),
        ('HRE', 'Hindu Religious Education', 'Core', False),
        ('ART', 'Creative Arts', 'Core', False),
        ('PHE', 'Physical & Health Education', 'Core', False),
        ('HSC', 'Home Science', 'Core', False),
    ],
    'Junior School (Grade 7-9)': {
        'core': [
            ('ENG', 'English', 'Core', False),
            ('KIS', 'Kiswahili / KSL', 'Core', False),
            ('MAT', 'Mathematics', 'Core', False),
            ('INT', 'Integrated Science', 'Core', False),
            ('SST', 'Social Studies', 'Core', False),
            ('CRE', 'Christian Religious Education', 'Core', False),
            ('IRE', 'Islamic Religious Education', 'Core', False),
            ('HRE', 'Hindu Religious Education', 'Core', False),
            ('AGR', 'Agriculture', 'Core', False),
            ('PTS', 'Pre-Technical Studies', 'Core', False),
            ('LSE', 'Life Skills Education', 'Core', False),
            ('HEA', 'Health Education', 'Core', False),
            ('SPE', 'Sports & Physical Education', 'Core', False),
            ('VIA', 'Visual Arts', 'Core', False),
            ('PFA', 'Performing Arts', 'Core', False),
        ],
        'optional': [
            ('CSC', 'Computer Science', 'Optional', True),
            ('FRE', 'French', 'Optional', True),
            ('GER', 'German', 'Optional', True),
            ('ARB', 'Arabic', 'Optional', True),
            ('KSL', 'Kenyan Sign Language', 'Optional', True),
        ]
    }
}

SUBJECTS_BY_LEVEL = {
    'Lower Primary (Grade 1-3)': [name for _, name, _, _ in SUBJECT_CATALOG['Lower Primary (Grade 1-3)']],
    'Upper Primary (Grade 4-6)': [name for _, name, _, _ in SUBJECT_CATALOG['Upper Primary (Grade 4-6)']],
    'Junior School (Grade 7-9)': {
        'core': [name for _, name, _, _ in SUBJECT_CATALOG['Junior School (Grade 7-9)']['core']],
        'optional': [name for _, name, _, _ in SUBJECT_CATALOG['Junior School (Grade 7-9)']['optional']],
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
EXAM_TYPES = ['Opener', 'Mid-Term', 'End-Term']
DEFAULT_EXAM_TYPE = 'End-Term'
EMAIL_SETTING_KEYS = [
    'smtp_host',
    'smtp_port',
    'smtp_username',
    'smtp_password',
    'smtp_sender_name',
    'smtp_use_tls',
]

CLASS_SUBJECTS_DONE_KEY = 'class_subjects_done_map'
ALL_SUBJECT_LEVEL = 'All Levels'
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


def get_grade_code(mark):
    """Return the CBC performance code for a numeric mark."""
    try:
        mark = float(mark)
    except (TypeError, ValueError):
        return 'IE'
    if mark >= 80: return 'EE'
    if mark >= 70: return 'ME'
    if mark >= 60: return 'AE'
    if mark >= 50: return 'BE'
    return 'IE'


def ensure_letterhead_assets():
    """Refresh extracted letterhead assets when the DOCX template changes."""
    docx_path = os.path.join('assets', 'letterhead.docx')
    png_path = os.path.join('assets', 'letterhead.png')
    footer_png_path = os.path.join('assets', 'letterhead_footer.png')
    json_path = os.path.join('assets', 'letterhead.json')
    if not os.path.exists(docx_path):
        return png_path if os.path.exists(png_path) else None

    try:
        needs_extract = (
            not os.path.exists(png_path) or
            not os.path.exists(footer_png_path) or
            not os.path.exists(json_path) or
            os.path.getmtime(docx_path) > os.path.getmtime(footer_png_path) or
            os.path.getmtime(docx_path) > os.path.getmtime(png_path) or
            os.path.getmtime(docx_path) > os.path.getmtime(json_path)
        )
        if needs_extract:
            extract_letterhead(docx_path=docx_path, output_dir='assets')
    except Exception as exc:
        print(f'Letterhead extraction failed: {exc}')

    return png_path if os.path.exists(png_path) else None


def get_letterhead_assets():
    """Return extracted header/footer image paths and text metadata."""
    header_path = ensure_letterhead_assets()
    footer_path = os.path.join('assets', 'letterhead_footer.png')
    json_path = os.path.join('assets', 'letterhead.json')
    data = {}

    try:
        if os.path.exists(json_path):
            with open(json_path, 'r', encoding='utf-8') as handle:
                data = json.load(handle)
    except Exception as exc:
        print(f'Letterhead metadata read failed: {exc}')

    return {
        'header_path': header_path if header_path and os.path.exists(header_path) else None,
        'footer_path': footer_path if os.path.exists(footer_path) else None,
        'header_lines': [str(line).strip() for line in data.get('header_lines', []) if str(line).strip()],
        'footer_lines': [str(line).strip() for line in data.get('footer_lines', []) if str(line).strip()],
    }


def get_processed_letterhead_image(image_path, section='header'):
    """Return a cleaned PIL image for rendering letterhead assets."""
    if not image_path or not os.path.exists(image_path):
        return None

    img = Image.open(image_path).copy()
    if section == 'header':
        # Trim the thin separator line embedded at the bottom of the header image.
        trim_bottom = max(6, int(img.height * 0.09))
        if img.height - trim_bottom > 20:
            img = img.crop((0, 0, img.width, img.height - trim_bottom))
    return img


def get_letterhead_print_lines():
    """Return text lines for printed report headers/footers."""
    default_lines = [
        'MT OLIVES ADVENTIST SCHOOL',
        'Sajin Close, Along Ngong-Matasia Road, Next to Oryx Petrol Station, Ngong',
        'school@mountolivessda.org | +254 788 700073 | https://mountolivessda.org/',
        'In God We Excel',
    ]

    assets = get_letterhead_assets()
    lines = assets['header_lines'] + assets['footer_lines']
    if lines:
        return lines
    return default_lines

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


def make_card(parent, padx=20, pady=16, theme='cream', **grid_or_pack):
    """Return a bordered card frame; ``theme`` selects from ``CARD_THEMES``."""
    ob, ib = _card_colors(theme)
    outer = tk.Frame(parent, bg=ob)
    inner = tk.Frame(outer, bg=ib, padx=padx, pady=pady)
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


def scrollable_frame_both(parent, bg=CONTENT_BG):
    """Return (canvas, v_scrollbar, inner_frame, h_scrollbar) for two-way scrolling."""
    outer = tk.Frame(parent, bg=bg)
    outer.pack(fill='both', expand=True)

    canvas = tk.Canvas(outer, bg=bg, highlightthickness=0)
    vbar = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
    hbar = ttk.Scrollbar(parent, orient='horizontal', command=canvas.xview)
    canvas.configure(yscrollcommand=vbar.set, xscrollcommand=hbar.set)

    vbar.pack(side='right', fill='y')
    canvas.pack(side='left', fill='both', expand=True)
    hbar.pack(fill='x', side='bottom')

    inner = tk.Frame(canvas, bg=bg)
    win = canvas.create_window((0, 0), window=inner, anchor='nw')

    def _update_scrollregion(event=None):
        canvas.configure(scrollregion=canvas.bbox('all'))

    def _on_resize(e):
        min_width = inner.winfo_reqwidth()
        canvas.itemconfig(win, width=max(e.width, min_width))
        _update_scrollregion()

    canvas.bind('<Configure>', _on_resize)
    inner.bind('<Configure>', _update_scrollregion)

    def _on_mousewheel(e):
        canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units')

    def _on_shift_mousewheel(e):
        canvas.xview_scroll(int(-1 * (e.delta / 120)), 'units')

    canvas.bind_all('<MouseWheel>', _on_mousewheel)
    canvas.bind_all('<Shift-MouseWheel>', _on_shift_mousewheel)

    return canvas, vbar, inner, hbar


def short_subject_name(subject):
    """Compact subject labels for dense table headers."""
    subject = str(subject or '').strip()
    short_map = {
        'Kiswahili / Kenyan Sign Language': 'Kiswahili /\nKSL',
        'Integrated Science': 'Integrated\nScience',
        'Health Education': 'Health\nEducation',
        'Pre-Technical Studies': 'Pre-Technical\nStudies',
        'Social Studies': 'Social\nStudies',
        'Religious Education (CRE/IRE/HRE)': 'Religious\nEducation',
        'Sports & Physical Education': 'Sports &\nPHE',
        'Life Skills Education': 'Life Skills\nEducation',
        'Visual Arts': 'Visual\nArts',
        'Performing Arts': 'Performing\nArts',
        'Foreign Languages (French, German, Arabic)': 'Foreign\nLanguages',
        'Science & Technology': 'Science &\nTechnology',
        'Christian Religious Education (CRE)': 'CRE',
        'Islamic Religious Education (IRE)': 'IRE',
        'Hindu Religious Education (HRE)': 'HRE',
        'English Language Activities': 'English\nActivities',
        'Kiswahili Language Activities': 'Kiswahili\nActivities',
        'Mathematical Activities': 'Math\nActivities',
        'Environmental Activities': 'Environmental\nActivities',
        'Religious Education Activities': 'Religious\nActivities',
        'Movement & Creative Activities': 'Movement &\nCreative',
    }
    return short_map.get(subject, subject)


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
        background=CARD_BG,
        foreground=TEXT_PRIMARY,
        bordercolor=BORDER_CLR,
        lightcolor=BORDER_CLR,
        darkcolor=BORDER_CLR,
        padding=8,
        font=(FF, 10),
    )


class AdvancedDataTable:
    """Reusable Treeview table with sorting, search, and pagination."""

    def __init__(self, parent, columns, page_size=20, search_label='Search', selectmode='browse',
                 enable_select_all=False):
        self.parent = parent
        self.columns = columns
        self.page_size = max(1, int(page_size))
        self.sort_column = None
        self.sort_desc = False
        self.current_page = 0
        self.rows = []
        self.filtered_rows = []
        self.visible_rows = []
        self.row_index = {}
        self.search_var = tk.StringVar()
        self.enable_select_all = bool(enable_select_all)
        self.select_all_var = tk.BooleanVar(value=False)

        controls = tk.Frame(parent, bg=parent.cget('bg'))
        controls.pack(fill='x', padx=12, pady=(10, 8))

        tk.Label(
            controls,
            text=f'{search_label}:',
            bg=parent.cget('bg'),
            fg=TEXT_SECONDARY,
            font=(FF, 10, 'bold')
        ).pack(side='left', padx=(0, 6))
        self.search_entry = ttk.Entry(controls, textvariable=self.search_var, style='App.TEntry', width=28)
        self.search_entry.pack(side='left', ipady=4)
        self.search_var.trace_add('write', lambda *_: self.apply_filters(reset_page=True))
        self.search_entry.bind('<Return>', lambda _e: self.apply_filters(reset_page=True))

        self.search_btn = tk.Button(
            controls,
            text='🔍 Search',
            bg=OLIVE_PRIMARY,
            fg='white',
            activebackground=OLIVE_DARK,
            activeforeground='white',
            font=(FF, 9, 'bold'),
            padx=10,
            pady=4,
            relief='flat',
            cursor='hand2',
            command=lambda: self.apply_filters(reset_page=True)
        )
        self.search_btn.pack(side='left', padx=(8, 0))

        if self.enable_select_all:
            self.select_all_cb = tk.Checkbutton(
                controls,
                text='Select all filtered',
                variable=self.select_all_var,
                bg=parent.cget('bg'),
                fg=TEXT_SECONDARY,
                activebackground=parent.cget('bg'),
                activeforeground=TEXT_PRIMARY,
                selectcolor=CARD_BG,
                font=(FF, 9, 'bold'),
                command=self._on_select_all_toggle
            )
            self.select_all_cb.pack(side='left', padx=(12, 0))

        self.status_label = tk.Label(
            controls,
            text='0 records',
            bg=parent.cget('bg'),
            fg=TEXT_SECONDARY,
            font=(FF, 9)
        )
        self.status_label.pack(side='right')

        tree_frame = tk.Frame(parent, bg=parent.cget('bg'))
        tree_frame.pack(fill='both', expand=True, padx=12, pady=(0, 8))

        col_keys = [c['key'] for c in columns]
        self.tree = ttk.Treeview(
            tree_frame, columns=col_keys, show='headings',
            style='App.Treeview', selectmode=selectmode
        )
        for col in columns:
            key = col['key']
            heading = col.get('title', key)
            width = col.get('width', 120)
            anchor = col.get('anchor', 'w')
            self.tree.heading(key, text=heading, command=lambda k=key: self.toggle_sort(k))
            self.tree.column(key, width=width, anchor=anchor)

        sb = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview, style='App.Vertical.TScrollbar')
        self.tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.tree.pack(fill='both', expand=True)

        pager = tk.Frame(parent, bg=parent.cget('bg'))
        pager.pack(fill='x', padx=12, pady=(0, 10))

        self.prev_btn = tk.Button(
            pager, text='< Prev', bg='#eef2ff', fg=TEXT_PRIMARY, relief='flat',
            font=(FF, 9, 'bold'), padx=12, pady=4, command=self.prev_page, cursor='hand2'
        )
        self.prev_btn.pack(side='left')

        self.next_btn = tk.Button(
            pager, text='Next >', bg='#eef2ff', fg=TEXT_PRIMARY, relief='flat',
            font=(FF, 9, 'bold'), padx=12, pady=4, command=self.next_page, cursor='hand2'
        )
        self.next_btn.pack(side='left', padx=6)

        self.page_label = tk.Label(pager, text='Page 1/1', bg=parent.cget('bg'), fg=TEXT_SECONDARY, font=(FF, 9))
        self.page_label.pack(side='left', padx=10)

        tk.Label(pager, text='Rows:', bg=parent.cget('bg'), fg=TEXT_SECONDARY, font=(FF, 9, 'bold')).pack(side='right')
        self.page_size_var = tk.StringVar(value=str(self.page_size))
        self.page_size_cb = ttk.Combobox(
            pager, textvariable=self.page_size_var, values=['10', '20', '50', '100'],
            state='readonly', width=5, style='App.TCombobox'
        )
        self.page_size_cb.pack(side='right', padx=(0, 6))
        self.page_size_cb.bind('<<ComboboxSelected>>', self._on_page_size_changed)

    def set_rows(self, rows, reset_page=True):
        self.rows = list(rows or [])
        self.apply_filters(reset_page=reset_page)

    def get_selected(self):
        selected = self.tree.selection()
        if not selected:
            return None
        row = self.row_index.get(selected[0])
        if row:
            return row
        return None

    def get_selected_iids(self):
        if self.enable_select_all and self.select_all_var.get():
            return [row.get('iid') for row in self.filtered_rows if row.get('iid')]
        return list(self.tree.selection())

    def toggle_sort(self, key):
        if self.sort_column == key:
            self.sort_desc = not self.sort_desc
        else:
            self.sort_column = key
            self.sort_desc = False
        self.apply_filters(reset_page=True)

    def prev_page(self):
        if self.current_page > 0:
            self.current_page -= 1
            self._render_page()

    def next_page(self):
        total_pages = self._total_pages()
        if self.current_page < total_pages - 1:
            self.current_page += 1
            self._render_page()

    def apply_filters(self, reset_page=False):
        query = self.search_var.get().strip().lower()
        if query:
            filtered = []
            for row in self.rows:
                haystack = row.get('search', '')
                if not haystack:
                    haystack = ' '.join(str(v) for v in row.get('values', ()))
                if query in haystack.lower():
                    filtered.append(row)
            self.filtered_rows = filtered
        else:
            self.filtered_rows = list(self.rows)

        if self.sort_column:
            self.filtered_rows.sort(
                key=lambda r: self._sort_value(r, self.sort_column),
                reverse=self.sort_desc
            )

        if reset_page:
            self.current_page = 0
        else:
            self.current_page = min(self.current_page, max(0, self._total_pages() - 1))
        self._render_page()

    def _on_page_size_changed(self, _event=None):
        try:
            self.page_size = max(1, int(self.page_size_var.get()))
        except Exception:
            self.page_size = 20
        self.apply_filters(reset_page=True)

    def _on_select_all_toggle(self):
        if self.select_all_var.get():
            visible_iids = [row.get('iid') for row in self.visible_rows if row.get('iid')]
            if visible_iids:
                self.tree.selection_set(visible_iids)
        else:
            self.tree.selection_remove(self.tree.selection())

    def _sort_value(self, row, key):
        value_map = row.get('value_map', {})
        value = value_map.get(key, '')
        if isinstance(value, (int, float)):
            return (0, value)
        text = str(value).strip()
        try:
            num = float(text)
            return (0, num)
        except Exception:
            return (1, text.lower())

    def _total_pages(self):
        if not self.filtered_rows:
            return 1
        return (len(self.filtered_rows) + self.page_size - 1) // self.page_size

    def _render_page(self):
        start = self.current_page * self.page_size
        end = start + self.page_size
        self.visible_rows = self.filtered_rows[start:end]

        self.row_index.clear()
        for item in self.tree.get_children():
            self.tree.delete(item)

        for idx, row in enumerate(self.visible_rows):
            row_iid = row.get('iid') or f'row_{start + idx}'
            self.tree.insert('', 'end', iid=row_iid, values=row.get('values', ()), tags=row.get('tags', ()))
            self.row_index[row_iid] = row

        if self.enable_select_all and self.select_all_var.get():
            visible_iids = [row.get('iid') for row in self.visible_rows if row.get('iid')]
            if visible_iids:
                self.tree.selection_set(visible_iids)

        total = len(self.filtered_rows)
        page_num = self.current_page + 1
        pages = self._total_pages()
        shown_from = 0 if total == 0 else start + 1
        shown_to = min(end, total)
        self.status_label.config(text=f'{shown_from}-{shown_to} of {total} records')
        self.page_label.config(text=f'Page {page_num}/{pages}')

        self.prev_btn.config(state='normal' if self.current_page > 0 else 'disabled')
        self.next_btn.config(state='normal' if self.current_page < pages - 1 else 'disabled')


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
        self.sidebar_collapsed = False
        self.sidebar_host = None
        self._topbar_clock_job = None
        self._topbar_clock_label = None
        self._topbar_bar = None
        self._topbar_clock_visible = True
        self._topbar_clock_generation = 0
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

        self._ensure_default_catalog_records()
        self._ensure_default_grading_scales()
        setup_treeview_style()
        self.show_login()

    # ------------------- utilities -------------------
    def _generate_short_label(self, text, kind='subject'):
        text = str(text or '').strip()
        if not text:
            return ''

        if kind == 'class':
            match = re.search(r'grade\s*(\d+)', text, re.I)
            if match:
                return f"G{match.group(1)}"
            return ''.join(word[0].upper() for word in text.split()[:3])

        if kind == 'teacher':
            parts = [part for part in re.split(r'\s+', text) if part]
            if len(parts) >= 2:
                return ''.join(part[0].upper() for part in parts[:2])
            return text[:4].upper()

        compact = short_subject_name(text).replace('\n', ' / ').strip()
        if len(compact) <= 16:
            return compact
        words = [word for word in re.split(r'[^A-Za-z0-9]+', text) if word]
        if len(words) > 1:
            return ''.join(word[0].upper() for word in words[:4])
        return text[:16]

    def _ensure_default_catalog_records(self):
        for level, classes in CLASSES_BY_LEVEL.items():
            for class_name in classes:
                if not db.get_class_by_name(class_name):
                    db.add_class(class_name, level, abbreviation=self._generate_short_label(class_name, 'class'))

        for level in LEVELS:
            level_subjects = SUBJECT_CATALOG.get(level, [])
            if isinstance(level_subjects, dict):
                catalog = list(level_subjects.get('core', [])) + list(level_subjects.get('optional', []))
            else:
                catalog = list(level_subjects)

            for code, name, category, is_optional in catalog:
                if not db.get_subject_by_name(name, level):
                    db.add_subject(
                        name,
                        level,
                        category,
                        is_optional=is_optional,
                        abbreviation=code,
                        code=code
                    )

    def _replace_subject_catalog_with_defaults(self):
        subject_rows = []
        for level in LEVELS:
            level_subjects = SUBJECT_CATALOG.get(level, [])
            if isinstance(level_subjects, dict):
                catalog = list(level_subjects.get('core', [])) + list(level_subjects.get('optional', []))
            else:
                catalog = list(level_subjects)
            for code, name, category, is_optional in catalog:
                subject_rows.append({
                    'code': code,
                    'name': name,
                    'level': level,
                    'category': category,
                    'is_optional': is_optional,
                })
        db.replace_subject_catalog(subject_rows)

    def _ensure_default_grading_scales(self):
        default_scale = [
            ('EE', 'Exceeding Expectations', 80, 100, 1),
            ('ME', 'Meeting Expectations', 70, 79, 2),
            ('AE', 'Approaching Expectations', 60, 69, 3),
            ('BE', 'Below Expectations', 50, 59, 4),
            ('IE', 'Inadequate', 0, 49, 5),
        ]
        all_classes = []
        for classes in CLASSES_BY_LEVEL.values():
            for class_name in classes:
                if class_name not in all_classes:
                    all_classes.append(class_name)
        for row in db.get_all_classes():
            if row.get('name') and row['name'] not in all_classes:
                all_classes.append(row['name'])

        for class_name in all_classes:
            if db.get_grading_scales(class_name):
                continue
            for code, label, min_mark, max_mark, order in default_scale:
                db.add_grading_scale(class_name, min_mark, max_mark, code, label, order)

    def _get_class_meta(self, class_name):
        return db.get_class_by_name(class_name) or {}

    def _get_subject_meta(self, subject, class_name=''):
        level = self._get_level_for_class(class_name) if class_name else None
        meta = db.get_subject_by_name(subject, level)
        if meta:
            return meta
        if level:
            meta = db.get_subject_by_name(subject, ALL_SUBJECT_LEVEL)
            if meta:
                return meta
        return db.get_subject_by_name(subject) or {}

    def _get_teacher_label(self, teacher):
        if isinstance(teacher, dict):
            abbreviation = str(teacher.get('abbreviation', '') or '').strip()
            return abbreviation or teacher.get('full_name') or teacher.get('username') or ''
        return self._generate_short_label(str(teacher or ''), 'teacher')

    def _get_class_label(self, class_name):
        meta = self._get_class_meta(class_name)
        abbreviation = str(meta.get('abbreviation', '') or '').strip()
        return abbreviation or class_name

    def _get_subject_label(self, subject, class_name='', multiline=False):
        meta = self._get_subject_meta(subject, class_name)
        label = str(meta.get('abbreviation', '') or '').strip()
        if not label:
            label = self._generate_short_label(subject, 'subject')
        return label.replace(' / ', '\n') if multiline else label

    def _get_class_grading_scale(self, class_name):
        if class_name in ('', ALL_SCHOOL_LEVEL, 'All'):
            return []
        scales = db.get_grading_scales(class_name)
        if scales:
            return scales
        return []

    def _get_grade_code_for_class(self, mark, class_name=''):
        try:
            mark = float(mark)
        except (TypeError, ValueError):
            return 'IE'
        for scale in self._get_class_grading_scale(class_name):
            if float(scale.get('min_mark', 0)) <= mark <= float(scale.get('max_mark', 0)):
                return scale.get('grade_code', 'IE')
        return get_grade_code(mark)

    def _get_grade_name_for_class(self, grade_code, class_name=''):
        for scale in self._get_class_grading_scale(class_name):
            if scale.get('grade_code') == grade_code:
                return scale.get('grade_name') or GRADE_LABELS.get(grade_base_code(grade_code), grade_code)
        return GRADE_LABELS.get(grade_base_code(grade_code), grade_code)

    def _get_grade_color(self, grade_code):
        return GRADE_COLORS.get(str(grade_code or '').upper(), GRADE_COLORS.get(grade_base_code(grade_code), GREEN))

    def _get_subject_color(self, subject, class_name=''):
        key = self._normalize_key(self._get_subject_label(subject, class_name) or subject)
        if not key:
            return SUBJECT_PALETTE[0]
        index = sum(ord(ch) for ch in key) % len(SUBJECT_PALETTE)
        return SUBJECT_PALETTE[index]

    def _get_subject_colors(self, subject, class_name=''):
        base = self._get_subject_color(subject, class_name)
        return {
            'base': base,
            'soft': _mix_hex(base, '#ffffff', 0.82),
            'mid': _mix_hex(base, '#ffffff', 0.65),
            'border': _mix_hex(base, '#223022', 0.18),
            'text': '#ffffff',
            'dark_text': _mix_hex(base, '#102018', 0.35),
        }

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
        return self._get_subjects_for_level(self.current_level)
    
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
        if self.current_level == ALL_SCHOOL_LEVEL:
            classes = []
            for level_classes in CLASSES_BY_LEVEL.values():
                for class_name in level_classes:
                    if class_name not in classes:
                        classes.append(class_name)
            return classes
        return CLASSES_BY_LEVEL.get(self.current_level, CLASSES)
    
    def get_current_grading(self):
        """Get grading system for the current CBC level"""
        if self.current_level == ALL_SCHOOL_LEVEL:
            return GRADING_BY_LEVEL['Junior School (Grade 7-9)']
        return GRADING_BY_LEVEL.get(self.current_level, GRADING_BY_LEVEL['Junior School (Grade 7-9)'])

    def _get_subjects_for_level(self, level):
        """Flatten level subject config into a clean ordered list."""
        if level == ALL_SCHOOL_LEVEL:
            subjects = []
            seen = set()
            for level_name in LEVELS:
                for subject in self._get_subjects_for_level(level_name):
                    if subject not in seen:
                        subjects.append(subject)
                        seen.add(subject)
            return subjects
        level_subjects = SUBJECTS_BY_LEVEL.get(level, SUBJECTS)
        if isinstance(level_subjects, dict):
            return list(level_subjects.get('core', []))
        return list(level_subjects)

    def _get_subjects_for_selected_class(self, class_name, term='One', exam_type=DEFAULT_EXAM_TYPE):
        """Return the right subject list for a concrete class selection."""
        if class_name and class_name not in ('All', ALL_SCHOOL_LEVEL):
            return self._get_subjects_for_class(class_name, term, exam_type)
        return self.get_current_subjects()

    def _get_level_for_class(self, class_name):
        """Resolve the CBC level for a class name."""
        for level, classes in CLASSES_BY_LEVEL.items():
            if class_name in classes:
                return level

        for row in db.get_all_classes():
            if row.get('name') == class_name:
                return row.get('level')

        return self.current_level

    def _get_subjects_for_class(self, class_name, term='One', exam_type=DEFAULT_EXAM_TYPE, for_reporting=False):
        """Return ordered class subject list, honoring class-specific 'subjects done' settings."""
        subjects = self._get_subject_pool_for_class(class_name)

        mapping = self._get_class_subjects_done_map()
        configured = mapping.get(class_name, [])
        if configured:
            subject_set = set(subjects)
            ordered = [subject for subject in configured if subject in subject_set]
            for subject in configured:
                if subject not in ordered:
                    ordered.append(subject)
            return ordered

        if for_reporting:
            # Automatic report mode: show only subjects already entered for this class/term/exam.
            done_subjects = self._get_done_subjects_from_marks(class_name, term, exam_type)
            return done_subjects

        # Entry mode fallback: include class pool plus any ad-hoc subjects already captured in marks.
        subject_set = set(subjects)
        for student in db.get_students_by_class(class_name):
            for subject in db.get_student_marks(student['id'], term, exam_type).keys():
                if subject not in subject_set:
                    subjects.append(subject)
                    subject_set.add(subject)

        return subjects

    def _get_subjects_for_scope(self, class_name, term='One', exam_type=DEFAULT_EXAM_TYPE, results=None):
        """Return ordered subject columns for a report scope."""
        if class_name != 'All':
            return self._get_subjects_for_class(class_name, term, exam_type, for_reporting=True)

        subjects = []
        subject_set = set()
        rows = results or []
        if not rows:
            rows = self._get_ranked_results(class_name, term, exam_type)
        for result in rows:
            for subject in result.get('subjects', []):
                if subject not in subject_set:
                    subjects.append(subject)
                    subject_set.add(subject)
        return subjects

    def _get_subject_mode_badge(self, class_name, term='One', exam_type=DEFAULT_EXAM_TYPE):
        """Return (label, bg, fg) for report subject mode badge."""
        mapping = self._get_class_subjects_done_map()
        if class_name and class_name != 'All':
            if mapping.get(class_name):
                return ('Mode: Manual Subjects', '#14532d', 'white')
            auto_subjects = self._get_done_subjects_from_marks(class_name, term, exam_type)
            if auto_subjects:
                return ('Mode: Automatic (From Marks)', '#1d4ed8', 'white')
            return ('Mode: Automatic (No Marks Yet)', '#92400e', 'white')

        classes = self.get_current_classes()
        manual = 0
        auto = 0
        for cls in classes:
            if mapping.get(cls):
                manual += 1
            else:
                auto += 1
        if manual and auto:
            return (f'Mode: Mixed ({manual} Manual / {auto} Auto)', '#6b21a8', 'white')
        if manual and not auto:
            return ('Mode: Manual Subjects (All Classes)', '#14532d', 'white')
        return ('Mode: Automatic (All Classes)', '#1d4ed8', 'white')

    def _get_ranked_results(self, class_name, term='One', exam_type=DEFAULT_EXAM_TYPE):
        """Build whole-class ranked results with subject marks and summary fields."""
        if class_name == 'All':
            allowed_classes = set(self.get_current_classes())
            students = [
                student for student in db.get_all_students()
                if student.get('class') in allowed_classes
            ]
        else:
            students = db.get_students_by_class(class_name)

        class_subjects = {}
        results = []

        for student in students:
            student_class = student.get('class', '')
            if student_class not in class_subjects:
                class_subjects[student_class] = self._get_subjects_for_class(
                    student_class, term, exam_type, for_reporting=True
                )

            subjects = class_subjects.get(student_class, [])
            marks = db.get_student_marks(student['id'], term, exam_type)
            total = sum(int(marks.get(subject, 0) or 0) for subject in subjects)
            average = round(total / len(subjects), 1) if subjects else 0
            grade = self._get_grade_code_for_class(average, student_class)
            class_level = self._get_level_for_class(student_class)

            results.append({
                'student': student,
                'marks': marks,
                'subjects': subjects,
                'exam_type': exam_type,
                'subject_count': len(subjects),
                'possible_total': len(subjects) * 100,
                'total': total,
                'average': average,
                'grade': grade,
                'level': self._get_grade_name_for_class(grade, student_class),
                'class_level': class_level,
            })

        results.sort(key=lambda row: (-row['total'], -row['average'], row['student']['name'].lower()))

        last_key = None
        current_position = 0
        for index, row in enumerate(results, start=1):
            rank_key = (row['total'], row['average'])
            if rank_key != last_key:
                current_position = index
                last_key = rank_key
            row['position'] = current_position

        return results

    def _get_stream_names_for_class(self, class_name):
        class_row = db.get_class_by_name(class_name or '')
        if not class_row:
            return []
        streams = db.get_streams_for_class(class_row['id'])
        return [stream.get('name', '').strip() for stream in streams if stream.get('name', '').strip()]

    def _get_selected_report_stream(self):
        if not hasattr(self, 'rc_stream_cb'):
            return ''
        stream = (self.rc_stream_cb.get() or '').strip()
        return '' if stream == 'All Streams' else stream

    def _get_report_card_results(self):
        results = self._get_ranked_results(
            self.rc_cls_cb.get(),
            self.rc_term_cb.get(),
            self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE,
        )
        selected_stream = self._get_selected_report_stream()
        if selected_stream:
            results = [
                result for result in results
                if result.get('student', {}).get('stream', '').strip() == selected_stream
            ]
        return results

    def _refresh_report_card_streams(self, reload_results=True):
        if not hasattr(self, 'rc_stream_cb'):
            return
        current = (self.rc_stream_cb.get() or '').strip()
        values = ['All Streams'] + self._get_stream_names_for_class(self.rc_cls_cb.get())
        self.rc_stream_cb['values'] = values
        if current and current in values:
            self.rc_stream_cb.set(current)
        else:
            self.rc_stream_cb.set(values[0] if values else 'All Streams')
        if reload_results:
            self._load_rc()

    def _normalize_text(self, value):
        text = str(value or '').strip().lower()
        text = re.sub(r'\s+', ' ', text)
        return text

    def _normalize_key(self, value):
        text = self._normalize_text(value)
        return re.sub(r'[^a-z0-9]+', '', text)

    def _sheet_title_to_class_name(self, title):
        title_norm = self._normalize_text(title)
        match = re.search(r'grade\s*(\d+)', title_norm)
        if match:
            return f"Grade {int(match.group(1))}"
        return None

    def _generate_admission_no(self, class_name, student_name):
        class_code = self._normalize_key(class_name).upper()[:6] or 'CLASS'
        name_code = self._normalize_key(student_name).upper()[:4] or 'STU'
        while True:
            candidate = f"AUTO-{class_code}-{name_code}-{uuid.uuid4().hex[:5].upper()}"
            if not db.get_student_by_admission_no(candidate):
                return candidate

    def _map_sheet_subject(self, raw_subject, class_name=''):
        raw = str(raw_subject or '').strip()
        if not raw:
            return ''

        classes = self._get_subjects_for_class(class_name, TERMS[0]) if class_name else []
        class_lookup = {self._normalize_key(subject): subject for subject in classes}
        raw_key = self._normalize_key(raw)
        if raw_key in class_lookup:
            return class_lookup[raw_key]

        alias_map = {
            'eng': 'English',
            'english': 'English',
            'lang': 'English Language Activities',
            'math': 'Mathematics',
            'maths': 'Mathematics',
            'mat': 'Mathematics',
            'mathematicalactivities': 'Mathematical Activities',
            'kis': 'Kiswahili / Kenyan Sign Language',
            'kiswahili': 'Kiswahili / Kenyan Sign Language',
            'englishlanguageactivities': 'English Language Activities',
            'kiswahililanguageactivities': 'Kiswahili Language Activities',
            'intsci': 'Integrated Science',
            'integratedscience': 'Integrated Science',
            'sci': 'Science & Technology',
            'science': 'Science & Technology',
            'scitech': 'Science & Technology',
            'scienceandtechnology': 'Science & Technology',
            'environ': 'Environmental Activities',
            'env': 'Environmental Activities',
            'creative': 'Creative Arts',
            'carts': 'Creative Arts',
            'ca': 'Creative Arts',
            'casports': 'Sports & Physical Education',
            'agri': 'Agriculture',
            'agrinut': 'Agriculture',
            'sst': 'Social Studies',
            'socialstudies': 'Social Studies',
            'cre': 'Christian Religious Education (CRE)',
            'pretech': 'Pre-Technical Studies',
            'french': 'Foreign Languages (French, German, Arabic)',
        }

        mapped = alias_map.get(raw_key, raw)
        mapped_key = self._normalize_key(mapped)
        if mapped_key in class_lookup:
            return class_lookup[mapped_key]

        for subject in classes:
            subject_key = self._normalize_key(subject)
            if raw_key and (raw_key in subject_key or subject_key in raw_key):
                return subject

        return mapped

    def _find_assessment_header_row(self, worksheet):
        for row_idx in range(1, min(10, worksheet.max_row) + 1):
            values = [self._normalize_text(cell) for cell in next(worksheet.iter_rows(min_row=row_idx, max_row=row_idx, values_only=True))]
            if 'learner' in values and ('no.' in values or 'no' in values):
                return row_idx
        return None

    def _parse_assessment_sheet(self, worksheet):
        header_row = self._find_assessment_header_row(worksheet)
        if not header_row:
            return None

        class_name = self._sheet_title_to_class_name(worksheet.title)
        if not class_name:
            return None

        excluded = {'no', 'no.', 'learner', 'total', 'total ', 'avg', 'average', 'psn', 'position', 'level'}
        header_values = [cell for cell in next(worksheet.iter_rows(min_row=header_row, max_row=header_row, values_only=True))]

        subject_columns = []
        for col_idx, cell in enumerate(header_values, start=1):
            label = self._normalize_text(cell)
            if col_idx <= 2 or not label or label in excluded:
                continue
            subject_columns.append((col_idx, self._map_sheet_subject(cell, class_name)))

        if not subject_columns:
            return None

        students = []
        for row in worksheet.iter_rows(min_row=header_row + 2, values_only=True):
            learner = str(row[1] or '').strip() if len(row) > 1 else ''
            if not learner:
                continue

            marks = {}
            for col_idx, subject in subject_columns:
                raw = row[col_idx - 1] if len(row) >= col_idx else None
                if raw in (None, ''):
                    continue
                try:
                    marks[subject] = min(100, max(0, int(float(raw))))
                except (TypeError, ValueError):
                    continue

            if marks:
                students.append({'name': learner, 'marks': marks})

        if not students:
            return None

        return {
            'sheet_name': worksheet.title,
            'class_name': class_name,
            'students': students,
        }

    def _open_progress_dialog(self, title, message):
        dialog = tk.Toplevel(self.root)
        dialog.title(title)
        dialog.geometry('420x150')
        dialog.configure(bg=CONTENT_BG)
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        dialog.protocol('WM_DELETE_WINDOW', lambda: None)

        pr_bo, pr_bi = _card_colors('azure')
        outer = tk.Frame(dialog, bg=pr_bo)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=pr_bi, padx=22, pady=20)
        card.pack(padx=1, pady=1)

        tk.Label(card, text=title, bg=pr_bi, fg=TEXT_PRIMARY,
                 font=(FF, 13, 'bold')).pack(anchor='w')
        status_label = tk.Label(card, text=message, bg=pr_bi, fg=TEXT_SECONDARY,
                                font=(FF, 10), wraplength=360, justify='left')
        status_label.pack(anchor='w', pady=(8, 10))

        percent_label = tk.Label(card, text='0%', bg=pr_bi, fg=GREEN,
                                 font=(FF, 11, 'bold'))
        percent_label.pack(anchor='e')

        progress = ttk.Progressbar(card, orient='horizontal', length=360, mode='determinate', maximum=100)
        progress.pack(fill='x')
        dialog.update_idletasks()
        return dialog, status_label, percent_label, progress

    def _update_progress_dialog(self, dialog, status_label, percent_label, progress, current, total, message):
        total = max(1, int(total or 1))
        current = min(total, max(0, int(current)))
        percent = int((current / total) * 100)
        status_label.config(text=message)
        percent_label.config(text=f'{percent}%')
        progress['value'] = percent
        dialog.update_idletasks()
    
    def set_level(self, level):
        """Set the current CBC level and update subjects/classes"""
        if level in LEVELS or level == ALL_SCHOOL_LEVEL:
            self.current_level = level
            # Update legacy compatibility variables
            global SUBJECTS, CLASSES
            SUBJECTS = self._get_subjects_for_level(level)
            CLASSES = self.get_current_classes()
            return True
        return False
    
    def load_logo(self):
        try:
            img = Image.open('moas.jpg').resize((60, 60))
            return ImageTk.PhotoImage(img)
        except:
            return None

    def _clear(self):
        self._topbar_clock_generation += 1
        self._topbar_clock_job = None
        self._topbar_clock_label = None
        self._topbar_bar = None
        self._topbar_clock_visible = True
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
        hover_bg = OLIVE_MID if bg in (BLUE, ORANGE, PURPLE, GREEN, OLIVE_PRIMARY, LEMON_ACCENT) else OLIVE_DARK
        b.bind('<Enter>', lambda e: b.config(bg=hover_bg))
        b.bind('<Leave>', lambda e: b.config(bg=bg))
        return b

    # ─────────────────────── Auth screen ───────────────────────────────────
    _AUTH_BG   = CREAM_BG
    _CARD_W    = 400
    _BLUE      = OLIVE_PRIMARY
    _DARK_BLUE = OLIVE_DARK
    _FIELD_BG  = LEMON_SOFT
    _TEXT      = TEXT_PRIMARY
    _GRAY_T    = TEXT_SECONDARY

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
        _rr(canvas, PAD+3, PAD+5, CW+PAD+3, CARD_H+PAD+5, 16, '#c8cfac')
        # card
        _rr(canvas, PAD, PAD, CW+PAD, CARD_H+PAD, 16, CARD_BG)

        # content frame embedded in canvas
        frame = tk.Frame(canvas, bg=CARD_BG)
        canvas.create_window(CW // 2 + PAD, PAD + 1, window=frame,
                             anchor='n', width=CW - 40)
        self._auth_frame = frame

        # ── Logo ────────────────────────────────────────────────────────────
        if self.logo_img:
            tk.Label(frame, image=self.logo_img, bg=CARD_BG).pack(pady=20)
        else:
            logo = tk.Canvas(frame, width=56, height=56, bg=CARD_BG, highlightthickness=0)
            logo.pack(pady=(24, 12))
            _rr(logo, 0, 0, 56, 56, 12, self._BLUE)
            logo.create_text(28, 28, text='\U0001f393', font=(FF, 22), fill='white')

        # ── Title ───────────────────────────────────────────────────────────
        tk.Label(frame, text='MOAS', bg=CARD_BG, fg=self._TEXT,
                 font=(FF, 16, 'bold')).pack(pady=(5, 3))
        tk.Label(frame, text='School Management System', bg=CARD_BG,
                 fg=self._GRAY_T, font=(FF, 11)).pack(pady=(0, 18))

        # ── Tab switcher ────────────────────────────────────────────────────
        TH = 42
        tab_c = tk.Canvas(frame, height=TH, bg=CARD_BG, highlightthickness=0)
        tab_c.pack(fill='x', pady=(0, 14))

        def draw_tabs(w):
            tab_c.delete('all')
            _rr(tab_c, 0, 0, w, TH, 10, self._FIELD_BG)
            half = w // 2
            # active indicator
            if tab == 'login':
                _rr(tab_c, 4, 4, half - 2, TH - 4, 7, CARD_BG)
            else:
                _rr(tab_c, half + 2, 4, w - 4, TH - 4, 7, CARD_BG)
            # Login label
            lo = tk.Label(tab_c,
                          text='Login',
                          bg=CARD_BG if tab == 'login' else self._FIELD_BG,
                          fg=self._TEXT if tab == 'login' else self._GRAY_T,
                          font=(FF, 11, 'bold' if tab == 'login' else 'normal'),
                          cursor='hand2')
            tab_c.create_window(half // 2, TH // 2, window=lo, width=half - 10)
            lo.bind('<Button-1>', lambda e: self._switch_tab('login'))
            # Sign Up label
            su = tk.Label(tab_c,
                          text='Sign Up',
                          bg=CARD_BG if tab == 'signup' else self._FIELD_BG,
                          fg=self._TEXT if tab == 'signup' else self._GRAY_T,
                          font=(FF, 11, 'bold' if tab == 'signup' else 'normal'),
                          cursor='hand2')
            tab_c.create_window(half + half // 2, TH // 2, window=su, width=half - 10)
            su.bind('<Button-1>', lambda e: self._switch_tab('signup'))

        tab_c.bind('<Configure>', lambda e: draw_tabs(e.width))

        # ── Form fields ─────────────────────────────────────────────────────
        self._auth_entries = {}

        def mk_field(label, key, show=''):
            tk.Label(frame, text=label, bg=CARD_BG, fg=self._TEXT,
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

        btn_c = tk.Canvas(frame, height=46, bg=CARD_BG, highlightthickness=0,
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
        btn_c.bind('<Enter>',     lambda e: draw_btn(OLIVE_MID))
        btn_c.bind('<Leave>',     lambda e: draw_btn(self._BLUE))

        # ── Forgot password / bottom spacer ─────────────────────────────────
        if tab == 'login':
            fp = tk.Label(frame, text='Forgot password?', bg=CARD_BG,
                          fg=self._BLUE, font=(FF, 10), cursor='hand2')
            fp.pack(pady=(12, 24))
            fp.bind('<Button-1>', lambda e: messagebox.showinfo(
                'Forgot Password',
                'Please contact your administrator to reset your password.'))
            
            # Demo Mode button
            tk.Button(frame, text='Demo Mode (Skip Login)', command=self.show_main, 
                     bg=LEMON_SOFT, fg=TEXT_PRIMARY, width=20, pady=5).pack(pady=5)
        else:
            tk.Frame(frame, bg=CARD_BG, height=24).pack()

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

        self.sidebar_host = tk.Frame(bottom, bg=SIDEBAR_BG)
        self.sidebar_host.pack(side='left', fill='y')
        self._build_sidebar(self.sidebar_host)

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
        elif icon_type == 'menu':
            c.create_line(4, 6, 18, 6, fill=FG, width=2, capstyle='round')
            c.create_line(4, 11, 18, 11, fill=FG, width=2, capstyle='round')
            c.create_line(4, 16, 18, 16, fill=FG, width=2, capstyle='round')

    def _topbar_btn(self, parent, icon_type, label, bg, command=None, side='right'):
        """Icon-over-text button with crisp Canvas-drawn icons."""
        hover = TOPBAR_BTN_HOV

        f = tk.Frame(parent, bg=bg, padx=16, cursor='hand2')
        f.pack(side=side, fill='y')

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

    def _topbar_compact_btn(self, parent, icon_type, label, bg, command=None, side='right'):
        """Compact topbar button with icon and label on one line."""
        hover = TOPBAR_BTN_HOV

        f = tk.Frame(parent, bg=bg, padx=10, cursor='hand2')
        f.pack(side=side, fill='y')

        inner = tk.Frame(f, bg=bg)
        inner.pack(expand=True, pady=10)

        ic = tk.Canvas(inner, width=22, height=22, bg=bg, highlightthickness=0)
        ic.pack(side='left')
        self._draw_topbar_icon(ic, icon_type, bg)

        txt = tk.Label(inner, text=label, bg=bg, fg=TOPBAR_ICON_FG,
                       font=(FF, 8, 'bold'))
        txt.pack(side='left', padx=(6, 0))

        all_w = [f, inner, ic, txt]
        if command:
            for w in all_w:
                w.bind('<Button-1>', lambda e: command())

        def _on_enter(e):
            for w in [f, inner, txt]:
                w.config(bg=hover)
            ic.config(bg=hover)
            self._draw_topbar_icon(ic, icon_type, hover)

        def _on_leave(e):
            for w in [f, inner, txt]:
                w.config(bg=bg)
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
        self._topbar_bar = bar
        bar.bind('<Configure>', self._update_topbar_responsive)
        self._topbar_compact_btn(bar, 'menu', 'Menu', TOPBAR_BTN_BG,
                                 command=self._toggle_sidebar, side='left')

        # ── Right buttons – packed right→left so visual order = left→right ─
        # About
        self._topbar_compact_btn(bar, 'about', 'About', TOPBAR_BTN_BG,
                                 command=lambda: messagebox.showinfo(
                                     'About', 'School MIS v1.0\nMT Olives Adventist School'))
        # Log Out  (License removed)
        self._topbar_compact_btn(bar, 'logout', 'Log Out', TOPBAR_BTN_BG,
                                 command=self.logout)

        # ── User section (purple) ──────────────────────────────────────────
        usr_f = tk.Frame(bar, bg=TOPBAR_USER_BG, padx=10, cursor='hand2')
        usr_f.pack(side='right', fill='y')

        # Crisp Canvas person icon
        av = tk.Canvas(usr_f, width=28, height=28,
                       bg=TOPBAR_USER_BG, highlightthickness=0)
        av.pack(side='left', pady=11)
        av.create_oval(7, 1, 21, 15, fill='white', outline='')       # head
        av.create_arc(1, 13, 27, 31, start=0, extent=180,
                      fill='white', outline='')                        # shoulders

        tk.Label(usr_f, text=f'  admin | {role}',
                 bg=TOPBAR_USER_BG, fg='white',
                 font=(FF, 9, 'bold')).pack(side='left', pady=14)

        # ── Academic year badge (green) ───────────────────────────────────
        yr_f = tk.Frame(bar, bg=TOPBAR_YR_BG, padx=10, cursor='hand2')
        yr_f.pack(side='right', fill='y', padx=(0, 2))

        # Canvas checkmark
        ck = tk.Canvas(yr_f, width=22, height=18,
                       bg=TOPBAR_YR_BG, highlightthickness=0)
        ck.pack(side='left', pady=16)
        ck.create_line(2, 9, 8, 15, fill='white', width=2.5,
                       capstyle='round', joinstyle='round')
        ck.create_line(8, 15, 20, 4, fill='white', width=2.5,
                       capstyle='round', joinstyle='round')

        tk.Label(yr_f, text=acad_year, bg=TOPBAR_YR_BG, fg='white',
                 font=(FF, 7, 'bold')).pack(side='left', padx=(4, 0), pady=16)

        # ── CBC Level Selector ────────────────────────────────────────────────
        level_frame = tk.Frame(bar, bg=OLIVE_PRIMARY, padx=8, cursor='hand2')
        level_frame.pack(side='right', fill='y', padx=(8, 2))
        
        tk.Label(level_frame, text='CBC:', bg=OLIVE_PRIMARY, fg='white',
                 font=(FF, 8, 'bold')).pack(side='left', pady=14)
        
        # Level dropdown
        self.level_var = tk.StringVar(value=self.current_level)
        level_cb = ttk.Combobox(level_frame, textvariable=self.level_var, 
                                values=[ALL_SCHOOL_LEVEL] + LEVELS, state='readonly', width=18)
        level_cb.pack(side='left', padx=(4, 0), pady=11)
        level_cb.bind('<<ComboboxSelected>>', lambda e: self._on_level_change())
        
        # ── Live clock ────────────────────────────────────────────────────
        dt_lbl = tk.Label(bar, bg=TOPBAR_RIGHT_BG, fg='#c8e6c9',
                          font=(FF, 9), padx=10)
        dt_lbl.pack(side='right', fill='y', pady=14)
        self._topbar_clock_label = dt_lbl
        self._topbar_clock_visible = True
        self._update_topbar_responsive()
        self._topbar_clock_generation += 1
        self._tick_topbar_clock(self._topbar_clock_generation)

    def _tick_topbar_clock(self, generation=None):
        try:
            if generation is not None and generation != self._topbar_clock_generation:
                self._topbar_clock_job = None
                return
            if not self._topbar_clock_label or not self._topbar_clock_label.winfo_exists():
                self._topbar_clock_job = None
                return
            self._topbar_clock_label.config(
                text=datetime.now().strftime('%a, %d %b %Y  %H:%M:%S')
            )
            self._topbar_clock_job = self.root.after(
                1000, lambda gen=generation: self._tick_topbar_clock(gen)
            )
        except Exception:
            self._topbar_clock_job = None

    def _update_topbar_responsive(self, event=None):
        """Hide lower-priority topbar elements when the window gets narrow."""
        if not self._topbar_bar or not self._topbar_bar.winfo_exists():
            return
        if not self._topbar_clock_label or not self._topbar_clock_label.winfo_exists():
            return

        width = event.width if event is not None else self._topbar_bar.winfo_width()
        should_show_clock = width >= 1500

        if should_show_clock and not self._topbar_clock_visible:
            self._topbar_clock_label.pack(side='right', fill='y', pady=14)
            self._topbar_clock_visible = True
        elif not should_show_clock and self._topbar_clock_visible:
            self._topbar_clock_label.pack_forget()
            self._topbar_clock_visible = False

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
                f'Switched to: {new_level}\n\nClasses available:\n' +
                '\n'.join(f'• {c}' for c in self.get_current_classes()[:6]) +
                ('...' if len(self.get_current_classes()) > 6 else '') +
                f'\n\nSubjects updated to:\n' + 
                '\n'.join(f'• {s}' for s in self.get_current_subjects()[:5]) +
                ('...' if len(self.get_current_subjects()) > 5 else ''))

    def _toggle_sidebar(self):
        """Collapse or expand the sidebar."""
        self.sidebar_collapsed = not self.sidebar_collapsed
        self._rebuild_sidebar()

    def _rebuild_sidebar(self):
        """Re-render the sidebar in its current collapsed/expanded state."""
        if not self.sidebar_host or not self.sidebar_host.winfo_exists():
            return
        for child in self.sidebar_host.winfo_children():
            child.destroy()
        self._build_sidebar(self.sidebar_host)
        if self.active_nav:
            self._set_nav(self.active_nav)

    def _build_sidebar(self, parent):
        sidebar_width = 72 if self.sidebar_collapsed else 220
        sb = tk.Frame(parent, bg=SIDEBAR_BG, width=sidebar_width)
        sb.pack(side='left', fill='y')
        sb.pack_propagate(False)

        # --- logo ---
        logo_row = tk.Frame(sb, bg=SIDEBAR_BG)
        logo_row.pack(fill='x', padx=8 if self.sidebar_collapsed else 14, pady=(8, 8))

        if self.logo_img:
            tk.Label(logo_row, image=self.logo_img, bg=SIDEBAR_BG).pack(side='left')
        else:
            icon_c = tk.Canvas(logo_row, width=40, height=40, bg=SIDEBAR_BG, highlightthickness=0)
            icon_c.pack(side='left')
            icon_c.create_oval(0, 0, 40, 40, fill=SIDEBAR_ACTIVE, outline=SIDEBAR_ACTIVE)
            icon_c.create_text(20, 20, text='\U0001f393', font=(FF, 17))

        if not self.sidebar_collapsed:
            title_box = tk.Frame(logo_row, bg=SIDEBAR_BG)
            title_box.pack(side='left', padx=8)
            tk.Label(title_box, text='MT OLIVES', bg=SIDEBAR_BG, fg='white',
                     font=(FF, 10, 'bold')).pack(anchor='w')
            tk.Label(title_box, text='ADVENTIST SCHOOL', bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                     font=(FF, 7)).pack(anchor='w')

        tk.Frame(sb, bg='#2E7D32', height=1).pack(fill='x', padx=8 if self.sidebar_collapsed else 14, pady=8)

        # --- nav items ---
        nav_cfg = self._get_role_based_nav()
        nav_wrap = tk.Frame(sb, bg=SIDEBAR_BG)
        nav_wrap.pack(fill='both', expand=True, padx=(4, 2) if self.sidebar_collapsed else (8, 4), pady=(0, 6))

        nav_canvas = tk.Canvas(nav_wrap, bg=SIDEBAR_BG, highlightthickness=0, bd=0)
        nav_scroll = ttk.Scrollbar(nav_wrap, orient='vertical', command=nav_canvas.yview)
        nav_box = tk.Frame(nav_canvas, bg=SIDEBAR_BG)

        nav_box.bind(
            '<Configure>',
            lambda e: nav_canvas.configure(scrollregion=nav_canvas.bbox('all'))
        )

        nav_window = nav_canvas.create_window((0, 0), window=nav_box, anchor='nw')

        def _resize_sidebar_nav(event):
            nav_canvas.itemconfigure(nav_window, width=event.width)

        def _scroll_sidebar_nav(event):
            nav_canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')

        nav_canvas.bind('<Configure>', _resize_sidebar_nav)
        nav_canvas.bind('<MouseWheel>', _scroll_sidebar_nav)
        nav_box.bind('<MouseWheel>', _scroll_sidebar_nav)

        nav_canvas.configure(yscrollcommand=nav_scroll.set)
        nav_canvas.pack(side='left', fill='both', expand=True)
        if not self.sidebar_collapsed:
            nav_scroll.pack(side='right', fill='y')
        self.nav_frames = {}

        for item in nav_cfg:
            if item[0] == 'section':
                self._nav_section(nav_box, item[1])
            else:
                icon, label, cmd = item
                self._nav_item(nav_box, icon, label, cmd)

        # --- bottom ---
        bot = tk.Frame(sb, bg=SIDEBAR_BG)
        bot.pack(side='bottom', fill='x', padx=8 if self.sidebar_collapsed else 14, pady=12)
        tk.Frame(bot, bg='#2E7D32', height=1).pack(fill='x', pady=(0, 8))

        uname = self.current_user.get('username', 'admin') if self.current_user else 'admin'
        email = uname if '@' in uname else uname + '@school.ac'
        if not self.sidebar_collapsed:
            tk.Label(bot, text=email, bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                     font=(FF, 9), anchor='w').pack(fill='x', pady=(0, 6))

        so = tk.Frame(bot, bg=SIDEBAR_BG, cursor='hand2')
        so.pack(fill='x')
        sign_out_text = '\u2192' if self.sidebar_collapsed else '\u2192  Sign Out'
        lbl_so = tk.Label(
            so,
            text=sign_out_text,
            bg=SIDEBAR_BG,
            fg=SIDEBAR_TEXT,
            font=(FF, 10),
            anchor='center' if self.sidebar_collapsed else 'w',
            padx=4,
            pady=5
        )
        lbl_so.pack(fill='x')

        for w in (so, lbl_so):
            w.bind('<Button-1>', lambda e: self.logout())
            w.bind('<Enter>', lambda e: [so.config(bg=SIDEBAR_HOVER), lbl_so.config(bg=SIDEBAR_HOVER)])
            w.bind('<Leave>', lambda e: [so.config(bg=SIDEBAR_BG), lbl_so.config(bg=SIDEBAR_BG)])

    def _get_role_based_nav(self):
        """Get navigation items based on user role."""
        role = getattr(self, 'user_role', 'admin')
        if role == 'admin':
            return [
                ('section', 'Main', None),
                ('🏠', 'Dashboard', self.show_dashboard),

                ('section', 'Academics', None),
                ('🏫', 'Classes', self.show_settings_classes),
                ('📚', 'Subjects', self.show_settings_subjects),
                ('👩‍🏫', 'Teachers', self.show_settings_teachers),

                ('section', 'Students', None),
                ('👥', 'Students', self.show_students),
                ('📝', 'Enter Marks', self.show_marks_entry),
                ('📊', 'Results', self.show_reports),
                ('📄', 'Report Cards', self.show_report_cards),
                ('🎓', 'Promotions', self.show_promotions),

                ('section', 'Analytics', None),
                ('📈', 'Charts', self.show_charts),
                ('📊', 'Exam Analytics', self.show_exam_analytics),
                ('📏', 'Grading Scale', self.show_settings_grading),

                ('section', 'System', None),
                ('⚙️', 'Settings', self.show_settings),
            ]

        if role == 'teacher':
            return [
                ('section', 'Main', None),
                ('🏠', 'Dashboard', self.show_dashboard),
                ('section', 'Students', None),
                ('📝', 'Enter Marks', self.show_marks_entry),
            ]

        if role == 'class_teacher':
            return [
                ('section', 'Main', None),
                ('🏠', 'Dashboard', self.show_dashboard),
                ('section', 'Students', None),
                ('👥', 'My Students', self.show_class_students),
                ('📝', 'Enter Marks', self.show_marks_entry),
                ('💬', 'Add Comments', self.show_add_comments),
                ('📄', 'Report Cards', self.show_report_cards),
            ]

        return [
            ('section', 'Main', None),
            ('🏠', 'Dashboard', self.show_dashboard),
        ]

    def _nav_section(self, parent, title: str):
        if self.sidebar_collapsed:
            tk.Frame(parent, bg='#2E7D32', height=1).pack(fill='x', padx=10, pady=(6, 4))
            return
        tk.Label(
            parent,
            text=title.upper(),
            bg=SIDEBAR_BG,
            fg='#c8d6a2',
            font=(FF, 8, 'bold'),
            anchor='w',
            padx=10,
            pady=6,
        ).pack(fill='x', pady=(5, 1))

    def _nav_item(self, parent, icon: str, label: str, cmd):
        frame = tk.Frame(parent, bg=SIDEBAR_BG, cursor='hand2')
        frame.pack(fill='x', pady=1)

        row = tk.Frame(frame, bg=SIDEBAR_BG, padx=6 if self.sidebar_collapsed else 10, pady=6)
        row.pack(fill='x')

        ico = tk.Label(row, text=icon, bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                       font=(FF, 12), width=2, anchor='center')
        ico.pack(side='left')
        txt = tk.Label(
            row,
            text='' if self.sidebar_collapsed else f'  {label}',
            bg=SIDEBAR_BG,
            fg=SIDEBAR_TEXT,
            font=(FF, 10),
            anchor='w'
        )
        if not self.sidebar_collapsed:
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

    def _all_tree_items(self, tree, parent=''):
        items = []
        for iid in tree.get_children(parent):
            items.append(iid)
            items.extend(self._all_tree_items(tree, iid))
        return items

    def _select_all_tree_rows(self, tree):
        items = self._all_tree_items(tree)
        if items:
            tree.selection_set(items)

    def _clear_tree_selection(self, tree):
        selected = tree.selection()
        if selected:
            tree.selection_remove(selected)

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

    # ==================== CLASSES ====================
    def show_classes(self):
        """Show classes inside the main content area with exam history and results."""
        self.clear_frame()
        self._set_nav('Classes')
        self._page_header('Classes', 'Browse stored classes, exam sessions, and detailed class results')

        classes_data = db.get_all_classes_exam_history()
        classes_map = {row.get('class_name', ''): row for row in classes_data}

        controls = tk.Frame(self.content_frame, bg=CONTENT_BG)
        controls.pack(fill='x', pady=(0, 12))

        tk.Label(controls, text='Class:', bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(side='left', padx=(10, 4))
        class_names = [row.get('class_name', '') for row in classes_data]
        class_var = tk.StringVar(value=class_names[0] if class_names else '')
        class_cb = ttk.Combobox(controls, textvariable=class_var, values=class_names,
                                state='readonly', style='App.TCombobox', width=18)
        class_cb.pack(side='left', ipady=4)

        tk.Label(controls, text='Stream:', bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(side='left', padx=(16, 4))
        stream_var = tk.StringVar(value='All Streams')
        stream_cb = ttk.Combobox(controls, textvariable=stream_var, values=['All Streams'],
                                 state='readonly', style='App.TCombobox', width=14)
        stream_cb.pack(side='left', ipady=4)

        tk.Label(controls, text='Search:', bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(side='left', padx=(16, 4))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(controls, textvariable=search_var, style='App.TEntry', width=26)
        search_entry.pack(side='left', ipady=4)
        tk.Button(
            controls, text='🔍 Search', bg=OLIVE_PRIMARY, fg='white',
            activebackground=OLIVE_DARK, activeforeground='white',
            font=(FF, 9, 'bold'), padx=10, pady=4, relief='flat', cursor='hand2',
            command=lambda: apply_search()
        ).pack(side='left', padx=(8, 0))

        tk.Label(controls, text='Exam Type:', bg=CONTENT_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(side='left', padx=(16, 4))
        exam_filter_var = tk.StringVar(value='All')
        exam_filter_cb = ttk.Combobox(
            controls,
            textvariable=exam_filter_var,
            values=['All'] + EXAM_TYPES,
            state='readonly',
            style='App.TCombobox',
            width=12,
        )
        exam_filter_cb.pack(side='left', ipady=4)

        actions_top = tk.Frame(controls, bg=CONTENT_BG)
        actions_top.pack(side='right')

        body = tk.Frame(self.content_frame, bg=CONTENT_BG)
        body.pack(fill='both', expand=True)
        body.columnconfigure(0, weight=4)
        body.columnconfigure(1, weight=5)
        body.rowconfigure(0, weight=1)

        l_bo, l_bi = _card_colors('mint')
        left_outer = tk.Frame(body, bg=l_bo)
        left_outer.grid(row=0, column=0, sticky='nsew', padx=(0, 8))
        left_card = tk.Frame(left_outer, bg=l_bi)
        left_card.pack(fill='both', expand=True, padx=1, pady=1)

        tk.Label(left_card, text='Stored Classes', bg=l_bi, fg=TEXT_PRIMARY,
                 font=(FF, 12, 'bold')).pack(anchor='w', padx=14, pady=(12, 6))

        class_tree_frame = tk.Frame(left_card, bg=l_bi)
        class_tree_frame.pack(fill='both', expand=True, padx=12, pady=(0, 12))
        class_cols = ('class', 'level', 'stream', 'students', 'exams', 'avg')
        class_tree = ttk.Treeview(class_tree_frame, columns=class_cols, show='headings', style='App.Treeview')
        class_tree.heading('class', text='Class')
        class_tree.heading('level', text='Level')
        class_tree.heading('stream', text='Stream')
        class_tree.heading('students', text='Students')
        class_tree.heading('exams', text='Exams')
        class_tree.heading('avg', text='Latest Avg')
        class_tree.column('class', width=120, anchor='w')
        class_tree.column('level', width=150, anchor='w')
        class_tree.column('stream', width=85, anchor='center')
        class_tree.column('students', width=70, anchor='center')
        class_tree.column('exams', width=70, anchor='center')
        class_tree.column('avg', width=85, anchor='center')
        class_scroll = ttk.Scrollbar(class_tree_frame, orient='vertical', command=class_tree.yview,
                                     style='App.Vertical.TScrollbar')
        class_tree.configure(yscrollcommand=class_scroll.set)
        class_tree.pack(side='left', fill='both', expand=True)
        class_scroll.pack(side='right', fill='y')

        rt_bo, rt_bi = _card_colors('sky')
        right_outer = tk.Frame(body, bg=rt_bo)
        right_outer.grid(row=0, column=1, sticky='nsew')
        right_card = tk.Frame(right_outer, bg=rt_bi)
        right_card.pack(fill='both', expand=True, padx=1, pady=1)

        summary = tk.Frame(right_card, bg=rt_bi, padx=16, pady=14)
        summary.pack(fill='x')
        class_title = tk.Label(summary, text='Select a class', bg=rt_bi, fg=TEXT_PRIMARY,
                               font=(FF, 14, 'bold'))
        class_title.pack(anchor='w')
        class_meta = tk.Label(summary, text='Choose a class from the list or use the class selector above.',
                              bg=rt_bi, fg=TEXT_SECONDARY, font=(FF, 10))
        class_meta.pack(anchor='w', pady=(4, 10))

        stats_row = tk.Frame(summary, bg=rt_bi)
        stats_row.pack(fill='x')
        stat_labels = {}
        for key, title, theme in [
            ('students', 'Students', 'mint'),
            ('exams', 'Exam Sessions', 'sky'),
            ('avg', 'Latest Average', 'sand'),
        ]:
            _, bg = _card_colors(theme)
            tile = tk.Frame(stats_row, bg=bg, padx=12, pady=10)
            tile.pack(side='left', padx=(0, 10))
            tk.Label(tile, text=title, bg=bg, fg=TEXT_SECONDARY, font=(FF, 9, 'bold')).pack(anchor='w')
            stat_labels[key] = tk.Label(tile, text='-', bg=bg, fg=TEXT_PRIMARY, font=(FF, 14, 'bold'))
            stat_labels[key].pack(anchor='w', pady=(4, 0))

        history_wrap = tk.Frame(right_card, bg=rt_bi, padx=16, pady=10)
        history_wrap.pack(fill='both', expand=True)
        tk.Label(history_wrap, text='Exam History', bg=rt_bi, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', pady=(0, 8))

        history_frame = tk.Frame(history_wrap, bg=rt_bi)
        history_frame.pack(fill='both', expand=True)
        exam_cols = ('term', 'exam_type', 'students', 'subjects', 'average')
        exam_tree = ttk.Treeview(history_frame, columns=exam_cols, show='headings', style='App.Treeview')
        exam_tree.heading('term', text='Term')
        exam_tree.heading('exam_type', text='Exam')
        exam_tree.heading('students', text='Students')
        exam_tree.heading('subjects', text='Subjects')
        exam_tree.heading('average', text='Class Avg')
        exam_tree.column('term', width=80, anchor='center')
        exam_tree.column('exam_type', width=120, anchor='center')
        exam_tree.column('students', width=80, anchor='center')
        exam_tree.column('subjects', width=80, anchor='center')
        exam_tree.column('average', width=90, anchor='center')
        exam_scroll = ttk.Scrollbar(history_frame, orient='vertical', command=exam_tree.yview,
                                    style='App.Vertical.TScrollbar')
        exam_tree.configure(yscrollcommand=exam_scroll.set)
        exam_tree.pack(side='left', fill='both', expand=True)
        exam_scroll.pack(side='right', fill='y')

        def format_avg(value):
            return '-' if value is None else f'{value}%'

        sort_state = {'classes': {}, 'history': {}}

        def _coerce_sort_value(raw_value):
            value = str(raw_value or '').strip()
            if value in ('', '-', 'No exam history'):
                return (0, '')
            if value.endswith('%'):
                try:
                    return (1, float(value[:-1]))
                except ValueError:
                    pass
            try:
                return (1, float(value))
            except ValueError:
                return (2, value.lower())

        def _sort_tree(tree, column, group_key):
            descending = sort_state[group_key].get(column, False)
            rows = [(tree.set(item, column), item) for item in tree.get_children('')]
            rows.sort(key=lambda row: _coerce_sort_value(row[0]), reverse=descending)
            for index, (_, item_id) in enumerate(rows):
                tree.move(item_id, '', index)
            sort_state[group_key][column] = not descending

        def normalize_exam_type(raw_exam_type):
            exam_value = str(raw_exam_type or '').strip().lower()
            if exam_value in ('opener', 'opening'):
                return 'Opener'
            if exam_value in ('mid-term', 'midterm', 'mid term'):
                return 'Mid-Term'
            if exam_value in ('end-term', 'endterm', 'end term'):
                return 'End-Term'
            return raw_exam_type or DEFAULT_EXAM_TYPE

        def get_stream_options(class_info):
            options = []
            class_name = class_info.get('class_name', '')
            class_row = class_info.get('class') or {}

            inline_stream = str(class_info.get('stream', '') or class_row.get('stream', '')).strip()
            if inline_stream and inline_stream != '-':
                options.append(inline_stream)

            class_id = class_row.get('id')
            if class_id:
                for row in db.get_streams_for_class(class_id):
                    name = str(row.get('name', '')).strip()
                    if name and name not in options:
                        options.append(name)

            try:
                conn = db.get_connection()
                cursor = conn.cursor()
                cursor.execute(
                    "SELECT DISTINCT stream FROM students WHERE class = ? AND stream IS NOT NULL AND TRIM(stream) != '' ORDER BY stream",
                    (class_name,)
                )
                for row in cursor.fetchall():
                    name = str(row['stream']).strip()
                    if name and name not in options:
                        options.append(name)
            except Exception:
                pass
            finally:
                try:
                    conn.close()
                except Exception:
                    pass

            return ['All Streams'] + options

        def clear_history_panel(message='Choose a class from the list or use the class selector above.'):
            class_title.config(text='Select a class')
            class_meta.config(text=message)
            for key in stat_labels:
                stat_labels[key].config(text='-')
            for item in exam_tree.get_children():
                exam_tree.delete(item)

        def get_selected_class_info():
            selected = class_tree.selection()
            if not selected:
                return None
            return classes_map.get(selected[0])

        def load_class_history(class_name):
            class_info = classes_map.get(class_name)
            if not class_info:
                clear_history_panel()
                return
            class_var.set(class_name)
            stream_options = get_stream_options(class_info)
            stream_cb['values'] = stream_options
            selected_stream = (stream_var.get() or '').strip()
            if selected_stream not in stream_options:
                selected_stream = 'All Streams'
                stream_var.set(selected_stream)

            selected_stream = stream_var.get().strip()
            stream_text = f" | Stream: {selected_stream}" if selected_stream != 'All Streams' else ''
            class_title.config(text=class_info.get('class_name', ''))
            class_meta.config(text=f"{class_info.get('level', '')}{stream_text}")

            for item in exam_tree.get_children():
                exam_tree.delete(item)

            all_exams = class_info.get('exams', [])
            selected_exam_type = exam_filter_var.get().strip()
            allowed_student_ids = None
            selected_stream = stream_var.get().strip()
            if selected_stream and selected_stream != 'All Streams':
                stream_students = [
                    s for s in db.get_students_by_class_and_stream(class_name, selected_stream)
                    if not self._is_summary_student_name(s.get('name'))
                ]
                allowed_student_ids = {s.get('id') for s in stream_students if s.get('id')}
                stat_labels['students'].config(text=str(len(stream_students)))
            else:
                stat_labels['students'].config(text=str(class_info.get('student_count', 0)))

            exams = []
            for index, exam in enumerate(all_exams):
                exam_type = normalize_exam_type(exam.get('exam_type', ''))
                if selected_exam_type != 'All' and exam_type != selected_exam_type:
                    continue
                exam_copy = dict(exam)
                exam_copy['_index'] = index
                exam_copy['exam_type'] = exam_type
                exams.append(exam_copy)

            visible_exam_count = 0
            latest_avg = None
            if not exams:
                stat_labels['exams'].config(text='0')
                stat_labels['avg'].config(text='-')
                if selected_exam_type == 'All':
                    exam_tree.insert('', 'end', values=('-', 'No exam history', '-', '-', '-'))
                else:
                    exam_tree.insert('', 'end', values=('-', f'No {selected_exam_type} history', '-', '-', '-'))
                return

            for exam in exams:
                term = exam.get('term', '')
                exam_type = exam.get('exam_type', '')
                results = db.get_class_exam_details(class_name, term, exam_type)
                if allowed_student_ids is not None:
                    results = [row for row in results if row.get('student_id') in allowed_student_ids]
                    if not results:
                        continue
                visible_exam_count += 1
                subject_count = len(results[0].get('marks', {})) if results else 0
                avg = round(sum(row.get('average', 0) for row in results) / len(results), 1) if results else None
                if latest_avg is None:
                    latest_avg = avg
                exam_tree.insert('', 'end', iid=f"{class_name}::{exam.get('_index', 0)}", values=(
                    term,
                    exam_type,
                    len(results),
                    subject_count,
                    format_avg(avg),
                ))
            stat_labels['exams'].config(text=str(visible_exam_count))
            stat_labels['avg'].config(text=format_avg(latest_avg))
            if visible_exam_count == 0:
                exam_tree.insert('', 'end', values=('-', 'No exam history for selected stream', '-', '-', '-'))

        def select_class_in_tree(class_name):
            if not class_name or class_name not in class_tree.get_children():
                return
            class_tree.selection_set(class_name)
            class_tree.focus(class_name)
            class_tree.see(class_name)
            load_class_history(class_name)

        def populate_classes(filtered_rows=None):
            rows = filtered_rows if filtered_rows is not None else classes_data
            visible_names = [row.get('class_name', '') for row in rows]
            class_cb['values'] = visible_names
            for item in class_tree.get_children():
                class_tree.delete(item)
            for cls_info in rows:
                class_tree.insert('', 'end', iid=cls_info['class_name'], values=(
                    cls_info.get('class_name', ''),
                    cls_info.get('level', ''),
                    cls_info.get('stream', '') or '-',
                    cls_info.get('student_count', 0),
                    cls_info.get('exam_count', 0),
                    format_avg(cls_info.get('latest_avg')),
                ))
            if rows:
                current = class_var.get().strip()
                target = current if current in visible_names else rows[0]['class_name']
                select_class_in_tree(target)
            else:
                class_var.set('')
                clear_history_panel('No classes match your search.')

        def open_selected_exam(event=None):
            class_info = get_selected_class_info()
            if not class_info:
                return
            exam_sel = exam_tree.selection()
            selected_stream = (stream_var.get() or '').strip()
            stream_filter = '' if selected_stream == 'All Streams' else selected_stream
            if not exam_sel:
                if class_info.get('exams'):
                    payload = dict(class_info)
                    if stream_filter:
                        payload['stream_filter'] = stream_filter
                    self._show_class_exam_details(payload)
                return
            exam_id = exam_sel[0]
            if '::' not in exam_id:
                return
            exam_index = int(exam_id.split('::', 1)[1])
            exams = class_info.get('exams', [])
            if 0 <= exam_index < len(exams):
                payload = dict(class_info)
                payload['exams'] = [exams[exam_index]]
                if stream_filter:
                    payload['stream_filter'] = stream_filter
                self._show_class_exam_details(payload)

        def apply_search(*_args):
            term = search_var.get().strip().lower()
            if not term:
                populate_classes()
                return
            filtered = []
            for cls_info in classes_data:
                haystack = ' '.join([
                    cls_info.get('class_name', ''),
                    cls_info.get('level', ''),
                    cls_info.get('stream', ''),
                ]).lower()
                if term in haystack:
                    filtered.append(cls_info)
            populate_classes(filtered)

        def refresh_page():
            nonlocal classes_data
            classes_data = db.get_all_classes_exam_history()
            classes_map.clear()
            classes_map.update({row.get('class_name', ''): row for row in classes_data})
            class_names = [row.get('class_name', '') for row in classes_data]
            class_cb['values'] = class_names
            populate_classes()

        search_var.trace_add('write', apply_search)
        search_entry.bind('<Return>', lambda _e: apply_search())
        class_tree.heading('class', text='Class', command=lambda: _sort_tree(class_tree, 'class', 'classes'))
        class_tree.heading('level', text='Level', command=lambda: _sort_tree(class_tree, 'level', 'classes'))
        class_tree.heading('stream', text='Stream', command=lambda: _sort_tree(class_tree, 'stream', 'classes'))
        class_tree.heading('students', text='Students', command=lambda: _sort_tree(class_tree, 'students', 'classes'))
        class_tree.heading('exams', text='Exams', command=lambda: _sort_tree(class_tree, 'exams', 'classes'))
        class_tree.heading('avg', text='Latest Avg', command=lambda: _sort_tree(class_tree, 'avg', 'classes'))

        exam_tree.heading('term', text='Term', command=lambda: _sort_tree(exam_tree, 'term', 'history'))
        exam_tree.heading('exam_type', text='Exam', command=lambda: _sort_tree(exam_tree, 'exam_type', 'history'))
        exam_tree.heading('students', text='Students', command=lambda: _sort_tree(exam_tree, 'students', 'history'))
        exam_tree.heading('subjects', text='Subjects', command=lambda: _sort_tree(exam_tree, 'subjects', 'history'))
        exam_tree.heading('average', text='Class Avg', command=lambda: _sort_tree(exam_tree, 'average', 'history'))

        class_cb.bind('<<ComboboxSelected>>', lambda e: select_class_in_tree(class_var.get().strip()))
        stream_cb.bind('<<ComboboxSelected>>',
                       lambda e: load_class_history(class_tree.selection()[0]) if class_tree.selection() else clear_history_panel())
        exam_filter_cb.bind('<<ComboboxSelected>>',
                            lambda e: load_class_history(class_tree.selection()[0]) if class_tree.selection() else clear_history_panel())
        class_tree.bind('<<TreeviewSelect>>', lambda e: load_class_history(class_tree.selection()[0]) if class_tree.selection() else clear_history_panel())
        exam_tree.bind('<Double-1>', open_selected_exam)

        tk.Button(actions_top, text='View Detailed Results', bg=OLIVE_PRIMARY, fg='white',
                  font=(FF, 10, 'bold'), padx=16, pady=8, command=open_selected_exam).pack(side='left', padx=(0, 8))
        tk.Button(actions_top, text='Refresh', bg=BLUE, fg='white',
                  font=(FF, 10, 'bold'), padx=16, pady=8, command=refresh_page).pack(side='left')

        populate_classes()

    def _show_class_exam_details(self, class_info):
        """Show detailed exam results for a specific class."""
        class_name = class_info.get('class_name', '')
        exams = class_info.get('exams', [])
        stream_filter = (class_info.get('stream_filter') or '').strip()
        
        if not exams:
            messagebox.showinfo('No Exams', f'No exam results available for {class_name}')
            return
        
        # Create dialog for selecting exam
        dialog = tk.Toplevel(self.root)
        dialog.title(f'Exam Results - {class_name}')
        dialog.geometry('1200x700')
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Header
        header = tk.Frame(dialog, bg=OLIVE_PRIMARY, padx=20, pady=15)
        header.pack(fill='x')
        tk.Label(header, text=f'Exam Results: {class_name}', bg=OLIVE_PRIMARY, fg='white',
                font=(FF, 14, 'bold')).pack(side='left')
        if stream_filter:
            tk.Label(header, text=f'Stream: {stream_filter}', bg=OLIVE_PRIMARY, fg='white',
                     font=(FF, 11, 'bold')).pack(side='right')
        
        # Exam selector
        selector_frame = tk.Frame(dialog, bg=CONTENT_BG, padx=20, pady=10)
        selector_frame.pack(fill='x')
        
        tk.Label(selector_frame, text='Select Exam:', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                font=(FF, 11)).pack(side='left', padx=(0, 10))
        
        exam_options = [f"Term {e.get('term', '')} - {e.get('exam_type', '')}" for e in exams]
        exam_var = tk.StringVar(value=exam_options[0] if exam_options else '')
        
        exam_combo = ttk.Combobox(selector_frame, textvariable=exam_var, values=exam_options,
                                   state='readonly', font=(FF, 10), width=25)
        exam_combo.pack(side='left', padx=(0, 10))
        
        # Results table frame
        table_frame = tk.Frame(dialog, bg=CONTENT_BG, padx=20, pady=10)
        table_frame.pack(fill='both', expand=True)
        
        # Create treeview - will be configured dynamically
        tree = ttk.Treeview(table_frame, show='headings', height=20)
        
        # Scrollbar
        tree_scroll = ttk.Scrollbar(table_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=tree_scroll.set)
        
        tree.pack(side='left', fill='both', expand=True)
        tree_scroll.pack(side='right', fill='y')
        
        def configure_tree(subjects):
            """Configure treeview columns based on available subjects."""
            # Clear existing columns
            tree.delete(*tree.get_children())
            for col in tree['columns']:
                tree.column(col, width=0)
                tree.heading(col, text='')
            
            # Define columns
            cols = ['position', 'name', 'admission'] + subjects + ['total', 'average', 'grade']
            tree['columns'] = cols
            
            # Configure column headings and widths
            tree.heading('position', text='#')
            tree.column('position', width=40, anchor='center')
            
            tree.heading('name', text='Student Name')
            tree.column('name', width=150, anchor='w')
            
            tree.heading('admission', text='Adm No')
            tree.column('admission', width=90, anchor='center')
            
            # Dynamic subject columns
            for subj in subjects:
                # Shorten long subject names for display
                short_name = subj[:12] if len(subj) > 12 else subj
                tree.heading(subj, text=short_name)
                tree.column(subj, width=55, anchor='center')
            
            tree.heading('total', text='Total')
            tree.column('total', width=60, anchor='center')
            
            tree.heading('average', text='Avg')
            tree.column('average', width=60, anchor='center')
            
            tree.heading('grade', text='Grade')
            tree.column('grade', width=50, anchor='center')
        
        def load_results():
            # Clear existing
            for item in tree.get_children():
                tree.delete(item)
            
            # Get selected exam
            selected_idx = exam_combo.current()
            if selected_idx < 0 or selected_idx >= len(exams):
                return
            
            exam = exams[selected_idx]
            term = exam.get('term', 'One')
            exam_type = exam.get('exam_type', 'End-Term')
            
            # Get results
            results = db.get_class_exam_details(class_name, term, exam_type)
            if stream_filter:
                stream_students = [
                    s for s in db.get_students_by_class_and_stream(class_name, stream_filter)
                    if not self._is_summary_student_name(s.get('name'))
                ]
                allowed_ids = {s.get('id') for s in stream_students if s.get('id')}
                results = [row for row in results if row.get('student_id') in allowed_ids]
                results.sort(key=lambda row: row.get('total', 0), reverse=True)
                for idx, row in enumerate(results, start=1):
                    row['position'] = idx
            
            # Get subjects from first result if available
            subjects = []
            if results:
                subjects = list(results[0].get('marks', {}).keys())
            
            # Configure tree with subjects
            configure_tree(subjects)
            
            # Insert into tree
            for r in results:
                marks = r.get('marks', {})
                values = [
                    r.get('position', ''),
                    r.get('student_name', ''),
                    r.get('admission_no', '')
                ]
                
                # Add marks for each subject
                for subj in subjects:
                    values.append(marks.get(subj, '-'))
                
                values.extend([
                    r.get('total', ''),
                    r.get('average', ''),
                    r.get('grade', '')
                ])
                tree.insert('', 'end', values=values)
        
        # Load button
        def on_exam_select(event=None):
            load_results()
        
        exam_combo.bind('<<ComboboxSelected>>', on_exam_select)
        
        tk.Button(selector_frame, text='Load Results', bg=OLIVE_PRIMARY, fg='white',
                 font=(FF, 10), padx=15, pady=5, command=load_results).pack(side='left')
        
        # Close button
        close_frame = tk.Frame(dialog, bg=CONTENT_BG, pady=10)
        close_frame.pack(fill='x')
        tk.Button(close_frame, text='Close', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 11), padx=20, pady=8, command=dialog.destroy).pack()
        
        # Load initial results
        load_results()

    # ==================== SETTINGS ====================
    def show_settings(self, initial_tab='classes', nav_label='Settings', show_tabs=True):
        """Show settings/admin page for managing classes, subjects, and teachers."""
        self.clear_frame()
        self._set_nav(nav_label)
        page_title = 'Settings' if show_tabs else nav_label
        self._page_header(page_title, 'Manage school configuration')

        tabs_bar = None
        if show_tabs:
            # Custom RGB tab bar
            tabs_outer = tk.Frame(self.content_frame, bg='#d1d5db', padx=1, pady=1)
            tabs_outer.pack(fill='x', padx=20, pady=(10, 0))
            tabs_bar = tk.Frame(tabs_outer, bg='#e5e7eb')
            tabs_bar.pack(fill='x', padx=1, pady=1)

        content_outer = tk.Frame(self.content_frame, bg='#d1d5db', padx=1, pady=1)
        top_pad = (0, 10) if show_tabs else (10, 10)
        content_outer.pack(fill='both', expand=True, padx=20, pady=top_pad)
        content_stack = tk.Frame(content_outer, bg=CONTENT_BG)
        content_stack.pack(fill='both', expand=True, padx=1, pady=1)

        tab_defs = [
            ('classes', '🏫 Classes', '#2563eb', self._build_classes_tab),
            ('subjects', '📚 Subjects', '#0f766e', self._build_subjects_tab),
            ('teachers', '👩‍🏫 Teachers', '#7c3aed', self._build_teachers_settings_tab),
            ('grading', '📏 Grading Scale', '#ea580c', self._build_grading_scale_tab),
        ]

        tab_buttons = {}
        tab_frames = {}
        neutral_bg = '#475569'

        for key, label, _color, builder in tab_defs:
            frame = tk.Frame(content_stack, bg=CONTENT_BG)
            frame.place(relx=0, rely=0, relwidth=1, relheight=1)
            builder(frame)
            tab_frames[key] = frame

        def activate_tab(active_key):
            for key, _label, color, _builder in tab_defs:
                if key == active_key:
                    tab_frames[key].lift()
                    if show_tabs:
                        btn = tab_buttons[key]
                        btn.configure(bg=color, fg='white', relief='sunken', bd=2)
                else:
                    if show_tabs:
                        btn = tab_buttons[key]
                        btn.configure(bg=neutral_bg, fg='white', relief='raised', bd=1)

        if show_tabs:
            for key, label, color, _builder in tab_defs:
                btn = tk.Button(
                    tabs_bar,
                    text=f'  {label}  ',
                    bg=neutral_bg,
                    fg='white',
                    activebackground=color,
                    activeforeground='white',
                    font=(FF, 10, 'bold'),
                    padx=10,
                    pady=8,
                    relief='raised',
                    bd=1,
                    cursor='hand2',
                    command=lambda k=key: activate_tab(k)
                )
                btn.pack(side='left', padx=2, pady=2)
                tab_buttons[key] = btn

        start_tab = initial_tab if initial_tab in tab_frames else 'classes'
        activate_tab(start_tab)

    def show_settings_classes(self):
        self.show_settings(initial_tab='classes', nav_label='Classes', show_tabs=False)

    def show_settings_streams(self):
        self.show_settings(initial_tab='classes', nav_label='Classes', show_tabs=False)

    def show_settings_subjects(self):
        self.show_settings(initial_tab='subjects', nav_label='Subjects', show_tabs=False)

    def show_settings_teachers(self):
        self.show_settings(initial_tab='teachers', nav_label='Teachers', show_tabs=False)

    def show_settings_grading(self):
        self.show_settings(initial_tab='grading', nav_label='Grading Scale', show_tabs=False)

    def show_settings_assignments(self):
        self.show_settings(initial_tab='teachers', nav_label='Teachers', show_tabs=False)

    def _build_classes_tab(self, parent, initial_scope='Both'):
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)

        tk.Button(
            toolbar, text='+ Add Class', bg=GREEN, fg='white',
            font=(FF, 10), padx=12, pady=5,
            command=lambda: self._open_class_dialog(on_save=refresh_tree)
        ).pack(side='left', padx=5)
        tk.Button(
            toolbar, text='+ Add Stream', bg='#0891b2', fg='white',
            font=(FF, 10), padx=12, pady=5,
            command=lambda: self._open_stream_dialog(on_save=refresh_tree)
        ).pack(side='left', padx=5)
        tk.Button(
            toolbar, text='Edit Selected', bg=BLUE, fg='white',
            font=(FF, 10), padx=12, pady=5, command=lambda: edit_selected()
        ).pack(side='left', padx=5)
        tk.Button(
            toolbar, text='Delete Selected', bg='#e74c3c', fg='white',
            font=(FF, 10), padx=12, pady=5, command=lambda: delete_selected()
        ).pack(side='left', padx=5)
        tk.Button(
            toolbar, text='Class Subjects Done', bg='#0f766e', fg='white',
            font=(FF, 10), padx=12, pady=5, command=self._open_class_subjects_done_dialog
        ).pack(side='left', padx=5)
        tk.Button(
            toolbar, text='Export CSV', bg='#1d4ed8', fg='white',
            font=(FF, 10), padx=12, pady=5, command=lambda: self._export_unified_classes_streams(tree)
        ).pack(side='left', padx=5)
        tk.Button(
            toolbar, text='Refresh', bg='#666', fg='white',
            font=(FF, 10), padx=12, pady=5, command=lambda: refresh_tree()
        ).pack(side='left', padx=5)

        filters = tk.Frame(parent, bg=CONTENT_BG)
        filters.pack(fill='x', padx=10, pady=(0, 4))
        tk.Label(filters, text='View:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=(0, 4))
        scope_var = tk.StringVar(value=initial_scope if initial_scope in ('Both', 'Classes Only', 'Streams Only') else 'Both')
        scope_cb = ttk.Combobox(
            filters, textvariable=scope_var, values=['Both', 'Classes Only', 'Streams Only'],
            state='readonly', style='App.TCombobox', width=14
        )
        scope_cb.pack(side='left', ipady=3)
        tk.Label(filters, text='Search:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=(14, 4))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(filters, textvariable=search_var, style='App.TEntry', width=30)
        search_entry.pack(side='left', ipady=3)
        tk.Button(
            filters, text='🔍 Search', bg=OLIVE_PRIMARY, fg='white',
            activebackground=OLIVE_DARK, activeforeground='white',
            font=(FF, 9, 'bold'), padx=10, pady=4, relief='flat', cursor='hand2',
            command=lambda: filtered(search_var.get(), scope_var.get())
        ).pack(side='left', padx=(8, 0))

        list_frame = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)

        cols = ('type', 'label', 'abbr', 'level', 'parent', 'students', 'subjects')
        tree = ttk.Treeview(list_frame, columns=cols, show='tree headings', style='App.Treeview', selectmode='extended')
        tree.heading('#0', text='Item')
        tree.heading('type', text='Type')
        tree.heading('label', text='Name')
        tree.heading('abbr', text='Short Label')
        tree.heading('level', text='Level')
        tree.heading('parent', text='Parent Class')
        tree.heading('students', text='Students')
        tree.heading('subjects', text='Subjects')

        tree.column('#0', width=200, anchor='w')
        tree.column('type', width=95, anchor='center')
        tree.column('label', width=170, anchor='w')
        tree.column('abbr', width=105, anchor='center')
        tree.column('level', width=210, anchor='w')
        tree.column('parent', width=150, anchor='w')
        tree.column('students', width=80, anchor='center')
        tree.column('subjects', width=80, anchor='center')

        select_all_var = tk.BooleanVar(value=False)
        select_row = tk.Frame(list_frame, bg=CARD_BG)
        select_row.pack(fill='x', padx=10, pady=(8, 0))
        tk.Checkbutton(
            select_row,
            text='Select all classes/streams',
            variable=select_all_var,
            bg=CARD_BG,
            fg=TEXT_SECONDARY,
            activebackground=CARD_BG,
            activeforeground=TEXT_PRIMARY,
            selectcolor=CARD_BG,
            font=(FF, 9, 'bold'),
            command=lambda: self._select_all_tree_rows(tree) if select_all_var.get() else self._clear_tree_selection(tree)
        ).pack(side='left')

        yscroll = ttk.Scrollbar(list_frame, orient='vertical', command=tree.yview, style='App.Vertical.TScrollbar')
        xscroll = ttk.Scrollbar(list_frame, orient='horizontal', command=tree.xview, style='App.Horizontal.TScrollbar')
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        tree.pack(side='left', fill='both', expand=True, padx=(10, 0), pady=10)
        yscroll.pack(side='right', fill='y', pady=(10, 0))
        xscroll.pack(side='bottom', fill='x', padx=10, pady=(0, 10))

        sort_state = {}
        class_cache = {}
        stream_cache = {}

        def _is_summary(student_row):
            return self._is_summary_student_name(student_row.get('name'))

        def _subject_count_for_class(class_name):
            try:
                return len(self._get_subjects_for_class(class_name, 'One', DEFAULT_EXAM_TYPE, for_reporting=True))
            except Exception:
                return 0

        def _stream_subject_count(class_name, stream_name):
            subjects = set()
            students = [s for s in db.get_students_by_class_and_stream(class_name, stream_name) if not _is_summary(s)]
            for student in students:
                sid = student.get('id')
                if not sid:
                    continue
                try:
                    conn = db.get_connection()
                    cursor = conn.cursor()
                    cursor.execute(
                        "SELECT DISTINCT subject FROM marks WHERE student_id = ? AND subject IS NOT NULL AND TRIM(subject) != ''",
                        (sid,)
                    )
                    for row in cursor.fetchall():
                        subjects.add(str(row['subject']).strip())
                except Exception:
                    pass
                finally:
                    try:
                        conn.close()
                    except Exception:
                        pass
            return len(subjects)

        def _coerce_sort_value(value):
            s = str(value or '').strip()
            if s in ('', '-', 'Class', 'Stream'):
                return (0, '')
            try:
                return (1, float(s))
            except ValueError:
                return (2, s.lower())

        def _flatten_rows():
            rows = []
            for cid in tree.get_children(''):
                rows.append(cid)
                for sid in tree.get_children(cid):
                    rows.append(sid)
            return rows

        def sort_by(column):
            reverse = sort_state.get(column, False)
            items = _flatten_rows()
            keyed = [(tree.set(iid, column), iid, tree.parent(iid)) for iid in items]
            keyed.sort(key=lambda row: (_coerce_sort_value(row[0]), row[2] != ''), reverse=reverse)
            # Rebuild top-level and child ordering while keeping hierarchy
            top = [iid for _v, iid, parent in keyed if parent == '']
            for idx, iid in enumerate(top):
                tree.move(iid, '', idx)
                children = [sid for _v, sid, p in keyed if p == iid]
                for cidx, sid in enumerate(children):
                    tree.move(sid, iid, cidx)
            sort_state[column] = not reverse

        def filtered(term, scope):
            q = term.strip().lower()
            for item in tree.get_children():
                tree.delete(item)

            for class_id, cls in class_cache.items():
                class_name = cls.get('name', '')
                class_level = cls.get('level', '')
                class_abbr = cls.get('abbreviation', '') or self._generate_short_label(class_name, 'class')
                class_students = [s for s in db.get_students_by_class_and_stream(class_name, '') if not _is_summary(s)]
                class_streams = stream_cache.get(class_id, [])
                stream_count_total = len(class_streams)
                class_subjects = _subject_count_for_class(class_name)

                class_haystack = ' '.join([class_name, class_level, class_abbr]).lower()
                class_match = (not q) or (q in class_haystack)
                include_class_row = scope in ('Both', 'Classes Only') and class_match
                include_stream_children = scope in ('Both', 'Streams Only')

                parent_iid = f'class::{class_id}'
                if include_class_row:
                    class_label = f'{class_name} ({stream_count_total} streams)'
                    tree.insert(
                        '', 'end', iid=parent_iid, text=f'🏫 {class_label}',
                        values=('Class', class_name, class_abbr, class_level, '-', len(class_students), class_subjects)
                    )

                stream_rows = []
                for stream in class_streams:
                    stream_name = stream.get('name', '')
                    stream_students = [s for s in db.get_students_by_class_and_stream(class_name, stream_name) if not _is_summary(s)]
                    stream_subjects = _stream_subject_count(class_name, stream_name)
                    stream_haystack = ' '.join([stream_name, class_name, class_level]).lower()
                    if q and q not in stream_haystack:
                        continue
                    stream_rows.append((stream, len(stream_students), stream_subjects))

                if include_stream_children and stream_rows:
                    attach_parent = parent_iid if include_class_row else ''
                    for stream, student_count, subject_count in stream_rows:
                        sid = stream.get('id', '')
                        tree.insert(
                            attach_parent, 'end', iid=f'stream::{sid}', text=f'🧩 {stream.get("name", "")}',
                            values=('Stream', stream.get('name', ''), '-', class_level, class_name, student_count, subject_count)
                        )

        def refresh_tree():
            nonlocal class_cache, stream_cache
            classes = db.get_all_classes()
            class_cache = {row.get('id', ''): row for row in classes}
            stream_cache = {}
            for class_id in class_cache:
                stream_cache[class_id] = db.get_streams_for_class(class_id)
            filtered(search_var.get(), scope_var.get())

        def selected_payload():
            sel = tree.selection()
            if not sel:
                return None, None
            iid = sel[0]
            if iid.startswith('class::'):
                class_id = iid.split('::', 1)[1]
                return 'class', class_cache.get(class_id)
            if iid.startswith('stream::'):
                stream_id = iid.split('::', 1)[1]
                for class_id, rows in stream_cache.items():
                    found = next((row for row in rows if row.get('id') == stream_id), None)
                    if found:
                        payload = dict(found)
                        payload['class_id'] = class_id
                        payload['class_name'] = class_cache.get(class_id, {}).get('name', '')
                        return 'stream', payload
            return None, None

        def edit_selected():
            selected = tree.selection()
            if len(selected) != 1:
                messagebox.showwarning('Select One', 'Please select exactly one class/stream to edit')
                return
            kind, payload = selected_payload()
            if not kind:
                messagebox.showwarning('Select', 'Please select a class or stream to edit')
                return
            if kind == 'class':
                self._open_class_dialog(payload, on_save=refresh_tree)
                return
            self._open_stream_dialog(payload, on_save=refresh_tree)

        def delete_selected():
            selected = list(tree.selection())
            if not selected:
                messagebox.showwarning('Select', 'Please select one or more classes/streams to delete')
                return

            class_ids = set()
            stream_ids = set()
            for iid in selected:
                if iid.startswith('class::'):
                    class_ids.add(iid.split('::', 1)[1])
                elif iid.startswith('stream::'):
                    stream_ids.add(iid.split('::', 1)[1])

            # Skip streams whose parent class is already selected (class delete cascades stream delete)
            if class_ids and stream_ids:
                covered_streams = set()
                for class_id in class_ids:
                    for stream in stream_cache.get(class_id, []):
                        sid = stream.get('id')
                        if sid:
                            covered_streams.add(sid)
                stream_ids = {sid for sid in stream_ids if sid not in covered_streams}

            if not class_ids and not stream_ids:
                messagebox.showwarning('Select', 'No valid classes/streams selected.')
                return

            total = len(class_ids) + len(stream_ids)
            if not messagebox.askyesno('Confirm Delete', f'Delete {total} selected item(s)?'):
                return

            errors = []
            # Delete classes first (cascades streams)
            for class_id in class_ids:
                if not db.delete_class(class_id):
                    errors.append(f'Class {class_id[:8]}')
            for stream_id in stream_ids:
                if not db.delete_stream(stream_id):
                    errors.append(f'Stream {stream_id[:8]}')

            if errors:
                messagebox.showerror('Delete Failed', 'Some items could not be deleted:\n' + '\n'.join(errors[:10]))
            else:
                messagebox.showinfo('Deleted', f'Deleted {total} item(s).')
            refresh_tree()

        for col, text in [
            ('type', 'Type'),
            ('label', 'Name'),
            ('abbr', 'Short Label'),
            ('level', 'Level'),
            ('parent', 'Parent Class'),
            ('students', 'Students'),
            ('subjects', 'Subjects'),
        ]:
            tree.heading(col, text=text, command=lambda c=col: sort_by(c))

        tree.bind('<Double-1>', lambda _e: edit_selected())
        scope_cb.bind('<<ComboboxSelected>>', lambda _e: filtered(search_var.get(), scope_var.get()))
        search_entry.bind('<Return>', lambda _e: filtered(search_var.get(), scope_var.get()))
        search_var.trace_add('write', lambda *_: filtered(search_var.get(), scope_var.get()))
        refresh_tree()

    def _load_classes(self, tree):
        for item in tree.get_children():
            tree.delete(item)
        
        classes = db.get_all_classes()
        for cls in classes:
            tree.insert('', 'end', iid=cls.get('id', ''), values=(
                cls.get('id', '')[:8],
                cls.get('name', ''),
                cls.get('abbreviation', '') or self._generate_short_label(cls.get('name', ''), 'class'),
                cls.get('level', ''),
                cls.get('stream', '') or '-'
            ))

    def _open_class_dialog(self, class_row=None, on_save=None):
        dialog = tk.Toplevel(self.root)
        dialog.title('Edit Class' if class_row else 'Add Class')
        dialog.geometry('430x390')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text='Class Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Short Label / Abbreviation:', font=(FF, 11)).pack(pady=(15, 5))
        abbr_entry = tk.Entry(dialog, font=(FF, 11))
        abbr_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Level:', font=(FF, 11)).pack(pady=(15, 5))
        level_var = tk.StringVar()
        level_cb = ttk.Combobox(dialog, textvariable=level_var, values=LEVELS, state='readonly', font=(FF, 10))
        level_cb.pack(fill='x', padx=20)

        tk.Label(dialog, text='Stream (optional):', font=(FF, 11)).pack(pady=(15, 5))
        stream_entry = tk.Entry(dialog, font=(FF, 11))
        stream_entry.pack(fill='x', padx=20)

        if class_row:
            name_entry.insert(0, class_row.get('name', ''))
            abbr_entry.insert(0, class_row.get('abbreviation', ''))
            level_var.set(class_row.get('level', ''))
            stream_entry.insert(0, class_row.get('stream', '') or '')
        else:
            level_var.set(self.current_level if self.current_level in LEVELS else LEVELS[0])

        def save():
            name = name_entry.get().strip()
            abbreviation = abbr_entry.get().strip()
            level = level_var.get().strip()
            stream = stream_entry.get().strip()

            if not name or not level:
                messagebox.showerror('Error', 'Class name and level are required')
                return

            if class_row:
                success, msg = db.update_class(class_row['id'], name, level, stream or None, abbreviation)
            else:
                success, msg = db.add_class(name, level, stream or None, abbreviation)
            if success:
                dialog.destroy()
                if callable(on_save):
                    on_save()
                else:
                    self.show_settings()
            else:
                messagebox.showerror('Error', msg)

        tk.Button(dialog, text='Save', bg=GREEN, fg='white',
                  font=(FF, 11), padx=20, pady=8, command=save).pack(pady=(12, 20))

    def _edit_class_dialog(self, tree):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select', 'Please select a class to edit')
            return
        class_id = selected[0]
        class_row = next((row for row in db.get_all_classes() if row.get('id') == class_id), None)
        if not class_row:
            messagebox.showerror('Error', 'Could not load the selected class')
            return
        self._open_class_dialog(class_row)
    
    def _delete_class(self, tree):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select', 'Please select a class to delete')
            return
        
        if not messagebox.askyesno('Confirm', 'Delete this class?'):
            return
        
        class_id = selected[0]
        
        if db.delete_class(class_id):
            messagebox.showinfo('Success', 'Class deleted')
            self._load_classes(tree)
        else:
            messagebox.showerror('Error', 'Failed to delete class')
    
    def _build_streams_tab(self, parent):
        """Unified classes + streams view focused on streams."""
        self._build_classes_tab(parent, initial_scope='Streams Only')
    
    def _open_stream_dialog(self, stream_row=None, on_save=None):
        """Dialog to add or edit a stream."""
        classes = db.get_all_classes()
        if not classes:
            messagebox.showwarning('No Classes', 'Please add classes first')
            return
        
        dialog = tk.Toplevel(self.root)
        dialog.title('Edit Stream' if stream_row else 'Add Stream')
        dialog.geometry('420x290')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        tk.Label(dialog, text='Stream Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Class:', font=(FF, 11)).pack(pady=(15, 5))
        class_var = tk.StringVar()
        class_cb = ttk.Combobox(dialog, textvariable=class_var, 
                               values=[c['name'] for c in classes], state='readonly', font=(FF, 10))
        class_cb.pack(fill='x', padx=20)

        if stream_row:
            name_entry.insert(0, stream_row.get('name', ''))
            class_name = stream_row.get('class_name', '')
            if not class_name and stream_row.get('class_id'):
                class_name = next((c.get('name', '') for c in classes if c.get('id') == stream_row.get('class_id')), '')
            class_var.set(class_name if class_name else classes[0]['name'])
        else:
            class_var.set(classes[0]['name'])
        
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
            
            if stream_row:
                success, msg = db.update_stream(stream_row.get('id', ''), name, class_id)
            else:
                success, msg = db.add_stream(name, class_id)
            if success:
                messagebox.showinfo('Success', msg)
                dialog.destroy()
                if callable(on_save):
                    on_save()
                else:
                    self.show_settings()
            else:
                messagebox.showerror('Error', msg)
        
        btn_row = tk.Frame(dialog, bg=dialog.cget('bg'))
        btn_row.pack(fill='x', pady=(16, 20), padx=20)
        tk.Button(btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=dialog.destroy).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(side='left')

    def _add_stream_dialog(self):
        """Backward-compatible add stream entry point."""
        self._open_stream_dialog()

    def _export_unified_classes_streams(self, tree):
        """Export unified classes/streams table to CSV."""
        path = filedialog.asksaveasfilename(
            title='Export Classes & Streams',
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile='classes_streams.csv'
        )
        if not path:
            return
        try:
            with open(path, 'w', newline='', encoding='utf-8') as fh:
                writer = csv.writer(fh)
                writer.writerow(['Type', 'Name', 'Short Label', 'Level', 'Parent Class', 'Students', 'Subjects'])
                for parent_id in tree.get_children(''):
                    parent_vals = tree.item(parent_id, 'values')
                    writer.writerow(list(parent_vals))
                    for child_id in tree.get_children(parent_id):
                        child_vals = tree.item(child_id, 'values')
                        writer.writerow(list(child_vals))
            messagebox.showinfo('Export Complete', f'Classes & Streams exported to:\n{path}')
        except Exception as exc:
            messagebox.showerror('Export Failed', f'Could not export CSV.\n\n{exc}')
    
    def _build_subjects_tab(self, parent):
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)
        
        tk.Button(toolbar, text='+ Add Subject', bg=GREEN, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._open_subject_dialog()).pack(side='left', padx=5)
        tk.Button(toolbar, text='Reset To Default List', bg=PURPLE, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._reset_subject_catalog(tree)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Edit Selected', bg=BLUE, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._edit_subject_dialog(tree)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Delete Selected', bg='#e74c3c', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._delete_subject(tree)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Class Subjects Done', bg='#0f766e', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=self._open_class_subjects_done_dialog).pack(side='left', padx=5)
        tk.Button(toolbar, text='Export Subjects', bg='#1d4ed8', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._export_subjects_table(tree)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Print Subjects', bg='#7c3aed', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._print_subjects_table(tree)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Refresh', bg='#666', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._load_subjects(tree)).pack(side='left', padx=5)

        search_row = tk.Frame(parent, bg=CONTENT_BG)
        search_row.pack(fill='x', padx=10, pady=(0, 4))
        tk.Label(search_row, text='Search:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=(0, 6))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_row, textvariable=search_var, style='App.TEntry', width=30)
        search_entry.pack(side='left', ipady=3)
        tk.Button(
            search_row, text='🔍 Search', bg=OLIVE_PRIMARY, fg='white',
            activebackground=OLIVE_DARK, activeforeground='white',
            font=(FF, 9, 'bold'), padx=10, pady=4, relief='flat', cursor='hand2',
            command=lambda: self._load_subjects(tree, search_var.get())
        ).pack(side='left', padx=(8, 0))

        list_frame = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        cols = ('id', 'name', 'abbr', 'level', 'category', 'optional', 'classes_done')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings', style='App.Treeview', selectmode='extended')
        tree.heading('id', text='Code')
        tree.heading('name', text='Subject Name')
        tree.heading('abbr', text='Short Label')
        tree.heading('level', text='Level')
        tree.heading('category', text='Category')
        tree.heading('optional', text='Optional')
        tree.heading('classes_done', text='Classes Taking This Subject')
        
        tree.column('id', width=70, anchor='center')
        tree.column('name', width=220)
        tree.column('abbr', width=120, anchor='center')
        tree.column('level', width=180)
        tree.column('category', width=120, anchor='center')
        tree.column('optional', width=90, anchor='center')
        tree.column('classes_done', width=250)

        select_all_var = tk.BooleanVar(value=False)
        select_row = tk.Frame(list_frame, bg=CARD_BG)
        select_row.pack(fill='x', padx=10, pady=(8, 0))
        tk.Checkbutton(
            select_row,
            text='Select all subjects',
            variable=select_all_var,
            bg=CARD_BG,
            fg=TEXT_SECONDARY,
            activebackground=CARD_BG,
            activeforeground=TEXT_PRIMARY,
            selectcolor=CARD_BG,
            font=(FF, 9, 'bold'),
            command=lambda: self._select_all_tree_rows(tree) if select_all_var.get() else self._clear_tree_selection(tree)
        ).pack(side='left')

        tree.pack(fill='both', expand=True, padx=10, pady=10)
        tree.bind('<Double-1>', lambda e: self._edit_subject_dialog(tree))
        self._subjects_col_labels = {
            'id': 'Code',
            'name': 'Subject Name',
            'abbr': 'Short Label',
            'level': 'Level',
            'category': 'Category',
            'optional': 'Optional',
            'classes_done': 'Classes Taking This Subject',
        }
        self._subjects_sort_state = {'column': None, 'reverse': False}
        self._configure_subjects_sorting(tree)
        search_entry.bind('<Return>', lambda _e: self._load_subjects(tree, search_var.get()))
        search_var.trace_add('write', lambda *_: self._load_subjects(tree, search_var.get()))
        self._load_subjects(tree, search_var.get())

    def _configure_subjects_sorting(self, tree):
        for col in tree['columns']:
            label = self._subjects_col_labels.get(col, col.title())
            tree.heading(col, text=label, command=lambda c=col: self._sort_subjects_tree(tree, c))

    def _subject_sort_value(self, col, value):
        text = str(value or '').strip()
        if col == 'optional':
            return 1 if text.lower() in ('yes', 'true', '1') else 0
        return text.lower()

    def _sort_subjects_tree(self, tree, col):
        state = getattr(self, '_subjects_sort_state', {'column': None, 'reverse': False})
        reverse = (state.get('column') == col and not state.get('reverse', False))
        rows = []
        for iid in tree.get_children(''):
            values = tree.item(iid).get('values', ())
            rows.append((iid, values))

        try:
            idx = list(tree['columns']).index(col)
        except ValueError:
            return

        rows.sort(key=lambda item: self._subject_sort_value(col, item[1][idx] if idx < len(item[1]) else ''), reverse=reverse)
        for position, (iid, _vals) in enumerate(rows):
            tree.move(iid, '', position)

        self._subjects_sort_state = {'column': col, 'reverse': reverse}
        for c in tree['columns']:
            base = self._subjects_col_labels.get(c, c.title())
            if c == col:
                base = f"{base} {'▼' if reverse else '▲'}"
            tree.heading(c, text=base, command=lambda key=c: self._sort_subjects_tree(tree, key))

    def _load_subjects(self, tree, search_query=''):
        for item in tree.get_children():
            tree.delete(item)

        # Get class subjects done mapping
        class_subjects_map = self._get_class_subjects_done_map()
        query = str(search_query or '').strip().lower()
        
        for subj in db.get_subjects_by_level():
            subject_name = subj.get('name', '')
            
            # Find which classes take this subject
            classes_taking = []
            for class_name, subjects in class_subjects_map.items():
                if subject_name in subjects:
                    classes_taking.append(class_name)
            
            # If no explicit mapping, check if subject is in the level's default subjects
            if not classes_taking:
                subject_level = subj.get('level', '')
                if subject_level in CLASSES_BY_LEVEL:
                    # Check if this subject is in the default subjects for this level
                    level_subjects = self._get_subjects_for_level(subject_level)
                    if subject_name in level_subjects:
                        classes_taking = CLASSES_BY_LEVEL[subject_level]
            
            classes_text = ', '.join(classes_taking) if classes_taking else 'Not assigned'

            if query:
                haystack = ' '.join([
                    str(subj.get('code', '') or subj.get('abbreviation', '')),
                    str(subj.get('name', '')),
                    str(subj.get('abbreviation', '') or subj.get('code', '')),
                    str(subj.get('level', '')),
                    str(subj.get('category', '')),
                    classes_text,
                ]).lower()
                if query not in haystack:
                    continue
            
            tree.insert('', 'end', iid=subj.get('id', ''), values=(
                subj.get('code', '') or subj.get('abbreviation', ''),
                subj.get('name', ''),
                subj.get('abbreviation', '') or subj.get('code', '') or self._generate_short_label(subj.get('name', ''), 'subject'),
                subj.get('level', ''),
                subj.get('category', ''),
                'Yes' if subj.get('is_optional') else 'No',
                classes_text
            ))

    def _export_subjects_table(self, tree):
        file_path = filedialog.asksaveasfilename(
            title='Export Subjects',
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile='subjects_list.csv'
        )
        if not file_path:
            return

        headers = [self._subjects_col_labels.get(col, col.title()) for col in tree['columns']]
        try:
            with open(file_path, 'w', newline='', encoding='utf-8') as f:
                writer = csv.writer(f)
                writer.writerow(headers)
                for iid in tree.get_children(''):
                    writer.writerow(tree.item(iid).get('values', ()))
            messagebox.showinfo('Export Complete', f'Subjects exported to:\n{file_path}')
        except Exception as exc:
            messagebox.showerror('Export Failed', f'Could not export subjects.\n\n{exc}')

    def _print_subjects_table(self, tree):
        headers = [self._subjects_col_labels.get(col, col.title()) for col in tree['columns']]
        rows = [tree.item(iid).get('values', ()) for iid in tree.get_children('')]
        if not rows:
            messagebox.showwarning('No Data', 'No subject rows available to print.')
            return

        col_widths = [len(h) for h in headers]
        for row in rows:
            for i, value in enumerate(row):
                col_widths[i] = max(col_widths[i], len(str(value)))

        def fmt_row(row_values):
            return ' | '.join(str(v).ljust(col_widths[i]) for i, v in enumerate(row_values))

        divider = '-+-'.join('-' * w for w in col_widths)
        lines = [
            'MT OLIVES ADVENTIST SCHOOL - SUBJECTS LIST',
            '',
            fmt_row(headers),
            divider,
        ]
        for row in rows:
            lines.append(fmt_row(row))

        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        print_path = os.path.join(tempfile.gettempdir(), f'subjects_print_{timestamp}.txt')
        try:
            with open(print_path, 'w', encoding='utf-8') as f:
                f.write('\n'.join(lines))
            try:
                os.startfile(print_path, 'print')
                messagebox.showinfo('Print Sent', 'Subjects list sent to printer.')
            except Exception:
                messagebox.showinfo('Print File Ready', f'Print file created:\n{print_path}')
        except Exception as exc:
            messagebox.showerror('Print Failed', f'Could not prepare print file.\n\n{exc}')

    def _get_class_subjects_done_map(self):
        raw = db.get_setting(CLASS_SUBJECTS_DONE_KEY, '{}')
        try:
            data = json.loads(raw or '{}')
        except Exception:
            data = {}
        cleaned = {}
        if isinstance(data, dict):
            for class_name, subjects in data.items():
                class_key = str(class_name or '').strip()
                if not class_key:
                    continue
                if isinstance(subjects, list):
                    cleaned[class_key] = [str(s).strip() for s in subjects if str(s).strip()]
        return cleaned

    def _save_class_subjects_done_map(self, mapping):
        payload = json.dumps(mapping or {}, ensure_ascii=True)
        db.set_setting(CLASS_SUBJECTS_DONE_KEY, payload)

    def _get_subject_pool_for_class(self, class_name):
        level = self._get_level_for_class(class_name)
        subjects = self._get_subjects_for_level(level)
        level_subjects = db.get_subjects_by_level(level)
        global_subjects = db.get_subjects_by_level(ALL_SUBJECT_LEVEL)
        combined_subjects = list(level_subjects or [])
        for row in (global_subjects or []):
            name = row.get('name', '')
            if name and not any(existing.get('name') == name for existing in combined_subjects):
                combined_subjects.append(row)

        if combined_subjects:
            custom_names = [row.get('name', '') for row in combined_subjects if row.get('name')]
            ordered = [subject for subject in subjects if subject in custom_names]
            for subject in custom_names:
                if subject not in ordered:
                    ordered.append(subject)
            subjects = ordered
        return subjects

    def _get_done_subjects_from_marks(self, class_name, term='One', exam_type=DEFAULT_EXAM_TYPE):
        present = set()
        for student in db.get_students_by_class(class_name):
            for subject in db.get_student_marks(student['id'], term, exam_type).keys():
                present.add(subject)

        if not present:
            return []

        ordered = []
        for subject in self._get_subject_pool_for_class(class_name):
            if subject in present:
                ordered.append(subject)
        for subject in sorted(present):
            if subject not in ordered:
                ordered.append(subject)
        return ordered

    def _open_class_subjects_done_dialog(self):
        dialog = tk.Toplevel(self.root)
        dialog.title('Class Subjects Done')
        dialog.geometry('860x620')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text='Class-specific subjects done for reports and report cards',
                 font=(FF, 11, 'bold')).pack(anchor='w', padx=20, pady=(16, 4))
        tk.Label(dialog, text='Only selected subjects will appear in Results/Report Cards for this class.',
                 font=(FF, 9), fg=TEXT_SECONDARY).pack(anchor='w', padx=20, pady=(0, 10))

        legend = tk.Frame(dialog, bg=dialog.cget('bg'))
        legend.pack(fill='x', padx=20, pady=(0, 10))

        def legend_chip(parent, color, text, fg='white'):
            chip = tk.Frame(parent, bg=color, padx=10, pady=4)
            chip.pack(side='left', padx=(0, 8))
            tk.Label(chip, text=text, bg=color, fg=fg, font=(FF, 9, 'bold')).pack()

        legend_chip(legend, '#16a34a', 'Add Subject')
        legend_chip(legend, '#f59e0b', 'Edit Subject', fg='#1f2937')
        legend_chip(legend, '#dc2626', 'Remove Subject')
        legend_chip(legend, '#2563eb', 'Auto From Marks')
        legend_chip(legend, '#7c3aed', 'Automatic Mode')

        classes = [row.get('name') for row in db.get_all_classes() if row.get('name')]
        if not classes:
            classes = self.get_current_classes()

        top = tk.Frame(dialog, bg=dialog.cget('bg'))
        top.pack(fill='x', padx=20, pady=(0, 8))

        tk.Label(top, text='Class:', font=(FF, 10, 'bold')).pack(side='left')
        class_var = tk.StringVar(value=classes[0] if classes else '')
        class_cb = ttk.Combobox(top, textvariable=class_var, values=classes, state='readonly', width=24)
        class_cb.pack(side='left', padx=(8, 14))

        tk.Label(top, text='Term:', font=(FF, 10, 'bold')).pack(side='left')
        term_var = tk.StringVar(value=TERMS[0])
        term_cb = ttk.Combobox(top, textvariable=term_var, values=TERMS, state='readonly', width=8)
        term_cb.pack(side='left', padx=(8, 8))

        tk.Label(top, text='Exam:', font=(FF, 10, 'bold')).pack(side='left')
        exam_var = tk.StringVar(value=DEFAULT_EXAM_TYPE)
        exam_cb = ttk.Combobox(top, textvariable=exam_var, values=EXAM_TYPES, state='readonly', width=12)
        exam_cb.pack(side='left', padx=(8, 0))

        list_wrap = tk.Frame(dialog, bg=CARD_BG, relief='flat', bd=1)
        list_wrap.pack(fill='both', expand=True, padx=20, pady=(4, 8))

        left_pane = tk.Frame(list_wrap, bg=CARD_BG)
        left_pane.pack(side='left', fill='both', expand=True, padx=(10, 6), pady=10)
        tk.Label(left_pane, text='Available Subjects', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold')).pack(anchor='w', pady=(0, 6))
        available_list = tk.Listbox(left_pane, selectmode='extended', font=(FF, 10), activestyle='none')
        available_list.pack(side='left', fill='both', expand=True)
        available_sb = ttk.Scrollbar(left_pane, orient='vertical', command=available_list.yview, style='App.Vertical.TScrollbar')
        available_sb.pack(side='right', fill='y')
        available_list.configure(yscrollcommand=available_sb.set)

        middle_pane = tk.Frame(list_wrap, bg=CARD_BG)
        middle_pane.pack(side='left', fill='y', padx=8, pady=10)

        right_pane = tk.Frame(list_wrap, bg=CARD_BG)
        right_pane.pack(side='left', fill='both', expand=True, padx=(6, 10), pady=10)
        tk.Label(right_pane, text='Subjects Done (Used in Reports)', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold')).pack(anchor='w', pady=(0, 6))
        selected_list = tk.Listbox(right_pane, selectmode='extended', font=(FF, 10), activestyle='none')
        selected_list.pack(side='left', fill='both', expand=True)
        selected_sb = ttk.Scrollbar(right_pane, orient='vertical', command=selected_list.yview, style='App.Vertical.TScrollbar')
        selected_sb.pack(side='right', fill='y')
        selected_list.configure(yscrollcommand=selected_sb.set)

        all_subjects = []
        selected_subjects = []

        def subject_line(subject, cls_name):
            label = self._get_subject_label(subject, cls_name).replace('\n', ' / ')
            return f'{label}  -  {subject}'

        def render_lists():
            cls_name = class_var.get().strip()
            available_list.delete(0, tk.END)
            selected_list.delete(0, tk.END)
            for subject in all_subjects:
                if subject not in selected_subjects:
                    available_list.insert(tk.END, subject_line(subject, cls_name))
            for subject in selected_subjects:
                selected_list.insert(tk.END, subject_line(subject, cls_name))

        def reload_subject_list(use_auto=False):
            nonlocal all_subjects, selected_subjects
            cls = class_var.get().strip()
            if not cls:
                return
            pool = self._get_subject_pool_for_class(cls)
            marks_subjects = self._get_done_subjects_from_marks(cls, term_var.get(), exam_var.get())
            mapping = self._get_class_subjects_done_map()
            configured = mapping.get(cls, [])
            all_subjects = list(pool)
            for subject in marks_subjects + configured:
                if subject not in all_subjects:
                    all_subjects.append(subject)

            selected = list(configured)
            if use_auto:
                selected = list(marks_subjects)
            elif not configured:
                selected = list(marks_subjects or pool)

            selected_subjects = [subject for subject in selected if subject in all_subjects] + [
                subject for subject in selected if subject not in all_subjects
            ]
            render_lists()

        class_cb.bind('<<ComboboxSelected>>', lambda _e: reload_subject_list())
        term_cb.bind('<<ComboboxSelected>>', lambda _e: reload_subject_list())
        exam_cb.bind('<<ComboboxSelected>>', lambda _e: reload_subject_list())

        def button(master, text, bg, fg='white', cmd=None):
            tk.Button(
                master, text=text, bg=bg, fg=fg, activebackground=bg, activeforeground=fg,
                font=(FF, 10, 'bold'), padx=10, pady=8, relief='flat', cursor='hand2',
                command=cmd
            ).pack(fill='x', pady=5)

        def add_subjects():
            chosen = [available_list.get(i) for i in available_list.curselection()]
            if not chosen:
                messagebox.showwarning('Select Subject', 'Select one or more available subjects to add.', parent=dialog)
                return
            cls = class_var.get().strip()
            for line in chosen:
                raw = line.split(' - ', 1)[1] if ' - ' in line else line
                raw = raw.strip()
                if raw and raw not in selected_subjects:
                    selected_subjects.append(raw)
                    if raw not in all_subjects:
                        all_subjects.append(raw)
            render_lists()

        def remove_subjects():
            chosen_idx = list(selected_list.curselection())
            if not chosen_idx:
                messagebox.showwarning('Select Subject', 'Select one or more subjects done to remove.', parent=dialog)
                return
            chosen_values = [selected_list.get(i) for i in chosen_idx]
            remove_raw = []
            for line in chosen_values:
                raw = line.split(' - ', 1)[1] if ' - ' in line else line
                remove_raw.append(raw.strip())
            selected_subjects[:] = [s for s in selected_subjects if s not in remove_raw]
            render_lists()

        def edit_subject():
            chosen_idx = list(selected_list.curselection())
            if len(chosen_idx) != 1:
                messagebox.showwarning('Select One', 'Select exactly one subject done to edit.', parent=dialog)
                return
            current_line = selected_list.get(chosen_idx[0])
            current_raw = current_line.split(' - ', 1)[1].strip() if ' - ' in current_line else current_line.strip()
            new_raw = simpledialog.askstring(
                'Edit Subject',
                'Update subject name (exact subject key used in reports):',
                initialvalue=current_raw,
                parent=dialog
            )
            if new_raw is None:
                return
            new_raw = new_raw.strip()
            if not new_raw:
                messagebox.showwarning('Invalid Value', 'Subject name cannot be empty.', parent=dialog)
                return
            selected_subjects[chosen_idx[0]] = new_raw
            if new_raw not in all_subjects:
                all_subjects.append(new_raw)
            render_lists()

        button(middle_pane, 'Add  >>', '#16a34a', 'white', add_subjects)
        button(middle_pane, 'Edit', '#f59e0b', '#1f2937', edit_subject)
        button(middle_pane, '<<  Remove', '#dc2626', 'white', remove_subjects)

        actions = tk.Frame(dialog, bg=dialog.cget('bg'))
        actions.pack(fill='x', padx=20, pady=(0, 12))

        def auto_pick_from_marks():
            reload_subject_list(use_auto=True)

        def clear_class_override():
            cls = class_var.get().strip()
            if not cls:
                return
            mapping = self._get_class_subjects_done_map()
            if cls in mapping:
                mapping.pop(cls, None)
                self._save_class_subjects_done_map(mapping)
            messagebox.showinfo('Updated', f'{cls} now uses automatic subjects (from entered marks).', parent=dialog)
            reload_subject_list()

        def save_selection():
            cls = class_var.get().strip()
            if not cls:
                messagebox.showwarning('Missing Class', 'Please select a class.', parent=dialog)
                return
            mapping = self._get_class_subjects_done_map()
            mapping[cls] = list(selected_subjects)
            self._save_class_subjects_done_map(mapping)
            messagebox.showinfo('Saved', f'Subjects done for {cls} saved ({len(selected_subjects)}).', parent=dialog)

        tk.Button(actions, text='Auto From Marks', bg='#2563eb', fg='white',
                  activebackground='#1d4ed8', activeforeground='white',
                  font=(FF, 10, 'bold'), padx=12, pady=8, command=auto_pick_from_marks).pack(side='left', padx=(0, 8))
        tk.Button(actions, text='Use Automatic Mode', bg='#7c3aed', fg='white',
                  activebackground='#6d28d9', activeforeground='white',
                  font=(FF, 10, 'bold'), padx=12, pady=8, command=clear_class_override).pack(side='left', padx=(0, 8))
        tk.Button(actions, text='Save Subjects Done', bg='#0f766e', fg='white',
                  activebackground='#0e675f', activeforeground='white',
                  font=(FF, 10, 'bold'), padx=12, pady=8, command=save_selection).pack(side='left')

        tk.Button(dialog, text='Close', bg='#cbd5e1', fg='#0f172a',
                  activebackground='#94a3b8', activeforeground='#0f172a',
                  font=(FF, 10, 'bold'), padx=22, pady=8, command=dialog.destroy).pack(pady=(0, 16))

        reload_subject_list()

    def _open_subject_dialog(self, subject_row=None):
        dialog = tk.Toplevel(self.root)
        dialog.title('Edit Subject' if subject_row else 'Add Subject')
        dialog.geometry('430x500')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text='Subject Code:', font=(FF, 11)).pack(pady=(20, 5))
        code_entry = tk.Entry(dialog, font=(FF, 11))
        code_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Subject Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Short Label / Abbreviation:', font=(FF, 11)).pack(pady=(15, 5))
        abbr_entry = tk.Entry(dialog, font=(FF, 11))
        abbr_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Level:', font=(FF, 11)).pack(pady=(15, 5))
        level_var = tk.StringVar()
        level_options = [ALL_SUBJECT_LEVEL] + list(LEVELS)
        level_cb = ttk.Combobox(dialog, textvariable=level_var, values=level_options, state='readonly', font=(FF, 10))
        level_cb.pack(fill='x', padx=20)

        tk.Label(dialog, text='Category:', font=(FF, 11)).pack(pady=(15, 5))
        category_entry = tk.Entry(dialog, font=(FF, 11))
        category_entry.pack(fill='x', padx=20)

        optional_var = tk.BooleanVar(value=False)
        tk.Checkbutton(dialog, text='Optional Subject', variable=optional_var,
                      font=(FF, 11), bg=dialog.cget('bg')).pack(pady=(12, 16))

        if subject_row:
            code_entry.insert(0, subject_row.get('code', '') or subject_row.get('abbreviation', ''))
            name_entry.insert(0, subject_row.get('name', ''))
            abbr_entry.insert(0, subject_row.get('abbreviation', ''))
            level_var.set(subject_row.get('level', ''))
            category_entry.insert(0, subject_row.get('category', ''))
            optional_var.set(bool(subject_row.get('is_optional')))
        else:
            level_var.set(self.current_level if self.current_level in LEVELS else LEVELS[0])
            category_entry.insert(0, 'Core')

        def save():
            code = code_entry.get().strip().upper()
            name = name_entry.get().strip()
            abbreviation = code or abbr_entry.get().strip().upper()
            level = level_var.get().strip()
            category = category_entry.get().strip()

            if not code or not name or not level or not category:
                messagebox.showerror('Error', 'Subject code, name, level, and category are required')
                return

            abbr_entry.delete(0, tk.END)
            abbr_entry.insert(0, abbreviation)
            if subject_row:
                success, msg = db.update_subject(subject_row['id'], name, level, category, optional_var.get(), abbreviation, code)
            else:
                success, msg = db.add_subject(name, level, category, optional_var.get(), abbreviation, code)
            if success:
                dialog.destroy()
                self.show_settings()
            else:
                messagebox.showerror('Error', msg)

        tk.Button(dialog, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(pady=(4, 20))

    def _edit_subject_dialog(self, tree):
        selected = tree.selection()
        if len(selected) != 1:
            messagebox.showwarning('Select One', 'Please select exactly one subject to edit')
            return
        subject_id = selected[0]
        subject_row = next((row for row in db.get_subjects_by_level() if row.get('id') == subject_id), None)
        if not subject_row:
            messagebox.showerror('Error', 'Could not load the selected subject')
            return
        self._open_subject_dialog(subject_row)

    def _delete_subject(self, tree):
        selected = list(tree.selection())
        if not selected:
            messagebox.showwarning('Select', 'Please select one or more subjects to delete')
            return
        if not messagebox.askyesno('Confirm', f'Delete {len(selected)} selected subject(s)?'):
            return
        failures = 0
        for subject_id in selected:
            if not db.delete_subject(subject_id):
                failures += 1
        self._load_subjects(tree)
        if failures:
            messagebox.showerror('Error', f'Failed to delete {failures} subject(s).')

    def _reset_subject_catalog(self, tree):
        if not messagebox.askyesno('Confirm Reset', 'Delete all current subjects and replace them with the new default subject list?'):
            return
        self._replace_subject_catalog_with_defaults()
        self._load_subjects(tree)
        messagebox.showinfo('Subjects Updated', 'All subject records were replaced with the new coded subject catalog.')

    def _build_teachers_settings_tab(self, parent):
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)

        tk.Button(toolbar, text='+ Add Teacher', bg=GREEN, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._open_teacher_dialog()).pack(side='left', padx=5)
        tk.Button(toolbar, text='Edit Selected', bg=BLUE, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._edit_teacher_dialog(tree)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Delete Selected', bg='#e74c3c', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._delete_teacher(tree, reload_callback=refresh_all)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Assign Subject', bg='#1d4ed8', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._assign_subject_dialog(on_saved=refresh_all)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Assign Class Teacher', bg=PURPLE, fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: self._assign_class_teacher_dialog(on_saved=refresh_all)).pack(side='left', padx=5)
        tk.Button(toolbar, text='Refresh', bg='#666', fg='white',
                 font=(FF, 10), padx=12, pady=5, command=lambda: refresh_all()).pack(side='left', padx=5)

        search_row = tk.Frame(parent, bg=CONTENT_BG)
        search_row.pack(fill='x', padx=10, pady=(0, 4))
        tk.Label(search_row, text='Search:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=(0, 6))
        search_var = tk.StringVar()
        search_entry = ttk.Entry(search_row, textvariable=search_var, style='App.TEntry', width=30)
        search_entry.pack(side='left', ipady=3)
        tk.Button(
            search_row, text='🔍 Search', bg=OLIVE_PRIMARY, fg='white',
            activebackground=OLIVE_DARK, activeforeground='white',
            font=(FF, 9, 'bold'), padx=10, pady=4, relief='flat', cursor='hand2',
            command=lambda: refresh_all()
        ).pack(side='left', padx=(8, 0))

        list_frame = tk.Frame(parent, bg=CARD_BG, relief='flat', bd=1)
        list_frame.pack(fill='both', expand=True, padx=10, pady=(5, 8))

        cols = ('id', 'abbr', 'name', 'username', 'role')
        tree = ttk.Treeview(list_frame, columns=cols, show='headings', style='App.Treeview', selectmode='extended')
        tree.heading('id', text='ID')
        tree.heading('abbr', text='Short Label')
        tree.heading('name', text='Full Name')
        tree.heading('username', text='Username')
        tree.heading('role', text='Role')

        tree.column('id', width=70, anchor='center')
        tree.column('abbr', width=100, anchor='center')
        tree.column('name', width=220)
        tree.column('username', width=150)
        tree.column('role', width=140, anchor='center')

        select_all_var = tk.BooleanVar(value=False)
        select_row = tk.Frame(list_frame, bg=CARD_BG)
        select_row.pack(fill='x', padx=10, pady=(8, 0))
        tk.Checkbutton(
            select_row,
            text='Select all teachers',
            variable=select_all_var,
            bg=CARD_BG,
            fg=TEXT_SECONDARY,
            activebackground=CARD_BG,
            activeforeground=TEXT_PRIMARY,
            selectcolor=CARD_BG,
            font=(FF, 9, 'bold'),
            command=lambda: self._select_all_tree_rows(tree) if select_all_var.get() else self._clear_tree_selection(tree)
        ).pack(side='left')

        tree.pack(fill='both', expand=True, padx=10, pady=10)
        tree.bind('<Double-1>', lambda e: self._edit_teacher_dialog(tree))

        assignments_wrap = tk.Frame(parent, bg=CONTENT_BG)
        assignments_wrap.pack(fill='both', expand=True, padx=10, pady=(0, 8))
        assignments_wrap.columnconfigure(0, weight=1)
        assignments_wrap.columnconfigure(1, weight=1)
        assignments_wrap.rowconfigure(0, weight=1)

        subject_card = tk.Frame(assignments_wrap, bg=CARD_BG, relief='flat', bd=1)
        subject_card.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        tk.Label(subject_card, text='Subject Teacher Assignments', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', padx=10, pady=(8, 4))
        subj_cols = ('teacher', 'subject', 'class')
        subj_tree = ttk.Treeview(subject_card, columns=subj_cols, show='headings', style='App.Treeview', height=8)
        subj_tree.heading('teacher', text='Teacher')
        subj_tree.heading('subject', text='Subject')
        subj_tree.heading('class', text='Class')
        subj_tree.column('teacher', width=180, anchor='w')
        subj_tree.column('subject', width=170, anchor='w')
        subj_tree.column('class', width=140, anchor='w')
        subj_tree.pack(fill='both', expand=True, padx=10, pady=6)
        tk.Button(subject_card, text='Remove Selected', bg='#dc2626', fg='white',
                  font=(FF, 9, 'bold'), padx=10, pady=5,
                  command=lambda: self._remove_assignment_from_tree(subj_tree, refresh_all)).pack(anchor='e', padx=10, pady=(0, 8))

        class_card = tk.Frame(assignments_wrap, bg=CARD_BG, relief='flat', bd=1)
        class_card.grid(row=0, column=1, sticky='nsew', padx=(6, 0))
        tk.Label(class_card, text='Class Teacher Assignments', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', padx=10, pady=(8, 4))
        class_cols = ('teacher', 'class')
        class_tree = ttk.Treeview(class_card, columns=class_cols, show='headings', style='App.Treeview', height=8)
        class_tree.heading('teacher', text='Teacher')
        class_tree.heading('class', text='Class')
        class_tree.column('teacher', width=210, anchor='w')
        class_tree.column('class', width=180, anchor='w')
        class_tree.pack(fill='both', expand=True, padx=10, pady=6)
        tk.Button(class_card, text='Remove Selected', bg='#dc2626', fg='white',
                  font=(FF, 9, 'bold'), padx=10, pady=5,
                  command=lambda: self._remove_assignment_from_tree(class_tree, refresh_all)).pack(anchor='e', padx=10, pady=(0, 8))

        def refresh_all():
            self._load_teachers_tree(tree, search_var.get())
            self._load_teacher_assignments_trees(subj_tree, class_tree)

        search_entry.bind('<Return>', lambda _e: refresh_all())
        search_var.trace_add('write', lambda *_: refresh_all())
        refresh_all()

    def _load_teachers_tree(self, tree, search_query=''):
        for item in tree.get_children():
            tree.delete(item)

        query = str(search_query or '').strip().lower()
        for teacher in db.get_all_teachers():
            role_label = 'Subject Teacher' if teacher.get('role') == 'teacher' else 'Class Teacher'
            if query:
                haystack = ' '.join([
                    teacher.get('id', '')[:8],
                    teacher.get('abbreviation', '') or self._generate_short_label(teacher.get('full_name', ''), 'teacher'),
                    teacher.get('full_name', ''),
                    teacher.get('username', ''),
                    role_label,
                ]).lower()
                if query not in haystack:
                    continue
            tree.insert('', 'end', iid=teacher.get('id', ''), values=(
                teacher.get('id', '')[:8],
                teacher.get('abbreviation', '') or self._generate_short_label(teacher.get('full_name', ''), 'teacher'),
                teacher.get('full_name', ''),
                teacher.get('username', ''),
                role_label
            ))

    def _load_teacher_assignments_trees(self, subj_tree, class_tree):
        for item in subj_tree.get_children():
            subj_tree.delete(item)
        for item in class_tree.get_children():
            class_tree.delete(item)

        for assignment in db.get_subject_teacher_assignments():
            subj_tree.insert('', 'end', iid=assignment.get('id', str(uuid.uuid4())), values=(
                self._get_teacher_label(assignment),
                self._get_subject_label(assignment.get('subject', ''), assignment.get('class_name', '')),
                self._get_class_label(assignment.get('class_name', ''))
            ))

        for assignment in db.get_class_teacher_assignments():
            class_tree.insert('', 'end', iid=assignment.get('id', str(uuid.uuid4())), values=(
                self._get_teacher_label(assignment),
                self._get_class_label(assignment.get('class_name', ''))
            ))

    def _remove_assignment_from_tree(self, tree, reload_callback=None):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select Assignment', 'Please select an assignment to remove')
            return
        if not messagebox.askyesno('Confirm', 'Remove selected assignment?'):
            return
        assignment_id = selected[0]
        if db.remove_assignment(assignment_id):
            if callable(reload_callback):
                reload_callback()
        else:
            messagebox.showerror('Error', 'Failed to remove assignment')

    def _open_teacher_dialog(self, teacher_row=None):
        dialog = tk.Toplevel(self.root)
        dialog.title('Edit Teacher' if teacher_row else 'Add Teacher')
        dialog.geometry('430x500')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        tk.Label(dialog, text='Full Name:', font=(FF, 11)).pack(pady=(20, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Short Label / Abbreviation:', font=(FF, 11)).pack(pady=(15, 5))
        abbr_entry = tk.Entry(dialog, font=(FF, 11))
        abbr_entry.pack(fill='x', padx=20)

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
        tk.Radiobutton(role_frame, text='Subject Teacher', variable=role_var, value='teacher', font=(FF, 10)).pack(side='left', padx=10)
        tk.Radiobutton(role_frame, text='Class Teacher', variable=role_var, value='class_teacher', font=(FF, 10)).pack(side='left', padx=10)

        if teacher_row:
            name_entry.insert(0, teacher_row.get('full_name', ''))
            abbr_entry.insert(0, teacher_row.get('abbreviation', ''))
            username_entry.insert(0, teacher_row.get('username', ''))
            role_var.set(teacher_row.get('role', 'teacher'))

        def save():
            name = name_entry.get().strip()
            abbreviation = abbr_entry.get().strip()
            username = username_entry.get().strip()
            password = password_entry.get().strip()
            role = role_var.get()

            if not name or not username or (not teacher_row and not password):
                messagebox.showerror('Error', 'Name, username, and password are required for new teachers')
                return

            if teacher_row:
                success, msg = db.update_teacher(teacher_row['id'], name, username, role, abbreviation, password)
            else:
                success, msg = db.add_teacher(name, username, password, role, abbreviation)
            if success:
                dialog.destroy()
                self.show_settings()
            else:
                messagebox.showerror('Error', msg)

        tk.Button(dialog, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(pady=(16, 24))

    def _edit_teacher_dialog(self, tree):
        selected = tree.selection()
        if len(selected) != 1:
            messagebox.showwarning('Select One Teacher', 'Please select exactly one teacher to edit')
            return
        teacher_id = selected[0]
        teacher_row = next((row for row in db.get_all_teachers() if row.get('id') == teacher_id), None)
        if not teacher_row:
            messagebox.showerror('Error', 'Could not load the selected teacher')
            return
        self._open_teacher_dialog(teacher_row)

    def _build_grading_scale_tab(self, parent):
        toolbar = tk.Frame(parent, bg=CONTENT_BG)
        toolbar.pack(fill='x', pady=10)

        tk.Label(toolbar, text='Class:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='left', padx=(5, 4))
        class_options = [row.get('name') for row in db.get_all_classes()] or self.get_current_classes()
        class_var = tk.StringVar(value=class_options[0] if class_options else '')
        class_cb = ttk.Combobox(toolbar, textvariable=class_var, values=class_options, state='readonly',
                                style='App.TCombobox', width=22)
        class_cb.pack(side='left', padx=(0, 10))

        tree = ttk.Treeview(parent, columns=('code', 'name', 'min', 'max', 'order'), show='headings', style='App.Treeview')
        tree.heading('code', text='Grade')
        tree.heading('name', text='Description')
        tree.heading('min', text='Min')
        tree.heading('max', text='Max')
        tree.heading('order', text='Order')
        tree.column('code', width=80, anchor='center')
        tree.column('name', width=220)
        tree.column('min', width=80, anchor='center')
        tree.column('max', width=80, anchor='center')
        tree.column('order', width=80, anchor='center')
        tree.pack(fill='both', expand=True, padx=10, pady=10)

        actions = tk.Frame(parent, bg=CONTENT_BG)
        actions.pack(fill='x', pady=(0, 10))
        tk.Button(actions, text='+ Add Band', bg=GREEN, fg='white',
                 font=(FF, 10), padx=12, pady=5,
                 command=lambda: self._open_grading_scale_dialog(class_var.get(), refresh=lambda: self._load_grading_scale_tree(tree, class_var.get()))).pack(side='left', padx=5)
        tk.Button(actions, text='Edit Selected', bg=BLUE, fg='white',
                 font=(FF, 10), padx=12, pady=5,
                 command=lambda: self._edit_grading_scale_dialog(tree, class_var.get())).pack(side='left', padx=5)
        tk.Button(actions, text='Delete Selected', bg='#e74c3c', fg='white',
                 font=(FF, 10), padx=12, pady=5,
                 command=lambda: self._delete_grading_scale(tree, class_var.get())).pack(side='left', padx=5)
        tk.Button(actions, text='Load 8-Band Template', bg='#7c3aed', fg='white',
                 font=(FF, 10), padx=12, pady=5,
                 command=lambda: self._apply_grade_band_template(class_var.get(), refresh=lambda: self._load_grading_scale_tree(tree, class_var.get()))).pack(side='left', padx=5)
        tk.Button(actions, text='Refresh', bg='#666', fg='white',
                 font=(FF, 10), padx=12, pady=5,
                 command=lambda: self._load_grading_scale_tree(tree, class_var.get())).pack(side='left', padx=5)

        class_cb.bind('<<ComboboxSelected>>', lambda e: self._load_grading_scale_tree(tree, class_var.get()))
        tree.bind('<Double-1>', lambda e: self._edit_grading_scale_dialog(tree, class_var.get()))
        self._load_grading_scale_tree(tree, class_var.get())

    def _load_grading_scale_tree(self, tree, class_name):
        for item in tree.get_children():
            tree.delete(item)
        for scale in db.get_grading_scales(class_name):
            tree.insert('', 'end', iid=scale.get('id', ''), values=(
                scale.get('grade_code', ''),
                scale.get('grade_name', '') or GRADE_LABELS.get(scale.get('grade_code', ''), scale.get('grade_code', '')),
                scale.get('min_mark', ''),
                scale.get('max_mark', ''),
                scale.get('sort_order', 0),
            ))

    def _open_grading_scale_dialog(self, class_name, scale_row=None, refresh=None):
        dialog = tk.Toplevel(self.root)
        dialog.title('Edit Grade Band' if scale_row else 'Add Grade Band')
        dialog.geometry('430x500')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)

        class_options = [row.get('name') for row in db.get_all_classes()] or self.get_current_classes()

        tk.Label(dialog, text='Class:', font=(FF, 11)).pack(pady=(20, 5))
        class_var = tk.StringVar(value=class_name)
        class_cb = ttk.Combobox(dialog, textvariable=class_var, values=class_options, state='readonly', style='App.TCombobox')
        class_cb.pack(fill='x', padx=20)

        tk.Label(dialog, text='Grade Code:', font=(FF, 11)).pack(pady=(15, 5))
        code_var = tk.StringVar()
        grade_code_options = [
            'EE', 'EE1', 'EE2',
            'ME', 'ME1', 'ME2',
            'AE', 'AE1', 'AE2',
            'BE', 'BE1', 'BE2',
            'IE'
        ]
        code_cb = ttk.Combobox(dialog, textvariable=code_var, values=grade_code_options, style='App.TCombobox')
        code_cb.pack(fill='x', padx=20)
        tk.Label(dialog, text='Tip: pick from list or type your own code (e.g. EE1, ME2, BE1).',
                 font=(FF, 9), fg=TEXT_SECONDARY, bg=dialog.cget('bg')).pack(anchor='w', padx=22, pady=(4, 0))

        tk.Label(dialog, text='Grade Name:', font=(FF, 11)).pack(pady=(15, 5))
        name_entry = tk.Entry(dialog, font=(FF, 11))
        name_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Minimum Mark:', font=(FF, 11)).pack(pady=(15, 5))
        min_entry = tk.Entry(dialog, font=(FF, 11))
        min_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Maximum Mark:', font=(FF, 11)).pack(pady=(15, 5))
        max_entry = tk.Entry(dialog, font=(FF, 11))
        max_entry.pack(fill='x', padx=20)

        tk.Label(dialog, text='Display Order:', font=(FF, 11)).pack(pady=(15, 5))
        order_entry = tk.Entry(dialog, font=(FF, 11))
        order_entry.pack(fill='x', padx=20)

        if scale_row:
            class_var.set(scale_row.get('class_name', class_name))
            code_var.set(scale_row.get('grade_code', ''))
            name_entry.insert(0, scale_row.get('grade_name', ''))
            min_entry.insert(0, str(scale_row.get('min_mark', '')))
            max_entry.insert(0, str(scale_row.get('max_mark', '')))
            order_entry.insert(0, str(scale_row.get('sort_order', 0)))
        else:
            code_var.set('EE')
            name_entry.insert(0, GRADE_LABELS.get('EE', 'Exceeding Expectations'))
            order_entry.insert(0, '1')

        def sync_grade_name(event=None):
            current = name_entry.get().strip()
            code = code_var.get().strip()
            if not current or current in GRADE_LABELS.values() or current == GRADE_LABELS.get(grade_base_code(code), ''):
                name_entry.delete(0, tk.END)
                name_entry.insert(0, GRADE_LABELS.get(grade_base_code(code), code or ''))

        code_cb.bind('<<ComboboxSelected>>', sync_grade_name)

        def save():
            try:
                min_mark = float(min_entry.get().strip())
                max_mark = float(max_entry.get().strip())
                sort_order = int(order_entry.get().strip() or '0')
            except ValueError:
                messagebox.showerror('Error', 'Min, max, and order must be numeric')
                return

            if not class_var.get().strip() or not code_var.get().strip():
                messagebox.showerror('Error', 'Class and grade code are required')
                return

            grade_code = code_var.get().strip()
            grade_name = name_entry.get().strip() or GRADE_LABELS.get(grade_base_code(grade_code), grade_code)
            if scale_row:
                success, msg = db.update_grading_scale(
                    scale_row['id'], class_var.get().strip(), min_mark, max_mark,
                    grade_code, grade_name, sort_order
                )
            else:
                success, msg = db.add_grading_scale(
                    class_var.get().strip(), min_mark, max_mark,
                    grade_code, grade_name, sort_order
                )

            if success:
                dialog.destroy()
                if refresh:
                    refresh()
            else:
                messagebox.showerror('Error', msg)

        btn_row = tk.Frame(dialog, bg=dialog.cget('bg'))
        btn_row.pack(fill='x', pady=(18, 20), padx=20)
        tk.Button(btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=dialog.destroy).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Save', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(side='left')

    def _apply_grade_band_template(self, class_name, refresh=None):
        class_name = str(class_name or '').strip()
        if not class_name:
            messagebox.showwarning('Missing Class', 'Please select a class first.')
            return

        if not messagebox.askyesno(
            'Apply Template',
            f'Apply 8-band template to {class_name}?\n\n'
            'This will replace existing grading bands for this class.'
        ):
            return

        default_bands = [
            (90, 100, 'EE1', 'Exceeding Expectation 1', 1),
            (76, 89, 'EE2', 'Exceeding Expectation 2', 2),
            (60, 75, 'ME1', 'Meeting Expectation 1', 3),
            (40, 59, 'ME2', 'Meeting Expectation 2', 4),
            (30, 39, 'AE1', 'Approaching Expectation 1', 5),
            (20, 29, 'AE2', 'Approaching Expectation 2', 6),
            (10, 19, 'BE1', 'Below Expectation 1', 7),
            (0, 9, 'BE2', 'Below Expectation 2', 8),
        ]

        for scale in db.get_grading_scales(class_name):
            db.delete_grading_scale(scale.get('id', ''))

        for min_mark, max_mark, code, name, order in default_bands:
            ok, msg = db.add_grading_scale(class_name, min_mark, max_mark, code, name, order)
            if not ok:
                messagebox.showerror('Template Failed', f'Could not add {code} band.\n\n{msg}')
                if refresh:
                    refresh()
                return

        if refresh:
            refresh()
        messagebox.showinfo('Template Applied', f'8-band grading template applied to {class_name}.')

    def _edit_grading_scale_dialog(self, tree, class_name):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select', 'Please select a grading band to edit')
            return
        scale_row = db.get_grading_scale(selected[0])
        if not scale_row:
            messagebox.showerror('Error', 'Could not load the selected grading band')
            return
        self._open_grading_scale_dialog(class_name, scale_row, refresh=lambda: self._load_grading_scale_tree(tree, class_name))

    def _delete_grading_scale(self, tree, class_name):
        selected = tree.selection()
        if not selected:
            messagebox.showwarning('Select', 'Please select a grading band to delete')
            return
        if not messagebox.askyesno('Confirm', 'Delete this grading band?'):
            return
        if db.delete_grading_scale(selected[0]):
            self._load_grading_scale_tree(tree, class_name)
        else:
            messagebox.showerror('Error', 'Failed to delete grading band')
    
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
                self._get_teacher_label(a),
                self._get_subject_label(a.get('subject', ''), a.get('class_name', '')),
                self._get_class_label(a.get('class_name', ''))
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
                self._get_teacher_label(a),
                self._get_class_label(a.get('class_name', ''))
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

        stats = db.get_statistics('One', DEFAULT_EXAM_TYPE)
        dashboard_classes = self.get_current_classes()
        all_students = db.get_all_students()
        recent_students = sorted(all_students, key=lambda row: row.get('created_at', ''), reverse=True)[:8]
        level = self.current_level

        # Top title and breadcrumb row
        top_row = tk.Frame(self.content_frame, bg=CONTENT_BG)
        top_row.pack(fill='x', pady=(2, 10))
        tk.Label(top_row, text='Dashboard', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 22, 'bold')).pack(side='left')
        crumb = tk.Frame(top_row, bg='#ecefe6', padx=10, pady=6)
        crumb.pack(side='right')
        tk.Label(crumb, text='Home', bg='#ecefe6', fg=TEXT_SECONDARY, font=(FF, 9)).pack(side='left')
        tk.Label(crumb, text=' / ', bg='#ecefe6', fg=TEXT_SECONDARY, font=(FF, 9)).pack(side='left')
        tk.Label(crumb, text='Dashboard', bg='#ecefe6', fg=TEXT_PRIMARY, font=(FF, 9, 'bold')).pack(side='left')

        # KPI cards row (preserve olive/lemon tone with RGB variation)
        cards_row = tk.Frame(self.content_frame, bg=CONTENT_BG)
        cards_row.pack(fill='x', pady=(0, 12))
        for i in range(4):
            cards_row.columnconfigure(i, weight=1)

        # Real KPI calculations (recent growth windows from student created_at)
        now = datetime.now()
        recent_28 = 0
        prev_28 = 0
        for student in all_students:
            created_at = str(student.get('created_at', '') or '').strip()
            if not created_at:
                continue
            dt = None
            try:
                dt = datetime.fromisoformat(created_at.replace('Z', '+00:00')).replace(tzinfo=None)
            except Exception:
                for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
                    try:
                        dt = datetime.strptime(created_at[:19], fmt)
                        break
                    except Exception:
                        continue
            if not dt:
                continue
            days = (now - dt).days
            if 0 <= days <= 27:
                recent_28 += 1
            elif 28 <= days <= 55:
                prev_28 += 1
        growth_pct = int(((recent_28 - prev_28) / prev_28) * 100) if prev_28 > 0 else (100 if recent_28 > 0 else 0)

        metrics = [
            ('Total Students', str(stats.get('students', 0)), f'{growth_pct:+d}% vs previous 28 days', '#58c7a6', '#2ea886', '👥'),
            ('New Students', str(recent_28), 'Joined in the last 28 days', '#e3c060', '#d39f2d', '🧑'),
            ('Current Classes', str(len(dashboard_classes)), 'Active classes', '#67a8e8', '#3f7ec1', '🏫'),
            ('Avg Score', str(stats.get('avg_score', 0)), f'Term One / {DEFAULT_EXAM_TYPE}', '#db6f8b', '#c6456d', '🎯'),
        ]

        for idx, (title, value, note, soft, strong, icon) in enumerate(metrics):
            outer = tk.Frame(cards_row, bg=strong)
            outer.grid(row=0, column=idx, padx=6, sticky='nsew')
            card = tk.Frame(outer, bg=soft, padx=14, pady=12)
            card.pack(fill='both', expand=True, padx=1, pady=1)

            head = tk.Frame(card, bg=soft)
            head.pack(fill='x')
            tk.Label(head, text=icon, bg=strong, fg='white', font=(FF, 10, 'bold'),
                     padx=8, pady=4).pack(side='left')
            tk.Label(head, text=title, bg=soft, fg='white', font=(FF, 10, 'bold')).pack(side='left', padx=8)

            tk.Label(card, text=value, bg=soft, fg='white', font=(FF, 18, 'bold')).pack(anchor='w', pady=(10, 2))
            tk.Label(card, text=note, bg=soft, fg='#f8fdf9', font=(FF, 8)).pack(anchor='w')

            progress = tk.Canvas(card, bg=soft, height=8, highlightthickness=0)
            progress.pack(fill='x', pady=(8, 0))
            progress.create_line(2, 4, 240, 4, fill=_mix_hex(strong, '#0f172a', 0.42), width=2)
            progress.create_line(2, 4, 130 + (idx * 20), 4, fill='#f5f5f5', width=2)

        # Analytics panels row (like reference: two large charts)
        charts_row = tk.Frame(self.content_frame, bg=CONTENT_BG)
        charts_row.pack(fill='both', expand=True, pady=(2, 4))
        charts_row.columnconfigure(0, weight=1)
        charts_row.columnconfigure(1, weight=1)
        charts_row.rowconfigure(0, weight=1)

        def make_chart_panel(parent, title, theme='sky'):
            bo, bi = _card_colors(theme)
            outer = tk.Frame(parent, bg=bo)
            body = tk.Frame(outer, bg=bi, padx=12, pady=10)
            body.pack(fill='both', expand=True, padx=1, pady=1)
            hdr = tk.Frame(body, bg=bi)
            hdr.pack(fill='x', pady=(0, 6))
            tk.Label(hdr, text=title, bg=bi, fg=TEXT_PRIMARY, font=(FF, 12, 'bold')).pack(side='left')
            tk.Label(hdr, text='• • •', bg=bi, fg=TEXT_SECONDARY, font=(FF, 10)).pack(side='right')
            plot_holder = tk.Frame(body, bg=bi)
            plot_holder.pack(fill='both', expand=True)
            return outer, plot_holder, bi

        panel_left, left_holder, left_bg = make_chart_panel(charts_row, 'Performance Trend', 'azure')
        panel_left.grid(row=0, column=0, sticky='nsew', padx=(0, 6))
        panel_right, right_holder, right_bg = make_chart_panel(charts_row, 'Enrollment Movement', 'mint')
        panel_right.grid(row=0, column=1, sticky='nsew', padx=(6, 0))

        # Chart 1: real exam-session performance trend
        exam_points = {}
        results_cache = {}
        for class_name in dashboard_classes:
            history_rows = db.get_class_exam_history(class_name) or []
            for row in history_rows:
                term = str(row.get('term', '') or '')
                exam_type = str(row.get('exam_type', '') or DEFAULT_EXAM_TYPE)
                created_at = str(row.get('created_at', '') or '').strip()
                point_key = (term, exam_type)
                key_dt = datetime.min
                if created_at:
                    try:
                        key_dt = datetime.fromisoformat(created_at.replace('Z', '+00:00')).replace(tzinfo=None)
                    except Exception:
                        pass
                cache_key = (class_name, term, exam_type)
                if cache_key not in results_cache:
                    class_results = self._get_ranked_results(class_name, term, exam_type)
                    class_avg = round(sum(r.get('average', 0) for r in class_results) / len(class_results), 1) if class_results else None
                    results_cache[cache_key] = class_avg
                class_avg = results_cache[cache_key]
                if class_avg is None:
                    continue
                if point_key not in exam_points:
                    exam_points[point_key] = {'dt': key_dt, 'values': []}
                exam_points[point_key]['dt'] = max(exam_points[point_key]['dt'], key_dt)
                exam_points[point_key]['values'].append(class_avg)

        ordered_points = sorted(
            [
                (key, info['dt'], round(sum(info['values']) / len(info['values']), 1))
                for key, info in exam_points.items() if info['values']
            ],
            key=lambda x: x[1]
        )
        if len(ordered_points) > 8:
            ordered_points = ordered_points[-8:]

        if ordered_points:
            labels = [f"T{item[0][0]}-{str(item[0][1]).split('-')[0][:3].upper()}" for item in ordered_points]
            y1 = [item[2] for item in ordered_points]
        else:
            labels = ['No Data']
            y1 = [float(stats.get('avg_score', 0) or 0)]
        y2 = [round((y1[i - 1] if i > 0 else y1[i]), 1) for i in range(len(y1))]
        if ordered_points and ordered_points[0][1] != datetime.min and ordered_points[-1][1] != datetime.min:
            trend_window = f"{ordered_points[0][1].strftime('%b %Y')} - {ordered_points[-1][1].strftime('%b %Y')}"
        else:
            trend_window = 'No dated exam sessions'
        trend_scope = f"Scope: {len(dashboard_classes)} classes | Window: {trend_window}"
        fig1, ax1 = plt.subplots(figsize=(5.6, 3.0))
        fig1.patch.set_facecolor(left_bg)
        ax1.set_facecolor(_mix_hex(left_bg, '#ffffff', 0.45))
        ax1.plot(labels, y1, color='#5561d8', linewidth=2.2, marker='o', markersize=4, label='Current')
        ax1.plot(labels, y2, color='#8d96a3', linewidth=2.2, marker='o', markersize=4, label='Previous')
        ax1.grid(axis='y', color='#d9dee8', linewidth=0.7)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.tick_params(axis='x', labelsize=8, colors=TEXT_SECONDARY)
        ax1.tick_params(axis='y', labelsize=8, colors=TEXT_SECONDARY)
        ax1.set_ylim(0, 100)
        ax1.set_title('Average Score Trend by Exam Session', fontsize=10, color=TEXT_PRIMARY, pad=10)
        ax1.text(0.01, 1.005, trend_scope, transform=ax1.transAxes, fontsize=8, color=TEXT_SECONDARY)
        ax1.legend(loc='upper left', bbox_to_anchor=(0.0, 0.98), fontsize=8, frameon=False)
        fig1.subplots_adjust(top=0.83, left=0.09, right=0.98, bottom=0.14)
        self.dashboard_chart_1 = FigureCanvasTkAgg(fig1, master=left_holder)
        self.dashboard_chart_1.draw()
        self.dashboard_chart_1.get_tk_widget().pack(fill='both', expand=True)
        plt.close(fig1)

        # Chart 2: real enrollment movement by month
        month_map = {}
        for student in all_students:
            created_at = str(student.get('created_at', '') or '').strip()
            if not created_at:
                continue
            dt = None
            try:
                dt = datetime.fromisoformat(created_at.replace('Z', '+00:00')).replace(tzinfo=None)
            except Exception:
                for fmt in ('%Y-%m-%d %H:%M:%S', '%Y-%m-%d'):
                    try:
                        dt = datetime.strptime(created_at[:19], fmt)
                        break
                    except Exception:
                        continue
            if not dt:
                continue
            key = dt.strftime('%Y-%m')
            month_map[key] = month_map.get(key, 0) + 1

        ordered_months = sorted(month_map.items(), key=lambda x: x[0])
        if len(ordered_months) > 8:
            ordered_months = ordered_months[-8:]

        if ordered_months:
            labels2 = [datetime.strptime(k, '%Y-%m').strftime('%b') for k, _ in ordered_months]
            y3 = [v for _, v in ordered_months]  # new students
            cumulative = 0
            y4 = []
            for value in y3:
                cumulative += value
                y4.append(cumulative)
            enroll_window = f"{datetime.strptime(ordered_months[0][0], '%Y-%m').strftime('%b %Y')} - {datetime.strptime(ordered_months[-1][0], '%Y-%m').strftime('%b %Y')}"
        else:
            # fallback: class distribution
            labels2 = [self._get_class_label(c) for c in dashboard_classes[:7]] or ['No Data']
            y3 = [len(db.get_students_by_class(c)) for c in dashboard_classes[:7]] or [0]
            cumulative = 0
            y4 = []
            for value in y3:
                cumulative += value
                y4.append(cumulative)
            enroll_window = 'Fallback class snapshot'
        enroll_scope = f"Scope: {len(all_students)} students | Window: {enroll_window}"
        fig2, ax2 = plt.subplots(figsize=(5.6, 3.0))
        fig2.patch.set_facecolor(right_bg)
        ax2.set_facecolor(_mix_hex(right_bg, '#ffffff', 0.45))
        ax2.plot(labels2, y3, color='#58c7a6', linewidth=2.0, label='New Students')
        ax2.fill_between(labels2, y3, color='#58c7a6', alpha=0.25)
        ax2.plot(labels2, y4, color='#7c8593', linewidth=2.0, label='Cumulative')
        ax2.fill_between(labels2, y4, color='#7c8593', alpha=0.18)
        ax2.grid(axis='y', color='#d9dee8', linewidth=0.7)
        ax2.spines['top'].set_visible(False)
        ax2.spines['right'].set_visible(False)
        ax2.tick_params(axis='x', labelsize=8, colors=TEXT_SECONDARY)
        ax2.tick_params(axis='y', labelsize=8, colors=TEXT_SECONDARY)
        upper = max(y3 + y4 + [5]) + 5
        ax2.set_ylim(0, upper)
        ax2.set_title('Enrollment Movement', fontsize=10, color=TEXT_PRIMARY, pad=10)
        ax2.text(0.01, 1.005, enroll_scope, transform=ax2.transAxes, fontsize=8, color=TEXT_SECONDARY)
        ax2.legend(loc='upper left', bbox_to_anchor=(0.0, 0.98), fontsize=8, frameon=False)
        fig2.subplots_adjust(top=0.83, left=0.09, right=0.98, bottom=0.14)
        self.dashboard_chart_2 = FigureCanvasTkAgg(fig2, master=right_holder)
        self.dashboard_chart_2.draw()
        self.dashboard_chart_2.get_tk_widget().pack(fill='both', expand=True)
        plt.close(fig2)

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
                 font=(FF, 10), padx=15, pady=5, command=lambda: self._open_teacher_dialog()).pack(side='left', padx=5)
        tk.Button(toolbar, text='Edit Selected', bg=BLUE, fg='white',
                 font=(FF, 10), padx=15, pady=5, command=lambda: self._edit_teacher_dialog(tree)).pack(side='left', padx=5)
        
        tk.Button(toolbar, text='Refresh', bg='#666', fg='white',
                 font=(FF, 10), padx=15, pady=5, command=self.show_teachers).pack(side='left', padx=5)
        
        # Teachers list
        tf_bo, tf_bi = _card_colors('azure')
        tf_outer = tk.Frame(self.content_frame, bg=tf_bo)
        tf_outer.pack(fill='both', expand=True, pady=4)
        teachers_frame = tk.Frame(tf_outer, bg=tf_bi)
        teachers_frame.pack(fill='both', expand=True, padx=1, pady=1)
        
        self.teachers_table = AdvancedDataTable(
            teachers_frame,
            columns=[
                {'key': 'id', 'title': 'ID', 'width': 90, 'anchor': 'center'},
                {'key': 'abbr', 'title': 'Short Label', 'width': 110, 'anchor': 'center'},
                {'key': 'name', 'title': 'Full Name', 'width': 200, 'anchor': 'w'},
                {'key': 'username', 'title': 'Username', 'width': 130, 'anchor': 'w'},
                {'key': 'role', 'title': 'Role', 'width': 130, 'anchor': 'center'},
                {'key': 'assignments', 'title': 'Assignments', 'width': 420, 'anchor': 'w'},
            ],
            page_size=12,
            search_label='Search teachers',
            selectmode='extended',
            enable_select_all=True
        )
        tree = self.teachers_table.tree
        
        # Load teachers
        teachers = db.get_all_teachers()
        rows = []
        for teacher in teachers:
            # Get assignments
            assignments = []
            subject_assignments = db.get_subject_teacher_assignments()
            class_assignments = db.get_class_teacher_assignments()
            
            for sa in subject_assignments:
                if sa['teacher_id'] == teacher['id']:
                    assignments.append(
                        f"{self._get_subject_label(sa['subject'], sa['class_name'])} ({self._get_class_label(sa['class_name'])})"
                    )
            
            for ca in class_assignments:
                if ca['teacher_id'] == teacher['id']:
                    assignments.append(f"Class Teacher: {self._get_class_label(ca['class_name'])}")
            
            role_label = 'Subject Teacher' if teacher.get('role') == 'teacher' else 'Class Teacher'
            assignment_text = ', '.join(assignments) if assignments else 'No assignments'
            
            row_values = (
                teacher.get('id', '')[:8],
                teacher.get('abbreviation', '') or self._generate_short_label(teacher.get('full_name', ''), 'teacher'),
                teacher.get('full_name', ''),
                teacher.get('username', ''),
                role_label,
                assignment_text
            )
            rows.append({
                'iid': teacher.get('id', str(uuid.uuid4())),
                'values': row_values,
                'value_map': {
                    'id': row_values[0],
                    'abbr': row_values[1],
                    'name': row_values[2],
                    'username': row_values[3],
                    'role': row_values[4],
                    'assignments': row_values[5],
                },
                'search': ' '.join(str(v) for v in row_values),
            })
        self.teachers_table.set_rows(rows)
        tree.bind('<Double-1>', lambda e: self._edit_teacher_dialog(tree))
        
        # Action buttons
        action_frame = tk.Frame(teachers_frame, bg=tf_bi)
        action_frame.pack(fill='x', padx=10, pady=(0, 10))
        
        tk.Button(action_frame, text='Assign Subject', bg=BLUE, fg='white',
                 font=(FF, 10), padx=10, pady=5, command=self._assign_subject_dialog).pack(side='left', padx=5)
        
        tk.Button(action_frame, text='Assign Class Teacher', bg=PURPLE, fg='white',
                 font=(FF, 10), padx=10, pady=5, command=self._assign_class_teacher_dialog).pack(side='left', padx=5)
        
        tk.Button(action_frame, text='Delete', bg='#e74c3c', fg='white',
                 font=(FF, 10), padx=10, pady=5, command=lambda: self._delete_teacher(tree, reload_callback=self.show_teachers)).pack(side='left', padx=5)
    
    def _add_teacher_dialog(self):
        self._open_teacher_dialog()
    
    def _assign_subject_dialog(self, on_saved=None):
        """Dialog to assign a subject to a teacher"""
        teachers = db.get_all_teachers()
        if not teachers:
            messagebox.showwarning('No Teachers', 'Please add teachers first')
            return
        classes = [row.get('name') for row in db.get_all_classes()]
        if not classes:
            classes = self.get_current_classes()
        
        dialog = tk.Toplevel(self.root)
        dialog.title('Assign Subject to Teacher')
        dialog.geometry('430x380')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        tk.Label(dialog, text='Select Teacher:', font=(FF, 11)).pack(pady=(20, 5))
        teacher_names = [f"{self._get_teacher_label(t)} - {t['full_name']} ({t['username']})" for t in teachers]
        teacher_var = tk.StringVar()
        teacher_cb = ttk.Combobox(dialog, textvariable=teacher_var, values=teacher_names, 
                                  state='readonly', font=(FF, 10))
        teacher_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Select Class:', font=(FF, 11)).pack(pady=(15, 5))
        class_var = tk.StringVar()
        class_cb = ttk.Combobox(dialog, textvariable=class_var, values=classes,
                               state='readonly', font=(FF, 10))
        class_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Select Subject:', font=(FF, 11)).pack(pady=(15, 5))
        subject_var = tk.StringVar()
        subject_cb = ttk.Combobox(dialog, textvariable=subject_var, values=self.get_current_subjects(),
                                  state='readonly', font=(FF, 10))
        subject_cb.pack(fill='x', padx=20)

        def refresh_subject_options(event=None):
            selected_class = class_var.get().strip()
            values = self._get_subjects_for_selected_class(selected_class, TERMS[0]) if selected_class else self.get_current_subjects()
            subject_cb.configure(values=values)
            if subject_var.get() not in values:
                subject_var.set(values[0] if values else '')

        class_cb.bind('<<ComboboxSelected>>', refresh_subject_options)
        
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
                if callable(on_saved):
                    on_saved()
                else:
                    self.show_teachers()
            else:
                messagebox.showerror('Error', 'Failed to assign subject')
        
        btn_row = tk.Frame(dialog, bg=dialog.cget('bg'))
        btn_row.pack(fill='x', pady=(18, 20), padx=20)
        tk.Button(btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=dialog.destroy).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Assign Subject', bg=BLUE, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save_assignment).pack(side='left')
    
    def _assign_class_teacher_dialog(self, on_saved=None):
        """Dialog to assign a class teacher"""
        teachers = db.get_all_teachers()
        if not teachers:
            messagebox.showwarning('No Teachers', 'Please add teachers first')
            return
        classes = [row.get('name') for row in db.get_all_classes()]
        if not classes:
            classes = self.get_current_classes()
        
        dialog = tk.Toplevel(self.root)
        dialog.title('Assign Class Teacher')
        dialog.geometry('430x320')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        tk.Label(dialog, text='Select Teacher:', font=(FF, 11)).pack(pady=(20, 5))
        teacher_names = [f"{self._get_teacher_label(t)} - {t['full_name']} ({t['username']})" for t in teachers]
        teacher_var = tk.StringVar()
        teacher_cb = ttk.Combobox(dialog, textvariable=teacher_var, values=teacher_names,
                                  state='readonly', font=(FF, 10))
        teacher_cb.pack(fill='x', padx=20)
        
        tk.Label(dialog, text='Select Class:', font=(FF, 11)).pack(pady=(15, 5))
        class_var = tk.StringVar()
        class_cb = ttk.Combobox(dialog, textvariable=class_var, values=classes,
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
                if callable(on_saved):
                    on_saved()
                else:
                    self.show_teachers()
            else:
                messagebox.showerror('Error', 'Failed to assign class teacher')
        
        btn_row = tk.Frame(dialog, bg=dialog.cget('bg'))
        btn_row.pack(fill='x', pady=(18, 20), padx=20)
        tk.Button(btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=dialog.destroy).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Assign Class Teacher', bg=PURPLE, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save_assignment).pack(side='left')
    
    def _delete_teacher(self, tree, reload_callback=None):
        selected = list(tree.selection())
        if not selected:
            messagebox.showwarning('Select Teacher', 'Please select one or more teachers to delete')
            return
        
        if not messagebox.askyesno('Confirm Delete', f'Are you sure you want to delete {len(selected)} selected teacher(s)?'):
            return

        failed = 0
        for teacher_id in selected:
            if not db.delete_user(teacher_id):
                failed += 1

        if reload_callback:
            reload_callback()
        else:
            self.show_teachers()

        if failed:
            messagebox.showerror('Error', f'Failed to delete {failed} teacher(s).')
        else:
            messagebox.showinfo('Success', f'Deleted {len(selected)} teacher(s) successfully.')

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

        self.class_students_table = AdvancedDataTable(
            list_frame,
            columns=[
                {'key': 'adm', 'title': 'Admission No', 'width': 140, 'anchor': 'center'},
                {'key': 'name', 'title': 'Name', 'width': 240, 'anchor': 'w'},
                {'key': 'gender', 'title': 'Gender', 'width': 100, 'anchor': 'center'},
            ],
            page_size=15,
            search_label='Search learners'
        )
        self.class_students_tree = self.class_students_table.tree
        self._load_class_students(classes[0])
    
    def _load_class_students(self, class_name):
        """Load students for a specific class"""
        rows = []
        students = db.get_students_by_class(class_name)
        for s in students:
            values = (
                s.get('admission_no', ''),
                s.get('name', ''),
                s.get('gender', '')
            )
            rows.append({
                'iid': s.get('id', str(uuid.uuid4())),
                'values': values,
                'value_map': {'adm': values[0], 'name': values[1], 'gender': values[2]},
                'search': ' '.join(str(v) for v in values),
            })
        if hasattr(self, 'class_students_table'):
            self.class_students_table.set_rows(rows)
    
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

        self.comments_table = AdvancedDataTable(
            list_frame,
            columns=[
                {'key': 'adm', 'title': 'Admission No', 'width': 140, 'anchor': 'center'},
                {'key': 'name', 'title': 'Name', 'width': 220, 'anchor': 'w'},
                {'key': 'comment', 'title': 'Current Comment', 'width': 450, 'anchor': 'w'},
            ],
            page_size=15,
            search_label='Search comments'
        )
        self.comments_tree = self.comments_table.tree
        self.comments_class_var = class_var
        self.comments_term_var = term_var
    
    def _load_students_for_comments(self, class_name, term):
        """Load students for adding comments"""
        rows = []
        students = db.get_students_by_class(class_name)
        for s in students:
            # Get existing comment
            comment_data = db.get_student_comment(s['id'], term)
            comment = comment_data.get('comment_text', '') if comment_data else ''

            values = (
                s.get('admission_no', ''),
                s.get('name', ''),
                comment
            )
            rows.append({
                'iid': s.get('id', str(uuid.uuid4())),
                'values': values,
                'tags': (s.get('id', ''),),
                'value_map': {'adm': values[0], 'name': values[1], 'comment': values[2]},
                'search': ' '.join(str(v) for v in values),
            })
        if hasattr(self, 'comments_table'):
            self.comments_table.set_rows(rows)
        
        # Add double-click to edit
        self.comments_tree.bind('<Double-1>', lambda e: self._edit_comment(self.comments_tree, class_name, term))
    
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
        dialog.geometry('540x360')
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
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
        
        btn_row = tk.Frame(dialog, bg=dialog.cget('bg'))
        btn_row.pack(fill='x', pady=(0, 14), padx=20)
        tk.Button(btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=dialog.destroy).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Save Comment', bg=GREEN, fg='white',
                 font=(FF, 11), padx=20, pady=8, command=save).pack(side='left')

    # ==================== PROMOTIONS ====================
    def show_promotions(self):
        """Show student promotions management interface."""
        self.clear_frame()
        self._set_nav('Promotions')
        self._page_header('Student Promotions', 'Manage student promotions to next class/grade')
        
        # Main container with two columns
        main_container = tk.Frame(self.content_frame, bg=CONTENT_BG)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        main_container.columnconfigure(0, weight=1)
        main_container.columnconfigure(1, weight=1)
        
        # Left column - Promotion Settings
        left_frame = tk.Frame(main_container, bg=CARD_BG, relief='flat', bd=1)
        left_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 5))
        
        tk.Label(left_frame, text='Promotion Settings', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 14, 'bold')).pack(anchor='w', padx=15, pady=(15, 10))
        
        # Settings form
        settings_frame = tk.Frame(left_frame, bg=CARD_BG)
        settings_frame.pack(fill='x', padx=15, pady=5)
        
        # Get current settings
        settings = promotion_manager.get_settings()
        
        # Promotion date
        tk.Label(settings_frame, text='Promotion Date (MM-DD):', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(anchor='w', pady=(10, 2))
        self.promo_date_var = tk.StringVar(value=settings.get('promotion_date', '12-01'))
        promo_date_entry = tk.Entry(settings_frame, textvariable=self.promo_date_var,
                                    font=(FF, 10), width=20)
        promo_date_entry.pack(anchor='w', pady=(0, 10))
        
        # Minimum passing average
        tk.Label(settings_frame, text='Minimum Passing Average (%):', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(anchor='w', pady=(10, 2))
        self.min_avg_var = tk.StringVar(value=settings.get('min_passing_average', '50.0'))
        min_avg_entry = tk.Entry(settings_frame, textvariable=self.min_avg_var,
                                 font=(FF, 10), width=20)
        min_avg_entry.pack(anchor='w', pady=(0, 10))

        tk.Label(settings_frame, text='Academic Year Override (optional):', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(anchor='w', pady=(10, 2))
        self.promo_year_var = tk.StringVar(value=settings.get('promotion_academic_year', ''))
        promo_year_entry = tk.Entry(settings_frame, textvariable=self.promo_year_var,
                                    font=(FF, 10), width=20)
        promo_year_entry.pack(anchor='w', pady=(0, 10))
        
        # Auto-promote enabled
        self.auto_promote_var = tk.BooleanVar(value=settings.get('auto_promote_enabled', 'false').lower() == 'true')
        auto_promote_cb = tk.Checkbutton(settings_frame, text='Enable Auto-Promotion',
                                         variable=self.auto_promote_var,
                                         bg=CARD_BG, fg=TEXT_PRIMARY,
                                         font=(FF, 10), selectcolor=CARD_BG)
        auto_promote_cb.pack(anchor='w', pady=(10, 10))
        
        # Save settings button
        tk.Button(settings_frame, text='Save Settings', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=5,
                  command=self._save_promotion_settings).pack(anchor='w', pady=(10, 0))
        
        # Promotion status
        status_frame = tk.Frame(left_frame, bg=CARD_BG)
        status_frame.pack(fill='x', padx=15, pady=15)
        
        is_due = promotion_manager.is_promotion_due()
        academic_year = promotion_manager.get_current_academic_year()
        
        tk.Label(status_frame, text=f'Current Academic Year:', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(anchor='w')
        tk.Label(status_frame, text=academic_year, bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', pady=(0, 10))
        
        tk.Label(status_frame, text=f'Promotion Due:', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(anchor='w')
        status_color = '#2ecc71' if is_due else '#e74c3c'
        tk.Label(status_frame, text='Yes' if is_due else 'No', bg=CARD_BG, fg=status_color,
                 font=(FF, 11, 'bold')).pack(anchor='w', pady=(0, 10))

        tk.Label(status_frame, text='Configured Trigger Date:', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).pack(anchor='w')
        tk.Label(status_frame, text=settings.get('promotion_date', '12-01'), bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', pady=(0, 10))
        
        # Right column - Promotion Actions
        right_frame = tk.Frame(main_container, bg=CARD_BG, relief='flat', bd=1)
        right_frame.grid(row=0, column=1, sticky='nsew', padx=(5, 0))
        
        tk.Label(right_frame, text='Promotion Actions', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 14, 'bold')).pack(anchor='w', padx=15, pady=(15, 10))
        
        # Action buttons
        actions_frame = tk.Frame(right_frame, bg=CARD_BG)
        actions_frame.pack(fill='x', padx=15, pady=5)
        
        tk.Button(actions_frame, text='Preview Promotions', bg=BLUE, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=8,
                  command=self._preview_promotions).pack(fill='x', pady=(0, 10))
        
        tk.Button(actions_frame, text='Execute Promotions', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=8,
                  command=self._execute_promotions).pack(fill='x', pady=(0, 10))
        
        tk.Button(actions_frame, text='View Promotion History', bg=PURPLE, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=8,
                  command=self._view_promotion_history).pack(fill='x', pady=(0, 10))
        
        tk.Button(actions_frame, text='View Audit Log', bg=ORANGE, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=8,
                  command=self._view_promotion_audit_log).pack(fill='x', pady=(0, 10))
        
        # Statistics
        stats_frame = tk.Frame(right_frame, bg=CARD_BG)
        stats_frame.pack(fill='x', padx=15, pady=15)
        
        stats = promotion_manager.get_promotion_statistics()
        
        tk.Label(stats_frame, text='Promotion Statistics', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 12, 'bold')).pack(anchor='w', pady=(0, 10))
        
        tk.Label(stats_frame, text=f'Total Promotions: {stats.get("total_promotions", 0)}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w')
        tk.Label(stats_frame, text=f'Promoted: {stats.get("promoted_count", 0)}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w')
        tk.Label(stats_frame, text=f'Repeating: {stats.get("repeating_count", 0)}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w')
    
    def _save_promotion_settings(self):
        """Save promotion settings."""
        settings = {
            'promotion_date': self.promo_date_var.get().strip(),
            'min_passing_average': self.min_avg_var.get().strip(),
            'promotion_academic_year': self.promo_year_var.get().strip(),
            'auto_promote_enabled': 'true' if self.auto_promote_var.get() else 'false'
        }

        success, message = promotion_manager.update_settings(settings)
        if success:
            messagebox.showinfo('Success', message)
            self.show_promotions()
        else:
            messagebox.showerror('Error', message)
    
    def _preview_promotions(self):
        """Preview which students will be promoted or repeat."""
        preview = promotion_manager.get_promotion_preview()
        
        # Create preview window
        preview_window = tk.Toplevel(self.root)
        preview_window.title('Promotion Preview')
        preview_window.geometry('800x600')
        preview_window.configure(bg=CONTENT_BG)
        
        # Title
        tk.Label(preview_window, text='Promotion Preview', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 16, 'bold')).pack(pady=(20, 10))
        
        # Summary
        summary_frame = tk.Frame(preview_window, bg=CARD_BG, relief='flat', bd=1)
        summary_frame.pack(fill='x', padx=20, pady=10)
        
        tk.Label(summary_frame, text='Summary', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 12, 'bold')).pack(anchor='w', padx=15, pady=(10, 5))
        
        tk.Label(summary_frame, text=f'Eligible for Promotion: {len(preview["eligible"])}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w', padx=15)
        tk.Label(summary_frame, text=f'Will Repeat: {len(preview["repeating"])}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w', padx=15)
        tk.Label(summary_frame, text=f'No Exam Data: {len(preview["no_data"])}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w', padx=15, pady=(0, 10))
        tk.Label(summary_frame, text=f'Final Class / No Next Class: {len(preview["terminal"])}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w', padx=15)
        tk.Label(summary_frame, text=f'Already Processed This Year: {len(preview["already_processed"])}',
                 bg=CARD_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(anchor='w', padx=15, pady=(0, 10))
        
        # Notebook for tabs
        notebook = ttk.Notebook(preview_window)
        notebook.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Eligible students tab
        eligible_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(eligible_frame, text=f'Eligible ({len(preview["eligible"])})')
        
        if preview['eligible']:
            cols = ('name', 'admission_no', 'current_class', 'next_class', 'average')
            tree = ttk.Treeview(eligible_frame, columns=cols, show='headings', style='App.Treeview')
            tree.heading('name', text='Name')
            tree.heading('admission_no', text='Admission No')
            tree.heading('current_class', text='Current Class')
            tree.heading('next_class', text='Next Class')
            tree.heading('average', text='Average')
            
            for student in preview['eligible']:
                tree.insert('', 'end', values=(
                    student.get('name', ''),
                    student.get('admission_no', ''),
                    student.get('current_class', ''),
                    student.get('next_class', ''),
                    f"{student.get('average_marks', 0):.1f}%"
                ))
            
            tree.pack(fill='both', expand=True, padx=10, pady=10)
        else:
            tk.Label(eligible_frame, text='No students eligible for promotion',
                     bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(pady=50)
        
        # Repeating students tab
        repeating_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(repeating_frame, text=f'Repeating ({len(preview["repeating"])})')
        
        if preview['repeating']:
            cols = ('name', 'admission_no', 'current_class', 'average')
            tree = ttk.Treeview(repeating_frame, columns=cols, show='headings', style='App.Treeview')
            tree.heading('name', text='Name')
            tree.heading('admission_no', text='Admission No')
            tree.heading('current_class', text='Current Class')
            tree.heading('average', text='Average')
            
            for student in preview['repeating']:
                tree.insert('', 'end', values=(
                    student.get('name', ''),
                    student.get('admission_no', ''),
                    student.get('current_class', ''),
                    f"{student.get('average_marks', 0):.1f}%"
                ))
            
            tree.pack(fill='both', expand=True, padx=10, pady=10)
        else:
            tk.Label(repeating_frame, text='No students repeating',
                     bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(pady=50)

        no_data_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(no_data_frame, text=f'No Data ({len(preview["no_data"])})')

        if preview['no_data']:
            cols = ('name', 'admission_no', 'current_class')
            tree = ttk.Treeview(no_data_frame, columns=cols, show='headings', style='App.Treeview')
            tree.heading('name', text='Name')
            tree.heading('admission_no', text='Admission No')
            tree.heading('current_class', text='Current Class')

            for student in preview['no_data']:
                tree.insert('', 'end', values=(
                    student.get('name', ''),
                    student.get('admission_no', ''),
                    student.get('current_class', ''),
                ))

            tree.pack(fill='both', expand=True, padx=10, pady=10)
        else:
            tk.Label(no_data_frame, text='All students have marks for the latest exam session',
                     bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(pady=50)

        terminal_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(terminal_frame, text=f'Final Class ({len(preview["terminal"])})')

        if preview['terminal']:
            cols = ('name', 'admission_no', 'current_class', 'average', 'reason')
            tree = ttk.Treeview(terminal_frame, columns=cols, show='headings', style='App.Treeview')
            tree.heading('name', text='Name')
            tree.heading('admission_no', text='Admission No')
            tree.heading('current_class', text='Current Class')
            tree.heading('average', text='Average')
            tree.heading('reason', text='Reason')

            for student in preview['terminal']:
                tree.insert('', 'end', values=(
                    student.get('name', ''),
                    student.get('admission_no', ''),
                    student.get('current_class', ''),
                    f"{float(student.get('average_marks') or 0):.1f}%",
                    student.get('reason', ''),
                ))

            tree.pack(fill='both', expand=True, padx=10, pady=10)
        else:
            tk.Label(terminal_frame, text='Every passing class has a configured next class',
                     bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(pady=50)

        processed_frame = tk.Frame(notebook, bg=CONTENT_BG)
        notebook.add(processed_frame, text=f'Processed ({len(preview["already_processed"])})')

        if preview['already_processed']:
            cols = ('name', 'admission_no', 'current_class', 'status', 'to_class', 'promotion_date')
            tree = ttk.Treeview(processed_frame, columns=cols, show='headings', style='App.Treeview')
            tree.heading('name', text='Name')
            tree.heading('admission_no', text='Admission No')
            tree.heading('current_class', text='Current Class')
            tree.heading('status', text='Status')
            tree.heading('to_class', text='Processed To')
            tree.heading('promotion_date', text='Processed On')

            for student in preview['already_processed']:
                tree.insert('', 'end', values=(
                    student.get('name', ''),
                    student.get('admission_no', ''),
                    student.get('current_class', ''),
                    student.get('status', ''),
                    student.get('to_class', ''),
                    (student.get('promotion_date', '') or '')[:10],
                ))

            tree.pack(fill='both', expand=True, padx=10, pady=10)
        else:
            tk.Label(processed_frame, text='No students have been processed for this academic year yet',
                     bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10)).pack(pady=50)
        
        # Close button
        tk.Button(preview_window, text='Close', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=20, pady=8,
                  command=preview_window.destroy).pack(pady=20)
    
    def _execute_promotions(self):
        """Execute student promotions."""
        preview = promotion_manager.get_promotion_preview()
        total_to_process = len(preview['eligible']) + len(preview['repeating'])

        if total_to_process == 0:
            messagebox.showinfo(
                'Nothing To Process',
                'No students are ready for batch promotion.\n\n'
                f'No data: {len(preview["no_data"])}\n'
                f'Final class / no next class: {len(preview["terminal"])}\n'
                f'Already processed this year: {len(preview["already_processed"])}'
            )
            return

        # Confirm execution
        if not messagebox.askyesno('Confirm Promotion',
                                   'Are you sure you want to execute promotions?\n\n'
                                   f'Promote: {len(preview["eligible"])}\n'
                                   f'Repeat: {len(preview["repeating"])}\n'
                                   f'No data (skipped): {len(preview["no_data"])}\n'
                                   f'Final class / no next class (skipped): {len(preview["terminal"])}\n'
                                   f'Already processed this year (skipped): {len(preview["already_processed"])}\n\n'
                                   'This will update student classes and cannot be easily undone.'):
            return
        
        # Get current user ID
        user_id = self.current_user.get('id') if self.current_user else None
        
        # Execute promotion
        success, message, results = promotion_manager.execute_promotion(
            class_name=None,
            performed_by=user_id,
            dry_run=False
        )
        
        if success:
            messagebox.showinfo('Success',
                                f'Promotions completed successfully!\n\n'
                                f'Promoted: {results.get("promoted", 0)}\n'
                                f'Repeating: {results.get("repeating", 0)}\n'
                                f'No data skipped: {results.get("no_data", 0)}\n'
                                f'Final class skipped: {results.get("terminal", 0)}\n'
                                f'Already processed skipped: {results.get("already_processed", 0)}\n'
                                f'Failed: {results.get("failed", 0)}\n'
                                f'Batch ID: {results.get("batch_id", "n/a")}')
            # Refresh the page
            self.show_promotions()
        else:
            errors = results.get('errors', [])[:5] if isinstance(results, dict) else []
            error_text = '\n'.join(errors)
            messagebox.showerror(
                'Error',
                f'Promotion failed: {message}'
                + (f'\n\nDetails:\n{error_text}' if error_text else '')
            )
    
    def _view_promotion_history(self):
        """View promotion history."""
        history = promotion_manager.get_promotion_history()
        
        # Create history window
        history_window = tk.Toplevel(self.root)
        history_window.title('Promotion History')
        history_window.geometry('900x600')
        history_window.configure(bg=CONTENT_BG)
        
        # Title
        tk.Label(history_window, text='Promotion History', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 16, 'bold')).pack(pady=(20, 10))
        
        # Treeview
        cols = ('student_name', 'admission_no', 'from_class', 'to_class', 'status', 'academic_year', 'promotion_date', 'reason')
        tree = ttk.Treeview(history_window, columns=cols, show='headings', style='App.Treeview')
        tree.heading('student_name', text='Student Name')
        tree.heading('admission_no', text='Admission No')
        tree.heading('from_class', text='From Class')
        tree.heading('to_class', text='To Class')
        tree.heading('status', text='Status')
        tree.heading('academic_year', text='Academic Year')
        tree.heading('promotion_date', text='Promotion Date')
        tree.heading('reason', text='Reason')
        
        for record in history:
            tree.insert('', 'end', values=(
                record.get('student_name', ''),
                record.get('admission_no', ''),
                record.get('from_class', ''),
                record.get('to_class', ''),
                record.get('status', ''),
                record.get('academic_year', ''),
                record.get('promotion_date', '')[:10] if record.get('promotion_date') else '',
                record.get('reason', ''),
            ))
        
        tree.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Close button
        tk.Button(history_window, text='Close', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=20, pady=8,
                  command=history_window.destroy).pack(pady=20)
    
    def _view_promotion_audit_log(self):
        """View promotion audit log."""
        audit_log = promotion_manager.get_promotion_audit_log(limit=100)
        
        # Create audit log window
        audit_window = tk.Toplevel(self.root)
        audit_window.title('Promotion Audit Log')
        audit_window.geometry('900x600')
        audit_window.configure(bg=CONTENT_BG)
        
        # Title
        tk.Label(audit_window, text='Promotion Audit Log', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 16, 'bold')).pack(pady=(20, 10))
        
        # Treeview
        cols = ('batch_id', 'action', 'details', 'performed_by', 'performed_at')
        tree = ttk.Treeview(audit_window, columns=cols, show='headings', style='App.Treeview')
        tree.heading('batch_id', text='Batch ID')
        tree.heading('action', text='Action')
        tree.heading('details', text='Details')
        tree.heading('performed_by', text='Performed By')
        tree.heading('performed_at', text='Performed At')
        
        for log in audit_log:
            tree.insert('', 'end', values=(
                log.get('promotion_batch_id', ''),
                log.get('action', ''),
                log.get('details', ''),
                log.get('performed_by_name', 'System'),
                log.get('performed_at', '')[:19] if log.get('performed_at') else ''
            ))
        
        tree.pack(fill='both', expand=True, padx=20, pady=10)
        
        # Close button
        tk.Button(audit_window, text='Close', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=20, pady=8,
                  command=audit_window.destroy).pack(pady=20)

    # ==================== EXAM ANALYTICS ====================
    def show_exam_analytics(self):
        """Show exam analytics and comparison interface."""
        self.clear_frame()
        self._set_nav('Exam Analytics')
        self._page_header('Exam Analytics', 'Compare exam types and analyze deviations')
        
        # Main container
        main_container = tk.Frame(self.content_frame, bg=CONTENT_BG)
        main_container.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Top section - Filters and controls
        top_frame = tk.Frame(main_container, bg=CARD_BG, relief='flat', bd=1)
        top_frame.pack(fill='x', pady=(0, 10))
        
        tk.Label(top_frame, text='Analysis Filters', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 14, 'bold')).pack(anchor='w', padx=15, pady=(15, 10))
        
        # Filter controls
        filter_frame = tk.Frame(top_frame, bg=CARD_BG)
        filter_frame.pack(fill='x', padx=15, pady=5)
        
        # Class filter
        tk.Label(filter_frame, text='Class:', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.analytics_class_var = tk.StringVar(value='All Classes')
        class_options = ['All Classes'] + [c['name'] for c in db.get_all_classes()]
        class_cb = ttk.Combobox(filter_frame, textvariable=self.analytics_class_var,
                                values=class_options, state='readonly', width=20, style='App.TCombobox')
        class_cb.grid(row=0, column=1, padx=(0, 20))
        
        # Term filter
        tk.Label(filter_frame, text='Term:', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).grid(row=0, column=2, sticky='w', padx=(0, 10))
        self.analytics_term_var = tk.StringVar(value='All Terms')
        term_options = ['All Terms', 'One', 'Two', 'Three']
        term_cb = ttk.Combobox(filter_frame, textvariable=self.analytics_term_var,
                               values=term_options, state='readonly', width=15, style='App.TCombobox')
        term_cb.grid(row=0, column=3, padx=(0, 20))
        
        # Exam type filter
        tk.Label(filter_frame, text='Exam Type:', bg=CARD_BG, fg=TEXT_SECONDARY,
                 font=(FF, 10)).grid(row=0, column=4, sticky='w', padx=(0, 10))
        self.analytics_exam_type_var = tk.StringVar(value='All Types')
        available_exam_types = {
            row.get('exam_type', '')
            for row in db.get_available_exam_sessions()
            if str(row.get('exam_type', '')).strip()
        }
        exam_type_options = ['All Types'] + sorted(
            available_exam_types.union({'Opener', 'Mid-Term', 'End-Term'}),
            key=lambda value: {'Opener': 1, 'Quiz': 2, 'Assignment': 3, 'Mid-Term': 4, 'CAT': 5, 'End-Term': 6}.get(value, 999)
        )
        exam_type_cb = ttk.Combobox(filter_frame, textvariable=self.analytics_exam_type_var,
                                    values=exam_type_options, state='readonly', width=15, style='App.TCombobox')
        exam_type_cb.grid(row=0, column=5, padx=(0, 20))
        
        # Action buttons
        btn_frame = tk.Frame(top_frame, bg=CARD_BG)
        btn_frame.pack(fill='x', padx=15, pady=(10, 15))
        
        tk.Button(btn_frame, text='Run Analysis', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=5,
                  command=self._run_exam_analysis).pack(side='left', padx=(0, 10))
        tk.Button(btn_frame, text='Generate Report', bg=BLUE, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=5,
                  command=self._generate_analytics_report).pack(side='left', padx=(0, 10))
        tk.Button(btn_frame, text='Export Charts', bg=PURPLE, fg='white',
                  font=(FF, 10, 'bold'), padx=15, pady=5,
                  command=self._export_analytics_charts).pack(side='left')
        
        # Results section
        results_frame = tk.Frame(main_container, bg=CONTENT_BG)
        results_frame.pack(fill='both', expand=True)
        
        # Left column - Summary
        left_frame = tk.Frame(results_frame, bg=CARD_BG, relief='flat', bd=1)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        tk.Label(left_frame, text='Analysis Summary', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 12, 'bold')).pack(anchor='w', padx=15, pady=(15, 10))
        
        # Summary text
        self.analytics_summary_text = tk.Text(left_frame, wrap='word', font=(FF, 10),
                                              bg=CARD_BG, fg=TEXT_PRIMARY, height=15)
        self.analytics_summary_text.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        self.analytics_summary_text.insert('1.0', 'Click "Run Analysis" to generate exam analytics...')
        self.analytics_summary_text.config(state='disabled')
        
        # Right column - Deviation metrics
        right_frame = tk.Frame(results_frame, bg=CARD_BG, relief='flat', bd=1)
        right_frame.pack(side='right', fill='both', expand=True, padx=(5, 0))
        
        tk.Label(right_frame, text='Deviation Metrics', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 12, 'bold')).pack(anchor='w', padx=15, pady=(15, 10))
        
        # Deviation metrics treeview
        cols = ('metric', 'value', 'severity')
        self.analytics_tree = ttk.Treeview(right_frame, columns=cols, show='headings', style='App.Treeview', height=12)
        self.analytics_tree.heading('metric', text='Metric')
        self.analytics_tree.heading('value', text='Value')
        self.analytics_tree.heading('severity', text='Severity')
        
        self.analytics_tree.column('metric', width=200)
        self.analytics_tree.column('value', width=100)
        self.analytics_tree.column('severity', width=100)
        
        self.analytics_tree.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        # Bottom section - Subject deviation view
        bottom_frame = tk.Frame(main_container, bg=CARD_BG, relief='flat', bd=1)
        bottom_frame.pack(fill='x', pady=(10, 0))
        
        tk.Label(bottom_frame, text='Subject Deviation View', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 12, 'bold')).pack(anchor='w', padx=15, pady=(15, 10))
        
        # Subject deviation display area
        self.analytics_chart_frame = tk.Frame(bottom_frame, bg=CARD_BG)
        self.analytics_chart_frame.pack(fill='both', expand=True, padx=15, pady=(0, 15))
        
        subject_cols = ('subject', 'pair', 'baseline', 'comparison', 'deviation', 'pass_change')
        self.analytics_subject_tree = ttk.Treeview(
            self.analytics_chart_frame,
            columns=subject_cols,
            show='headings',
            style='App.Treeview',
            height=10
        )
        self.analytics_subject_tree.heading('subject', text='Subject')
        self.analytics_subject_tree.heading('pair', text='Exam Pair')
        self.analytics_subject_tree.heading('baseline', text='Baseline Mean')
        self.analytics_subject_tree.heading('comparison', text='Compared Mean')
        self.analytics_subject_tree.heading('deviation', text='Deviation')
        self.analytics_subject_tree.heading('pass_change', text='Pass Rate Shift')
        self.analytics_subject_tree.column('subject', width=150)
        self.analytics_subject_tree.column('pair', width=170)
        self.analytics_subject_tree.column('baseline', width=110, anchor='center')
        self.analytics_subject_tree.column('comparison', width=120, anchor='center')
        self.analytics_subject_tree.column('deviation', width=95, anchor='center')
        self.analytics_subject_tree.column('pass_change', width=110, anchor='center')
        self.analytics_subject_tree.pack(fill='both', expand=True)
    
    def _run_exam_analysis(self):
        """Run exam analysis based on selected filters."""
        # Get filter values
        class_name = self.analytics_class_var.get()
        term = self.analytics_term_var.get()
        exam_type = self.analytics_exam_type_var.get()
        
        # Convert 'All' options to None
        class_name = None if class_name == 'All Classes' else class_name
        term = None if term == 'All Terms' else term
        exam_type = None if exam_type == 'All Types' else exam_type
        
        # Get exam sessions
        sessions = exam_analytics.get_exam_sessions(
            class_name=class_name,
            term=term,
            exam_type=exam_type
        )
        
        if len(sessions) < 2:
            messagebox.showwarning('Insufficient Data',
                                 'At least 2 exam sessions are required for comparison analysis.')
            return
        
        # Run comparison
        comparison = exam_analytics.compare_exam_sessions(sessions)
        
        # Update summary
        self.analytics_summary_text.config(state='normal')
        self.analytics_summary_text.delete('1.0', 'end')
        
        summary = f"Exam Analysis Results\n{'='*50}\n\n"
        summary += f"Sessions Analyzed: {len(sessions)}\n"
        summary += f"Overall Similarity Score: {comparison.overall_similarity_score:.1f}/100\n\n"
        if comparison.anova_results:
            summary += f"ANOVA: {comparison.anova_results.get('interpretation', 'N/A')}\n"
        if comparison.time_series_results:
            summary += f"Trend: {comparison.time_series_results.get('interpretation', 'N/A')}\n"
        
        if comparison.exam_type_summaries:
            summary += "\nExam Type Summary:\n"
            for item in comparison.exam_type_summaries:
                summary += (
                    f"- {item['exam_type']}: mean {item['mean_score']:.1f}, "
                    f"pass {item['pass_rate']:.1f}%, difficulty {item['difficulty_index']:.3f}\n"
                )
            summary += "\n"
        
        summary += "Patterns Identified:\n"
        for pattern in comparison.patterns:
            summary += f"• {pattern}\n"
        
        summary += "\nAnomalies Detected:\n"
        if comparison.anomalies:
            for anomaly in comparison.anomalies:
                summary += f"• {anomaly}\n"
        else:
            summary += "• No significant anomalies detected\n"
        
        summary += "\nRecommendations:\n"
        for i, rec in enumerate(comparison.recommendations, 1):
            summary += f"{i}. {rec}\n"
        
        self.analytics_summary_text.insert('1.0', summary)
        self.analytics_summary_text.config(state='disabled')
        
        # Update deviation metrics treeview
        for item in self.analytics_tree.get_children():
            self.analytics_tree.delete(item)
        
        for dev in comparison.deviations:
            metric_name = dev.metric.value.replace('_', ' ').title()
            self.analytics_tree.insert('', 'end', values=(
                metric_name,
                f"{dev.value:.3f}",
                dev.severity.upper()
            ))
        
        for item in self.analytics_subject_tree.get_children():
            self.analytics_subject_tree.delete(item)
        
        for row in comparison.subject_deviation_rows:
            self.analytics_subject_tree.insert('', 'end', values=(
                row.get('subject', ''),
                f"{row.get('baseline_exam_type', '')} -> {row.get('comparison_exam_type', '')}",
                f"{row.get('baseline_mean', 0):.1f}",
                f"{row.get('comparison_mean', 0):.1f}",
                f"{row.get('score_deviation', 0):+.1f}",
                f"{row.get('pass_rate_deviation', 0):+.1f}%",
            ))
        
        # Store comparison for later use
        self.current_comparison = comparison
        
        messagebox.showinfo('Analysis Complete',
                           f'Analysis completed successfully!\n\n'
                           f'Sessions analyzed: {len(sessions)}\n'
                           f'Similarity score: {comparison.overall_similarity_score:.1f}/100')
    
    def _generate_analytics_report(self):
        """Generate and display a detailed analytics report."""
        if not hasattr(self, 'current_comparison'):
            messagebox.showwarning('No Analysis', 'Please run an analysis first.')
            return
        
        # Generate report
        report = exam_analytics.generate_comparison_report(self.current_comparison)
        
        # Create report window
        report_window = tk.Toplevel(self.root)
        report_window.title('Exam Analytics Report')
        report_window.geometry('800x600')
        report_window.configure(bg=CONTENT_BG)
        
        # Title
        tk.Label(report_window, text='Exam Analytics Report', bg=CONTENT_BG, fg=TEXT_PRIMARY,
                 font=(FF, 16, 'bold')).pack(pady=(20, 10))
        
        # Report text
        report_text = tk.Text(report_window, wrap='word', font=(FF, 10))
        report_text.pack(fill='both', expand=True, padx=20, pady=10)
        report_text.insert('1.0', report)
        report_text.config(state='disabled')
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(report_window, orient='vertical', command=report_text.yview)
        scrollbar.pack(side='right', fill='y', padx=(0, 20), pady=10)
        report_text.config(yscrollcommand=scrollbar.set)
        
        # Close button
        tk.Button(report_window, text='Close', bg=GREEN, fg='white',
                  font=(FF, 10, 'bold'), padx=20, pady=8,
                  command=report_window.destroy).pack(pady=20)
    
    def _export_analytics_charts(self):
        """Export analytics charts to files."""
        if not hasattr(self, 'current_comparison'):
            messagebox.showwarning('No Analysis', 'Please run an analysis first.')
            return
        
        # Ask for export directory
        export_dir = filedialog.askdirectory(title='Select Export Directory')
        if not export_dir:
            return
        
        try:
            # Create dashboard charts
            charts = exam_analytics.create_comparison_dashboard(
                self.current_comparison, export_dir
            )
            
            messagebox.showinfo('Export Complete',
                               f'Charts exported successfully to:\n{export_dir}\n\n'
                               f'Files created:\n' + '\n'.join(charts.values()))
        except Exception as e:
            messagebox.showerror('Export Error', f'Failed to export charts: {str(e)}')

    # ==================== STUDENTS ====================
    def show_students(self):
        self.clear_frame()
        self._set_nav('Students')
        self._page_header('Students', 'Manage student registrations')

        # toolbar
        tb = tk.Frame(self.content_frame, bg=CONTENT_BG)
        tb.pack(fill='x', pady=(0, 12))

        self._toolbar_btn(tb, '+ Add Student', self.add_student).pack(side='left')
        self._toolbar_btn(tb, 'Edit Selected', self.edit_student, bg=BLUE).pack(side='left', padx=10)
        self._toolbar_btn(tb, 'Delete Selected', self.delete_student, bg='#d9534f').pack(side='left')
        self._toolbar_btn(tb, '📥 Template', self.download_template, bg=ORANGE).pack(side='left', padx=10)
        self._toolbar_btn(tb, '📥 Import Excel', self.import_excel, bg=BLUE).pack(side='left')
        self._toolbar_btn(tb, '📤 Export Excel', self.export_excel, bg=PURPLE).pack(side='left', padx=10)

        class_options = [row.get('name', '') for row in db.get_all_classes() if row.get('name', '')]
        target_values = ['Use Class From File'] + class_options
        self.students_import_class_var = tk.StringVar(value='Use Class From File')
        tk.Label(tb, text='Import To Class:', bg=CONTENT_BG, fg=TEXT_SECONDARY, font=(FF, 10, 'bold')).pack(side='left', padx=(14, 6))
        self.students_import_class_cb = ttk.Combobox(
            tb, textvariable=self.students_import_class_var,
            values=target_values, state='readonly', width=22, style='App.TCombobox'
        )
        self.students_import_class_cb.pack(side='left', ipady=2)

        # table card
        st_bo, st_bi = _card_colors('peach')
        tc_outer = tk.Frame(self.content_frame, bg=st_bo)
        tc_outer.pack(fill='both', expand=True, pady=4)
        tc = tk.Frame(tc_outer, bg=st_bi)
        tc.pack(fill='both', expand=True, padx=1, pady=1)

        self.students_table = AdvancedDataTable(
            tc,
            columns=[
                {'key': 'adm_no', 'title': 'Admission No', 'width': 140, 'anchor': 'center'},
                {'key': 'name', 'title': 'Name', 'width': 250, 'anchor': 'w'},
                {'key': 'class', 'title': 'Class', 'width': 120, 'anchor': 'center'},
                {'key': 'stream', 'title': 'Stream', 'width': 110, 'anchor': 'center'},
                {'key': 'gender', 'title': 'Gender', 'width': 100, 'anchor': 'center'},
            ],
            page_size=20,
            search_label='Search students',
            selectmode='extended',
            enable_select_all=True
        )
        self.students_tree = self.students_table.tree

        self.students_tree.bind('<Double-Button-1>', lambda e: self.edit_student())
        self.students_tree.bind('<Button-3>', self._student_ctx)
        self.load_students()

    def load_students(self):
        rows = []
        for s in db.get_all_students():
            values = (
                s.get('admission_no', ''),
                s.get('name', ''),
                s.get('class', ''),
                s.get('stream', ''),
                s.get('gender', ''),
            )
            rows.append({
                'iid': s.get('id', str(uuid.uuid4())),
                'values': values,
                'tags': (s.get('id', ''), s.get('photo_path', '')),
                'value_map': {
                    'adm_no': values[0],
                    'name': values[1],
                    'class': values[2],
                    'stream': values[3],
                    'gender': values[4],
                },
                'search': ' '.join(str(v) for v in values),
            })
        if hasattr(self, 'students_table'):
            self.students_table.set_rows(rows, reset_page=False)

    def filter_students(self):
        if hasattr(self, 'students_table'):
            self.students_table.apply_filters(reset_page=True)

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
                        guardian_name='', parent_email='', stream='', on_save=None):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry('500x590')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        sd_bo, sd_bi = _card_colors('lilac')
        outer = tk.Frame(dlg, bg=sd_bo)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=sd_bi, padx=30, pady=24)
        card.pack(padx=1, pady=1)

        tk.Label(card, text=title, bg=sd_bi, fg=TEXT_PRIMARY,
                 font=(FF, 14, 'bold')).pack(anchor='w', pady=(0, 20))

        # Photo selection
        photo_frame = tk.Frame(card, bg=sd_bi)
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
            tk.Label(card, text=label_text, bg=sd_bi, fg=TEXT_SECONDARY,
                     font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
            e = ttk.Entry(card, style='App.TEntry')
            e.insert(0, default)
            e.pack(fill='x', ipady=4, pady=(0, 10))
            return e

        adm_e  = lbl_entry('Admission No', adm)
        name_e = lbl_entry('Full Name', name)
        guardian_e = lbl_entry('Guardian Name', guardian_name)
        parent_email_e = lbl_entry('Parent Email', parent_email)

        tk.Label(card, text='Class', bg=sd_bi, fg=TEXT_SECONDARY,
                 font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
        cls_cb = ttk.Combobox(card, values=self.get_current_classes(), state='readonly', style='App.TCombobox')
        cls_cb.set(cls)
        cls_cb.pack(fill='x', ipady=4, pady=(0, 10))

        tk.Label(card, text='Stream', bg=sd_bi, fg=TEXT_SECONDARY,
                 font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
        stream_cb = ttk.Combobox(card, style='App.TCombobox')
        stream_cb.pack(fill='x', ipady=4, pady=(0, 10))

        def refresh_streams():
            current = stream_cb.get().strip()
            values = self._get_stream_names_for_class(cls_cb.get())
            stream_cb['values'] = values
            if current and (current in values or not values):
                stream_cb.set(current)
            elif stream and stream in values:
                stream_cb.set(stream)
            elif values:
                stream_cb.set(values[0])
            else:
                stream_cb.set(current or stream)

        refresh_streams()
        cls_cb.bind('<<ComboboxSelected>>', lambda e: refresh_streams())

        tk.Label(card, text='Gender', bg=sd_bi, fg=TEXT_SECONDARY,
                 font=(FF, 10, 'bold'), anchor='w').pack(fill='x', pady=(0, 3))
        gen_cb = ttk.Combobox(card, values=['Male', 'Female'], state='readonly',
                              style='App.TCombobox')
        gen_cb.set(gender)
        gen_cb.pack(fill='x', ipady=4, pady=(0, 18))

        btn_row = tk.Frame(card, bg=sd_bi)
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
            on_save(
                adm_e.get().strip(),
                name_e.get().strip(),
                cls_cb.get(),
                gen_cb.get(),
                self.temp_photo_path,
                guardian_e.get().strip(),
                parent_email_e.get().strip(),
                stream_cb.get().strip(),
            )
            dlg.destroy()

        save = tk.Label(btn_row, text='Save', bg=BLUE, fg='white',
                        font=(FF, 10, 'bold'), padx=20, pady=8, cursor='hand2')
        save.pack(side='left')
        save.bind('<Button-1>', lambda e: do_save())

    def add_student(self):
        def on_save(adm, name, cls, gender, photo, guardian_name, parent_email, stream):
            db.add_student(name, cls, gender, adm, photo, guardian_name, parent_email, stream)
            self.load_students()
        self._student_dialog('Add Student', on_save=on_save)

    def edit_student(self):
        selected_iids = self.students_table.get_selected_iids() if hasattr(self, 'students_table') else list(self.students_tree.selection())
        if len(selected_iids) != 1:
            messagebox.showwarning('Select One Student', 'Please select exactly one student to edit')
            return
        sid = selected_iids[0]
        item = self.students_tree.item(sid)
        tags = item.get('tags', ())
        photo = tags[1] if len(tags) > 1 else ""
        student = next((row for row in db.get_all_students() if row.get('id') == sid), {})
        def on_save(a, n, c, g, p, guardian_name, parent_email, stream):
            db.update_student(sid, n, c, g, a, p, guardian_name, parent_email, stream)
            self.load_students()
        self._student_dialog(
            'Edit Student',
            student.get('admission_no', ''),
            student.get('name', ''),
            student.get('class', CLASSES[0]),
            student.get('gender', 'Male'),
            photo,
            student.get('guardian_name', ''),
            student.get('parent_email', ''),
            student.get('stream', ''),
            on_save=on_save
        )

    def delete_student(self):
        selected_iids = self.students_table.get_selected_iids() if hasattr(self, 'students_table') else list(self.students_tree.selection())
        if not selected_iids:
            messagebox.showwarning('Select Student', 'Please select one or more students to delete')
            return
        if not messagebox.askyesno('Confirm Delete', f'Delete {len(selected_iids)} selected student(s)?', parent=self.root):
            return
        failed = 0
        for sid in selected_iids:
            if not sid or not db.delete_student(sid):
                failed += 1
        self.load_students()
        if failed:
            messagebox.showerror('Delete Failed', f'Failed to delete {failed} student(s).')

    def import_excel(self):
        """Import students from an Excel file"""
        import_target_class = ''
        if hasattr(self, 'students_import_class_var'):
            selected_target = (self.students_import_class_var.get() or '').strip()
            if selected_target and selected_target != 'Use Class From File':
                import_target_class = selected_target

        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if not file_path:
            return
        
        try:
            workbook = pd.read_excel(file_path, sheet_name=None)
            if not isinstance(workbook, dict):
                workbook = {'Sheet1': workbook}

            col_aliases = {
                'admission_no': {'admission_no', 'admission no', 'adm_no', 'adm no', 'adm', 'reg_no', 'reg no'},
                'name': {'name', 'student name', 'student_name', 'full name', 'full_name', 'learner'},
                'class': {'class', 'grade', 'class_name'},
                'stream': {'stream', 'student_stream', 'class_stream'},
                'gender': {'gender', 'sex'},
                'photo_path': {'photo', 'photo_path', 'photo path', 'image', 'image_path'},
            }

            known_classes = [row.get('name', '') for row in db.get_all_classes() if row.get('name', '')]

            def clean_value(value):
                value = str(value or '').strip()
                return '' if value.lower() == 'nan' else value

            def resolve_sheet_class(sheet_name):
                norm_sheet = self._normalize_text(sheet_name)
                for cls_name in known_classes:
                    if self._normalize_text(cls_name) == norm_sheet:
                        return cls_name
                return ''

            prepared_sheets = []
            total_rows = 0
            for sheet_name, raw_df in workbook.items():
                if raw_df is None or raw_df.empty:
                    continue
                df = raw_df.copy()
                df.columns = [self._normalize_text(c) for c in df.columns]

                def find_col(alias_key):
                    aliases = col_aliases[alias_key]
                    return next((col for col in df.columns if col in aliases), None)

                name_col = find_col('name')
                if not name_col:
                    continue

                prepared_sheets.append({
                    'sheet_name': str(sheet_name),
                    'df': df,
                    'name_col': name_col,
                    'adm_col': find_col('admission_no'),
                    'class_col': find_col('class'),
                    'stream_col': find_col('stream'),
                    'gender_col': find_col('gender'),
                    'photo_col': find_col('photo_path'),
                    'sheet_default_class': resolve_sheet_class(str(sheet_name)),
                })
                total_rows += len(df.index)

            if not prepared_sheets:
                messagebox.showerror(
                    "Error",
                    "No valid worksheets found.\nEach sheet should include at least a learner name column."
                )
                return

            current_classes = self.get_current_classes()
            default_class = current_classes[0] if current_classes else 'Grade 1'

            def show_import_preview_dialog():
                top = tk.Toplevel(self.root)
                top.title('Import Preview')
                top.geometry('860x520')
                top.transient(self.root)
                top.grab_set()
                top.resizable(False, False)

                tk.Label(
                    top,
                    text='Workbook Import Preview',
                    bg=CONTENT_BG,
                    fg=TEXT_PRIMARY,
                    font=(FF, 12, 'bold')
                ).pack(anchor='w', padx=16, pady=(14, 4))

                scope_txt = import_target_class if import_target_class else 'Use class from each sheet/row'
                tk.Label(
                    top,
                    text=f'Selected import scope: {scope_txt}',
                    bg=CONTENT_BG,
                    fg=TEXT_SECONDARY,
                    font=(FF, 10)
                ).pack(anchor='w', padx=16, pady=(0, 10))

                frame = tk.Frame(top, bg=CARD_BG)
                frame.pack(fill='both', expand=True, padx=16, pady=(0, 12))
                cols = ('sheet', 'rows', 'name_col', 'class_col', 'mapped')
                tv = ttk.Treeview(frame, columns=cols, show='headings', style='App.Treeview')
                tv.heading('sheet', text='Sheet')
                tv.heading('rows', text='Rows')
                tv.heading('name_col', text='Name Column')
                tv.heading('class_col', text='Class Column')
                tv.heading('mapped', text='Mapped Class')
                tv.column('sheet', width=220, anchor='w')
                tv.column('rows', width=70, anchor='center')
                tv.column('name_col', width=150, anchor='w')
                tv.column('class_col', width=150, anchor='w')
                tv.column('mapped', width=220, anchor='w')
                sb = ttk.Scrollbar(frame, orient='vertical', command=tv.yview, style='App.Vertical.TScrollbar')
                tv.configure(yscrollcommand=sb.set)
                tv.pack(side='left', fill='both', expand=True, padx=(0, 4), pady=4)
                sb.pack(side='right', fill='y', pady=4)

                for pack in prepared_sheets:
                    mapped = import_target_class or pack.get('sheet_default_class') or f"From '{pack.get('class_col') or 'class'}' column / fallback {default_class}"
                    class_col_label = pack.get('class_col') or '-'
                    tv.insert('', 'end', values=(
                        pack.get('sheet_name', ''),
                        len(pack.get('df', [])),
                        pack.get('name_col', ''),
                        class_col_label,
                        mapped,
                    ))

                decision = {'ok': False}
                btn_row = tk.Frame(top, bg=CONTENT_BG)
                btn_row.pack(fill='x', padx=16, pady=(0, 14))
                tk.Button(
                    btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY, font=(FF, 10, 'bold'),
                    padx=18, pady=8, command=top.destroy
                ).pack(side='left')
                tk.Button(
                    btn_row, text='Proceed Import', bg=GREEN, fg='white', font=(FF, 10, 'bold'),
                    padx=18, pady=8, command=lambda: (decision.__setitem__('ok', True), top.destroy())
                ).pack(side='left', padx=(8, 0))

                self.root.wait_window(top)
                return decision['ok']

            if not show_import_preview_dialog():
                return

            progress_dialog, status_label, percent_label, progress = self._open_progress_dialog(
                'Importing Students', 'Preparing student rows...'
            )

            imported_count = 0
            updated_count  = 0
            generated_count = 0
            processed = 0
            for sheet_pack in prepared_sheets:
                sheet_name = sheet_pack['sheet_name']
                df = sheet_pack['df']
                name_col = sheet_pack['name_col']
                adm_col = sheet_pack['adm_col']
                class_col = sheet_pack['class_col']
                stream_col = sheet_pack['stream_col']
                gender_col = sheet_pack['gender_col']
                photo_col = sheet_pack['photo_col']
                sheet_default_class = sheet_pack['sheet_default_class']

                for index, (_, row) in enumerate(df.iterrows(), start=1):
                    self._update_progress_dialog(
                        progress_dialog, status_label, percent_label, progress,
                        processed, total_rows,
                        f'[{sheet_name}] Processing row {index}...'
                    )
                    processed += 1

                    name = clean_value(row.get(name_col, ''))
                    admission_no = clean_value(row.get(adm_col, '')) if adm_col else ''
                    cls = clean_value(row.get(class_col, '')) if class_col else ''
                    stream = clean_value(row.get(stream_col, '')) if stream_col else ''
                    gender = clean_value(row.get(gender_col, '')) if gender_col else ''
                    photo_path = clean_value(row.get(photo_col, '')) if photo_col else ''

                    if not name:
                        continue
                    if import_target_class:
                        cls = import_target_class
                    elif not cls and sheet_default_class:
                        cls = sheet_default_class
                    if not cls:
                        cls = default_class
                    if gender not in ('Male', 'Female'):
                        gender = 'Male'
                    if not admission_no:
                        admission_no = self._generate_admission_no(cls, name)
                        generated_count += 1

                    # Upsert: update existing student (preserving photo if not provided),
                    # or add new student.
                    existing = db.get_student_by_admission_no(admission_no)
                    if existing:
                        db.update_student(
                            existing['id'],
                            name,
                            cls,
                            gender,
                            admission_no,
                            photo_path,
                            existing.get('guardian_name', ''),
                            existing.get('parent_email', ''),
                            stream or existing.get('stream', '')
                        )
                        updated_count += 1
                    else:
                        db.add_student(name, cls, gender, admission_no, photo_path, '', '', stream)
                        imported_count += 1

            self._update_progress_dialog(
                progress_dialog, status_label, percent_label, progress,
                total_rows, total_rows,
                'Refreshing student list...'
            )
            self.load_students()
            progress_dialog.destroy()
            scope_note = f"\nImport target class: {import_target_class}" if import_target_class else "\nImport target class: From file values"
            messagebox.showinfo("Import Complete",
                                f"Done!  {imported_count} new student(s) added, "
                                f"{updated_count} updated.\n"
                                f"Generated admission numbers: {generated_count}\n"
                                f"Sheets processed: {len(prepared_sheets)}\n"
                                f"Existing photos were preserved where no new photo was supplied."
                                f"{scope_note}")
             
        except Exception as e:
            try:
                progress_dialog.destroy()
            except Exception:
                pass
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
                'name': ['John Doe', 'Jane Smith', 'Michael Johnson'],
                'class': ['Grade 7', 'Grade 8', 'Grade 9'],
                'stream': ['Blue', 'Red', ''],
                'admission_no (optional)': ['001', '', '003'],
                'gender (optional)': ['Male', '', 'Female'],
                'photo_path': ['', '', '']
            }
            df = pd.DataFrame(template_data)
            df.to_excel(file_path, index=False)
            messagebox.showinfo("Template Created", f"Template saved to: {file_path}")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to create template: {str(e)}")

    # ==================== ENTER MARKS ====================
    def show_marks_entry(self):
        """Show class selection page with all classes (Grade 1-9) and their streams"""
        self.clear_frame()
        self._set_nav('Enter Marks')
        self._page_header('Enter Marks', 'Select a class and stream to enter marks')

        def _stream_palette(stream_name):
            """Return (bg, fg, hover_bg, border) colors for a stream button."""
            name = (stream_name or '').strip().lower()
            if 'green' in name:
                return ('#2E7D32', '#FFFFFF', '#1B5E20', '#A5D6A7')
            if 'yellow' in name:
                return ('#F9A825', '#1F1F1F', '#F57F17', '#FFE082')
            if 'blue' in name:
                return ('#1976D2', '#FFFFFF', '#0D47A1', '#90CAF9')
            if 'red' in name:
                return ('#D32F2F', '#FFFFFF', '#B71C1C', '#EF9A9A')
            return (OLIVE_PRIMARY, '#FFFFFF', OLIVE_DARK, '#C5D29A')

        # Create scrollable container for class cards
        canvas_outer = tk.Frame(self.content_frame, bg=CONTENT_BG)
        canvas_outer.pack(fill='both', expand=True, padx=10, pady=10)

        canvas = tk.Canvas(canvas_outer, bg=CONTENT_BG, highlightthickness=0)
        scrollbar = ttk.Scrollbar(canvas_outer, orient='vertical', command=canvas.yview, style='App.Vertical.TScrollbar')
        scrollable_frame = tk.Frame(canvas, bg=CONTENT_BG)

        scrollable_frame.bind(
            '<Configure>',
            lambda e: canvas.configure(scrollregion=canvas.bbox('all'))
        )

        canvas_window = canvas.create_window((0, 0), window=scrollable_frame, anchor='nw')
        canvas.configure(yscrollcommand=scrollbar.set)

        scrollbar.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        # Mouse wheel scrolling
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), 'units')
        canvas.bind_all('<MouseWheel>', _on_mousewheel)

        # Keep scrollable content width synced with canvas so responsive grids can use full page width
        def _sync_canvas_width(event):
            canvas.itemconfigure(canvas_window, width=event.width)
        canvas.bind('<Configure>', _sync_canvas_width)

        # Get all classes from Grade 1-9
        all_classes = []
        for level_name, classes in CLASSES_BY_LEVEL.items():
            for class_name in classes:
                all_classes.append(class_name)

        # Group classes by grade level for better organization
        grade_levels = {
            'Lower Primary (Grade 1-3)': ['Grade 1', 'Grade 2', 'Grade 3'],
            'Upper Primary (Grade 4-6)': ['Grade 4', 'Grade 5', 'Grade 6'],
            'Junior School (Grade 7-9)': ['Grade 7', 'Grade 8', 'Grade 9']
        }

        card_themes_list = list(CARD_THEMES.keys())
        theme_idx = 0

        for level_name, classes in grade_levels.items():
            # Level header
            level_frame = tk.Frame(scrollable_frame, bg=CONTENT_BG)
            level_frame.pack(fill='x', pady=(10, 5), padx=10)
            
            tk.Label(level_frame, text=level_name, bg=CONTENT_BG, fg=OLIVE_PRIMARY,
                    font=(FF, 14, 'bold')).pack(anchor='w', pady=(0, 10))

            # Responsive cards grid (1-3 columns) per level
            cards_grid = tk.Frame(scrollable_frame, bg=CONTENT_BG)
            cards_grid.pack(fill='x', padx=10, pady=(0, 6))
            level_card_outers = []

            def _reflow_level_cards(event=None, frame=cards_grid, cards=level_card_outers):
                width = frame.winfo_width()
                if width and width < 760:
                    cols = 1
                elif width and width < 1180:
                    cols = 2
                else:
                    cols = 3
                for c in range(4):
                    frame.grid_columnconfigure(c, weight=0)
                for c in range(cols):
                    frame.grid_columnconfigure(c, weight=1, uniform=f"level_{level_name}")
                for i, card_outer in enumerate(cards):
                    card_outer.grid_forget()
                    row = i // cols
                    col = i % cols
                    card_outer.grid(row=row, column=col, padx=6, pady=6, sticky='nsew')

            # Create cards for each class in this level
            for class_name in classes:
                # Get class info and streams
                class_row = db.get_class_by_name(class_name)
                streams = []
                if class_row:
                    streams = db.get_streams_for_class(class_row['id'])
                
                # Get student count
                students = db.get_students_by_class(class_name)
                student_count = len([s for s in students if not self._is_summary_student_name(s.get('name'))])

                # Create class card
                theme_key = card_themes_list[theme_idx % len(card_themes_list)]
                theme_idx += 1
                border_clr, card_bg = _card_colors(theme_key)

                card_outer = tk.Frame(cards_grid, bg=border_clr, padx=2, pady=2)
                level_card_outers.append(card_outer)

                card = tk.Frame(card_outer, bg=card_bg, padx=20, pady=15)
                card.pack(fill='both', expand=True)

                # Class header
                header_frame = tk.Frame(card, bg=card_bg)
                header_frame.pack(fill='x', pady=(0, 10))

                tk.Label(header_frame, text=f"📚 {class_name}", bg=card_bg, fg=OLIVE_PRIMARY,
                        font=(FF, 13, 'bold')).pack(side='left')
                
                tk.Label(header_frame, text=f"{student_count} students", bg=card_bg, fg=TEXT_SECONDARY,
                        font=(FF, 10)).pack(side='right', padx=10)

                # Streams section
                if streams:
                    streams_frame = tk.Frame(card, bg=card_bg)
                    streams_frame.pack(fill='x', pady=5)

                    tk.Label(streams_frame, text="Streams:", bg=card_bg, fg=TEXT_SECONDARY,
                            font=(FF, 10, 'bold')).pack(anchor='w', pady=(0, 5))

                    # Create buttons for each stream
                    buttons_frame = tk.Frame(streams_frame, bg=card_bg)
                    buttons_frame.pack(fill='x')
                    stream_chips = []

                    def _reflow_stream_buttons(event=None, frame=buttons_frame, chips=stream_chips):
                        width = frame.winfo_width()
                        cols = 1 if width and width < 360 else 2
                        for c in range(4):
                            frame.grid_columnconfigure(c, weight=0)
                        for c in range(cols):
                            frame.grid_columnconfigure(c, weight=1)
                        for i, chip in enumerate(chips):
                            chip.grid_forget()
                            row = i // cols
                            col = i % cols
                            chip.grid(row=row, column=col, padx=5, pady=4, sticky='ew')

                    visible_idx = 0
                    for stream in streams:
                        stream_name = stream.get('name', '').strip()
                        if stream_name:
                            # Get student count for this stream
                            stream_students = db.get_students_by_class_and_stream(class_name, stream_name)
                            stream_count = len([s for s in stream_students if not self._is_summary_student_name(s.get('name'))])
                            stream_bg, stream_fg, stream_hover_bg, stream_border = _stream_palette(stream_name)

                            stream_chip = tk.Frame(buttons_frame, bg=stream_border, padx=1, pady=1)
                            stream_chips.append(stream_chip)

                            btn = tk.Button(
                                stream_chip,
                                text=f"{stream_name} ({stream_count})",
                                bg=stream_bg,
                                fg=stream_fg,
                                font=(FF, 10, 'bold'),
                                relief='flat',
                                bd=0,
                                padx=15,
                                pady=8,
                                cursor='hand2',
                                command=lambda c=class_name, s=stream_name: self._show_marks_entry_form(c, s)
                            )
                            btn.pack(fill='x')

                            # Hover effects
                            def on_enter(e, b=btn, c=stream_hover_bg):
                                b.configure(bg=c)
                            def on_leave(e, b=btn, c=stream_bg):
                                b.configure(bg=c)
                            btn.bind('<Enter>', on_enter)
                            btn.bind('<Leave>', on_leave)
                            visible_idx += 1

                    buttons_frame.bind('<Configure>', _reflow_stream_buttons)
                    _reflow_stream_buttons()
                else:
                    # No streams - show button for entire class
                    btn_frame = tk.Frame(card, bg=card_bg)
                    btn_frame.pack(fill='x', pady=5)

                    btn = tk.Button(
                        btn_frame,
                        text=f"Enter Marks for {class_name}",
                        bg=OLIVE_PRIMARY,
                        fg='white',
                        font=(FF, 11, 'bold'),
                        relief='flat',
                        padx=20,
                        pady=10,
                        cursor='hand2',
                        command=lambda c=class_name: self._show_marks_entry_form(c, '')
                    )
                    btn.pack(anchor='w')

                    # Hover effects
                    def on_enter(e, b=btn):
                        b.configure(bg=OLIVE_DARK)
                    def on_leave(e, b=btn):
                        b.configure(bg=OLIVE_PRIMARY)
                    btn.bind('<Enter>', on_enter)
                    btn.bind('<Leave>', on_leave)

            cards_grid.bind('<Configure>', _reflow_level_cards)
            _reflow_level_cards()

    def _show_marks_entry_form(self, class_name, stream=''):
        """Show the actual marks entry form for a specific class and stream"""
        self.clear_frame()
        self._set_nav('Enter Marks')
        
        # Build header text
        header_text = f'Enter Marks - {class_name}'
        if stream:
            header_text += f' (Stream {stream})'
        
        self._page_header('Enter Marks', header_text)

        # Back button
        back_frame = tk.Frame(self.content_frame, bg=CONTENT_BG)
        back_frame.pack(fill='x', pady=(0, 10))
        
        back_btn = tk.Button(
            back_frame,
            text='← Back to Class Selection',
            bg=OLIVE_MID,
            fg='white',
            font=(FF, 10, 'bold'),
            relief='flat',
            padx=15,
            pady=8,
            cursor='hand2',
            command=self.show_marks_entry
        )
        back_btn.pack(side='left', padx=10)

        # controls
        ctrl = tk.Frame(self.content_frame, bg=CONTENT_BG)
        ctrl.pack(fill='x', pady=(0, 12))

        def lbl(text): tk.Label(ctrl, text=text, bg=CONTENT_BG, fg=TEXT_SECONDARY,
                                 font=(FF, 10)).pack(side='left', padx=(10, 4))

        lbl('Class:')
        self.marks_class_cb = ttk.Combobox(ctrl, values=self.get_current_classes(), state='readonly',
                                           style='App.TCombobox', width=12)
        self.marks_class_cb.set(class_name)
        self.marks_class_cb.pack(side='left', ipady=4)

        # Store the selected stream
        self._selected_marks_stream = stream

        lbl('Term:')
        self.marks_term_cb = ttk.Combobox(ctrl, values=TERMS, state='readonly',
                                          style='App.TCombobox', width=10)
        self.marks_term_cb.set(TERMS[0])
        self.marks_term_cb.pack(side='left', ipady=4)

        lbl('Exam:')
        self.marks_exam_cb = ttk.Combobox(ctrl, values=EXAM_TYPES, state='readonly',
                                          style='App.TCombobox', width=12)
        self.marks_exam_cb.set(DEFAULT_EXAM_TYPE)
        self.marks_exam_cb.pack(side='left', ipady=4)

        self.marks_class_cb.bind('<<ComboboxSelected>>', lambda e: self._load_marks_table())
        self.marks_term_cb.bind('<<ComboboxSelected>>', lambda e: self._load_marks_table())
        self.marks_exam_cb.bind('<<ComboboxSelected>>', lambda e: self._load_marks_table())

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
        mk_bo, mk_bi = _card_colors('sand')
        tc_outer = tk.Frame(self.content_frame, bg=mk_bo)
        tc_outer.pack(fill='both', expand=True, pady=4)
        self.marks_card = tk.Frame(tc_outer, bg=mk_bi, padx=1, pady=1)
        self.marks_card.pack(fill='both', expand=True)
        self.marks_panel_bg = mk_bi

        # sticky header + scrollable rows grid
        self.marks_header_canvas = tk.Canvas(self.marks_card, bg=mk_bi, highlightthickness=0, height=118)
        self.marks_header_canvas.pack(fill='x', side='top')

        rows_outer = tk.Frame(self.marks_card, bg=mk_bi)
        rows_outer.pack(fill='both', expand=True, side='top')

        self.marks_scroll_canvas = tk.Canvas(rows_outer, bg=mk_bi, highlightthickness=0)
        marks_scroll_sb = ttk.Scrollbar(rows_outer, orient='vertical', command=self.marks_scroll_canvas.yview,
                                        style='App.Vertical.TScrollbar')
        self.marks_scroll_x = ttk.Scrollbar(self.marks_card, orient='horizontal', command=self._marks_xview,
                                            style='App.Vertical.TScrollbar')

        self.marks_scroll_canvas.configure(yscrollcommand=marks_scroll_sb.set, xscrollcommand=self._marks_xscroll_set)
        self.marks_header_canvas.configure(xscrollcommand=self._marks_xscroll_set)

        marks_scroll_sb.pack(side='right', fill='y')
        self.marks_scroll_canvas.pack(side='left', fill='both', expand=True)
        self.marks_scroll_x.pack(fill='x', side='bottom')

        self.marks_header_inner = tk.Frame(self.marks_header_canvas, bg=mk_bi)
        self.marks_inner = tk.Frame(self.marks_scroll_canvas, bg=mk_bi)

        self._marks_header_win = self.marks_header_canvas.create_window((0, 0), window=self.marks_header_inner, anchor='nw')
        self._marks_body_win = self.marks_scroll_canvas.create_window((0, 0), window=self.marks_inner, anchor='nw')

        def _update_header_scrollregion(_event=None):
            self.marks_header_canvas.configure(scrollregion=self.marks_header_canvas.bbox('all'))

        def _update_body_scrollregion(_event=None):
            self.marks_scroll_canvas.configure(scrollregion=self.marks_scroll_canvas.bbox('all'))

        def _on_header_canvas_resize(e):
            min_width = self.marks_header_inner.winfo_reqwidth()
            self.marks_header_canvas.itemconfig(self._marks_header_win, width=max(e.width, min_width))
            _update_header_scrollregion()

        def _on_body_canvas_resize(e):
            min_width = self.marks_inner.winfo_reqwidth()
            self.marks_scroll_canvas.itemconfig(self._marks_body_win, width=max(e.width, min_width))
            _update_body_scrollregion()

        self.marks_header_canvas.bind('<Configure>', _on_header_canvas_resize)
        self.marks_scroll_canvas.bind('<Configure>', _on_body_canvas_resize)
        self.marks_header_inner.bind('<Configure>', _update_header_scrollregion)
        self.marks_inner.bind('<Configure>', _update_body_scrollregion)

        self.marks_scroll_canvas.bind('<MouseWheel>', lambda e: self.marks_scroll_canvas.yview_scroll(int(-1 * (e.delta / 120)), 'units'))
        self.marks_scroll_canvas.bind('<Shift-MouseWheel>', lambda e: self._marks_xview('scroll', int(-1 * (e.delta / 120)), 'units'))
        
        self.marks_entries = {}  # sid -> {subject: Entry}
        self.student_widgets = []  # for double-click
        
        self._load_marks_table()

    def _is_summary_student_name(self, name):
        key = re.sub(r'[^a-z0-9]+', '', str(name or '').lower())
        return key in {
            'average', 'avg', 'mean', 'overallaverage', 'classaverage',
            'total', 'totals', 'grandtotal', 'position', 'psn', 'rank'
        }

    def _style_marks_entry(self, entry, row_index):
        base_bg = '#ffffff' if row_index % 2 == 0 else '#f6fbf6'
        entry.configure(
            bg=base_bg,
            fg='#163d19',
            insertbackground='#163d19',
            relief='flat',
            highlightthickness=1,
            highlightbackground='#cfe7d1',
            highlightcolor=SIDEBAR_ACTIVE,
            disabledbackground=base_bg,
            disabledforeground='#163d19',
            readonlybackground=base_bg,
            font=(FF, 10, 'bold'),
            selectbackground=SIDEBAR_ACTIVE,
            selectforeground='white',
        )
        return base_bg

    def _marks_xview(self, *args):
        """Scroll marks header and body horizontally in sync."""
        if hasattr(self, 'marks_header_canvas'):
            self.marks_header_canvas.xview(*args)
        if hasattr(self, 'marks_scroll_canvas'):
            self.marks_scroll_canvas.xview(*args)

    def _marks_xscroll_set(self, *args):
        """Mirror horizontal scroll state to shared scrollbar and header canvas."""
        if hasattr(self, 'marks_scroll_x'):
            self.marks_scroll_x.set(*args)
        if hasattr(self, 'marks_header_canvas'):
            self.marks_header_canvas.xview_moveto(args[0])

    def _load_marks_table(self):
        # clear existing
        if hasattr(self, 'marks_header_inner'):
            for w in self.marks_header_inner.winfo_children():
                w.destroy()
        for w in self.marks_inner.winfo_children():
            w.destroy()
        self.marks_entries.clear()
        self.student_widgets.clear()
        
        cls = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        exam_type = self.marks_exam_cb.get() or DEFAULT_EXAM_TYPE
        
        # Get students based on stream selection
        stream = getattr(self, '_selected_marks_stream', '')
        if stream:
            students = [s for s in db.get_students_by_class_and_stream(cls, stream) if not self._is_summary_student_name(s.get('name'))]
        else:
            students = [s for s in db.get_students_by_class(cls) if not self._is_summary_student_name(s.get('name'))]
        subjects = self._get_subjects_for_selected_class(cls, term, exam_type)
        student_col_width = 260
        subject_col_width = 96
        table_width = student_col_width + (len(subjects) * subject_col_width)
        
        mp_bg = getattr(self, 'marks_panel_bg', CARD_BG)
        if not students:
            if hasattr(self, 'marks_header_inner'):
                tk.Frame(self.marks_header_inner, bg=mp_bg, height=1).pack(fill='x')
            tk.Label(self.marks_inner, text='No students in this class', 
                     bg=mp_bg, fg=TEXT_SECONDARY, font=(FF, 12)).pack(pady=40)
            return

        header_wrap = tk.Frame(self.marks_header_inner, bg=mp_bg, width=table_width)
        header_wrap.pack(anchor='nw', pady=(0, 2))

        summary = tk.Frame(header_wrap, bg='#edf7ee', highlightthickness=1, highlightbackground='#cfe7d1')
        summary.pack(fill='x', pady=(0, 10))
        tk.Label(summary, text=f'{self._get_class_label(cls)}  •  Term {term}', bg='#edf7ee', fg=SIDEBAR_BG,
                 font=(FF, 11, 'bold')).pack(side='left', padx=14, pady=10)
        tk.Label(summary, text=f'{len(students)} learners', bg='#edf7ee', fg=GREEN,
                 font=(FF, 10, 'bold')).pack(side='left', padx=(0, 14))
        tk.Label(summary, text=f'{len(subjects)} subjects', bg='#edf7ee', fg=GREEN,
                 font=(FF, 10, 'bold')).pack(side='left')

        # header row
        hdr = tk.Frame(header_wrap, bg=SIDEBAR_BG, width=table_width, height=62)
        hdr.pack(anchor='nw', pady=(0, 8))
        hdr.pack_propagate(False)

        rows_wrap = tk.Frame(self.marks_inner, bg=mp_bg, width=table_width)
        rows_wrap.pack(anchor='nw', pady=(0, 2))

        student_hdr_cell = tk.Frame(hdr, width=student_col_width, bg=SIDEBAR_BG, height=62)
        student_hdr_cell.pack(side='left', fill='y')
        student_hdr_cell.pack_propagate(False)
        tk.Label(student_hdr_cell, text='Student', bg=SIDEBAR_BG, fg='white',
                 font=(FF, 11, 'bold'), anchor='w', padx=14, pady=12).pack(fill='both', expand=True)

        for sub in subjects:
            subject_style = self._get_subject_colors(sub, cls)
            sub_hdr_cell = tk.Frame(hdr, width=subject_col_width, bg=subject_style['base'], height=62)
            sub_hdr_cell.pack(side='left', fill='y')
            sub_hdr_cell.pack_propagate(False)
            tk.Label(
                sub_hdr_cell,
                text=self._get_subject_label(sub, cls, multiline=True),
                bg=subject_style['base'],
                fg=subject_style['text'],
                font=(FF, 9, 'bold'),
                anchor='center',
                justify='center',
                wraplength=subject_col_width - 10,
                pady=8
            ).pack(fill='both', expand=True)
        
        # student rows
        for row_index, s in enumerate(students):
            sid = s['id']
            row_bg = '#ffffff' if row_index % 2 == 0 else '#f6fbf6'
            row_frame = tk.Frame(
                rows_wrap,
                bg=row_bg,
                width=table_width,
                height=44,
                highlightthickness=1,
                highlightbackground='#e3efe4'
            )
            row_frame.pack(anchor='nw', pady=2)
            row_frame.pack_propagate(False)
            
            # student name (clickable)
            name_cell = tk.Frame(row_frame, width=student_col_width, bg=row_bg, height=44)
            name_cell.pack(side='left', fill='y')
            name_cell.pack_propagate(False)
            name_btn = tk.Label(name_cell, text=s['name'], bg=row_bg, fg=TEXT_PRIMARY,
                               font=(FF, 10, 'bold'), anchor='w', cursor='hand2',
                               padx=14, pady=10)
            name_btn.pack(fill='both', expand=True)
            name_btn.bind('<Button-1>', lambda e, sid=sid: self._edit_student_marks(sid))
            self.student_widgets.append(name_btn)
            
            # mark entries
            self.marks_entries[sid] = {}
            m = db.get_student_marks(sid, term, exam_type)
            for sub in subjects:
                e_frame = tk.Frame(row_frame, bg=row_bg, width=subject_col_width, height=44)
                e_frame.pack(side='left', fill='y')
                e_frame.pack_propagate(False)
                
                e = tk.Entry(e_frame, width=6, justify='center', bd=1, bg='white', fg='black')
                e.pack(fill='x', padx=10, pady=8)
                val = m.get(sub, '')
                e.insert(0, '' if val in (None, '') else str(val))
                e.bind('<KeyRelease>', lambda ev, s=sub, sid=sid: self._validate_mark(ev, sid, s))
                
                self.marks_entries[sid][sub] = e

        # Refresh scroll regions and keep header/body horizontally synchronized.
        self.marks_header_inner.update_idletasks()
        self.marks_inner.update_idletasks()
        self.marks_header_canvas.configure(scrollregion=self.marks_header_canvas.bbox('all'))
        self.marks_scroll_canvas.configure(scrollregion=self.marks_scroll_canvas.bbox('all'))
        self._marks_xview('moveto', '0')

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
        exam_type = self.marks_exam_cb.get() or DEFAULT_EXAM_TYPE
        cur = db.get_student_marks(sid, term, exam_type)
        subjects = self._get_subjects_for_selected_class(self.marks_class_cb.get(), term, exam_type)

        dlg = tk.Toplevel(self.root)
        dlg.title(f'Marks – {name}')
        dlg.geometry('460x420')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        em_bo, em_bi = _card_colors('mint')
        outer = tk.Frame(dlg, bg=em_bo)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=em_bi, padx=30, pady=22)
        card.pack(padx=1, pady=1)

        tk.Label(card, text=f'Marks for {name}  (Term {term} - {exam_type})',
                 bg=em_bi, fg=TEXT_PRIMARY, font=(FF, 13, 'bold')).pack(anchor='w', pady=(0, 16))

        entries: dict = {}
        grid = tk.Frame(card, bg=em_bi)
        grid.pack(fill='x')
        for i, sub in enumerate(subjects):
            r, c = divmod(i, 3)
            subject_style = self._get_subject_colors(sub, self.marks_class_cb.get())
            tk.Label(grid, text=sub, bg=subject_style['soft'], fg=subject_style['dark_text'],
                     font=(FF, 10, 'bold'), anchor='w', width=9).grid(
                         row=r*2, column=c, padx=6, pady=(6, 2), sticky='w')
            e = tk.Entry(grid, width=8, justify='center', bd=1, bg='white', fg='black')
            e.insert(0, '' if cur.get(sub, '') in (None, '') else str(cur.get(sub, '')))
            e.grid(row=r*2+1, column=c, padx=6, pady=(0, 6), ipady=2, sticky='ew')
            entries[sub] = e
        for c in range(3): grid.columnconfigure(c, weight=1)

        btn_row = tk.Frame(card, bg=em_bi)
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
            db.save_student_marks(sid, marks, term, exam_type)
            self._load_marks_table()
            dlg.destroy()

        sv = tk.Label(btn_row, text='Save Marks', bg=BLUE, fg='white',
                      font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        sv.pack(side='left')
        sv.bind('<Button-1>', lambda e: do_save())

    def save_marks(self):
        cls = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        exam_type = self.marks_exam_cb.get() or DEFAULT_EXAM_TYPE
        
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
                db.save_student_marks(sid, marks, term, exam_type)
        
        messagebox.showinfo('Success', f'Marks saved successfully for {term} - {exam_type}! ({saved_count} values updated)')

    # ── Marks import / template helpers ──────────────────────────────────────

    def download_marks_template(self):
        """Export a clean, colorful marks template (Student + subjects)."""
        cls  = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        exam_type = self.marks_exam_cb.get() or DEFAULT_EXAM_TYPE
        subjects = self._get_subjects_for_selected_class(cls, term, exam_type)
        students = db.get_students_by_class(cls)

        file_path = filedialog.asksaveasfilename(
            title='Save Marks Template',
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f"marks_template_{cls.replace(' ', '_')}_T{term}_{exam_type.replace('-', '_')}.xlsx"
        )
        if not file_path:
            return

        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.utils import get_column_letter

            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = cls.upper().replace(' ', '_')
            ws.sheet_view.showGridLines = False

            # Header theme inspired by your Enter Marks table
            base_fill = PatternFill('solid', fgColor='6F7C4A')  # olive for student column
            data_fill = PatternFill('solid', fgColor='F8FAFC')
            zebra_fill = PatternFill('solid', fgColor='EEF3EF')
            meta_fill = PatternFill('solid', fgColor='EEF3EF')
            hdr_font = Font(bold=True, color='FFFFFF', name='Calibri', size=11)
            data_font = Font(name='Calibri', size=10)
            ctr = Alignment(horizontal='center', vertical='center')
            left = Alignment(horizontal='left', vertical='center')
            thin = Side(style='thin', color='000000')
            border = Border(left=thin, right=thin, top=thin, bottom=thin)

            # Meta row (for human readability only)
            last_col = 2 + len(subjects)
            ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=max(2, last_col))
            meta = ws.cell(row=1, column=1, value=f'{cls} | Term {term} | {exam_type} | Fill marks 0-100 (blanks are allowed)')
            meta.fill = meta_fill
            meta.font = Font(bold=True, color='2F3B1A', name='Calibri', size=10)
            meta.alignment = left
            meta.border = border

            # Row 2 headers: optional admission_no + student + subjects
            ws.cell(row=2, column=1, value='admission_no').fill = base_fill
            ws.cell(row=2, column=1).font = hdr_font
            ws.cell(row=2, column=1).alignment = ctr
            ws.cell(row=2, column=1).border = border

            ws.cell(row=2, column=2, value='Student').fill = base_fill
            ws.cell(row=2, column=2).font = hdr_font
            ws.cell(row=2, column=2).alignment = left
            ws.cell(row=2, column=2).border = border

            subject_palette = [
                'F57C00',  # orange
                'E65100',  # deep orange
                'D81B60',  # pink
                '6D4C41',  # brown
                'FB8C00',  # amber
                '8E24AA',  # purple
                '3949AB',  # indigo
                'FF8F00',  # amber dark
                '9E9D24',  # olive
                '2E7D32',  # green
                '1E88E5',  # blue
                'C62828'   # red
            ]

            for i, subject in enumerate(subjects, start=3):
                fill = PatternFill('solid', fgColor=subject_palette[(i - 3) % len(subject_palette)])
                cell = ws.cell(row=2, column=i, value=str(subject).strip().upper())
                cell.fill = fill
                cell.font = hdr_font
                cell.alignment = ctr
                cell.border = border

            # Data rows
            start_row = 3
            for row_offset, student in enumerate(students):
                ri = start_row + row_offset
                existing = db.get_student_marks(student['id'], term, exam_type)
                row_fill = zebra_fill if row_offset % 2 else data_fill

                adm = (student.get('admission_no') or '').strip()
                ws.cell(row=ri, column=1, value=adm)
                ws.cell(row=ri, column=2, value=student.get('name', ''))

                for col_idx in range(1, last_col + 1):
                    cell = ws.cell(row=ri, column=col_idx)
                    cell.fill = row_fill
                    cell.font = data_font
                    cell.border = border
                    cell.alignment = left if col_idx == 2 else ctr

                for col_idx, subject in enumerate(subjects, start=3):
                    mark_val = existing.get(subject, '')
                    ws.cell(row=ri, column=col_idx, value=mark_val if mark_val != '' else '')
                    ws.cell(row=ri, column=col_idx).alignment = ctr

            ws.column_dimensions['A'].width = 18
            ws.column_dimensions['B'].width = 34
            for col_idx in range(3, last_col + 1):
                ws.column_dimensions[get_column_letter(col_idx)].width = 11
            ws.freeze_panes = 'C3'

            wb.save(file_path)
            messagebox.showinfo(
                'Template Saved',
                f'Template saved to:\n{file_path}\n\n'
                'Format: admission_no (optional), Student, then subject columns.\n'
                'You can leave any mark cell blank and import will skip blanks.'
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
        exam_type = self.marks_exam_cb.get() or DEFAULT_EXAM_TYPE
        subjects = self._get_subjects_for_selected_class(cls, term, exam_type)

        file_path = filedialog.askopenfilename(
            title='Select Marks File',
            filetypes=[('Excel files', '*.xlsx *.xls')]
        )
        if not file_path:
            return

        try:
            import openpyxl
            wb = openpyxl.load_workbook(file_path, data_only=True)
            progress_dialog, status_label, percent_label, progress = self._open_progress_dialog(
                'Importing Marks', 'Scanning workbook sheets...'
            )

            parsed_sheets = []
            total_sheets = len(wb.worksheets)
            for sheet_index, ws in enumerate(wb.worksheets, start=1):
                self._update_progress_dialog(
                    progress_dialog, status_label, percent_label, progress,
                    sheet_index - 1, total_sheets,
                    f'Scanning sheet {sheet_index} of {total_sheets}: {ws.title}'
                )
                parsed = self._parse_assessment_sheet(ws)
                if parsed:
                    parsed_sheets.append(parsed)

            if parsed_sheets:
                total_updated = 0
                total_created = 0
                affected_classes = []
                total_students = sum(len(parsed['students']) for parsed in parsed_sheets)
                processed_students = 0

                for parsed in parsed_sheets:
                    class_name = parsed['class_name']
                    affected_classes.append(class_name)
                    existing_students = db.get_students_by_class(class_name)
                    name_to_student = {
                        self._normalize_key(student['name']): student
                        for student in existing_students
                    }

                    for item in parsed['students']:
                        processed_students += 1
                        self._update_progress_dialog(
                            progress_dialog, status_label, percent_label, progress,
                            processed_students, total_students,
                            f'Importing {item["name"].strip()} into {class_name} ({processed_students}/{total_students})'
                        )
                        name_key = self._normalize_key(item['name'])
                        student = name_to_student.get(name_key)
                        if not student:
                            admission_no = self._generate_admission_no(class_name, item['name'])
                            student = db.add_student(item['name'].strip(), class_name, 'Male', admission_no, '')
                            name_to_student[name_key] = student
                            total_created += 1

                        current_marks = db.get_student_marks(student['id'], term, exam_type)
                        current_marks.update(item['marks'])
                        db.save_student_marks(student['id'], current_marks, term, exam_type)
                        total_updated += 1

                self._update_progress_dialog(
                    progress_dialog, status_label, percent_label, progress,
                    total_students, total_students,
                    'Refreshing marks grid...'
                )
                self._load_marks_table()
                progress_dialog.destroy()
                affected_classes = sorted(set(affected_classes))
                messagebox.showinfo(
                    'Import Complete',
                    f'Assessment workbook imported successfully.\n\n'
                    f'Classes: {", ".join(affected_classes)}\n'
                    f'Student records updated: {total_updated}\n'
                    f'New students created with auto admission numbers: {total_created}'
                )
                return

            df = pd.read_excel(file_path)
            total_rows = len(df.index)
            self._update_progress_dialog(
                progress_dialog, status_label, percent_label, progress,
                0, max(1, total_rows),
                'Workbook did not match assessment layout. Using flat table import...'
            )

            # ── Normalise all column names to lowercase-stripped for matching ──
            orig_cols = list(df.columns)
            df.columns = [str(c).strip().lower() for c in df.columns]

            # Accepted aliases for the two identifier fields
            ADM_ALIASES  = {'admission_no', 'adm_no', 'admission no', 'adm no',
                            'admission', 'adm', 'reg_no', 'reg no', 'regno'}
            NAME_ALIASES = {'name', 'student_name', 'student name', 'learner',
                            'learner name', 'full_name', 'fullname', 'pupil', 'student'}

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
                messagebox.showwarning(
                    'No Subject Columns Found',
                    f'No matching subject columns were found in this file.\n\n'
                    f'Expected subject columns (any capitalisation):\n'
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

            for index, (_, row) in enumerate(df.iterrows(), start=1):
                self._update_progress_dialog(
                    progress_dialog, status_label, percent_label, progress,
                    index - 1, total_rows,
                    f'Importing marks row {index} of {total_rows}...'
                )
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
                    db.save_student_marks(sid, marks, term, exam_type)
                    updated += 1
                else:
                    skipped += 1

            # Refresh on-screen grid
            self._update_progress_dialog(
                progress_dialog, status_label, percent_label, progress,
                total_rows, total_rows,
                'Refreshing marks grid...'
            )
            self._load_marks_table()
            progress_dialog.destroy()

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
            try:
                progress_dialog.destroy()
            except Exception:
                pass
            messagebox.showerror('Import Error', f'Failed to import marks:\n{exc}')

    # ==================== RESULTS ====================
    def show_reports(self):
        self.clear_frame()
        self._set_nav('Results')
        self._page_header('Results', 'View student performance and rankings')

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

        lbl('Exam:')
        self.rep_exam_cb = ttk.Combobox(ctrl, values=EXAM_TYPES, state='readonly',
                                        style='App.TCombobox', width=12)
        self.rep_exam_cb.set(DEFAULT_EXAM_TYPE)
        self.rep_exam_cb.pack(side='left', ipady=4)

        self.rep_cls_cb.bind('<<ComboboxSelected>>',  lambda e: self.load_reports())
        self.rep_term_cb.bind('<<ComboboxSelected>>', lambda e: self.load_reports())
        self.rep_exam_cb.bind('<<ComboboxSelected>>', lambda e: self.load_reports())

        self._toolbar_btn(ctrl, '\u2193  Export CSV', self.export_csv,
                          bg='#475569').pack(side='left', padx=16)
        self._toolbar_btn(ctrl, '\U0001F4CA  Spotlight Excel', self.export_spotlight_excel,
                          bg='#1B5E20').pack(side='left', padx=4)
        self.rep_mode_badge = tk.Label(
            ctrl, text='Mode: -', bg='#1d4ed8', fg='white',
            font=(FF, 9, 'bold'), padx=10, pady=5
        )
        self.rep_mode_badge.pack(side='right', padx=(10, 8))

        # subject averages strip
        sa_bo, sa_bi = _card_colors('lilac')
        subj_outer = tk.Frame(self.content_frame, bg=sa_bo)
        subj_outer.pack(fill='x', pady=(0, 10))
        subj_card = tk.Frame(subj_outer, bg=sa_bi, padx=16, pady=14)
        subj_card.pack(fill='both', expand=True, padx=1, pady=1)
        tk.Label(subj_card, text='Subject Averages', bg=sa_bi, fg=TEXT_PRIMARY,
                 font=(FF, 11, 'bold')).pack(anchor='w', pady=(0, 8))
        self.subj_row = tk.Frame(subj_card, bg=sa_bi)
        self.subj_row.pack(fill='x')

        # rankings table
        rt_bo, rt_bi = _card_colors('mint')
        tc_outer = tk.Frame(self.content_frame, bg=rt_bo)
        tc_outer.pack(fill='both', expand=True, pady=4)
        self.report_table_card = tk.Frame(tc_outer, bg=rt_bi)
        self.report_table_card.pack(fill='both', expand=True, padx=1, pady=1)
        self.report_table_bg = rt_bi
        self.report_table = None
        self.report_subject_columns = None

        self.load_reports()

    def load_reports(self):
        for w in self.subj_row.winfo_children():
            w.destroy()

        cls  = self.rep_cls_cb.get()
        term = self.rep_term_cb.get()
        exam_type = self.rep_exam_cb.get() or DEFAULT_EXAM_TYPE
        if hasattr(self, 'rep_mode_badge'):
            badge_text, badge_bg, badge_fg = self._get_subject_mode_badge(cls, term, exam_type)
            self.rep_mode_badge.config(text=badge_text, bg=badge_bg, fg=badge_fg)
        results = self._get_ranked_results(cls, term, exam_type)
        subjects = self._get_subjects_for_scope(cls, term, exam_type, results)

        if self.report_subject_columns != subjects:
            self.report_subject_columns = list(subjects)
            for widget in self.report_table_card.winfo_children():
                widget.destroy()

            columns = [
                {'key': 'pos', 'title': 'Pos', 'width': 60, 'anchor': 'center'},
                {'key': 'adm', 'title': 'Adm No', 'width': 105, 'anchor': 'center'},
                {'key': 'name', 'title': 'Student Name', 'width': 220, 'anchor': 'w'},
                {'key': 'class', 'title': 'Grade', 'width': 95, 'anchor': 'center'},
                {'key': 'class_level', 'title': 'CBC Level', 'width': 160, 'anchor': 'center'},
            ]
            for s in subjects:
                columns.append({'key': s, 'title': self._get_subject_label(s, cls), 'width': 88, 'anchor': 'center'})
            columns.extend([
                {'key': 'total', 'title': 'Total', 'width': 80, 'anchor': 'center'},
                {'key': 'avg', 'title': 'Average', 'width': 85, 'anchor': 'center'},
                {'key': 'grade', 'title': 'Grade', 'width': 70, 'anchor': 'center'},
                {'key': 'level', 'title': 'Level', 'width': 180, 'anchor': 'w'},
            ])

            self.report_table = AdvancedDataTable(
                self.report_table_card,
                columns=columns,
                page_size=20,
                search_label='Search results'
            )
            self.rep_tree = self.report_table.tree

        # subject averages
        subj_totals = {s: [] for s in subjects}
        for r in results:
            for s in subjects:
                if r['marks'].get(s) is not None:
                    subj_totals[s].append(r['marks'][s])

        for s in subjects:
            vals = subj_totals[s]
            avg  = round(sum(vals) / len(vals), 1) if vals else 0
            grade = self._get_grade_code_for_class(avg, cls if cls != 'All' else '')
            subject_style = self._get_subject_colors(s, cls)
            tile  = tk.Frame(self.subj_row, bg=subject_style['base'], padx=10, pady=8)
            tile.pack(side='left', padx=3, expand=True, fill='both')
            tk.Label(tile, text=self._get_subject_label(s, cls), bg=subject_style['base'], fg=subject_style['text'], font=(FF, 9, 'bold')).pack()
            tk.Label(tile, text=str(avg), bg=subject_style['base'], fg=subject_style['text'], font=(FF, 13, 'bold')).pack()
            tk.Label(tile, text=grade, bg=subject_style['base'], fg=subject_style['text'], font=(FF, 8)).pack()

        rows = []
        for r in results:
            vals = [r['position'], r['student']['admission_no'], r['student']['name'], self._get_class_label(r['student']['class']), r['class_level']]
            value_map = {
                'pos': r['position'],
                'adm': r['student']['admission_no'],
                'name': r['student']['name'],
                'class': self._get_class_label(r['student']['class']),
                'class_level': r['class_level'],
            }
            for s in subjects:
                mark_val = r['marks'].get(s, '-')
                vals.append(mark_val)
                value_map[s] = mark_val
            vals += [r['total'], r['average'], r['grade'], r['level']]
            value_map.update({'total': r['total'], 'avg': r['average'], 'grade': r['grade'], 'level': r['level']})
            rows.append({
                'iid': f"result_{r['student']['id']}",
                'values': tuple(vals),
                'value_map': value_map,
                'search': ' '.join(str(v) for v in vals),
            })
        if self.report_table:
            self.report_table.set_rows(rows)
        return

        # rows
        for r in results:
            vals = [r['position'], r['student']['name'], r['student']['class']]
            for s in self.get_current_subjects(): vals.append(r['marks'].get(s, '—'))
            vals += [r['total'], r['average'], r['grade']]
            self.rep_tree.insert('', 'end', values=vals)

    def export_csv(self):
        cls  = self.rep_cls_cb.get()
        term = self.rep_term_cb.get()
        exam_type = self.rep_exam_cb.get() or DEFAULT_EXAM_TYPE
        results = self._get_ranked_results(cls, term, exam_type)
        subjects = self._get_subjects_for_scope(cls, term, exam_type, results)
        subject_headers = [self._get_subject_label(s, cls if cls != 'All' else '') for s in subjects]
        fn = f"report_{cls.replace(' ', '_')}_term_{term}_{exam_type.replace('-', '_')}.csv"
        with open(fn, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['Position', 'Adm No', 'Name', 'Class', 'CBC Level'] + subject_headers + ['Total', 'Average', 'Grade', 'Level'])
            for r in results:
                row = [r['position'], r['student']['admission_no'],
                       r['student']['name'], r['student']['class'], r['class_level']]
                row += [r['marks'].get(s, '') for s in subjects]
                row += [r['total'], r['average'], r['grade'], r['level']]
                w.writerow(row)
        messagebox.showinfo('Exported', f'Report saved to {fn}')
        return

    # ==================== WESTERN SPOTLIGHT EXPORT ====================
    def export_spotlight_excel(self):
        """Open a dialog to configure and export the Western Spotlight class report."""
        # Pre-populate from Result page if available
        try:
            initial_cls  = self.rep_cls_cb.get()
            if initial_cls == 'All':
                initial_cls = CLASSES[0]
            initial_term = self.rep_term_cb.get()
            initial_exam = self.rep_exam_cb.get()
        except AttributeError:
            initial_cls  = CLASSES[0]
            initial_term = TERMS[0]
            initial_exam = DEFAULT_EXAM_TYPE

        dlg = tk.Toplevel(self.root)
        dlg.title('Export – Western Spotlight')
        dlg.geometry('400x340')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        ws_bo, ws_bi = _card_colors('sand')
        outer = tk.Frame(dlg, bg=ws_bo)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=ws_bi, padx=30, pady=24)
        card.pack(padx=1, pady=1)
        card.columnconfigure(1, weight=1)

        tk.Label(card, text='Western Spotlight Export', bg=ws_bi, fg=TEXT_PRIMARY,
                 font=(FF, 13, 'bold')).grid(row=0, column=0, columnspan=2,
                                              sticky='w', pady=(0, 18))

        def row_field(r, label, widget):
            tk.Label(card, text=label, bg=ws_bi, fg=TEXT_SECONDARY,
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

        exam_cb = ttk.Combobox(card, values=EXAM_TYPES, state='readonly',
                                style='App.TCombobox', width=20)
        exam_cb.set(initial_exam)
        row_field(3, 'Exam Type:', exam_cb)

        stream_var = tk.StringVar(value='GREEN')
        stream_e = ttk.Entry(card, textvariable=stream_var,
                              style='App.TEntry', width=20)
        row_field(4, 'Stream / Group:', stream_e)

        assess_cb = ttk.Combobox(card, values=['MID-TERM', 'END-TERM'],
                                  state='readonly', style='App.TCombobox', width=20)
        assess_cb.set('MID-TERM')
        row_field(5, 'Assessment:', assess_cb)

        year_var = tk.StringVar(value=str(datetime.now().year))
        year_e = ttk.Entry(card, textvariable=year_var,
                            style='App.TEntry', width=20)
        row_field(6, 'Year:', year_e)

        btn_row = tk.Frame(card, bg=ws_bi)
        btn_row.grid(row=7, column=0, columnspan=2, pady=(20, 0))

        cancel = tk.Label(btn_row, text='Cancel', bg='#e8f5e9', fg=TEXT_PRIMARY,
                          font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        cancel.pack(side='left', padx=(0, 8))
        cancel.bind('<Button-1>', lambda e: dlg.destroy())

        def do_export():
            selected_cls    = cls_cb.get()
            selected_term   = term_cb.get()
            selected_exam   = exam_cb.get()
            selected_stream = stream_var.get().strip() or 'GREEN'
            selected_assess = assess_cb.get()
            selected_year   = year_var.get().strip() or str(datetime.now().year)
            dlg.destroy()
            self._do_spotlight_export(selected_cls, selected_term, selected_exam,
                                      selected_stream, selected_assess, selected_year)

        export_btn = tk.Label(btn_row, text='Export Excel', bg='#1B5E20', fg='white',
                              font=(FF, 10, 'bold'), padx=18, pady=8, cursor='hand2')
        export_btn.pack(side='left')
        export_btn.bind('<Button-1>', lambda e: do_export())

    def _do_spotlight_export(self, cls, term, exam_type, stream, assess, year):
        """Generate the Western Spotlight Excel workbook and save it."""
        results = self._get_ranked_results(cls, term, exam_type)
        if not results:
            messagebox.showwarning('No Data',
                                   f'No results found for {cls}, Term {term}.\n'
                                   'Please enter marks first.')
            return

        file_path = filedialog.asksaveasfilename(
            title='Save – Western Spotlight',
            defaultextension='.xlsx',
            filetypes=[('Excel files', '*.xlsx')],
            initialfile=f"spotlight_{cls.replace(' ', '_')}_T{term}_{exam_type.replace('-', '_')}_{year}.xlsx"
        )
        if not file_path:
            return

        try:
            import openpyxl
            from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
            from openpyxl.drawing.image import Image as XLImage
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

            subjects   = self._get_subjects_for_scope(cls, term, exam_type, results)
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
            ROW_OFFSET = 0

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

            letterhead_path = ensure_letterhead_assets()
            if letterhead_path and os.path.exists(letterhead_path):
                try:
                    xl_img = XLImage(letterhead_path)
                    max_width = 980
                    if xl_img.width > max_width:
                        ratio = max_width / float(xl_img.width)
                        xl_img.width = int(xl_img.width * ratio)
                        xl_img.height = int(xl_img.height * ratio)
                    ws.add_image(xl_img, 'A1')
                    ROW_OFFSET = max(6, int(max(120, xl_img.height) / 20) + 1)
                    for row_idx in range(1, ROW_OFFSET + 1):
                        ws.row_dimensions[row_idx].height = 18
                except Exception as img_exc:
                    print(f'Excel letterhead load error: {img_exc}')

            # ── Title section (rows 1-5) ─────────────────────────────────────
            grade_num   = cls.replace('Grade ', '')
            grade_words = {'1':'ONE','2':'TWO','3':'THREE','4':'FOUR','5':'FIVE',
                           '6':'SIX','7':'SEVEN','8':'EIGHT','9':'NINE'}.get(grade_num, grade_num)

            title_row = ROW_OFFSET + 1
            blank_row_1 = title_row + 1
            subtitle_row = title_row + 2
            blank_row_2 = title_row + 3
            banner_row = title_row + 4
            header_row_1 = title_row + 5
            header_row_2 = title_row + 6
            first_data_row = title_row + 7

            fill_entire_row(title_row, BG_TITLE, FNT_YLW, height=26)
            merged_cell(title_row, C_NO, title_row, C_LAST,
                        'MT.  OLIVES ADVENTIST SCHOOL,  NGONG',
                        BG_TITLE, FNT_YLW, ALIGN_CTR)

            fill_entire_row(blank_row_1, BG_TITLE, FNT_YLW, height=8)
            merged_cell(blank_row_1, C_NO, blank_row_1, C_LAST, '', BG_TITLE, FNT_YLW)

            title3 = (f'GRADE {grade_words} ({grade_num}) {stream} '
                      f'TERM {term.upper()} {assess} ASSESSMENT REPORT {year}')
            fill_entire_row(subtitle_row, BG_TITLE, FNT_YLW, height=22)
            merged_cell(subtitle_row, C_NO, subtitle_row, C_LAST, title3, BG_TITLE, FNT_YLW, ALIGN_CTR)

            fill_entire_row(blank_row_2, BG_TITLE, FNT_YLW, height=8)
            merged_cell(blank_row_2, C_NO, blank_row_2, C_LAST, '', BG_TITLE, FNT_YLW)

            fill_entire_row(banner_row, BG_TITLE, FNT_YLW, height=22)
            merged_cell(banner_row, C_NO, banner_row, C_LAST,
                        'THE WESTERN SPOTLIGHT', BG_TITLE, FNT_YLW, ALIGN_CTR)

            # ── Column header rows (6 & 7) ───────────────────────────────────
            ws.row_dimensions[header_row_1].height = 20
            ws.row_dimensions[header_row_2].height = 18
            for c in range(1, C_LAST + 1):
                for r in (header_row_1, header_row_2):
                    apply_style(r, c, fill=BG_HDR, font=FNT_WHT,
                                align=ALIGN_CTR, border=BORDER)

            # NO.  – merged rows 6-7
            merged_cell(header_row_1, C_NO, header_row_2, C_NO, 'NO.', BG_HDR, FNT_WHT, ALIGN_CTR)
            for r in (header_row_1, header_row_2):
                ws.cell(r, C_NO).border = BORDER

            # LEARNER – merged rows 6-7
            merged_cell(header_row_1, C_NAME, header_row_2, C_NAME, 'LEARNER', BG_HDR, FNT_WHT, ALIGN_CTR)
            for r in (header_row_1, header_row_2):
                ws.cell(r, C_NAME).border = BORDER

            # Subject names row 6 (merged score+grade cols), row 7 sub-labels
            for i, subj in enumerate(subjects):
                sc  = C_S0 + i * 2      # score col
                gc  = sc + 1             # grade col
                lbl = self._get_subject_label(subj, cls).replace('\n', ' / ').upper()
                merged_cell(header_row_1, sc, header_row_1, gc, lbl, BG_HDR, FNT_WHT, ALIGN_CTR)
                for c in (sc, gc):
                    ws.cell(header_row_1, c).border = BORDER
                # Row 7: '100' under score, blank under grade
                c7 = ws.cell(header_row_2, sc)
                c7.value = '100'; c7.fill = BG_HDR; c7.font = FNT_WHT_S
                c7.alignment = ALIGN_CTR; c7.border = BORDER
                ws.cell(header_row_2, gc).border = BORDER

            # TOTAL
            ws.cell(header_row_1, C_TOTAL).value = 'TOTAL'
            ws.cell(header_row_2, C_TOTAL).value = str(total_max)
            for r in (header_row_1, header_row_2):
                apply_style(r, C_TOTAL, fill=BG_HDR, font=FNT_WHT,
                            align=ALIGN_CTR, border=BORDER)

            # AVERAGE
            ws.cell(header_row_1, C_AVG).value = 'AVERAGE'
            ws.cell(header_row_2, C_AVG).value = '100%'
            for r in (header_row_1, header_row_2):
                apply_style(r, C_AVG, fill=BG_HDR, font=FNT_WHT,
                            align=ALIGN_CTR, border=BORDER)

            # PSN – merged rows 6-7
            merged_cell(header_row_1, C_PSN, header_row_2, C_PSN, 'PSN', BG_HDR, FNT_WHT, ALIGN_CTR)
            for r in (header_row_1, header_row_2):
                ws.cell(r, C_PSN).border = BORDER

            # ── Data rows ────────────────────────────────────────────────────
            for idx, result in enumerate(results):
                r  = first_data_row + idx
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
            ws.freeze_panes = f'A{first_data_row}'
            header_lines = get_letterhead_print_lines()
            left_header = header_lines[0] if header_lines else 'MT OLIVES ADVENTIST SCHOOL'
            center_header = header_lines[1] if len(header_lines) > 1 else ''
            right_header = header_lines[2] if len(header_lines) > 2 else ''
            footer_line = header_lines[3] if len(header_lines) > 3 else 'In God We Excel'
            ws.oddHeader.left.text = left_header
            ws.oddHeader.center.text = center_header
            ws.oddHeader.right.text = right_header
            ws.oddFooter.center.text = f'{footer_line}    Page &[Page] of &[Pages]'
            ws.oddFooter.right.text = '&[Date] &[Time]'
            ws.page_margins.top = 0.6
            ws.page_margins.header = 0.3
            ws.page_margins.footer = 0.3

            wb.save(file_path)
            messagebox.showinfo('Export Complete',
                                f'Western Spotlight report saved:\n{file_path}')

        except ImportError:
            messagebox.showerror('Missing Library',
                                 'openpyxl is required. Run:\n  pip install openpyxl')
        except Exception as exc:
            messagebox.showerror('Export Error', f'Failed to generate file:\n{exc}')

    def generate_pdf_report(self, student_id=None):
        """Generate a PDF report card for a student."""
        if not student_id:
            sel = self.students_tree.selection()
            if not sel:
                messagebox.showwarning("Warning", "Please select a student first!")
                return
            item = self.students_tree.item(sel[0])
            student_id = item["tags"][0]

        cls = "Grade 7"
        term = "One"
        exam_type = DEFAULT_EXAM_TYPE
        if hasattr(self, 'rc_exam_cb'):
            exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        elif hasattr(self, 'marks_exam_cb'):
            exam_type = self.marks_exam_cb.get() or DEFAULT_EXAM_TYPE

        results = self._get_ranked_results(cls, term, exam_type)
        result = next((r for r in results if r['student']['id'] == student_id), None)

        if not result:
            for c in self.get_current_classes():
                if c == cls:
                    continue
                results = self._get_ranked_results(c, term, exam_type)
                result = next((r for r in results if r['student']['id'] == student_id), None)
                if result:
                    cls = c
                    break

        if not result:
            messagebox.showerror("Error", "Could not find results for this student in the current term.")
            return

        s = result['student']
        file_path = filedialog.asksaveasfilename(
            title="Save PDF Report",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=f"Report_{s['name'].replace(' ', '_')}_{cls.replace(' ', '_')}_{term}_{exam_type.replace('-', '_')}.pdf"
        )
        if not file_path:
            return

        self._build_report_card_pdf(result, len(results), term, exam_type, file_path)

    def _build_report_card_pdf(self, result, total_students, term, exam_type, file_path):
        """Build the styled PDF report card used by preview/print flows."""
        try:
            s = result['student']
            styles = getSampleStyleSheet()
            
            # Custom styles
            styles.add(ParagraphStyle(name='SchoolName', parent=styles['Heading1'], fontName='Helvetica-Bold', fontSize=20, textColor=colors.HexColor("#1b5e20"), alignment=1))
            styles.add(ParagraphStyle(name='SchoolInfo', parent=styles['Normal'], fontSize=9, textColor=colors.gray, alignment=1))
            styles.add(ParagraphStyle(name='ReportTitle', parent=styles['Heading2'], fontName='Helvetica-Bold', fontSize=16, textColor=colors.HexColor("#4CAF50"), alignment=1, spaceBefore=10, spaceAfter=10))
            styles.add(ParagraphStyle(name='FooterLine', parent=styles['Normal'], fontSize=8, textColor=colors.HexColor("#666666"), alignment=1))

            letterhead_assets = get_letterhead_assets()
            header_path = letterhead_assets['header_path']
            footer_path = letterhead_assets['footer_path']
            header_lines = letterhead_assets['header_lines']
            footer_lines = letterhead_assets['footer_lines']

            def _image_height(image_path, target_width, min_height, max_height):
                if not image_path or not os.path.exists(image_path):
                    return min_height
                try:
                    img = get_processed_letterhead_image(image_path, 'header' if image_path == header_path else 'footer')
                    if img is None:
                        return min_height
                    scaled = int(img.height * target_width / img.width)
                    return max(min_height, min(max_height, scaled))
                except Exception:
                    return min_height

            page_inner_width = A4[0] - 40
            header_height = _image_height(header_path, page_inner_width, 80, 120)
            footer_height = _image_height(footer_path, page_inner_width, 45, 90)
            top_margin = header_height + 28
            bottom_margin = footer_height + 28

            doc = SimpleDocTemplate(
                file_path,
                pagesize=A4,
                rightMargin=20,
                leftMargin=20,
                topMargin=top_margin,
                bottomMargin=bottom_margin,
            )
            
            elements = []

            if not header_path:
                elements.append(Paragraph("MT OLIVES ADVENTIST SCHOOL", styles['SchoolName']))
            divider = Table([['']], colWidths=[520], rowHeights=[2])
            divider.setStyle(TableStyle([
                ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#1b5e20")),
                ('LEFTPADDING', (0,0), (-1,-1), 0),
                ('RIGHTPADDING', (0,0), (-1,-1), 0),
                ('TOPPADDING', (0,0), (-1,-1), 0),
                ('BOTTOMPADDING', (0,0), (-1,-1), 0),
            ]))
            elements.append(divider)
            elements.append(Spacer(1, 12))
            border_table = Table([[Paragraph("LEARNER ASSESSMENT REPORT CARD", styles['ReportTitle'])]], colWidths=[260])
            border_table.setStyle(TableStyle([
                ('BOX', (0,0), (-1,-1), 1, colors.HexColor("#d32f2f")),
                ('LEFTPADDING', (0,0), (-1,-1), 14),
                ('RIGHTPADDING', (0,0), (-1,-1), 14),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('ALIGN', (0,0), (-1,-1), 'CENTER'),
            ]))
            elements.append(border_table)
            elements.append(Spacer(1, 12))

            # Student Info Grid
            stream_display = s.get('stream', '').strip() or s.get('admission_no', '')
            info_data = [
                ["NAME", s['name'].upper(), "GRADE", s['class']],
                ["STREAM", stream_display, "YEAR", "2026"],
                ["TERM", f"{term} / {exam_type}", "GENDER", s['gender']]
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
            subjects = result.get('subjects', self._get_subjects_for_class(s['class'], term, exam_type))
            
            for subj in subjects:
                mk = result['marks'].get(subj, 0)
                # Subject grade
                subj_grade = self._get_grade_code_for_class(mk, s.get('class', ''))
                subj_label = self._get_subject_label(subj, s.get('class', ''))
                marks_data.append([subj_label, mk, mk, self._get_grade_name_for_class(subj_grade, s.get('class', ''))])
            
            # Summary Rows
            possible = result.get('possible_total', len(subjects) * 100)
            marks_data.append(["Total Scores", f"{result['total']}/{possible}", f"{result['average']}/100", f"Level: {result['level']}"])
            marks_data.append(["Average Scores", f"{result['average']}/100", "", f"Position: {result['position']}"])

            marks_table = Table(marks_data, colWidths=[150, 100, 100, 170])
            marks_table_style = [
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#DDEBF7")),
                ('TEXTCOLOR', (0,0), (-1,0), colors.HexColor("#1b5e20")),
                ('FONTNAME', (0,0), (-1,0), 'Helvetica-Bold'),
                ('GRID', (0,0), (-1,-3), 0.5, colors.gray),
                ('FONTSIZE', (0,0), (-1,-1), 9),
                ('BOTTOMPADDING', (0,0), (-1,-1), 8),
                ('TOPPADDING', (0,0), (-1,-1), 8),
                ('BACKGROUND', (0,-2), (-1,-1), colors.HexColor("#D0F0E0")),
                ('FONTNAME', (0,-2), (-1,-1), 'Helvetica-Bold'),
            ]
            for row_index, subject in enumerate(subjects, start=1):
                subject_style = self._get_subject_colors(subject, s.get('class', ''))
                marks_table_style.extend([
                    ('BACKGROUND', (0, row_index), (0, row_index), colors.HexColor(subject_style['soft'])),
                    ('TEXTCOLOR', (0, row_index), (0, row_index), colors.HexColor(subject_style['dark_text'])),
                    ('TEXTCOLOR', (3, row_index), (3, row_index), colors.HexColor("#d32f2f")),
                    ('FONTNAME', (3, row_index), (3, row_index), 'Helvetica-Oblique'),
                ])
            marks_table.setStyle(TableStyle(marks_table_style))
            elements.append(marks_table)
            elements.append(Spacer(1, 10))

            # Performance Trend Chart
            drawing = Drawing(400, 150)
            bc = VerticalBarChart()
            bc.x = 50
            bc.y = 20
            bc.height = 100
            bc.width = 350
            bc.data = [[result['marks'].get(sub, 0) for sub in subjects]]
            bc.strokeColor = colors.black
            bc.valueAxis.valueMin = 0
            bc.valueAxis.valueMax = 100
            bc.valueAxis.valueStep = 25
            bc.categoryAxis.labels.boxAnchor = 'ne'
            bc.categoryAxis.labels.angle = 30
            bc.categoryAxis.categoryNames = subjects
            
            for i, subj in enumerate(subjects):
                subject_style = self._get_subject_colors(subj, s.get('class', ''))
                bc.bars[0, i].fillColor = colors.HexColor(subject_style['base'])

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

            footer_text = " | ".join(footer_lines) if footer_lines else "In God We Excel"

            def add_border(canvas, doc):
                canvas.saveState()
                canvas.setStrokeColor(colors.HexColor("#1b5e20"))
                canvas.setLineWidth(2)
                canvas.rect(10, 10, A4[0]-20, A4[1]-20)

                inner_x = doc.leftMargin
                inner_width = doc.width
                page_height = doc.pagesize[1]

                if header_path and os.path.exists(header_path):
                    header_img = get_processed_letterhead_image(header_path, 'header')
                    if header_img is not None:
                        canvas.drawImage(
                            ImageReader(header_img),
                            inner_x,
                            page_height - doc.topMargin + 8,
                            width=inner_width,
                            height=header_height,
                            preserveAspectRatio=False,
                            mask='auto',
                        )

                if footer_path and os.path.exists(footer_path):
                    footer_img = get_processed_letterhead_image(footer_path, 'footer')
                    if footer_img is not None:
                        canvas.drawImage(
                            ImageReader(footer_img),
                            inner_x,
                            16,
                            width=inner_width,
                            height=footer_height,
                            preserveAspectRatio=False,
                            mask='auto',
                        )
                    else:
                        canvas.setFont('Helvetica', 8)
                        canvas.setFillColor(colors.HexColor("#666666"))
                        canvas.drawCentredString(doc.pagesize[0] / 2, 24, footer_text)
                else:
                    canvas.setFont('Helvetica', 8)
                    canvas.setFillColor(colors.HexColor("#666666"))
                    canvas.drawCentredString(doc.pagesize[0] / 2, 24, footer_text)

                canvas.restoreState()

            doc.build(elements, onFirstPage=add_border, onLaterPages=add_border)
            return True
        except Exception as e:
            messagebox.showerror("Error", f"Failed to generate PDF: {str(e)}")
            return False

    def _get_email_settings(self):
        settings = db.get_settings(EMAIL_SETTING_KEYS)
        settings['smtp_port'] = settings.get('smtp_port', '') or '465'
        settings['smtp_use_tls'] = settings.get('smtp_use_tls', '0') or '0'
        return settings

    def _open_email_settings_dialog(self):
        settings = self._get_email_settings()
        dlg = tk.Toplevel(self.root)
        dlg.title('Email Settings')
        dlg.geometry('460x420')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()
        dlg.resizable(False, False)

        em_bo, em_bi = _card_colors('peach')
        outer = tk.Frame(dlg, bg=em_bo)
        outer.place(relx=0.5, rely=0.5, anchor='center')
        card = tk.Frame(outer, bg=em_bi, padx=26, pady=24)
        card.pack(padx=1, pady=1)

        tk.Label(card, text='Result Email Settings', bg=em_bi, fg=TEXT_PRIMARY,
                 font=(FF, 14, 'bold')).pack(anchor='w', pady=(0, 16))

        def entry_field(label, default='', show=None):
            tk.Label(card, text=label, bg=em_bi, fg=TEXT_SECONDARY,
                     font=(FF, 10, 'bold')).pack(anchor='w', pady=(0, 4))
            ent = ttk.Entry(card, style='App.TEntry', show=show)
            ent.insert(0, default)
            ent.pack(fill='x', ipady=4, pady=(0, 10))
            return ent

        host_e = entry_field('SMTP Host', settings.get('smtp_host', ''))
        port_e = entry_field('SMTP Port', settings.get('smtp_port', '465'))
        user_e = entry_field('SMTP Username', settings.get('smtp_username', ''))
        pass_e = entry_field('SMTP Password / App Password', settings.get('smtp_password', ''), show='*')
        sender_e = entry_field('Sender Name', settings.get('smtp_sender_name', 'School Results'))

        tls_var = tk.BooleanVar(value=settings.get('smtp_use_tls', '0') == '1')
        tk.Checkbutton(card, text='Use STARTTLS instead of SSL', variable=tls_var,
                       bg=em_bi, fg=TEXT_PRIMARY, activebackground=em_bi,
                       selectcolor=em_bi, font=(FF, 9)).pack(anchor='w', pady=(0, 12))

        btn_row = tk.Frame(card, bg=em_bi)
        btn_row.pack(fill='x')

        def save():
            db.set_setting('smtp_host', host_e.get().strip())
            db.set_setting('smtp_port', port_e.get().strip())
            db.set_setting('smtp_username', user_e.get().strip())
            db.set_setting('smtp_password', pass_e.get().strip())
            db.set_setting('smtp_sender_name', sender_e.get().strip())
            db.set_setting('smtp_use_tls', '1' if tls_var.get() else '0')
            messagebox.showinfo('Saved', 'Email settings saved.')
            dlg.destroy()

        tk.Button(btn_row, text='Cancel', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=dlg.destroy).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Save Settings', bg=BLUE, fg='white',
                 font=(FF, 10, 'bold'), padx=18, pady=8, command=save).pack(side='left')

    def _validate_email_setup(self):
        settings = self._get_email_settings()
        missing = [key for key in ('smtp_host', 'smtp_port', 'smtp_username', 'smtp_password') if not settings.get(key)]
        if missing:
            self._open_email_settings_dialog()
            return None
        return settings

    def _is_valid_email(self, email):
        return bool(re.match(r"[^@]+@[^@]+\.[^@]+", str(email or '').strip()))

    def _create_result_email_html(self, student, term, exam_type, settings):
        school_name = 'MT. OLIVES ADVENTIST SCHOOL, NGONG'
        guardian_name = student.get('guardian_name', '').strip() or 'Parent/Guardian'
        sender_name = settings.get('smtp_sender_name', '').strip() or school_name
        stream_html = ''
        if student.get('stream', '').strip():
            stream_html = f"<tr><td style=\"padding: 6px 12px 6px 0;\"><b>Stream</b></td><td>{student.get('stream', '')}</td></tr>"
        return f"""
        <html>
        <body style="font-family: Arial, sans-serif; color: #333333; line-height: 1.5; background: #f7f8ee; padding: 20px;">
            <div style="max-width: 640px; margin: 0 auto; background: #ffffff; border: 1px solid #d9e2c3; border-radius: 10px; overflow: hidden;">
                <div style="background: linear-gradient(135deg, #6B764B 0%, #889660 100%); color: #ffffff; padding: 20px 24px;">
                    <div style="font-size: 22px; font-weight: 700;">{school_name}</div>
                    <div style="font-size: 13px; opacity: 0.9; margin-top: 4px;">Official Student Result Delivery</div>
                </div>
                <div style="padding: 24px;">
                <h2 style="margin: 0 0 12px 0; color: #55603A;">Student Report Card</h2>
                <p>Dear {guardian_name},</p>
                <p>Please find attached the academic report for <b>{student.get('name', '')}</b>.</p>
                <table style="border-collapse: collapse; margin: 16px 0;">
                    <tr><td style="padding: 6px 12px 6px 0;"><b>Class</b></td><td>{student.get('class', '')}</td></tr>
                    {stream_html}
                    <tr><td style="padding: 6px 12px 6px 0;"><b>Term</b></td><td>{term}</td></tr>
                    <tr><td style="padding: 6px 12px 6px 0;"><b>Exam</b></td><td>{exam_type}</td></tr>
                    <tr><td style="padding: 6px 12px 6px 0;"><b>Admission No</b></td><td>{student.get('admission_no', '')}</td></tr>
                </table>
                <p>We appreciate your continued support in your child's education.</p>
                <br>
                <p>
                    Regards,<br>
                    <b>{sender_name}</b><br>
                    MT. OLIVES ADVENTIST SCHOOL, NGONG<br>
                    Nairobi, Kenya
                </p>
                </div>
            </div>
        </body>
        </html>
        """

    def _send_email_with_attachment(self, to_email, subject, body, attachment_path, settings, html_body=''):
        msg = EmailMessage()
        sender_name = settings.get('smtp_sender_name', '').strip() or 'School Results'
        username = settings.get('smtp_username', '').strip()
        msg['Subject'] = subject
        msg['From'] = f'{sender_name} <{username}>'
        msg['To'] = to_email.strip()
        msg.set_content(body)
        if html_body:
            msg.add_alternative(html_body, subtype='html')

        with open(attachment_path, 'rb') as handle:
            msg.add_attachment(handle.read(), maintype='application', subtype='pdf',
                               filename=os.path.basename(attachment_path))

        host = settings.get('smtp_host', '').strip()
        port = int(settings.get('smtp_port', '465') or 465)
        password = settings.get('smtp_password', '')
        use_tls = settings.get('smtp_use_tls', '0') == '1'

        if use_tls:
            with smtplib.SMTP(host, port, timeout=30) as server:
                server.starttls()
                server.login(username, password)
                server.send_message(msg)
        else:
            with smtplib.SMTP_SSL(host, port, timeout=30) as server:
                server.login(username, password)
                server.send_message(msg)

    def _build_temp_result_pdf(self, result, total_students, term, exam_type):
        temp_dir = tempfile.mkdtemp(prefix='result_mail_')
        student = result['student']
        pdf_path = os.path.join(
            temp_dir,
            f"Report_{student['name'].replace(' ', '_')}_{student['class'].replace(' ', '_')}_{term}_{exam_type.replace('-', '_')}.pdf"
        )
        ok = self._build_report_card_pdf(result, total_students, term, exam_type, pdf_path)
        return pdf_path if ok else None

    def _get_result_for_student(self, student_id, class_name, term, exam_type):
        results = self._get_ranked_results(class_name, term, exam_type)
        return next((r for r in results if r['student']['id'] == student_id), None), results

    def _retry_failed_email_logs(self, logs):
        settings = self._validate_email_setup()
        if not settings:
            return
        if not logs:
            messagebox.showinfo('No Failed Emails', 'There are no failed email records to retry.')
            return

        progress_dialog, status_label, percent_label, progress = self._open_progress_dialog(
            'Retrying Failed Emails',
            f'Preparing {len(logs)} failed email(s)...'
        )

        def worker():
            sent = 0
            failed = 0
            for index, log in enumerate(logs, start=1):
                pdf_path = None
                try:
                    student = db.get_student(log['student_id'])
                    if not student:
                        failed += 1
                        continue
                    result, results = self._get_result_for_student(
                        log['student_id'],
                        log.get('class_name', ''),
                        log['term'],
                        log.get('exam_type', DEFAULT_EXAM_TYPE)
                    )
                    if not result:
                        failed += 1
                        continue
                    self.root.after(0, lambda i=index, n=student['name']: self._update_progress_dialog(
                        progress_dialog, status_label, percent_label, progress, i - 1, len(logs),
                        f'{i}/{len(logs)} Retrying {n}...'
                    ))
                    pdf_path = self._build_temp_result_pdf(result, len(results), log['term'], log.get('exam_type', DEFAULT_EXAM_TYPE))
                    if not pdf_path:
                        failed += 1
                        continue
                    subject = f"{student['name']} Report Card - Term {log['term']} {log.get('exam_type', DEFAULT_EXAM_TYPE)}"
                    body = (
                        f"Dear {student.get('guardian_name', 'Parent/Guardian')},\n\n"
                        f"Please find attached the report card for {student['name']}.\n\n"
                        f"Class: {student['class']}\n"
                        f"Term: {log['term']}\n"
                        f"Exam: {log.get('exam_type', DEFAULT_EXAM_TYPE)}\n\n"
                        f"Regards,\n{settings.get('smtp_sender_name', 'School Results')}"
                    )
                    html_body = self._create_result_email_html(student, log['term'], log.get('exam_type', DEFAULT_EXAM_TYPE), settings)
                    self._send_email_with_attachment(log['recipient_email'], subject, body, pdf_path, settings, html_body)
                    db.log_email_delivery(student['id'], log['term'], log.get('exam_type', DEFAULT_EXAM_TYPE), log['recipient_email'], 'sent')
                    sent += 1
                except Exception as exc:
                    db.log_email_delivery(log['student_id'], log['term'], log.get('exam_type', DEFAULT_EXAM_TYPE), log['recipient_email'], 'failed', str(exc))
                    failed += 1
                finally:
                    if pdf_path and os.path.exists(pdf_path):
                        try:
                            shutil.rmtree(os.path.dirname(pdf_path), ignore_errors=True)
                        except Exception:
                            pass

            self.root.after(0, progress_dialog.destroy)
            self.root.after(0, lambda: messagebox.showinfo('Retry Complete', f'Sent: {sent}\nFailed: {failed}'))

        threading.Thread(target=worker, daemon=True).start()

    def _export_failed_email_logs(self):
        cls = self.rc_cls_cb.get()
        term = self.rc_term_cb.get()
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        stream = self._get_selected_report_stream()
        logs = db.get_email_logs(cls, term, exam_type, stream, 'failed')
        if not logs:
            messagebox.showinfo('No Failed Emails', 'No failed email logs found for this selection.')
            return
        file_path = filedialog.asksaveasfilename(
            title='Export Failed Emails',
            defaultextension='.csv',
            filetypes=[('CSV files', '*.csv')],
            initialfile=f'failed_emails_{cls.replace(" ", "_")}_{(stream or "all_streams").replace(" ", "_")}_{term}_{exam_type.replace("-", "_")}.csv'
        )
        if not file_path:
            return
        with open(file_path, 'w', newline='', encoding='utf-8') as handle:
            writer = csv.writer(handle)
            writer.writerow(['Student', 'Admission No', 'Class', 'Term', 'Exam Type', 'Recipient', 'Status', 'Error', 'Sent At'])
            for log in logs:
                writer.writerow([
                    log.get('student_name', ''),
                    log.get('admission_no', ''),
                    log.get('class_name', ''),
                    log.get('term', ''),
                    log.get('exam_type', ''),
                    log.get('recipient_email', ''),
                    log.get('status', ''),
                    log.get('error_message', ''),
                    log.get('sent_at', ''),
                ])
        messagebox.showinfo('Exported', f'Failed email log exported to {file_path}')

    def _show_failed_email_logs(self):
        cls = self.rc_cls_cb.get()
        term = self.rc_term_cb.get()
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        stream = self._get_selected_report_stream()
        logs = db.get_email_logs(cls, term, exam_type, stream, 'failed')

        dlg = tk.Toplevel(self.root)
        dlg.title('Failed Email Logs')
        dlg.geometry('860x420')
        dlg.configure(bg=CONTENT_BG)
        dlg.transient(self.root)
        dlg.grab_set()

        fl_bo, fl_bi = _card_colors('blossom')
        outer = tk.Frame(dlg, bg=fl_bo)
        outer.pack(fill='both', expand=True, padx=14, pady=14)
        card = tk.Frame(outer, bg=fl_bi)
        card.pack(fill='both', expand=True, padx=1, pady=1)

        top = tk.Frame(card, bg=fl_bi, padx=16, pady=14)
        top.pack(fill='x')
        stream_suffix = f' / {stream}' if stream else ''
        tk.Label(top, text=f'Failed Email Logs - {cls}{stream_suffix} / Term {term} / {exam_type}',
                 bg=fl_bi, fg=TEXT_PRIMARY, font=(FF, 13, 'bold')).pack(side='left')

        cols = ('student', 'email', 'error', 'sent_at')
        tree = ttk.Treeview(card, columns=cols, show='headings', style='App.Treeview')
        tree.heading('student', text='Student')
        tree.heading('email', text='Recipient Email')
        tree.heading('error', text='Error')
        tree.heading('sent_at', text='Time')
        tree.column('student', width=180)
        tree.column('email', width=180)
        tree.column('error', width=320)
        tree.column('sent_at', width=150)
        for log in logs:
            tree.insert('', 'end', iid=log['id'], values=(
                log.get('student_name', ''),
                log.get('recipient_email', ''),
                log.get('error_message', ''),
                log.get('sent_at', ''),
            ))
        tree.pack(fill='both', expand=True, padx=16, pady=(0, 12))

        btn_row = tk.Frame(card, bg=fl_bi, padx=16, pady=12)
        btn_row.pack(fill='x')

        def retry_selected():
            selected = tree.selection()
            picked = [log for log in logs if log['id'] in selected]
            if not picked:
                messagebox.showwarning('Select', 'Select one or more failed emails to retry.', parent=dlg)
                return
            dlg.destroy()
            self._retry_failed_email_logs(picked)

        def retry_all():
            dlg.destroy()
            self._retry_failed_email_logs(logs)

        tk.Button(btn_row, text='Retry Selected', bg=BLUE, fg='white',
                 font=(FF, 10, 'bold'), padx=16, pady=8, command=retry_selected).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Retry All', bg=GREEN, fg='white',
                 font=(FF, 10, 'bold'), padx=16, pady=8, command=retry_all).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Export Failed CSV', bg=PURPLE, fg='white',
                 font=(FF, 10, 'bold'), padx=16, pady=8, command=self._export_failed_email_logs).pack(side='left', padx=(0, 8))
        tk.Button(btn_row, text='Close', bg=LEMON_SOFT, fg=TEXT_PRIMARY,
                 font=(FF, 10, 'bold'), padx=16, pady=8, command=dlg.destroy).pack(side='right')

    def _send_result_email(self):
        name = self.rc_stu_cb.get()
        if not name:
            return
        settings = self._validate_email_setup()
        if not settings:
            return
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        term = self.rc_term_cb.get()
        results = self._get_report_card_results()
        result = next((r for r in results if r['student']['name'] == name), None)
        if not result:
            return
        student = result['student']
        recipient = student.get('parent_email', '').strip()
        if not recipient:
            messagebox.showwarning('Missing Email', 'This student has no parent email saved.')
            return
        if not self._is_valid_email(recipient):
            messagebox.showwarning('Invalid Email', f'Invalid parent email for {student["name"]}: {recipient}')
            return

        progress_dialog, status_label, percent_label, progress = self._open_progress_dialog(
            'Sending Result',
            f'Preparing result for {student["name"]}...'
        )

        def worker():
            pdf_path = None
            try:
                self.root.after(0, lambda: self._update_progress_dialog(
                    progress_dialog, status_label, percent_label, progress, 1, 3,
                    f'Generating PDF for {student["name"]}...'
                ))
                pdf_path = self._build_temp_result_pdf(result, len(results), term, exam_type)
                if not pdf_path:
                    self.root.after(0, progress_dialog.destroy)
                    return
                subject = f"{student['name']} Report Card - Term {term} {exam_type}"
                body = (
                    f"Dear {student.get('guardian_name', 'Parent/Guardian')},\n\n"
                    f"Please find attached the report card for {student['name']}.\n\n"
                    f"Class: {student['class']}\n"
                    f"{'Stream: ' + student.get('stream', '').strip() + chr(10) if student.get('stream', '').strip() else ''}"
                    f"Term: {term}\n"
                    f"Exam: {exam_type}\n\n"
                    f"Regards,\n{settings.get('smtp_sender_name', 'School Results')}"
                )
                html_body = self._create_result_email_html(student, term, exam_type, settings)
                self.root.after(0, lambda: self._update_progress_dialog(
                    progress_dialog, status_label, percent_label, progress, 2, 3,
                    f'Sending email to {recipient}...'
                ))
                self._send_email_with_attachment(recipient, subject, body, pdf_path, settings, html_body)
                db.log_email_delivery(student['id'], term, exam_type, recipient, 'sent')
                self.root.after(0, lambda: self._update_progress_dialog(
                    progress_dialog, status_label, percent_label, progress, 3, 3,
                    f'Email sent to {recipient}.'
                ))
                self.root.after(0, progress_dialog.destroy)
                self.root.after(0, lambda: messagebox.showinfo('Email Sent', f'Result sent to {recipient}'))
            except Exception as exc:
                db.log_email_delivery(student['id'], term, exam_type, recipient, 'failed', str(exc))
                self.root.after(0, progress_dialog.destroy)
                self.root.after(0, lambda: messagebox.showerror('Email Failed', str(exc)))
            finally:
                if pdf_path and os.path.exists(pdf_path):
                    try:
                        shutil.rmtree(os.path.dirname(pdf_path), ignore_errors=True)
                    except Exception:
                        pass

        threading.Thread(target=worker, daemon=True).start()

    def _send_all_results_email(self):
        settings = self._validate_email_setup()
        if not settings:
            return
        cls = self.rc_cls_cb.get()
        term = self.rc_term_cb.get()
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        selected_stream = self._get_selected_report_stream()
        results = self._get_report_card_results()
        if not results:
            scope = f'{cls} - {selected_stream}' if selected_stream else cls
            messagebox.showwarning('No Data', f'No students found for {scope}.')
            return

        progress_dialog, status_label, percent_label, progress = self._open_progress_dialog(
            'Sending Bulk Results',
            f'Preparing to send {len(results)} result(s)...'
        )

        def worker():
            sent = 0
            failed = 0
            missing = 0
            total = len(results)
            for index, result in enumerate(results, start=1):
                student = result['student']
                recipient = student.get('parent_email', '').strip()
                if not recipient:
                    missing += 1
                    self.root.after(0, lambda i=index, n=student['name']: self._update_progress_dialog(
                        progress_dialog, status_label, percent_label, progress, i, total,
                        f'{i}/{total} Skipping {n} - missing parent email.'
                    ))
                    continue
                if not self._is_valid_email(recipient):
                    failed += 1
                    db.log_email_delivery(student['id'], term, exam_type, recipient, 'failed', 'Invalid email format')
                    self.root.after(0, lambda i=index, n=student['name']: self._update_progress_dialog(
                        progress_dialog, status_label, percent_label, progress, i, total,
                        f'{i}/{total} Skipping {n} - invalid email address.'
                    ))
                    continue
                pdf_path = None
                try:
                    self.root.after(0, lambda i=index, n=student['name']: self._update_progress_dialog(
                        progress_dialog, status_label, percent_label, progress, i - 1, total,
                        f'{i}/{total} Generating PDF for {n}...'
                    ))
                    pdf_path = self._build_temp_result_pdf(result, len(results), term, exam_type)
                    if not pdf_path:
                        failed += 1
                        continue
                    subject = f"{student['name']} Report Card - Term {term} {exam_type}"
                    body = (
                        f"Dear {student.get('guardian_name', 'Parent/Guardian')},\n\n"
                        f"Please find attached the report card for {student['name']}.\n\n"
                        f"Class: {student['class']}\n"
                        f"{'Stream: ' + student.get('stream', '').strip() + chr(10) if student.get('stream', '').strip() else ''}"
                        f"Term: {term}\n"
                        f"Exam: {exam_type}\n\n"
                        f"Regards,\n{settings.get('smtp_sender_name', 'School Results')}"
                    )
                    html_body = self._create_result_email_html(student, term, exam_type, settings)
                    self.root.after(0, lambda i=index, n=student['name']: self._update_progress_dialog(
                        progress_dialog, status_label, percent_label, progress, i, total,
                        f'{i}/{total} Sending to {n}...'
                    ))
                    self._send_email_with_attachment(recipient, subject, body, pdf_path, settings, html_body)
                    db.log_email_delivery(student['id'], term, exam_type, recipient, 'sent')
                    sent += 1
                except Exception as exc:
                    db.log_email_delivery(student['id'], term, exam_type, recipient, 'failed', str(exc))
                    failed += 1
                finally:
                    if pdf_path and os.path.exists(pdf_path):
                        try:
                            shutil.rmtree(os.path.dirname(pdf_path), ignore_errors=True)
                        except Exception:
                            pass

            self.root.after(0, progress_dialog.destroy)
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    'Bulk Email Complete',
                    f'Sent: {sent}\nFailed: {failed}\nMissing Email: {missing}'
                )
            )

        threading.Thread(target=worker, daemon=True).start()

    # ==================== CBC INFO ====================
    def show_cbc_info(self):
        """Legacy entry point kept for compatibility; CBC info now lives on the dashboard."""
        self.show_dashboard()

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
        lbl('Exam:')
        self.ch_exam_cb = ttk.Combobox(ctrl, values=EXAM_TYPES, state='readonly',
                                       style='App.TCombobox', width=12)
        self.ch_exam_cb.set(DEFAULT_EXAM_TYPE)
        self.ch_exam_cb.pack(side='left', ipady=4)
        self.ch_cls_cb.bind('<<ComboboxSelected>>',  lambda e: self.load_charts())
        self.ch_term_cb.bind('<<ComboboxSelected>>', lambda e: self.load_charts())
        self.ch_exam_cb.bind('<<ComboboxSelected>>', lambda e: self.load_charts())

        ch_bo, ch_bi = _card_colors('sky')
        chart_outer = tk.Frame(self.content_frame, bg=ch_bo)
        chart_outer.pack(fill='both', expand=True, pady=4)
        chart_card = tk.Frame(chart_outer, bg=ch_bi, padx=10, pady=10)
        chart_card.pack(fill='both', expand=True, padx=1, pady=1)
        self._chart_card_bg = ch_bi

        self.fig, self.axes = plt.subplots(2, 2, figsize=(12, 8),
                                           constrained_layout=True)
        self.fig.set_facecolor(ch_bi)
        self.canvas = FigureCanvasTkAgg(self.fig, master=chart_card)
        self.canvas.get_tk_widget().pack(fill='both', expand=True)
        self.load_charts()

    def load_charts(self):
        cls  = self.ch_cls_cb.get()
        term = self.ch_term_cb.get()
        exam_type = self.ch_exam_cb.get() or DEFAULT_EXAM_TYPE
        results = self._get_ranked_results(cls, term, exam_type)
        subjects = self._get_subjects_for_scope(cls, term, exam_type, results)
        plot_panel = _mix_hex(getattr(self, '_chart_card_bg', CARD_BG), '#ffffff', 0.5)
        for ax in self.axes.flat:
            ax.clear()
            ax.set_facecolor(plot_panel)

        subj_totals = {s: [] for s in subjects}
        for r in results:
            for s in subjects:
                if r['marks'].get(s): subj_totals[s].append(r['marks'][s])

        avgs   = [round(sum(subj_totals[s]) / len(subj_totals[s]), 1) if subj_totals[s] else 0 for s in subjects]
        colors = [self._get_subject_color(subject, cls) for subject in subjects]

        subject_labels = [self._get_subject_label(subject, cls if cls != 'All' else '') for subject in subjects]
        bars0 = self.axes[0, 0].bar(subject_labels, avgs, color=colors, edgecolor='none', width=0.6)
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

        grade_counts = {}
        for r in results:
            grade_counts[r['grade']] = grade_counts.get(r['grade'], 0) + 1
        grades = [g for g, v in grade_counts.items() if v > 0]
        counts = [grade_counts[g] for g in grades]
        if counts:
            pie_colors = [self._get_grade_color(g) for g in grades]
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
            for c in self.get_current_classes():
                cr = self._get_ranked_results(c, term, exam_type)
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
        rc_classes = self.get_current_classes()
        self.rc_cls_cb.set(rc_classes[0] if rc_classes else '')
        self.rc_cls_cb.pack(side='left', ipady=4)
        lbl('Stream:')
        self.rc_stream_cb = ttk.Combobox(ctrl, state='readonly',
                                         style='App.TCombobox', width=14)
        self.rc_stream_cb.pack(side='left', ipady=4)
        lbl('Term:')
        self.rc_term_cb = ttk.Combobox(ctrl, values=TERMS, state='readonly',
                                       style='App.TCombobox', width=10)
        self.rc_term_cb.set(TERMS[0])
        self.rc_term_cb.pack(side='left', ipady=4)
        lbl('Exam:')
        self.rc_exam_cb = ttk.Combobox(ctrl, values=EXAM_TYPES, state='readonly',
                                       style='App.TCombobox', width=12)
        self.rc_exam_cb.set(DEFAULT_EXAM_TYPE)
        self.rc_exam_cb.pack(side='left', ipady=4)
        lbl('Student:')
        self.rc_stu_cb = ttk.Combobox(ctrl, state='readonly',
                                      style='App.TCombobox', width=22)
        self.rc_stu_cb.pack(side='left', ipady=4)

        self.rc_cls_cb.bind('<<ComboboxSelected>>',  lambda e: self._refresh_report_card_streams())
        self.rc_stream_cb.bind('<<ComboboxSelected>>', lambda e: self._load_rc())
        self.rc_term_cb.bind('<<ComboboxSelected>>', lambda e: self._load_rc())
        self.rc_exam_cb.bind('<<ComboboxSelected>>', lambda e: self._load_rc())
        self.rc_stu_cb.bind('<<ComboboxSelected>>',  lambda e: self._display_rc())

        self._toolbar_btn(ctrl, '\U0001f5a8  Print', self._print_rc).pack(side='left', padx=14)
        self._toolbar_btn(ctrl, 'Email Result', self._send_result_email, bg=ORANGE).pack(side='left', padx=4)
        self._toolbar_btn(ctrl, 'Email All', self._send_all_results_email, bg=PURPLE).pack(side='left', padx=4)
        self._toolbar_btn(ctrl, 'Failed Emails', self._show_failed_email_logs, bg='#8B5CF6').pack(side='left', padx=4)
        self._toolbar_btn(ctrl, 'Email Settings', self._open_email_settings_dialog, bg='#475569').pack(side='left', padx=4)
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

        self._refresh_report_card_streams(reload_results=False)
        self._load_rc()

    def _load_rc(self):
        results = self._get_report_card_results()
        names = [r['student']['name'] for r in results]
        self.rc_stu_cb['values'] = names
        if names:
            self.rc_stu_cb.current(0)
            self._display_rc()
        else:
            self.rc_stu_cb.set('')
            for w in self._rc_paper.winfo_children():
                w.destroy()

    def _display_rc(self):
        name = self.rc_stu_cb.get()
        if not name: return
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        results = self._get_report_card_results()
        result  = next((r for r in results if r['student']['name'] == name), None)
        if not result: return
        for w in self._rc_paper.winfo_children():
            w.destroy()
        self._render_report_card(self._rc_paper, result, len(results),
                                 self.rc_term_cb.get(), exam_type)

    def _render_report_card(self, parent, result, total_students, term, exam_type=DEFAULT_EXAM_TYPE):
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
            return self._get_grade_code_for_class(m, s.get('class', ''))

        subjects = result.get('subjects', self._get_subjects_for_class(s.get('class', ''), term, exam_type))

        # School Header - Try to use letterhead image
        letterhead_assets = get_letterhead_assets()
        header_path = letterhead_assets.get('header_path')
        using_letterhead_header = False
        
        if header_path and os.path.exists(header_path):
            try:
                header_img = get_processed_letterhead_image(header_path, 'header')
                if header_img is None:
                    raise ValueError('Processed header image unavailable')
                header_width = 620
                header_img = header_img.resize((header_width, int(header_img.height * header_width / header_img.width)), Image.LANCZOS)
                header_photo = ImageTk.PhotoImage(header_img)
                header_label = tk.Label(parent, image=header_photo, bg='white')
                header_label.image = header_photo  # Keep reference
                header_label.pack(pady=(0, 8))
                using_letterhead_header = True
            except Exception as e:
                print(f'Failed to load letterhead: {e}')
                # Fallback to text header
                tk.Label(parent, text='MT OLIVES ADVENTIST SCHOOL',
                         bg='white', fg=SCH_BLUE, font=(FF, 15, 'bold')).pack()
        else:
            # Fallback to text header
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
        
        if not using_letterhead_header:
            tk.Label(parent, text='In God We Excel',
                     bg='white', fg='#999', font=(FF, 9, 'italic')).pack(pady=(0, 8))
        tk.Frame(parent, bg=SCH_BLUE, height=2).pack(fill='x', pady=(0, 12))

        title_border = tk.Frame(parent, bg=RED_ACC)
        title_border.pack(pady=(0, 14))
        title_inner = tk.Frame(title_border, bg='white', padx=30, pady=7)
        title_inner.pack(padx=2, pady=2)
        tk.Label(title_inner, text='LEARNER ASSESSMENT REPORT CARD',
                 bg='white', fg=TTL_ORG, font=(FF, 12, 'bold')).pack()

        # Student Info grid
        grade_num = s.get('class', 'Grade 7').replace('Grade ', '')
        stream_display = s.get('stream', '').strip() or s.get('admission_no', '')
        info_rows = [
            ('NAME',   s['name'],         'GRADE',  grade_num),
            ('STREAM', stream_display, 'YEAR',   '2026'),
            ('TERM',   f'{term} / {exam_type}', 'GENDER', s['gender']),
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
        for i, subj in enumerate(subjects):
            mk    = marks.get(subj, 0)
            grade = get_grade(mk)
            subject_style = self._get_subject_colors(subj, s.get('class', ''))
            row_bg = subject_style['soft'] if i % 2 == 0 else subject_style['mid']
            subj_label = self._get_subject_label(subj, s.get('class', ''))
            tk.Label(tbl, text=subj_label, bg=row_bg, fg=subject_style['dark_text'], font=(FF, 10, 'bold'),
                     padx=10, pady=6, anchor='w'
                     ).grid(row=i + 1, column=0, sticky='nsew', padx=1, pady=1)
            for col in (1, 2):
                tk.Label(tbl, text=str(mk), bg=row_bg, fg='#333', font=(FF, 10),
                         padx=10, pady=6
                         ).grid(row=i + 1, column=col, sticky='nsew', padx=1, pady=1)
            tk.Label(tbl, text=self._get_grade_name_for_class(grade, s.get('class', '')), bg=row_bg,
                     fg=PERF_CLR.get(grade_base_code(grade), PERF_CLR['IE']), font=(FF, 10, 'italic'),
                     padx=10, pady=6, anchor='w'
                     ).grid(row=i + 1, column=3, sticky='nsew', padx=1, pady=1)

        total_marks = result['total']
        avg_score   = result['average']
        grade_ov    = result['grade']
        pos         = result['position']
        possible    = result.get('possible_total', len(subjects) * 100)
        base        = len(subjects) + 1
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
        chart_subject_labels = [self._get_subject_label(sub, s.get('class', '')) for sub in subjects]
        mark_vals  = [marks.get(sub, 0) for sub in subjects]
        bar_colors = [self._get_subject_color(sub, s.get('class', '')) for sub in subjects]
        ax.bar(chart_subject_labels, mark_vals, color=bar_colors, edgecolor='none', width=0.55)
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

        # Footer contacts should sit at the very bottom and span wider.
        footer_path = letterhead_assets.get('footer_path')
        if footer_path and os.path.exists(footer_path):
            try:
                footer_img = Image.open(footer_path)
                footer_width = 620
                footer_img = footer_img.resize((footer_width, int(footer_img.height * footer_width / footer_img.width)), Image.LANCZOS)
                footer_photo = ImageTk.PhotoImage(footer_img)
                footer_label = tk.Label(parent, image=footer_photo, bg='white')
                footer_label.image = footer_photo  # Keep reference
                footer_label.pack(pady=(10, 0))
            except Exception as e:
                print(f'Failed to load letterhead footer: {e}')

    def _gen_rc_text(self, result, total, term, exam_type=DEFAULT_EXAM_TYPE):
        """Plain-text fallback used for printing."""
        s, m = result['student'], result['marks']
        subjects = result.get('subjects', self._get_subjects_for_class(s.get('class', ''), term, exam_type))
        possible = result.get('possible_total', len(subjects) * 100)
        lines = [
            '=' * 62,
            '      MT OLIVES ADVENTIST SCHOOL',
            '          LEARNER ASSESSMENT REPORT CARD',
            '=' * 62,
            f"  Name    : {s['name']:<20}  Grade  : {s.get('class','').replace('Grade ','')}",
            f"  Stream  : {s['admission_no']:<20}  Year   : 2026",
            f"  Term    : {f'{term} / {exam_type}':<20}  Gender : {s['gender']}",
            '-' * 62,
            f"  {'Subject':<16} {'Marks':>6}  {'Avg':>6}  Performance Level",
            '-' * 62,
        ]
        for sub in subjects:
            mk = m.get(sub, 0)
            g  = self._get_grade_code_for_class(mk, s.get('class', ''))
            lines.append(f"  {sub:<16} {mk:>6}  {mk:>6}  {self._get_grade_name_for_class(g, s.get('class', ''))}")
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
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        cls = self.rc_cls_cb.get()
        term = self.rc_term_cb.get()
        results = self._get_report_card_results()
        result  = next((r for r in results if r['student']['name'] == name), None)
        if not result: return
        s = result['student']
        stream = self._get_selected_report_stream() or s.get('stream', '').strip()
        file_path = filedialog.asksaveasfilename(
            title='Save Report Card PDF',
            defaultextension='.pdf',
            filetypes=[('PDF files', '*.pdf')],
            initialfile=f"Report_{s['name'].replace(' ', '_')}_{cls.replace(' ', '_')}_{(stream or 'all_streams').replace(' ', '_')}_{term}_{exam_type.replace('-', '_')}.pdf"
        )
        if not file_path:
            return
        if self._build_report_card_pdf(result, len(results), term, exam_type, file_path):
            messagebox.showinfo('Done', f'Report card PDF saved to {file_path}')

    def _print_all_rc(self):
        cls  = self.rc_cls_cb.get()
        term = self.rc_term_cb.get()
        exam_type = self.rc_exam_cb.get() or DEFAULT_EXAM_TYPE
        selected_stream = self._get_selected_report_stream()
        results = self._get_report_card_results()
        if not results:
            messagebox.showwarning('No Data', 'No students found'); return
        output_dir = filedialog.askdirectory(title='Select Folder for Report Card PDFs')
        if not output_dir:
            return

        saved = 0
        for r in results:
            student = r['student']
            file_path = os.path.join(
                output_dir,
                f"Report_{student['name'].replace(' ', '_')}_{cls.replace(' ', '_')}_{(selected_stream or student.get('stream', '').strip() or 'all_streams').replace(' ', '_')}_{term}_{exam_type.replace('-', '_')}.pdf"
            )
            if self._build_report_card_pdf(r, len(results), term, exam_type, file_path):
                saved += 1

        if saved:
            messagebox.showinfo('Done', f'{saved} report card PDF(s) saved to {output_dir}')


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




