# -*- coding: utf-8 -*-
"""
Main application file for School Report Management System
Python Tkinter Desktop Application – Modern Redesigned UI
"""

import os
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from PIL import Image, ImageTk
from database import db
import matplotlib
matplotlib.use('TkAgg')
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
import csv
from datetime import datetime

# ====================== DESIGN TOKENS ======================
# Logo Color Theme: Green-based
SIDEBAR_BG        = '#1b5e20'       # Dark Olive Green (darker shade)
SIDEBAR_ACTIVE    = '#4CAF50'       # Light Green
SIDEBAR_HOVER     = '#2E7D32'       # Dark Olive Green
SIDEBAR_TEXT      = '#a5d6a7'       # Light green text
SIDEBAR_TEXT_ACT  = '#ffffff'       # White

CONTENT_BG  = '#ffffff'             # White
CARD_BG     = '#ffffff'             # White
BORDER_CLR  = '#e0e0e0'             # Light gray border

TEXT_PRIMARY   = '#333333'          # Dark Gray / Black
TEXT_SECONDARY = '#666666'          # Medium Gray

BLUE   = '#4CAF50'                  # Light Green (logo theme)
GREEN  = '#2E7D32'                   # Dark Olive Green
ORANGE = '#f59e0b'                   # Orange (kept)
PURPLE = '#6366f1'                   # Purple (kept)

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

# ====================== CONSTANTS ==========================
SUBJECTS = ['Math', 'Eng', 'Kis', 'Int Sci', 'Agri', 'SST', 'CRE', 'CIA', 'Pre-Tech']
CLASSES  = ['Grade 7', 'Grade 8', 'Grade 9']
TERMS    = ['One', 'Two', 'Three']


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
        self.root.title('School Report Management System')
        self.root.geometry('1280x760')
        self.root.configure(bg=CONTENT_BG)
        self.root.minsize(960, 620)

        self.current_user = None
        self.nav_frames: dict = {}
        self.active_nav: str = ''
        
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

        setup_treeview_style()
        self.show_login()

    # ------------------- utilities -------------------
    def _clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    def clear_frame(self):
        for w in self.content_frame.winfo_children():
            w.destroy()

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
        logo = tk.Canvas(frame, width=56, height=56, bg='white', highlightthickness=0)
        logo.pack(pady=(24, 12))
        _rr(logo, 0, 0, 56, 56, 12, self._BLUE)
        logo.create_text(28, 28, text='\U0001f393', font=(FF, 22), fill='white')

        # ── Title ───────────────────────────────────────────────────────────
        tk.Label(frame, text='SchoolReport', bg='white', fg=self._TEXT,
                 font=(FF, 21, 'bold')).pack(pady=(0, 3))
        tk.Label(frame, text='School Report Management System', bg='white',
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
            self.show_main_window()
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

    def login(self):
        """Legacy – kept for compatibility; auth now handled by _do_login."""
        pass

    # ------------------- main layout -------------------
    def show_main_window(self):
        self._clear()

        wrapper = tk.Frame(self.root, bg=CONTENT_BG)
        wrapper.pack(fill='both', expand=True)

        self._build_sidebar(wrapper)

        # content scroller
        content_wrapper = tk.Frame(wrapper, bg=CONTENT_BG)
        content_wrapper.pack(side='left', fill='both', expand=True)

        self.content_frame = tk.Frame(content_wrapper, bg=CONTENT_BG)
        self.content_frame.pack(fill='both', expand=True, padx=32, pady=28)

        self.show_dashboard()

    def _build_sidebar(self, parent):
        sb = tk.Frame(parent, bg=SIDEBAR_BG, width=238)
        sb.pack(side='left', fill='y')
        sb.pack_propagate(False)

        # --- logo ---
        logo_row = tk.Frame(sb, bg=SIDEBAR_BG)
        logo_row.pack(fill='x', padx=18, pady=(22, 10))

        icon_c = tk.Canvas(logo_row, width=40, height=40, bg=SIDEBAR_BG, highlightthickness=0)
        icon_c.pack(side='left')
        icon_c.create_oval(0, 0, 40, 40, fill=SIDEBAR_ACTIVE, outline=SIDEBAR_ACTIVE)
        icon_c.create_text(20, 20, text='\U0001f393', font=(FF, 17))

        title_box = tk.Frame(logo_row, bg=SIDEBAR_BG)
        title_box.pack(side='left', padx=10)
        tk.Label(title_box, text='SchoolReport', bg=SIDEBAR_BG, fg='white',
                 font=(FF, 13, 'bold')).pack(anchor='w')
        tk.Label(title_box, text='Management System', bg=SIDEBAR_BG, fg=SIDEBAR_TEXT,
                 font=(FF, 9)).pack(anchor='w')

        tk.Frame(sb, bg='#2E7D32', height=1).pack(fill='x', padx=18, pady=10)

        # --- nav items ---
        nav_cfg = [
            ('⊞',  'Dashboard',   self.show_dashboard),
            ('◉',  'Students',    self.show_students),
            ('\u270e', 'Enter Marks', self.show_marks_entry),
            ('≡',  'Reports',     self.show_reports),
            ('▦',  'Charts',      self.show_charts),
            ('☰',  'Report Cards', self.show_report_cards),
        ]
        nav_box = tk.Frame(sb, bg=SIDEBAR_BG)
        nav_box.pack(fill='x', padx=10)
        self.nav_frames = {}

        for icon, label, cmd in nav_cfg:
            self._nav_item(nav_box, icon, label, cmd)

        # --- bottom ---
        bot = tk.Frame(sb, bg=SIDEBAR_BG)
        bot.pack(side='bottom', fill='x', padx=18, pady=18)
        tk.Frame(bot, bg='#2E7D32', height=1).pack(fill='x', pady=(0, 12))

        uname = self.current_user.get('username', 'admin')
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
        tk.Label(title_frame, text='School Report Management System overview', bg=CONTENT_BG, fg=TEXT_SECONDARY,
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

        # ---- grading scale card ----
        gc_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        gc_outer.pack(fill='x', pady=8)
        gc = tk.Frame(gc_outer, bg=CARD_BG, padx=24, pady=20)
        gc.pack(fill='both', expand=True, padx=1, pady=1)

        tk.Label(gc, text='Grading Scale', bg=CARD_BG, fg=TEXT_PRIMARY,
                 font=(FF, 13, 'bold')).pack(anchor='w', pady=(0, 14))

        row = tk.Frame(gc, bg=CARD_BG)
        row.pack(fill='x')

        grade_data = [
            ('EE', '80–100', 'Exceeding\nExpectations',   GRADE_COLORS['EE']),
            ('ME', '70–79',  'Meeting\nExpectations',     GRADE_COLORS['ME']),
            ('AE', '60–69',  'Approaching\nExpectations', GRADE_COLORS['AE']),
            ('BE', '50–59',  'Below\nExpectations',       GRADE_COLORS['BE']),
            ('IE', '0–49',   'Inadequate',                GRADE_COLORS['IE']),
        ]
        for code, rng, desc, clr in grade_data:
            tile = tk.Frame(row, bg=clr, padx=14, pady=14)
            tile.pack(side='left', padx=5, expand=True, fill='both')
            tk.Label(tile, text=code, bg=clr, fg='white',
                     font=(FF, 18, 'bold')).pack()
            tk.Label(tile, text=rng, bg=clr, fg='white',
                     font=(FF, 9)).pack()
            tk.Label(tile, text=desc, bg=clr, fg='white',
                     font=(FF, 8), justify='center').pack()

    # ==================== STUDENTS ====================
    def show_students(self):
        self.clear_frame()
        self._set_nav('Students')
        self._page_header('Students', 'Manage student registrations')

        # toolbar
        tb = tk.Frame(self.content_frame, bg=CONTENT_BG)
        tb.pack(fill='x', pady=(0, 12))

        self._toolbar_btn(tb, '+ Add Student', self.add_student).pack(side='left')

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
                tags=(s['id'],))

    def filter_students(self):
        q = self.student_search.get()
        for i in self.students_tree.get_children():
            self.students_tree.delete(i)
        rows = db.search_students(q) if q else db.get_all_students()
        for s in rows:
            self.students_tree.insert('', 'end',
                values=(s['admission_no'], s['name'], s['class'], s['gender']),
                tags=(s['id'],))

    def _student_ctx(self, event):
        item = self.students_tree.identify('item', event.x, event.y)
        if item:
            self.students_tree.selection_set(item)
            m = tk.Menu(self.root, tearoff=0)
            m.add_command(label='Edit',   command=self.edit_student)
            m.add_command(label='Delete', command=self.delete_student)
            m.post(event.x_root, event.y_root)

    def _student_dialog(self, title, adm='', name='', cls=CLASSES[0], gender='Male',
                        on_save=None):
        dlg = tk.Toplevel(self.root)
        dlg.title(title)
        dlg.geometry('400x370')
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
        cls_cb = ttk.Combobox(card, values=CLASSES, state='readonly', style='App.TCombobox')
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
            on_save(adm_e.get().strip(), name_e.get().strip(), cls_cb.get(), gen_cb.get())
            dlg.destroy()

        save = tk.Label(btn_row, text='Save', bg=BLUE, fg='white',
                        font=(FF, 10, 'bold'), padx=20, pady=8, cursor='hand2')
        save.pack(side='left')
        save.bind('<Button-1>', lambda e: do_save())

    def add_student(self):
        def on_save(adm, name, cls, gender):
            db.add_student(name, cls, gender, adm)
            self.load_students()
        self._student_dialog('Add Student', on_save=on_save)

    def edit_student(self):
        sel = self.students_tree.selection()
        if not sel: return
        item = self.students_tree.item(sel[0])
        adm, name, cls, gender = item['values']
        sid = item['tags'][0]
        def on_save(a, n, c, g):
            db.update_student(sid, n, c, g, a)
            self.load_students()
        self._student_dialog('Edit Student', adm, name, cls, gender, on_save=on_save)

    def delete_student(self):
        sel = self.students_tree.selection()
        if not sel: return
        sid = self.students_tree.item(sel[0])['tags'][0]
        if messagebox.askyesno('Confirm Delete', 'Delete this student?', parent=self.root):
            db.delete_student(sid)
            self.load_students()

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
        self.marks_class_cb = ttk.Combobox(ctrl, values=CLASSES, state='readonly',
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

        # table card
        tc_outer = tk.Frame(self.content_frame, bg=BORDER_CLR)
        tc_outer.pack(fill='both', expand=True, pady=4)
        tc = tk.Frame(tc_outer, bg=CARD_BG)
        tc.pack(fill='both', expand=True, padx=1, pady=1)

        cols = ['student'] + SUBJECTS
        self.marks_tree = ttk.Treeview(tc, columns=cols, show='headings',
                                       style='App.Treeview')
        self.marks_tree.heading('student', text='Student')
        self.marks_tree.column('student', width=180, anchor='w')
        for s in SUBJECTS:
            self.marks_tree.heading(s, text=s)
            self.marks_tree.column(s, width=72, anchor='center')

        sb = ttk.Scrollbar(tc, orient='vertical', command=self.marks_tree.yview,
                           style='App.Vertical.TScrollbar')
        self.marks_tree.configure(yscrollcommand=sb.set)
        sb.pack(side='right', fill='y')
        self.marks_tree.pack(fill='both', expand=True)

        self.marks_tree.bind('<Double-Button-1>', lambda e: self._edit_student_marks())
        self._load_marks_table()

    def _load_marks_table(self):
        for i in self.marks_tree.get_children():
            self.marks_tree.delete(i)
        cls  = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        for s in db.get_students_by_class(cls):
            m = db.get_student_marks(s['id'], term)
            vals = [s['name']] + [m.get(sub, '') for sub in SUBJECTS]
            self.marks_tree.insert('', 'end', values=vals, tags=(s['id'],))

    def _edit_student_marks(self):
        sel = self.marks_tree.selection()
        if not sel: return
        item  = self.marks_tree.item(sel[0])
        sid   = item['tags'][0]
        name  = item['values'][0]
        term  = self.marks_term_cb.get()
        cur   = db.get_student_marks(sid, term)

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
        for i, sub in enumerate(SUBJECTS):
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
        cls  = self.marks_class_cb.get()
        term = self.marks_term_cb.get()
        students = db.get_students_by_class(cls)
        for s, iid in zip(students, self.marks_tree.get_children()):
            vals = self.marks_tree.item(iid)['values']
            marks = {}
            for j, sub in enumerate(SUBJECTS):
                try:
                    v = vals[j + 1]
                    if v != '':
                        marks[sub] = min(100, max(0, int(v)))
                except (ValueError, IndexError):
                    pass
            db.save_student_marks(s['id'], marks, term)
        messagebox.showinfo('Success', 'Marks saved successfully!')

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
        self.rep_cls_cb = ttk.Combobox(ctrl, values=['All'] + CLASSES, state='readonly',
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

        cols = ['pos', 'name', 'class'] + SUBJECTS + ['total', 'avg', 'grade']
        self.rep_tree = ttk.Treeview(tc, columns=cols, show='headings', style='App.Treeview')

        spec = [('pos','Pos',45,'center'), ('name','Name',170,'w'), ('class','Class',90,'center')]
        for col, txt, w, anchor in spec:
            self.rep_tree.heading(col, text=txt)
            self.rep_tree.column(col, width=w, anchor=anchor)
        for s in SUBJECTS:
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
        subj_totals = {s: [] for s in SUBJECTS}
        for r in results:
            for s in SUBJECTS:
                if r['marks'].get(s): subj_totals[s].append(r['marks'][s])

        for s in SUBJECTS:
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
            for s in SUBJECTS: vals.append(r['marks'].get(s, '—'))
            vals += [r['total'], r['average'], r['grade']]
            self.rep_tree.insert('', 'end', values=vals)

    def export_csv(self):
        cls  = self.rep_cls_cb.get()
        term = self.rep_term_cb.get()
        results = db.calculate_results(cls, term)
        fn = f"report_{cls.replace(' ', '_')}_term_{term}.csv"
        with open(fn, 'w', newline='') as f:
            w = csv.writer(f)
            w.writerow(['Position', 'Adm No', 'Name', 'Class'] + SUBJECTS + ['Total', 'Average', 'Grade'])
            for r in results:
                row = [r['position'], r['student']['admission_no'],
                       r['student']['name'], r['student']['class']]
                row += [r['marks'].get(s, 0) for s in SUBJECTS]
                row += [r['total'], r['average'], r['grade']]
                w.writerow(row)
        messagebox.showinfo('Exported', f'Report saved to {fn}')

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
        self.ch_cls_cb = ttk.Combobox(ctrl, values=['All'] + CLASSES, state='readonly',
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

        self.fig, self.axes = plt.subplots(2, 2, figsize=(12, 7))
        self.fig.set_facecolor(CARD_BG)
        self.fig.tight_layout(pad=3.0)
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

        subj_totals = {s: [] for s in SUBJECTS}
        for r in results:
            for s in SUBJECTS:
                if r['marks'].get(s): subj_totals[s].append(r['marks'][s])

        avgs   = [round(sum(subj_totals[s]) / len(subj_totals[s]), 1) if subj_totals[s] else 0 for s in SUBJECTS]
        colors = [GRADE_COLORS['EE'] if a >= 80 else GRADE_COLORS['ME'] if a >= 70 else
                  GRADE_COLORS['AE'] if a >= 60 else GRADE_COLORS['BE'] if a >= 50 else
                  GRADE_COLORS['IE'] for a in avgs]

        self.axes[0, 0].bar(SUBJECTS, avgs, color=colors, edgecolor='none', width=0.6)
        self.axes[0, 0].set_title('Subject Averages', fontweight='bold', color=TEXT_PRIMARY, pad=10)
        self.axes[0, 0].set_ylim(0, 100)
        self.axes[0, 0].set_ylabel('Average Marks', color=TEXT_SECONDARY)
        self.axes[0, 0].tick_params(axis='x', rotation=45, labelcolor=TEXT_SECONDARY)
        self.axes[0, 0].tick_params(axis='y', labelcolor=TEXT_SECONDARY)
        self.axes[0, 0].spines['top'].set_visible(False)
        self.axes[0, 0].spines['right'].set_visible(False)

        grade_counts = {g: 0 for g in GRADE_COLORS}
        for r in results: grade_counts[r['grade']] += 1
        grades = [g for g, v in grade_counts.items() if v > 0]
        counts = [grade_counts[g] for g in grades]
        if counts:
            pie_colors = [GRADE_COLORS[g] for g in grades]
            wedges, texts, autotexts = self.axes[0, 1].pie(
                counts, labels=grades, autopct='%1.1f%%',
                colors=pie_colors, startangle=90, pctdistance=0.75)
            for t in autotexts: t.set_color('white')
            self.axes[0, 1].set_title('Grade Distribution', fontweight='bold', color=TEXT_PRIMARY, pad=10)

        top5 = results[:5]
        if top5:
            names = [r['student']['name'].split()[0] for r in top5]
            avgs5 = [r['average'] for r in top5]
            top_colors = [BLUE, GREEN, ORANGE, PURPLE, GRADE_COLORS['IE']][:len(top5)]
            bars = self.axes[1, 0].barh(names, avgs5, color=top_colors, edgecolor='none', height=0.5)
            self.axes[1, 0].set_title('Top Students', fontweight='bold', color=TEXT_PRIMARY, pad=10)
            self.axes[1, 0].set_xlim(0, 100)
            self.axes[1, 0].set_xlabel('Average Marks', color=TEXT_SECONDARY)
            self.axes[1, 0].tick_params(labelcolor=TEXT_SECONDARY)
            self.axes[1, 0].spines['top'].set_visible(False)
            self.axes[1, 0].spines['right'].set_visible(False)

        if cls == 'All':
            cls_perf = {}
            for c in CLASSES:
                cr = db.calculate_results(c, term)
                cls_perf[c] = round(sum(r['average'] for r in cr) / len(cr), 1) if cr else 0
            self.axes[1, 1].bar(list(cls_perf.keys()), list(cls_perf.values()),
                                color=BLUE, edgecolor='none', width=0.5)
            self.axes[1, 1].set_title('Class Performance', fontweight='bold', color=TEXT_PRIMARY, pad=10)
            self.axes[1, 1].set_ylim(0, 100)
            self.axes[1, 1].set_ylabel('Average Marks', color=TEXT_SECONDARY)
            self.axes[1, 1].tick_params(labelcolor=TEXT_SECONDARY)
            self.axes[1, 1].spines['top'].set_visible(False)
            self.axes[1, 1].spines['right'].set_visible(False)
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
        self.rc_cls_cb = ttk.Combobox(ctrl, values=CLASSES, state='readonly',
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

        SCH_BLUE = '#1565C0'
        RED_ACC  = '#c62828'
        TTL_ORG  = '#e65100'
        MINT_BG  = '#e8f5e9'
        HDR_BG   = '#e3f2fd'
        PERF_CLR = {'EE': '#1b5e20', 'ME': '#0d47a1',
                    'AE': '#bf360c', 'BE': '#bf360c', 'IE': '#b71c1c'}

        def get_grade(m):
            return ('EE' if m >= 80 else 'ME' if m >= 70 else
                    'AE' if m >= 60 else 'BE' if m >= 50 else 'IE')

        # School Header
        tk.Label(parent, text='VISION PRIMARY AND JUNIOR SCHOOLS',
                 bg='white', fg=SCH_BLUE, font=(FF, 15, 'bold')).pack()
        tk.Label(parent, text='P.O. BOX 54 KENYA',
                 bg='white', fg='#666', font=(FF, 9)).pack()
        tk.Label(parent, text='visionprimaryschool@gmail.com',
                 bg='white', fg=SCH_BLUE, font=(FF, 9)).pack()
        tk.Label(parent, text='+254718481515/+254718481515',
                 bg='white', fg='#666', font=(FF, 9)).pack()
        tk.Label(parent, text='Education For Excellence',
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
        for i, subj in enumerate(SUBJECTS):
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
        possible    = len(SUBJECTS) * 100
        base        = len(SUBJECTS) + 1
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
        mark_vals  = [marks.get(sub, 0) for sub in SUBJECTS]
        bar_colors = [GRADE_COLORS[get_grade(m)] for m in mark_vals]
        ax.bar(SUBJECTS, mark_vals, color=bar_colors, edgecolor='none', width=0.55)
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
        possible = len(SUBJECTS) * 100
        lines = [
            '=' * 62,
            '      VISION PRIMARY AND JUNIOR SCHOOLS',
            '          LEARNER ASSESSMENT REPORT CARD',
            '=' * 62,
            f"  Name    : {s['name']:<20}  Grade  : {s.get('class','').replace('Grade ','')}",
            f"  Stream  : {s['admission_no']:<20}  Year   : 2026",
            f"  Term    : {term:<20}  Gender : {s['gender']}",
            '-' * 62,
            f"  {'Subject':<16} {'Marks':>6}  {'Avg':>6}  Performance Level",
            '-' * 62,
        ]
        for sub in SUBJECTS:
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
    
    app  = SchoolReportApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()
