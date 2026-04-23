import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from database import db
from datetime import datetime


class StudentsTab:
    def __init__(self, parent, app):
        self.parent = parent
        self.app = app
        self.students_table = None
        self.build_ui()

    def build_ui(self):
        # Filters row
        filters = tk.Frame(self.parent, bg=self.app.CONTENT_BG, padx=12, pady=8)
        filters.pack(fill="x", pady=(0, 10))

        tk.Label(
            filters,
            text="Class:",
            bg=self.app.CONTENT_BG,
            fg=self.app.TEXT_SECONDARY,
            font=(self.app.FF, 10),
        ).pack(side="left", padx=(0, 4))
        self.class_var = tk.StringVar()
        self.class_cb = ttk.Combobox(
            filters,
            textvariable=self.class_var,
            style="App.TCombobox",
            width=20,
            state="readonly",
        )
        self.class_cb.pack(side="left", padx=(0, 12))
        self.class_cb.bind("<<ComboboxSelected>>", self.on_filter_change)

        # Load class options
        class_names = [
            row.get("name", "") for row in db.get_all_classes() if row.get("name")
        ]
        self.class_cb["values"] = ["All Classes"] + class_names
        self.class_var.set("All Classes")

        tk.Label(
            filters,
            text="Stream:",
            bg=self.app.CONTENT_BG,
            fg=self.app.TEXT_SECONDARY,
            font=(self.app.FF, 10),
        ).pack(side="left", padx=(0, 4))
        self.stream_var = tk.StringVar(value="All Streams")
        self.stream_cb = ttk.Combobox(
            filters,
            textvariable=self.stream_var,
            style="App.TCombobox",
            width=18,
            state="readonly",
        )
        self.stream_cb.pack(side="left", padx=(0, 12))
        self.stream_cb.bind("<<ComboboxSelected>>", self.on_filter_change)

        # Action buttons
        btn_frame = tk.Frame(self.parent, bg=self.app.CONTENT_BG, padx=12, pady=8)
        btn_frame.pack(fill="x", pady=(0, 8))

        self.print_btn = tk.Button(
            btn_frame,
            text="🖨️ Print Student List PDF",
            bg=self.app.PURPLE,
            fg="white",
            font=(self.app.FF, 11, "bold"),
            padx=18,
            pady=8,
            command=self.print_student_list_pdf,
            cursor="hand2",
        )
        self.print_btn.pack(side="left")

        refresh_btn = tk.Button(
            btn_frame,
            text="🔄 Refresh",
            bg=self.app.BLUE,
            fg="white",
            font=(self.app.FF, 10),
            padx=16,
            pady=8,
            command=self.refresh_students,
            cursor="hand2",
        )
        refresh_btn.pack(side="left", padx=(12, 0))

        view_info = tk.Label(
            btn_frame,
            text="Class filter changes refresh the stream list and table.",
            bg=self.app.CONTENT_BG,
            fg=self.app.TEXT_SECONDARY,
            font=(self.app.FF, 9),
        )
        view_info.pack(side="right")

        # Students table
        self.students_frame = tk.Frame(
            self.parent, bg=self.app.CARD_BG, relief="sunken", bd=1
        )
        self.students_frame.pack(fill="both", expand=True, pady=8)

        columns = [
            {"key": "sno", "title": "#", "width": 50, "anchor": "center"},
            {
                "key": "admission_no",
                "title": "Adm No",
                "width": 110,
                "anchor": "center",
            },
            {"key": "name", "title": "Student Name", "width": 200, "anchor": "w"},
            {"key": "class", "title": "Class", "width": 110, "anchor": "center"},
            {"key": "stream", "title": "Stream", "width": 100, "anchor": "center"},
            {"key": "gender", "title": "Gnd", "width": 60, "anchor": "center"},
            {"key": "guardian_name", "title": "Guardian", "width": 150, "anchor": "w"},
            {
                "key": "parent_email",
                "title": "Contact Email",
                "width": 180,
                "anchor": "w",
            },
        ]
        self.students_table = self.app.AdvancedDataTable(
            self.students_frame,
            columns,
            page_size=25,
            search_label="Search students",
            enable_select_all=True,
        )

        self.refresh_students()
        self.refresh_streams()

    def refresh_students(self):
        """Load students for current filters with serial numbers."""
        class_name = self.class_var.get().strip()
        stream_name = self.stream_var.get().strip()

        # Determine which students to load
        if not class_name or class_name == "All Classes":
            students = db.get_all_students()
        elif stream_name and stream_name != "All Streams":
            students = db.get_students_by_class_and_stream(class_name, stream_name)
        else:
            students = db.get_students_by_class(class_name)

        rows = []
        for idx, student in enumerate(students, 1):
            gender_letter = (
                student.get("gender", "")[:1] if student.get("gender") else ""
            )
            row_data = {
                "iid": student.get("id", ""),
                "values": (
                    str(idx),  # Serial number
                    student.get("admission_no", ""),
                    student.get("name", ""),
                    student.get("class", ""),
                    student.get("stream", ""),
                    gender_letter,
                    student.get("guardian_name", "") or "",
                    student.get("parent_email", "") or "",
                ),
                "value_map": {
                    "sno": str(idx),
                    "admission_no": student.get("admission_no", ""),
                    "name": student.get("name", ""),
                    "class": student.get("class", ""),
                    "stream": student.get("stream", ""),
                    "gender": student.get("gender", ""),
                    "guardian_name": student.get("guardian_name", "") or "",
                    "parent_email": student.get("parent_email", "") or "",
                },
                "search": " ".join(
                    [
                        str(student.get("admission_no", "")),
                        student.get("name", ""),
                        student.get("class", ""),
                        student.get("stream", ""),
                        student.get("gender", ""),
                        student.get("guardian_name", "") or "",
                        student.get("parent_email", "") or "",
                    ]
                ),
            }
            rows.append(row_data)

        self.students_table.set_rows(rows)

    def refresh_streams(self):
        """Update stream combobox based on selected class."""
        class_name = self.class_var.get().strip()
        if class_name and class_name != "All Classes":
            class_row = db.get_class_by_name(class_name)
            streams = db.get_streams_for_class(class_row["id"]) if class_row else []
        else:
            streams = []

        stream_names = ["All Streams"] + [s["name"] for s in streams]
        self.stream_cb["values"] = stream_names
        if self.stream_var.get() not in stream_names:
            self.stream_var.set("All Streams")

    def on_filter_change(self, event=None):
        """Handle filter changes."""
        self.refresh_streams()
        self.refresh_students()

    def print_student_list_pdf(self):
        """Generate and save PDF student list."""
        class_name = self.class_var.get().strip()
        stream_name = self.stream_var.get().strip()
        if stream_name == "All Streams":
            stream_name = ""

        if not class_name or class_name == "All Classes":
            messagebox.showwarning("Warning", "Please select a class first.")
            return

        # Generate filename
        timestamp = datetime.now().strftime("%Y%m%d-%H%M")
        filename = f"Student_List_{class_name.replace(' ', '_')}"
        if stream_name:
            filename += f"_{stream_name}"
        filename += f"_{timestamp}.pdf"

        file_path = filedialog.asksaveasfilename(
            title="Save Student List PDF",
            defaultextension=".pdf",
            filetypes=[("PDF files", "*.pdf")],
            initialfile=filename,
        )
        if not file_path:
            return

        # Generate PDF
        self.app.generate_student_list_pdf(class_name, stream_name, file_path)
