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
        # Control panel styled like the Subjects page toolbar.
        controls = self.app._toolbar_panel(self.parent, pady=(0, 12), padx=10)
        btn_frame = self.app._toolbar_row(controls, pady=(0, 10))
        filters = self.app._toolbar_row(controls, pady=0)

        self.print_btn = self.app._toolbar_btn(
            btn_frame,
            "Print Student List PDF",
            self.print_student_list_pdf,
            bg=self.app.PURPLE,
        )
        self.print_btn.pack(side="left", padx=(4, 8))

        self.app._toolbar_btn(
            btn_frame,
            "+ Add Student",
            self.app.add_student,
            bg=self.app.GREEN,
        ).pack(side="left", padx=4)

        self.app._toolbar_btn(
            btn_frame,
            "Edit Selected",
            self.app.edit_student,
            bg=self.app.BLUE,
        ).pack(side="left", padx=4)

        self.app._toolbar_btn(
            btn_frame,
            "Delete Selected",
            self.app.delete_student,
            bg="#e74c3c",
        ).pack(side="left", padx=4)

        self.app._toolbar_btn(
            btn_frame,
            "Template",
            self.app.download_template,
            bg=self.app.ORANGE,
        ).pack(side="left", padx=4)

        self.app._toolbar_btn(
            btn_frame,
            "Import Excel",
            self.app.import_excel,
            bg="#7c3aed",
        ).pack(side="left", padx=4)

        self.app._toolbar_btn(
            btn_frame,
            "Export Excel",
            self.app.export_excel,
            bg="#0f766e",
        ).pack(side="left", padx=4)

        self.app._toolbar_btn(
            btn_frame,
            "Refresh",
            self.refresh_all,
            bg="#64748b",
        ).pack(side="left", padx=4)

        self.app._toolbar_label(filters, "Class:", left=4)
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

        self.app._toolbar_label(filters, "Import as:")
        self.import_class_var = tk.StringVar(value="Use Class From File")
        self.import_class_cb = ttk.Combobox(
            filters,
            textvariable=self.import_class_var,
            values=["Use Class From File"] + class_names,
            style="App.TCombobox",
            width=20,
            state="readonly",
        )
        self.import_class_cb.pack(side="left", padx=(0, 12))
        self.app.students_import_class_var = self.import_class_var

        self.app._toolbar_label(filters, "Stream:")
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

        view_info = tk.Label(
            filters,
            text="Use Class and Stream filters above to narrow the table.",
            bg=self.app.CONTENT_BG,
            fg=self.app.TEXT_SECONDARY,
            font=(self.app.FF, 9),
        )
        view_info.pack(side="left", padx=(8, 0))

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
        self.refresh_class_options()
        class_name = self.class_var.get().strip()
        stream_name = self.stream_var.get().strip()

        # Determine which students to load
        if not class_name or class_name == "All Classes":
            students = db.get_all_students()
        elif stream_name and stream_name != "All Streams":
            students = [
                student
                for student in db.get_students_by_class(class_name)
                if self._stream_matches(student.get("stream", ""), stream_name)
            ]
        else:
            students = db.get_students_by_class(class_name)

        if (
            (not class_name or class_name == "All Classes")
            and stream_name
            and stream_name != "All Streams"
        ):
            students = [
                student
                for student in students
                if self._stream_matches(student.get("stream", ""), stream_name)
            ]

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

    def _stream_matches(self, saved_stream, selected_stream):
        """Compare stream labels in a forgiving way for imported student data."""
        return self.app._normalize_key(saved_stream) == self.app._normalize_key(
            selected_stream
        )

    def refresh_class_options(self):
        """Refresh class choices without disturbing current filters."""
        class_names = [
            row.get("name", "") for row in db.get_all_classes() if row.get("name")
        ]
        class_values = ["All Classes"] + class_names
        import_values = ["Use Class From File"] + class_names
        self.class_cb["values"] = class_values
        self.import_class_cb["values"] = import_values
        if self.class_var.get() not in class_values:
            self.class_var.set("All Classes")
        if self.import_class_var.get() not in import_values:
            self.import_class_var.set("Use Class From File")

    def refresh_streams(self):
        """Update stream combobox based on selected class."""
        class_name = self.class_var.get().strip()
        current = self.stream_var.get().strip()
        stream_names = []

        def add_stream_name(value):
            name = str(value or "").strip()
            if not name:
                return
            existing_keys = {self.app._normalize_key(item) for item in stream_names}
            if self.app._normalize_key(name) not in existing_keys:
                stream_names.append(name)

        if class_name and class_name != "All Classes":
            class_row = db.get_class_by_name(class_name)
            streams = db.get_streams_for_class(class_row["id"]) if class_row else []
            for stream in streams:
                add_stream_name(stream.get("name", ""))
            for student in db.get_students_by_class(class_name):
                add_stream_name(student.get("stream", ""))
        else:
            for student in db.get_all_students():
                add_stream_name(student.get("stream", ""))

        values = ["All Streams"] + sorted(
            stream_names, key=lambda item: item.strip().lower()
        )
        self.stream_cb["values"] = values
        current_key = self.app._normalize_key(current)
        matched_current = next(
            (
                value
                for value in values
                if self.app._normalize_key(value) == current_key
            ),
            "",
        )
        if matched_current:
            self.stream_var.set(matched_current)
        else:
            self.stream_var.set("All Streams")

    def on_filter_change(self, event=None):
        """Handle filter changes."""
        if event is None or event.widget == self.class_cb:
            self.refresh_streams()
        self.refresh_students()

    def refresh_all(self):
        """Refresh filters and table together."""
        self.refresh_class_options()
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
