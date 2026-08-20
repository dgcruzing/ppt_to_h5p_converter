import os
import tkinter as tk
from tkinter import filedialog, ttk
import tkinter.font as tkfont


class ConverterView(tk.Tk):

    def __init__(self):
        super().__init__()
        self.title("PowerPoint to H5P Batch Converter")
        self.geometry("750x600")
        self.minsize(650, 500)

        self._log_visible = False

        self._configure_styles()
        self._build_ui()

    def _configure_styles(self):
        self.style = ttk.Style()
        self.style.theme_use("clam")
        self.style.configure("Treeview", rowheight=25)
        self.style.configure("Treeview.Heading", font=("Helvetica", 9, "bold"))
        self.style.configure("TButton", padding=4)

    def _build_ui(self):
        self.columnconfigure(0, weight=1)
        self.rowconfigure(1, weight=1)  # Queue table expands

        self._build_toolbar()
        self._build_queue_table()
        self._build_settings_panel()
        self._build_execution_bar()
        self._build_log_drawer()

    # --- UI Section Builders ---

    def _build_toolbar(self):
        toolbar = ttk.Frame(self, padding=(10, 8, 10, 0))
        toolbar.grid(row=0, column=0, sticky="ew")

        self.add_btn = ttk.Button(
            toolbar, text="+ Add PPTX Files", command=self._on_add_files
        )
        self.add_btn.pack(side="left", padx=(0, 5))

        self.remove_btn = ttk.Button(
            toolbar, text="- Remove Selected", command=self._on_remove_selected
        )
        self.remove_btn.pack(side="left", padx=5)

        self.clear_btn = ttk.Button(
            toolbar, text="Clear Queue", command=self._on_clear_queue
        )
        self.clear_btn.pack(side="left", padx=5)

    def _build_queue_table(self):
        frame = ttk.Frame(self, padding=10)
        frame.grid(row=1, column=0, sticky="nsew")
        frame.columnconfigure(0, weight=1)
        frame.rowconfigure(0, weight=1)

        columns = ("filename", "resolution", "output_path", "status")
        self.queue_tree = ttk.Treeview(
            frame, columns=columns, show="headings", selectmode="extended"
        )

        self.queue_tree.heading("filename", text="Source File")
        self.queue_tree.heading("resolution", text="Resolution")
        self.queue_tree.heading("output_path", text="Output Path")
        self.queue_tree.heading("status", text="Status")

        self.queue_tree.column(
            "filename", width=220, minwidth=150, anchor="w", stretch=False
        )
        self.queue_tree.column(
            "resolution", width=100, minwidth=80, anchor="center", stretch=False
        )
        self.queue_tree.column(
            "output_path", width=350, minwidth=200, anchor="w", stretch=False
        )
        self.queue_tree.column(
            "status", width=120, minwidth=100, anchor="center", stretch=False
        )

        v_scrollbar = ttk.Scrollbar(
            frame, orient="vertical", command=self.queue_tree.yview
        )
        self.queue_tree.configure(yscrollcommand=v_scrollbar.set)

        h_scrollbar = ttk.Scrollbar(
        frame, orient="horizontal", command=self.queue_tree.xview
        )
        self.queue_tree.configure(xscrollcommand=h_scrollbar.set)

        self.queue_tree.grid(row=0, column=0, sticky="nsew")
        v_scrollbar.grid(row=0, column=1, sticky="ns")
        h_scrollbar.grid(row=1, column=0, sticky="ew")

        self.queue_tree.bind("<Double-1>", self._on_tree_double_click)

    def _autofit_column(self, col_id: str):
        """Adjusts a single Treeview column's width based on its longest text entry or header.

            col_id: The column identifier returned by identify_column (e.g. '#1',
            '#2') or column name ('filename').
        """
        font = tkfont.Font()
        padding = 2

        # Resolve column name from symbolic ID (e.g. '#1' -> 'filename')
        if col_id.startswith("#"):
            col_index = int(col_id.replace("#", "")) - 1
            if col_index < 0 or col_index >= len(self.queue_tree["columns"]):
                return
            col_name = self.queue_tree["columns"][col_index]
        else:
            col_name = col_id
            col_index = list(self.queue_tree["columns"]).index(col_name)

        # Start with header text measurement
        header_text = self.queue_tree.heading(col_name, "text")
        max_width = font.measure(header_text)

        # Compare against cell values in that specific column
        for item_id in self.queue_tree.get_children():
            values = self.queue_tree.item(item_id, "values")
            if col_index < len(values):
                cell_text = str(values[col_index])
                cell_width = font.measure(cell_text)
                if cell_width > max_width:
                    max_width = cell_width

        # Apply updated width only to the target column
        self.queue_tree.column(col_name, width=max_width)

    def _on_tree_double_click(self, event):
        """Event handler for double-clicking Treeview column headers or separators."""
        region = self.queue_tree.identify_region(event.x, event.y)

        if region in ("separator", "heading"):
            # Identify the exact column clicked (e.g. '#1', '#2', etc.)
            col_id = self.queue_tree.identify_column(event.x)
            if col_id:
                self._autofit_column(col_id)

    def _build_settings_panel(self):
        self.output_dir_var = tk.StringVar()
        self.output_dir_var.trace_add("write", self._on_output_dir_changed)


        panel = ttk.LabelFrame(
            self, text=" Export Settings ", padding=(10, 8)
        )
        panel.grid(row=2, column=0, sticky="ew", padx=0, pady=(0, 0))
        panel.columnconfigure(1, weight=1)

        # Default Output Folder
        ttk.Label(panel, text="Output Directory:").grid(
            row=0, column=0, sticky="w", padx=(0, 5)
        )
        self.output_dir_entry = ttk.Entry(panel)
        self.output_dir_entry.grid(
            row=0, column=1, sticky="ew", padx=5, pady=2
        )
        ttk.Button(
            panel, text="Browse...", command=self._on_browse_output_dir
        ).grid(row=0, column=2, padx=(5, 0))

        # Global Resolution
        ttk.Label(panel, text="Default Resolution:").grid(
            row=1, column=0, sticky="w", padx=(0, 5), pady=(5, 0)
        )
        self.resolution_var = tk.StringVar(value="1280x720")
        res_combo = ttk.Combobox(
            panel,
            textvariable=self.resolution_var,
            values=["1280x720", "1920x1080"],
            state="readonly",
            width=12,
        )
        res_combo.grid(
            row=1, column=1, sticky="w", padx=5, pady=(5, 0)
        )
        self.output_dir_entry = ttk.Entry(panel, textvariable=self.output_dir_var)

    def _build_execution_bar(self):
        bar = ttk.Frame(self, padding=10)
        bar.grid(row=3, column=0, sticky="ew")
        bar.columnconfigure(0, weight=1)

        # Progress and Status Text
        self.status_label = ttk.Label(
            bar, text="Ready", font=("Helvetica", 9, "italic")
        )
        self.status_label.grid(row=0, column=0, sticky="w", pady=(0, 4))

        self.progress_bar = ttk.Progressbar(
            bar, orient="horizontal", mode="determinate"
        )
        self.progress_bar.grid(
            row=1, column=0, sticky="ew", padx=(0, 10)
        )

        # Action Buttons Frame
        btn_frame = ttk.Frame(bar)
        btn_frame.grid(row=1, column=1, sticky="e")

        self.convert_btn = ttk.Button(
            btn_frame,
            text="Convert Queue",
            command=self._on_convert_clicked,
            width=15,
        )
        self.convert_btn.pack(side="left", padx=2)

        self.cancel_btn = ttk.Button(
            btn_frame, text="Cancel", state="disabled", width=10
        )
        self.cancel_btn.pack(side="left", padx=2)

    def _build_log_drawer(self):
        self.log_frame = ttk.Frame(self, padding=(10, 0, 10, 10))
        self.log_frame.grid(row=4, column=0, sticky="ew")
        self.log_frame.columnconfigure(0, weight=1)

        self.toggle_log_btn = ttk.Button(
            self.log_frame,
            text="▲ Show Activity Log",
            command=self._toggle_log_drawer,
        )
        self.toggle_log_btn.grid(row=0, column=0, sticky="w", pady=(0, 5))

        # Log content (Hidden by default)
        self.log_content_frame = ttk.Frame(self.log_frame)

        self.log_text = tk.Text(
            self.log_content_frame, height=6, wrap="word", font=("Consolas", 8)
        )
        log_scroll = ttk.Scrollbar(
            self.log_content_frame,
            orient="vertical",
            command=self.log_text.yview,
        )
        self.log_text.configure(yscrollcommand=log_scroll.set)

        self.log_text.pack(side="left", fill="both", expand=True)
        log_scroll.pack(side="right", fill="y")

    # --- Internal Event Handlers ---
    def _on_output_dir_changed(self, *args):
        new_dir = self.output_dir_var.get().strip()
        for item_id in self.queue_tree.get_children():
            values = list(self.queue_tree.item(item_id, "values"))
            filename = values[0]
            out_name = os.path.splitext(filename)[0] + ".h5p"
            
            # Recalculate output path
            values[2] = os.path.join(new_dir, out_name) if new_dir else ""
            self.queue_tree.item(item_id, values=values)

    def _on_add_files(self):
        files = filedialog.askopenfilenames(
            title="Select PowerPoint Presentations",
            filetypes=[("PowerPoint Files", "*.pptx")],
        )
        if files:
            default_dir = self.output_dir_entry.get().strip()
            res = self.resolution_var.get()
            for filepath in files:
                filename = os.path.basename(filepath)
                out_name = os.path.splitext(filename)[0] + ".h5p"
                
                # Use global output dir if specified, otherwise default to source file's directory
                if default_dir:
                    out_path = os.path.join(default_dir, out_name)
                else:
                    source_dir = os.path.dirname(filepath)
                    out_path = os.path.join(source_dir, out_name)

                self.queue_tree.insert(
                    "",
                    "end",
                    values=(filename, res, out_path, "Pending"),
                    tags=(filepath,),
                )

    def _on_remove_selected(self):
        selected_items = self.queue_tree.selection()
        for item in selected_items:
            self.queue_tree.delete(item)

    def _on_clear_queue(self):
        for item in self.queue_tree.get_children():
            self.queue_tree.delete(item)

    def _on_browse_output_dir(self):
        directory = filedialog.askdirectory(title="Select Output Directory")
        if directory:
            self.output_dir_entry.delete(0, tk.END)
            self.output_dir_entry.insert(0, directory)

    def _toggle_log_drawer(self):
        if self._log_visible:
            self.log_content_frame.grid_forget()
            self.toggle_log_btn.config(text="▲ Show Activity Log")
            self._log_visible = False
        else:
            self.log_content_frame.grid(
                row=1, column=0, sticky="ew", pady=(2, 0)
            )
            self.toggle_log_btn.config(text="▼ Hide Activity Log")
            self._log_visible = True

    def _on_convert_clicked(self):
        # Stub for action handling — overridden or bound by Controller
        pass

    # --- Public View API (For Controller / Presenter Integration) ---

    def get_queue_data(self):
        """Returns structured data of all items in the queue."""
        items = []
        for item_id in self.queue_tree.get_children():
            values = self.queue_tree.item(item_id, "values")
            full_path = self.queue_tree.item(item_id, "tags")[0]
            items.append(
                {
                    "item_id": item_id,
                    "source_path": full_path,
                    "resolution": values[1],
                    "output_path": values[2],
                    "status": values[3],
                }
            )
        return items

    def update_item_status(self, item_id: str, status: str):
        """Updates the status column for a specific item (e.g. 'Converting...', 'Completed', 'Error')."""
        current_values = list(self.queue_tree.item(item_id, "values"))
        current_values[3] = status
        self.queue_tree.item(item_id, values=current_values)

    def update_progress(self, current: int, total: int, status_text: str = ""):
        """Updates the global progress bar and execution status text."""
        self.progress_bar["maximum"] = total
        self.progress_bar["value"] = current
        if status_text:
            self.status_label.config(text=status_text)

    def append_log(self, message: str):
        """Appends a new entry to the collapsible activity log text area."""
        self.log_text.config(state="normal")
        self.log_text.insert("end", f"{message}\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")

    def set_processing_state(self, is_processing: bool):
        """Enables/disables UI elements based on current operational state."""
        state = "disabled" if is_processing else "normal"
        self.add_btn.config(state=state)
        self.remove_btn.config(state=state)
        self.clear_btn.config(state=state)
        self.convert_btn.config(state=state)
        self.cancel_btn.config(
            state="normal" if is_processing else "disabled"
        )


if __name__ == "__main__":
    app = ConverterView()
    app.mainloop()