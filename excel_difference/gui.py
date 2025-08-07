#!/usr/bin/env python3
"""
GUI for Excel difference generator.
"""

import io
import sys
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext, ttk

from excel_difference.excel_diff import excel_diff


class ConsoleRedirector:
    """Redirect console output to a text widget."""

    def __init__(self, text_widget):
        self.text_widget = text_widget
        self.buffer = io.StringIO()

    def write(self, text):
        self.buffer.write(text)
        self.text_widget.insert(tk.END, text)
        self.text_widget.see(tk.END)
        self.text_widget.update_idletasks()

    def flush(self):
        pass


class ExcelDiffGUI:
    """Main GUI class for Excel difference generator."""

    def __init__(self, root):
        self.root = root
        self.root.title("Excel Difference Generator")
        self.root.geometry("800x600")
        self.root.minsize(600, 400)

        # Variables
        self.file1_path = tk.StringVar()
        self.file2_path = tk.StringVar()
        self.output_path = tk.StringVar()
        self.key_column = tk.IntVar(value=1)

        self.setup_ui()

    def setup_ui(self):
        """Setup the user interface."""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(4, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame, text="Excel Difference Generator", font=("Arial", 16, "bold")
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # File 1 selection
        ttk.Label(main_frame, text="First Excel File:").grid(
            row=1, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.file1_path, width=50).grid(
            row=1, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5
        )
        ttk.Button(main_frame, text="Browse", command=self.browse_file1).grid(
            row=1, column=2, pady=5
        )

        # File 2 selection
        ttk.Label(main_frame, text="Second Excel File:").grid(
            row=2, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.file2_path, width=50).grid(
            row=2, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5
        )
        ttk.Button(main_frame, text="Browse", command=self.browse_file2).grid(
            row=2, column=2, pady=5
        )

        # Output file selection
        ttk.Label(main_frame, text="Output File:").grid(
            row=3, column=0, sticky=tk.W, pady=5
        )
        ttk.Entry(main_frame, textvariable=self.output_path, width=50).grid(
            row=3, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5
        )
        ttk.Button(main_frame, text="Browse", command=self.browse_output).grid(
            row=3, column=2, pady=5
        )

        # Key column selection
        ttk.Label(main_frame, text="Key Column:").grid(
            row=4, column=0, sticky=tk.W, pady=5
        )
        key_frame = ttk.Frame(main_frame)
        key_frame.grid(row=4, column=1, sticky=(tk.W, tk.E), padx=(5, 5), pady=5)

        ttk.Label(key_frame, text="Column to match rows (1-based):").pack(side=tk.LEFT)
        ttk.Spinbox(
            key_frame, from_=1, to=100, textvariable=self.key_column, width=10
        ).pack(side=tk.LEFT, padx=(5, 0))

        # Run button
        self.run_button = ttk.Button(
            main_frame,
            text="Generate Difference",
            command=self.run_difference,
            style="Accent.TButton",
        )
        self.run_button.grid(row=5, column=0, columnspan=3, pady=20)

        # Console output
        ttk.Label(main_frame, text="Console Output:").grid(
            row=6, column=0, sticky=tk.W, pady=(20, 5)
        )

        # Console text area
        self.console_text = scrolledtext.ScrolledText(main_frame, height=15, width=80)
        self.console_text.grid(
            row=7, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5
        )

        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(
            main_frame, textvariable=self.status_var, relief=tk.SUNKEN
        )
        status_bar.grid(
            row=8, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0)
        )

    def browse_file1(self):
        """Browse for first Excel file."""
        filename = filedialog.askopenfilename(
            title="Select First Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.file1_path.set(filename)

    def browse_file2(self):
        """Browse for second Excel file."""
        filename = filedialog.askopenfilename(
            title="Select Second Excel File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
        )
        if filename:
            self.file2_path.set(filename)

    def browse_output(self):
        """Browse for output file."""
        filename = filedialog.asksaveasfilename(
            title="Save Output File",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
        )
        if filename:
            self.output_path.set(filename)

    def run_difference(self):
        """Run the difference generation in a separate thread."""
        # Validate inputs
        if not self.file1_path.get():
            messagebox.showerror("Error", "Please select the first Excel file.")
            return

        if not self.file2_path.get():
            messagebox.showerror("Error", "Please select the second Excel file.")
            return

        if not self.output_path.get():
            messagebox.showerror("Error", "Please select an output file.")
            return

        # Check if input files exist
        if not Path(self.file1_path.get()).exists():
            messagebox.showerror(
                "Error", f"File '{self.file1_path.get()}' does not exist."
            )
            return

        if not Path(self.file2_path.get()).exists():
            messagebox.showerror(
                "Error", f"File '{self.file2_path.get()}' does not exist."
            )
            return

        # Disable run button and update status
        self.run_button.config(state="disabled")
        self.status_var.set("Processing...")

        # Clear console
        self.console_text.delete(1.0, tk.END)

        # Redirect console output
        old_stdout = sys.stdout
        old_stderr = sys.stderr
        console_redirector = ConsoleRedirector(self.console_text)
        sys.stdout = console_redirector
        sys.stderr = console_redirector

        # Run in separate thread
        thread = threading.Thread(target=self._run_difference_thread)
        thread.daemon = True
        thread.start()

        # Restore console output
        sys.stdout = old_stdout
        sys.stderr = old_stderr

    def _run_difference_thread(self):
        """Run difference generation in background thread."""
        try:
            print("Starting Excel difference generation...")
            print(f"File 1: {self.file1_path.get()}")
            print(f"File 2: {self.file2_path.get()}")
            print(f"Output: {self.output_path.get()}")
            print("-" * 50)

            excel_diff(
                self.file1_path.get(),
                self.file2_path.get(),
                self.output_path.get(),
                self.key_column.get(),
            )

            print("Successfully generated difference file!")
            self.root.after(0, lambda: self.status_var.set("Completed successfully"))
            self.root.after(
                0,
                lambda: messagebox.showinfo(
                    "Success",
                    f"Successfully generated difference file:\n"
                    f"{self.output_path.get()}",
                ),
            )

        except Exception as e:
            error_msg = f"Error: {str(e)}"
            print(error_msg)
            self.root.after(0, lambda: self.status_var.set("Error occurred"))
            self.root.after(0, lambda: messagebox.showerror("Error", error_msg))

        finally:
            # Re-enable run button
            self.root.after(0, lambda: self.run_button.config(state="normal"))


def main():
    """Main function to run the GUI."""
    root = tk.Tk()
    ExcelDiffGUI(root)
    root.mainloop()


if __name__ == "__main__":
    main()
