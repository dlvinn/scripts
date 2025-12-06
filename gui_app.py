#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Arabic Document Tools GUI
A graphical user interface for the Arabic Document Tools suite.
"""

import sys
import os
import threading
import queue
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from pathlib import Path
from datetime import datetime

# Import tool classes
# We need to add the current directory to sys.path to ensure imports work
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

try:
    from docx_format_fixer import ArabicDocxFixer
    from docx_generator import ArabicDocumentGenerator
    from odf_to_docx_converter import ODFToDocxConverter
    from discover_mojibake import scan_folder as scan_mojibake
    from add_date_to_footer import ODFFooterModifier
    from replace_with_fixed import FixedFileReplacer
except ImportError as e:
    print(f"Error importing modules: {e}")
    sys.exit(1)


class TextRedirector(object):
    """Redirects stdout to a tkinter Text widget"""
    def __init__(self, widget, queue):
        self.widget = widget
        self.queue = queue

    def write(self, str):
        self.queue.put(str)

    def flush(self):
        pass


class ArabicToolsApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Arabic Document Tools")
        self.root.geometry("900x700")

        # Configure style
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Create tabs
        self.tab_control = ttk.Notebook(root)

        self.tab_fixer = ttk.Frame(self.tab_control)
        self.tab_generator = ttk.Frame(self.tab_control)
        self.tab_converter = ttk.Frame(self.tab_control)
        self.tab_mojibake = ttk.Frame(self.tab_control)
        self.tab_footer = ttk.Frame(self.tab_control)
        self.tab_tools = ttk.Frame(self.tab_control)

        self.tab_control.add(self.tab_fixer, text='Format Fixer')
        self.tab_control.add(self.tab_generator, text='Doc Generator')
        self.tab_control.add(self.tab_converter, text='ODF Converter')
        self.tab_control.add(self.tab_mojibake, text='Mojibake Scanner')
        self.tab_control.add(self.tab_footer, text='Footer Modifier')
        self.tab_control.add(self.tab_tools, text='Tools')

        self.tab_control.pack(expand=1, fill="both")

        # Setup the tabs
        self.setup_fixer_tab()
        self.setup_generator_tab()
        self.setup_converter_tab()
        self.setup_mojibake_tab()
        self.setup_footer_tab()
        self.setup_tools_tab()

        # Console Output
        self.create_console_output()

        # Queue for thread communication
        self.log_queue = queue.Queue()

        # Redirect stdout
        self.original_stdout = sys.stdout
        sys.stdout = TextRedirector(self.console, self.log_queue)

        # Start checking the queue
        self.root.after(100, self.check_queue)

    def create_console_output(self):
        frame = ttk.LabelFrame(self.root, text="Log Output")
        frame.pack(fill="both", expand=True, padx=10, pady=10)

        self.console = scrolledtext.ScrolledText(frame, height=12, state='disabled')
        self.console.pack(fill="both", expand=True, padx=5, pady=5)

        # Clear log button
        ttk.Button(frame, text="Clear Log", command=self.clear_log).pack(anchor="e", padx=5, pady=2)

    def clear_log(self):
        self.console.configure(state='normal')
        self.console.delete(1.0, tk.END)
        self.console.configure(state='disabled')

    def check_queue(self):
        """Check the queue for new log messages"""
        while not self.log_queue.empty():
            msg = self.log_queue.get()
            self.console.configure(state='normal')
            self.console.insert(tk.END, msg)
            self.console.see(tk.END)
            self.console.configure(state='disabled')
        self.root.after(100, self.check_queue)

    def run_in_thread(self, target, *args):
        """Run a function in a separate thread"""
        t = threading.Thread(target=target, args=args)
        t.daemon = True
        t.start()

    # ==========================================
    # Format Fixer Tab
    # ==========================================
    def setup_fixer_tab(self):
        frame = ttk.Frame(self.tab_fixer, padding=20)
        frame.pack(fill="both", expand=True)

        # Folder Selection
        ttk.Label(frame, text="Select Folder to Fix:").grid(row=0, column=0, sticky="w", pady=5)
        self.fixer_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.fixer_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse...", command=lambda: self.browse_folder(self.fixer_path_var)).grid(row=0, column=2, padx=5, pady=5)

        # Options
        self.fixer_recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Recursive (Include Subfolders)", variable=self.fixer_recursive_var).grid(row=1, column=0, columnspan=2, sticky="w", pady=5)

        self.fixer_dryrun_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Dry Run (Preview only, no changes)", variable=self.fixer_dryrun_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)

        self.fixer_encoding_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Fix Encoding (Mojibake)", variable=self.fixer_encoding_var).grid(row=3, column=0, columnspan=2, sticky="w", pady=5)

        # Action Button
        ttk.Button(frame, text="Run Format Fixer", command=self.run_fixer).grid(row=4, column=0, columnspan=3, pady=20)

    def run_fixer(self):
        folder_path = self.fixer_path_var.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder first.")
            return

        recursive = self.fixer_recursive_var.get()
        dry_run = self.fixer_dryrun_var.get()
        fix_encoding = self.fixer_encoding_var.get()

        self.clear_log()
        print(f"Starting Format Fixer on: {folder_path}")
        print(f"Options: Recursive={recursive}, Dry Run={dry_run}, Fix Encoding={fix_encoding}")

        def _run():
            try:
                fixer = ArabicDocxFixer(folder_path, dry_run=dry_run, fix_encoding=fix_encoding)
                fixer.process_folder(recursive=recursive)
                print("\nJob Completed!")
            except Exception as e:
                print(f"\nError: {e}")
                import traceback
                traceback.print_exc(file=sys.stdout)

        self.run_in_thread(_run)

    # ==========================================
    # Document Generator Tab
    # ==========================================
    def setup_generator_tab(self):
        frame = ttk.Frame(self.tab_generator, padding=20)
        frame.pack(fill="both", expand=True)

        # Logo
        ttk.Label(frame, text="Logo Path (Optional):").grid(row=0, column=0, sticky="w", pady=5)
        self.gen_logo_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.gen_logo_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse...", command=self.browse_logo).grid(row=0, column=2, padx=5, pady=5)

        # Company Name
        ttk.Label(frame, text="Company Name:").grid(row=1, column=0, sticky="w", pady=5)
        self.gen_company_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.gen_company_var, width=40).grid(row=1, column=1, padx=5, pady=5)

        # Output File
        ttk.Label(frame, text="Output Filename:").grid(row=2, column=0, sticky="w", pady=5)
        self.gen_output_var = tk.StringVar(value="generated_doc.docx")
        ttk.Entry(frame, textvariable=self.gen_output_var, width=40).grid(row=2, column=1, padx=5, pady=5)

        # Content Source
        ttk.Label(frame, text="Content Source:").grid(row=3, column=0, sticky="w", pady=5)
        self.gen_source_var = tk.StringVar(value="text")
        ttk.Radiobutton(frame, text="Direct Text", variable=self.gen_source_var, value="text", command=self.toggle_gen_source).grid(row=3, column=1, sticky="w")
        ttk.Radiobutton(frame, text="From File", variable=self.gen_source_var, value="file", command=self.toggle_gen_source).grid(row=3, column=2, sticky="w")

        # File Selection (Hidden by default)
        self.gen_file_frame = ttk.Frame(frame)
        self.gen_file_frame.grid(row=4, column=0, columnspan=3, sticky="ew")
        ttk.Label(self.gen_file_frame, text="Content File:").grid(row=0, column=0, sticky="w", pady=5)
        self.gen_file_var = tk.StringVar()
        ttk.Entry(self.gen_file_frame, textvariable=self.gen_file_var, width=40).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(self.gen_file_frame, text="Browse...", command=self.browse_content_file).grid(row=0, column=2, padx=5, pady=5)

        # Content Text Area
        self.gen_content_frame = ttk.Frame(frame)
        self.gen_content_frame.grid(row=5, column=0, columnspan=3, sticky="ew")
        ttk.Label(self.gen_content_frame, text="Content:").grid(row=0, column=0, sticky="nw", pady=5)
        self.gen_content_text = tk.Text(self.gen_content_frame, height=10, width=40)
        self.gen_content_text.grid(row=0, column=1, columnspan=2, padx=5, pady=5, sticky="ew")

        # Action Button
        ttk.Button(frame, text="Generate Document", command=self.run_generator).grid(row=6, column=0, columnspan=3, pady=20)

        self.toggle_gen_source() # Initial state

    def toggle_gen_source(self):
        source = self.gen_source_var.get()
        if source == "file":
            self.gen_content_frame.grid_remove()
            self.gen_file_frame.grid()
        else:
            self.gen_file_frame.grid_remove()
            self.gen_content_frame.grid()

    def browse_logo(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.png *.jpg *.jpeg")])
        if file_path:
            self.gen_logo_var.set(file_path)

    def browse_content_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Text files", "*.txt"), ("All files", "*.*")])
        if file_path:
            self.gen_file_var.set(file_path)

    def run_generator(self):
        logo_path = self.gen_logo_var.get()
        company_name = self.gen_company_var.get()
        output_path = self.gen_output_var.get()
        source = self.gen_source_var.get()

        content = ""
        if source == "text":
            content = self.gen_content_text.get("1.0", tk.END).strip()
        else:
            file_path = self.gen_file_var.get()
            if not file_path:
                messagebox.showerror("Error", "Please select a content file.")
                return
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    content = f.read()
            except Exception as e:
                messagebox.showerror("Error", f"Could not read file: {e}")
                return

        if not company_name and not content:
            messagebox.showerror("Error", "Please provide Company Name or Content.")
            return

        self.clear_log()

        def _run():
            try:
                generator = ArabicDocumentGenerator(
                    logo_path=logo_path if logo_path else None,
                    company_name=company_name,
                    content=content,
                    output_path=output_path
                )
                generated_path = generator.generate()
                print(f"\nDocument generated at: {generated_path}")
            except Exception as e:
                print(f"\nError: {e}")
                import traceback
                traceback.print_exc(file=sys.stdout)

        self.run_in_thread(_run)

    # ==========================================
    # ODF Converter Tab
    # ==========================================
    def setup_converter_tab(self):
        frame = ttk.Frame(self.tab_converter, padding=20)
        frame.pack(fill="both", expand=True)

        # Mode Selection
        self.conv_mode_var = tk.StringVar(value="folder")
        ttk.Radiobutton(frame, text="Convert Folder", variable=self.conv_mode_var, value="folder", command=self.toggle_conv_mode).grid(row=0, column=0, sticky="w", pady=5)
        ttk.Radiobutton(frame, text="Convert Single File", variable=self.conv_mode_var, value="file", command=self.toggle_conv_mode).grid(row=0, column=1, sticky="w", pady=5)

        # Path Selection
        ttk.Label(frame, text="Path:").grid(row=1, column=0, sticky="w", pady=5)
        self.conv_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.conv_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse...", command=self.browse_conv_path).grid(row=1, column=2, padx=5, pady=5)

        # Options
        self.conv_recursive_var = tk.BooleanVar(value=True)
        self.conv_recursive_check = ttk.Checkbutton(frame, text="Recursive (Include Subfolders)", variable=self.conv_recursive_var)
        self.conv_recursive_check.grid(row=2, column=0, columnspan=2, sticky="w", pady=5)

        # Action Button
        ttk.Button(frame, text="Run Converter", command=self.run_converter).grid(row=3, column=0, columnspan=3, pady=20)

    def toggle_conv_mode(self):
        mode = self.conv_mode_var.get()
        if mode == "file":
            self.conv_recursive_check.state(['disabled'])
        else:
            self.conv_recursive_check.state(['!disabled'])

    def browse_conv_path(self):
        mode = self.conv_mode_var.get()
        if mode == "file":
            path = filedialog.askopenfilename(filetypes=[("ODF Text", "*.odt")])
        else:
            path = filedialog.askdirectory()

        if path:
            self.conv_path_var.set(path)

    def run_converter(self):
        path = self.conv_path_var.get()
        mode = self.conv_mode_var.get()

        if not path:
            messagebox.showerror("Error", "Please select a file or folder.")
            return

        self.clear_log()

        recursive = self.conv_recursive_var.get()

        def _run():
            try:
                converter = ODFToDocxConverter()
                if mode == "file":
                    converter.convert_file(path)
                else:
                    converter.convert_folder(path, recursive=recursive)
                print("\nConversion Job Completed!")
            except Exception as e:
                print(f"\nError: {e}")
                import traceback
                traceback.print_exc(file=sys.stdout)

        self.run_in_thread(_run)

    # ==========================================
    # Mojibake Scanner Tab
    # ==========================================
    def setup_mojibake_tab(self):
        frame = ttk.Frame(self.tab_mojibake, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Scan folder for encoding issues (Mojibake)").grid(row=0, column=0, columnspan=3, sticky="w", pady=10)

        # Folder Selection
        ttk.Label(frame, text="Select Folder:").grid(row=1, column=0, sticky="w", pady=5)
        self.moji_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.moji_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse...", command=lambda: self.browse_folder(self.moji_path_var)).grid(row=1, column=2, padx=5, pady=5)

        # Options
        self.moji_recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Recursive (Include Subfolders)", variable=self.moji_recursive_var).grid(row=2, column=0, columnspan=2, sticky="w", pady=5)

        # Action Button
        ttk.Button(frame, text="Scan for Mojibake", command=self.run_mojibake_scan).grid(row=3, column=0, columnspan=3, pady=20)

    def run_mojibake_scan(self):
        folder_path = self.moji_path_var.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder first.")
            return

        recursive = self.moji_recursive_var.get()

        self.clear_log()
        print(f"Starting Mojibake Scan on: {folder_path}")

        def _run():
            try:
                scan_mojibake(folder_path, recursive=recursive)
                print("\nScan Completed!")
            except Exception as e:
                print(f"\nError: {e}")
                import traceback
                traceback.print_exc(file=sys.stdout)

        self.run_in_thread(_run)

    # ==========================================
    # Footer Modifier Tab
    # ==========================================
    def setup_footer_tab(self):
        frame = ttk.Frame(self.tab_footer, padding=20)
        frame.pack(fill="both", expand=True)

        ttk.Label(frame, text="Add/Update Date in ODF Footers").grid(row=0, column=0, columnspan=3, sticky="w", pady=10)

        # Folder Selection
        ttk.Label(frame, text="Select Folder:").grid(row=1, column=0, sticky="w", pady=5)
        self.footer_path_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.footer_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(frame, text="Browse...", command=lambda: self.browse_folder(self.footer_path_var)).grid(row=1, column=2, padx=5, pady=5)

        # Options
        self.footer_recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(frame, text="Recursive", variable=self.footer_recursive_var).grid(row=2, column=0, sticky="w", pady=5)

        self.footer_dryrun_var = tk.BooleanVar(value=False)
        ttk.Checkbutton(frame, text="Dry Run (Preview only)", variable=self.footer_dryrun_var).grid(row=2, column=1, sticky="w", pady=5)

        # Date Format
        ttk.Label(frame, text="Date Format:").grid(row=3, column=0, sticky="w", pady=5)
        self.footer_format_var = tk.StringVar(value="%Y-%m-%d")
        ttk.Entry(frame, textvariable=self.footer_format_var, width=30).grid(row=3, column=1, padx=5, pady=5, sticky="w")
        ttk.Label(frame, text="(e.g. %Y-%m-%d)").grid(row=3, column=2, sticky="w", padx=5)

        # Custom Text
        ttk.Label(frame, text="Custom Text:").grid(row=4, column=0, sticky="w", pady=5)
        self.footer_text_var = tk.StringVar()
        ttk.Entry(frame, textvariable=self.footer_text_var, width=50).grid(row=4, column=1, columnspan=2, padx=5, pady=5, sticky="w")
        ttk.Label(frame, text="(Leave empty to use Date)").grid(row=5, column=1, sticky="w", padx=5)

        # Action Button
        ttk.Button(frame, text="Update Footers", command=self.run_footer_update).grid(row=6, column=0, columnspan=3, pady=20)

    def run_footer_update(self):
        folder_path = self.footer_path_var.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder first.")
            return

        recursive = self.footer_recursive_var.get()
        dry_run = self.footer_dryrun_var.get()
        date_format = self.footer_format_var.get()
        custom_text = self.footer_text_var.get() or None

        self.clear_log()

        def _run():
            try:
                modifier = ODFFooterModifier(date_format=date_format, date_text=custom_text, dry_run=dry_run)
                modifier.process_folder(folder_path, recursive=recursive)
                print("\nFooter Update Completed!")
            except Exception as e:
                print(f"\nError: {e}")
                import traceback
                traceback.print_exc(file=sys.stdout)

        self.run_in_thread(_run)

    # ==========================================
    # Tools / Cleanup Tab
    # ==========================================
    def setup_tools_tab(self):
        frame = ttk.Frame(self.tab_tools, padding=20)
        frame.pack(fill="both", expand=True)

        # --- Replace With Fixed ---
        group = ttk.LabelFrame(frame, text="Replace Original Files with Fixed Versions", padding=10)
        group.pack(fill="x", pady=10)

        ttk.Label(group, text="Folder:").grid(row=0, column=0, sticky="w", pady=5)
        self.repl_path_var = tk.StringVar()
        ttk.Entry(group, textvariable=self.repl_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(group, text="Browse...", command=lambda: self.browse_folder(self.repl_path_var)).grid(row=0, column=2, padx=5, pady=5)

        self.repl_recursive_var = tk.BooleanVar(value=True)
        ttk.Checkbutton(group, text="Recursive", variable=self.repl_recursive_var).grid(row=1, column=0, sticky="w", pady=5)

        self.repl_dryrun_var = tk.BooleanVar(value=True) # Default to Dry Run for safety
        ttk.Checkbutton(group, text="Dry Run (Safe Mode)", variable=self.repl_dryrun_var).grid(row=1, column=1, sticky="w", pady=5)

        ttk.Button(group, text="Replace Files", command=self.run_replace).grid(row=2, column=0, columnspan=3, pady=10)

    def run_replace(self):
        folder_path = self.repl_path_var.get()
        if not folder_path:
            messagebox.showerror("Error", "Please select a folder first.")
            return

        recursive = self.repl_recursive_var.get()
        dry_run = self.repl_dryrun_var.get()

        if not dry_run:
            if not messagebox.askyesno("Confirm Deletion",
                                       "WARNING: This will permanently DELETE original files and replace them with fixed versions.\n\nAre you sure you want to proceed?"):
                return

        self.clear_log()

        def _run():
            try:
                replacer = FixedFileReplacer(dry_run=dry_run)
                replacer.process_folder(folder_path, recursive=recursive)
                print("\nReplacement Job Completed!")
            except Exception as e:
                print(f"\nError: {e}")
                import traceback
                traceback.print_exc(file=sys.stdout)

        self.run_in_thread(_run)

    # ==========================================
    # Shared Helpers
    # ==========================================
    def browse_folder(self, string_var):
        folder = filedialog.askdirectory()
        if folder:
            string_var.set(folder)


def main():
    root = tk.Tk()
    app = ArabicToolsApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
