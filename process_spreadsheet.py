import os
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import pandas as pd
from bs4 import BeautifulSoup
from fuzzywuzzy import fuzz


def clean_html(html_content):
    if isinstance(html_content, str) and os.path.exists(html_content):
        with open(html_content, 'r', encoding='utf-8') as file:
            content = file.read()
    elif isinstance(html_content, str):
        content = html_content
    else:
        raise ValueError("Expected a string for html_content, got: {}".format(type(html_content).__name__))

    # Using BeautifulSoup to extract text from HTML
    soup = BeautifulSoup(content, "html.parser")
    cleaned_text = soup.get_text()

    # Remove unwanted characters and normalize spaces
    cleaned_text = cleaned_text.replace('&nbsp;', ' ').replace(';', ' ')
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()

    return cleaned_text


def match_public_body(row, opss_chart):
    public_body_input = row['public_body'].strip().lower()
    best_match_score = max(
        fuzz.ratio(public_body_input, opss_entry.strip().lower()) for opss_entry in opss_chart['public_body'])
    match_status = "MATCH" if best_match_score >= 90 else "REVIEW"
    return f"{match_status} ({best_match_score}%)"


class DataProcessorApp:
    def __init__(self, root):
        self.root = root
        root.title('OSBA Policy Tool')
        root.geometry("300x300")

        # Initialize file path attributes
        self.input_file_path = None
        self.opss_file_path = None

        # Label for the input Excel file
        tk.Label(root, text="Select the Input Excel File:").pack()
        self.input_file_label = tk.Label(root, text="No file selected")
        self.input_file_label.pack()
        tk.Button(root, text="Browse", command=self.load_input_file).pack()

        # Label for the OPSS Chart file
        tk.Label(root, text="Select the OPSS Chart File:").pack()
        self.opss_file_label = tk.Label(root, text="No file selected")
        self.opss_file_label.pack()
        tk.Button(root, text="Browse", command=self.load_opss_file).pack()

        tk.Button(root, text="Process Data", command=self.process_data).pack()
        self.progress = ttk.Progressbar(root, orient="horizontal", length=200, mode="determinate")
        self.progress.pack(pady=20)

    def load_input_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")))
        if file_path:
            self.input_file_path = file_path  # Save the file path
            self.input_file_label.config(text=os.path.basename(file_path))
        else:
            self.input_file_label.config(text="No file selected")

    def load_opss_file(self):
        file_path = filedialog.askopenfilename(filetypes=(("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")))
        if file_path:
            self.opss_file_path = file_path  # Save the file path
            self.opss_file_label.config(text=os.path.basename(file_path))
        else:
            self.opss_file_label.config(text="No file selected")

    def process_data(self):
        # Check if both files are selected before processing
        if not self.input_file_path or not self.opss_file_path:
            messagebox.showwarning("Warning", "Please select both files before processing.")
            return
        threading.Thread(target=self.actual_data_processing).start()

    def actual_data_processing(self):
        try:
            # Initial progress update
            self.root.after(0, lambda: self.update_progress(5))

            # Step 1: Load raw data
            raw_data = pd.read_excel(self.input_file_path, header=None)
            header_row = raw_data.dropna(how='all', axis=0).index[0]
            processed_data = pd.read_excel(self.input_file_path, header=header_row)
            processed_data.columns = [col.lower() for col in processed_data.columns]
            self.root.after(0, lambda: self.update_progress(20))  # Update progress

            # Step 2: Drop unnecessary columns and rename 'number' to 'code'
            columns_to_drop = ['unnamed: 0', 'unnamed: 10', 'helper column', 'section', 'status']
            processed_data.drop(columns_to_drop, axis=1, inplace=True, errors='ignore')
            if 'number' in processed_data.columns:
                processed_data.rename(columns={'number': 'code'}, inplace=True)
            self.root.after(0, lambda: self.update_progress(35))  # Update progress

            # Step 3: Load OPSS chart and clean data
            opss_chart = pd.read_excel(self.opss_file_path, header=0)
            opss_chart.columns = [col.lower() for col in opss_chart.columns]
            processed_data['public_body'] = processed_data['public_body'].astype(str).apply(clean_html)
            opss_chart['public_body'] = opss_chart['public_body'].astype(str).apply(clean_html)
            self.root.after(0, lambda: self.update_progress(50))  # Update progress

            # Step 4: Check for 'N' and clear them
            education_columns = ['department of education', 'state board', 'state superintendent',
                                 'superintendent of public instruction']
            processed_data = processed_data[~(processed_data[education_columns] == 'N').all(axis=1)]
            for col in education_columns:
                processed_data[col] = processed_data[col].replace('N', '')
            self.root.after(0, lambda: self.update_progress(65))  # Update progress

            # Step 5: Apply fuzzy matching
            processed_data['body match (%)'] = processed_data.apply(lambda row: match_public_body(row, opss_chart),
                                                                    axis=1)
            self.root.after(0, lambda: self.update_progress(80))  # Update progress

            # Step 6: Process recommendations
            opss_dict = opss_chart.set_index('code')['recommendation'].to_dict()
            processed_data['recommendation'] = processed_data['code'].apply(lambda x: opss_dict.get(x, "REVIEW"))
            self.root.after(0, lambda: self.update_progress(95))  # Update progress

            # Final processing steps...
            self.root.after(0, lambda: self.update_progress(100))  # Ensure it reaches 100% after all processing

            # Delay the save file dialog to ensure users see the 100% completion
            self.root.after(500, lambda: self.save_file(processed_data))  # Delay to ensure users see the 100%
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")
            self.root.after(0, lambda: self.update_progress(0))  # Reset progress on error

    def update_progress(self, value):
        self.progress['value'] = value
        self.root.update_idletasks()

    def save_file(self, processed_data):
        file_options = {
            'defaultextension': '.xlsx',
            'filetypes': [('Excel files', '*.xlsx'), ('All files', '*.*')],
            'initialdir': os.path.expanduser('~'),
            'title': 'Save processed file'
        }
        output_file_path = filedialog.asksaveasfilename(**file_options)
        if output_file_path:
            processed_data.to_excel(output_file_path, index=False)
            messagebox.showinfo("Success", f"Data processed successfully and saved to {output_file_path}.")
        else:
            messagebox.showwarning("Cancelled", "Save operation was cancelled.")

        # Reset progress bar to 0% after save dialog is completed, regardless of outcome
        self.root.after(0, lambda: self.update_progress(0))


def main():
    root = tk.Tk()
    app = DataProcessorApp(root)  # Ignore Warn: tkinter GUI will not build without this line.
    root.mainloop()


if __name__ == "__main__":
    main()
