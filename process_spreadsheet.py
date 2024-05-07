import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog

import pandas as pd
from bs4 import BeautifulSoup
from fuzzywuzzy import fuzz


def clean_html(html_content):
    if os.path.exists(html_content):
        with open(html_content, 'r', encoding='utf-8') as file:
            content = file.read()
    else:
        content = html_content

    # Using BeautifulSoup to extract text
    soup = BeautifulSoup(content, "html.parser")
    cleaned_text = soup.get_text()

    # Replace non-breaking spaces, remove semicolons, and normalize spaces
    cleaned_text = cleaned_text.replace('&nbsp;', ' ').replace(';', ' ')
    cleaned_text = re.sub(r'\s+', ' ', cleaned_text).strip()

    return cleaned_text


def match_public_body(row, opss_chart):
    # Lowercase and strip both compared texts
    public_body_input = row['public_body'].strip().lower()
    # Using fuzzy matching to determine the best match score
    best_match_score = max(
        fuzz.ratio(public_body_input, opss_entry.strip().lower()) for opss_entry in opss_chart['public_body'])
    # Format the output to include the match status and the percentage
    match_status = "MATCH" if best_match_score >= 90 else "REVIEW"
    return f"{match_status} ({best_match_score}%)"


class DataProcessorApp:
    def __init__(self, root):
        self.root = root
        root.title('OSBA Policy Tool')
        root.geometry("250x275")

        # Set up the GUI layout
        tk.Label(root, text="Select the Input Excel File:").pack()
        tk.Button(root, text="Browse", command=self.load_input_file).pack()
        tk.Label(root, text="Select the OPSS Chart File:").pack()
        tk.Button(root, text="Browse", command=self.load_opss_file).pack()
        tk.Button(root, text="Process Data", command=self.process_data).pack()

        self.input_file_path = None
        self.opss_file_path = None

    def load_input_file(self):
        self.input_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if self.input_file_path:
            messagebox.showinfo("File Selected", f"Selected file: {self.input_file_path}")

    def load_opss_file(self):
        self.opss_file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
        if self.opss_file_path:
            messagebox.showinfo("File Selected", f"Selected file: {self.opss_file_path}")

    def process_data(self):
        if not self.input_file_path or not self.opss_file_path:
            messagebox.showwarning("File Not Selected",
                                   "Please select both the input and OPSS chart files before processing.")
            return

        try:
            raw_data = pd.read_excel(self.input_file_path, header=None)
            header_row = raw_data.dropna(how='all', axis=0).index[0]
            processed_data = pd.read_excel(self.input_file_path, header=header_row)
            processed_data.columns = [col.lower() for col in processed_data.columns]

            # Extend the list of columns to drop
            columns_to_drop = ['unnamed: 0', 'unnamed: 10', 'helper column', 'section', 'status']
            processed_data.drop(columns_to_drop, axis=1, inplace=True, errors='ignore')

            # Rename 'number' column to 'code' before any operations that use 'code'
            if 'number' in processed_data.columns:
                processed_data.rename(columns={'number': 'code'}, inplace=True)

            # Load OPSS chart data, assuming the headers are correctly set
            opss_chart = pd.read_excel(self.opss_file_path, header=0)
            opss_chart.columns = [col.lower() for col in opss_chart.columns]

            # Ensure all relevant text fields are cleaned
            if 'public_body' in processed_data.columns:
                processed_data['public_body'] = processed_data['public_body'].astype(str).apply(clean_html)

            # Load and clean OPSS chart data
            opss_chart = pd.read_excel(self.opss_file_path, header=0)
            opss_chart.columns = [col.lower() for col in opss_chart.columns]
            opss_chart['public_body'] = opss_chart['public_body'].astype(str).apply(clean_html)

            # Apply fuzzy matching and capture the match confidence
            processed_data['body match (%)'] = processed_data.apply(lambda row: match_public_body(row, opss_chart),
                                                                    axis=1)

            # Columns to check for 'N' and subsequently clear 'N'
            education_columns = ['department of education', 'state board', 'state superintendent',
                                 'superintendent of public instruction']
            processed_data = processed_data[~(processed_data[education_columns] == 'N').all(axis=1)]

            # Replace remaining 'N's with empty strings in specified columns
            for col in education_columns:
                processed_data[col] = processed_data[col].replace('N', '')

            # Process codes and recommendations
            opss_dict = opss_chart.set_index('code')['recommendation'].to_dict()
            processed_data['recommendation'] = processed_data['code'].apply(lambda x: opss_dict.get(x, "REVIEW"))

            # Prompt user for output file name
            output_file_name = simpledialog.askstring("Output File Name", "Enter the output file save name:",
                                                      parent=self.root)
            if output_file_name:
                if not output_file_name.endswith('.xlsx'):
                    output_file_name += '.xlsx'
                output_file_path = os.path.join(os.path.expanduser('~'), 'Documents', output_file_name)
                processed_data.to_excel(output_file_path, index=False)
                messagebox.showinfo("Success",
                                    f'Data processed and recommendations added successfully. '
                                    f'File saved to "{output_file_path}".')
            else:
                messagebox.showwarning("Cancelled", "File save name was not provided. The process was cancelled.")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


def main():
    root = tk.Tk()
    app = DataProcessorApp(root)  # IGNORE WARN: Without this GUI will not build correctly.
    root.mainloop()


if __name__ == "__main__":
    main()
