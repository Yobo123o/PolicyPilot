import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import pandas as pd
from bs4 import BeautifulSoup
import os
import warnings


def clean_html(html_content):
    # Check if html_content is a filename
    if os.path.exists(html_content):
        # Open the file and pass the file handle to BeautifulSoup
        with open(html_content, 'r', encoding='utf-8') as file:
            soup = BeautifulSoup(file, "html.parser")
    else:
        # Assume html_content is already HTML markup and pass it directly to BeautifulSoup
        soup = BeautifulSoup(html_content, "html.parser")

    # Extract text from the BeautifulSoup object
    cleaned_text = soup.get_text()
    cleaned_text = cleaned_text.replace('&nbsp;', ' ')  # Replace '&nbsp;' with a space

    # Suppress MarkupResemblesLocatorWarning
    warnings.simplefilter("ignore", category=UserWarning)

    return cleaned_text


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

        # Variables to store file paths
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
        try:
            # Dynamically determine the header row for the input file
            # Attempt to open the file without setting the header to read raw data
            raw_data = pd.read_excel(self.input_file_path, header=None)
            # Find the first non-empty row to set as header
            header_row = raw_data.dropna(how='all', axis=0).index[0]
            processed_data = pd.read_excel(self.input_file_path, header=header_row)
            processed_data.columns = [col.lower() for col in processed_data.columns]

            # DEBUG:
            # print("Columns after setting header dynamically and converting to lower case:", processed_data.columns)

            # Load OPSS chart data, assuming the headers are correctly set
            opss_chart = pd.read_excel(self.opss_file_path, header=0)
            opss_chart.columns = [col.lower() for col in opss_chart.columns]

            # DEBUG:
            # print("OPSS chart columns after setting header:", opss_chart.columns)

            # Rename 'number' column to 'code' before any operations that use 'code'
            if 'number' in processed_data.columns:
                processed_data.rename(columns={'number': 'code'}, inplace=True)
            else:
                messagebox.showerror("Error", "Missing 'number' column in input file.")
                return

            # Ensure all data in columns expected to contain text are strings, and clean HTML
            if 'public_body' in processed_data.columns:
                processed_data['public_body'] = processed_data['public_body'].astype(str).apply(clean_html)

            # Convert potential float or NaN values in other columns to string if necessary
            text_columns = ['department of education', 'state board', 'state superintendent',
                            'superintendent of public instruction']
            for col in text_columns:
                if col in processed_data.columns:
                    processed_data[col] = processed_data[col].fillna('').apply(lambda x: str(x).replace('N', ''))

            # Drop 'section' and 'status' columns from processed data as not needed
            if 'section' in processed_data.columns or 'status' in processed_data.columns:
                processed_data.drop(['unnamed: 0', 'section', 'status', 'helper column', 'unnamed: 10'], axis=1,
                                    inplace=True, errors='ignore')

            # Clean HTML from public_body column if it exists
            if 'public_body' in processed_data.columns:
                processed_data['public_body'] = processed_data['public_body'].apply(clean_html)

            # Drop rows where specified columns only contain 'N'
            columns_to_check = ['department of education', 'state board', 'state superintendent',
                                'superintendent of public instruction']
            processed_data = processed_data[~(processed_data[columns_to_check].fillna('') == 'N').all(axis=1)]

            # Replace 'N' with empty string in specified columns
            for col in columns_to_check:
                if col in processed_data.columns:
                    processed_data[col] = processed_data[col.replace('N', '')]

            # Create body match column based on public_body comparison
            processed_data['body match'] = processed_data.apply(lambda row: "MATCH" if any(
                opss_chart.loc[opss_chart['code'] == row['code'], 'public_body'].fillna('').str.lower() == row[
                    'public_body'].lower()) else "REVIEW", axis=1)

            # Process codes and recommendations
            opss_dict = opss_chart.set_index('code')['recommendation'].to_dict()
            processed_data['recommendation'] = processed_data['code'].apply(
                lambda x: opss_dict.get(x, "REVIEW"))

            # Prompt user for output file name
            output_file_name = simpledialog.askstring("Output File Name", "Enter the output file save name:",
                                                      parent=self.root)
            if output_file_name:
                # Ensure the output file name has .xlsx extension
                if not output_file_name.endswith('.xlsx'):
                    output_file_name += '.xlsx'
                documents_folder = os.path.join(os.path.expanduser('~'), 'Documents')
                output_file_path = os.path.join(documents_folder, output_file_name)

                # Save the final data to an Excel file with the user-defined name
                processed_data.to_excel(output_file_path, index=False)
                messagebox.showinfo("Success",
                                    f'Data processed and recommendations added successfully. '
                                    f'File saved to \'{output_file_path}\'.')
            else:
                messagebox.showwarning("Cancelled", "File save name was not provided. The process was cancelled.")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred: {str(e)}")


def main():
    root = tk.Tk()
    app = DataProcessorApp(root)  # Instantiate your GUI application
    root.mainloop()


if __name__ == "__main__":
    main()
