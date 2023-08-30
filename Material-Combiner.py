import tkinter as tk
from tkinter import filedialog
import os
import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows

def select_file():
    global file_path
    root.update_idletasks()  # Force the window to update
    file_path = filedialog.askopenfilename(title="Select your Excel file", filetypes=[("Excel files", "*.xlsx *.xls *.xlsm")])
    if not file_path:
        print("No file selected. Exiting.")
        status_label.config(text="No file selected.")
    else:
        status_label.config(text="Processing File...")
        root.update_idletasks()  # Force the window to update
        # Load the Excel file
        xl = pd.ExcelFile(file_path)

        # Create an empty DataFrame to store the combined data
        combined_data = pd.DataFrame()

        # Loop through the sheet names, combining the ones that match the pattern
        for sheet_name in xl.sheet_names:
            if (sheet_name.startswith('P') or sheet_name.startswith('S')) and sheet_name[1:].isdigit():
                sheet_data = pd.read_excel(file_path, sheet_name=sheet_name)
                # Trim spaces from all the columns
                sheet_data = sheet_data.applymap(lambda x: x.strip() if isinstance(x, str) else x)
                combined_data = combined_data._append(sheet_data, ignore_index=True)

        # Separate rows with 'R/I' in 'Work Function' into two rows: one with 'R' and one with 'I'
        rows_to_append = []
        for index, row in combined_data.iterrows():
            work_function = row['Work Function']
            if work_function == 'R/I':
                row_r = row.copy()
                row_i = row.copy()
                row_r['Work Function'] = 'R'
                row_i['Work Function'] = 'I'
                rows_to_append.append(row_r)
                rows_to_append.append(row_i)
            else:
                rows_to_append.append(row)
        combined_data = pd.DataFrame(rows_to_append)

        # Delete rows where the 'E' column is NaN or empty
        combined_data.dropna(subset=['Quantity'], inplace=True)
        combined_data.drop('Points', axis=1, inplace=True)

        # Group by the first and second columns and sum the 'Quantity'
        combined_data['Quantity'] = combined_data['Quantity'].astype(float)
        combined_data = combined_data.groupby(['Material Name', 'Work Function'], as_index=False).agg({'Quantity': 'sum', 'PART NO.': 'first', 'BIN NO.': 'first'})

        # Create a new workbook and worksheet
        wb = Workbook()
        ws = wb.active

        # Add the combined data to the worksheet
        for r_idx, row in enumerate(dataframe_to_rows(combined_data, index=False, header=True), 1):
            for c_idx, value in enumerate(row, 1):
                ws.cell(row=r_idx, column=c_idx, value=value)

        # Adjust the column widths to fit the contents
        for col in ws.columns:
            max_length = 0
            column = [cell for cell in col]
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(cell.value)
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column[0].column_letter].width = adjusted_width

        # Save the workbook in the same directory as the selected file
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx",
                                                 filetypes=[("Excel files", "*.xlsx")],
                                                 initialdir=os.path.dirname(file_path),
                                                 title="Save the combined file as")
        wb.save(save_path)
        print(f"Sheets combined successfully! Saved as {save_path}")
        root.destroy()


# Initialize Tkinter root window
root = tk.Tk()
root.title("Excel File Combiner")
root.geometry("400x200")

# Add a button to trigger file selection
button = tk.Button(root, text="Select file of materials to combine.", command=select_file, height=3, width=50)
button.pack(pady=20)

# Add a label to display the status
status_label = tk.Label(root, text="", width=50)
status_label.pack(pady=10)

root.mainloop()
