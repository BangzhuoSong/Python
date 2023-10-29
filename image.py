import tkinter as tk
from tkinter import filedialog
import pandas as pd
import openpyxl
from openpyxl.drawing.image import Image
import requests
from io import BytesIO

def is_valid_url(url):
    if isinstance(url, str) and (url.startswith('http://') or url.startswith('https://')):
        return True
    return False

def process_excel(file_path):
    df = pd.read_excel(file_path)

    workbook = openpyxl.load_workbook(file_path)
    sheet = workbook.active

    row = 2
    for url in df['Unnamed: 17']:
        if is_valid_url(url):
            try:
                response = requests.get(url)
                image_stream = BytesIO(response.content)
                img = Image(image_stream)

                # Adjust image dimensions
                img.width = 140
                img.height = 180

                cell_col = df.columns.get_loc('Unnamed: 17') + 1
                col_letter = openpyxl.utils.get_column_letter(cell_col)

                # Adjust cell dimensions
                sheet.column_dimensions[col_letter].width = 150
                sheet.row_dimensions[row].height = 190

                sheet.add_image(img, f'{col_letter}{row}')
            except requests.RequestException as e:
                print(f"Error fetching image from URL {url}: {e}")
        row += 1

    workbook.save(file_path)


def process_file():
    file_path = filedialog.askopenfilename(title="Select Excel file",
                                       filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls")])

    if file_path:  # if a file was selected
        try:
            process_excel(file_path)
            tk.messagebox.showinfo("Success", "File processed successfully!")
        except Exception as e:
            tk.messagebox.showerror("Error", f"An error occurred: {str(e)}")

# Create basic window
app = tk.Tk()
app.title("Excel Image Processor")
app.geometry("400x150")

# Add a button to open file and process it
process_button = tk.Button(app, text="Process Excel File", command=process_file, height=2, width=20)
process_button.pack(pady=50)

# Run the app
app.mainloop()
