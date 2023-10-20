import tkinter as tk
from tkinter import filedialog
import pandas as pd
import pandas as pd
import os
import pandas as pd

def convert_html_to_xlsx(file_path):
    # Trying to read HTML tables from the file
    try:
        tables = pd.read_html(file_path)
        # Assuming the first table is what you want
        df = tables[0]
        
        new_file_path = os.path.splitext(file_path)[0] + '.xlsx'
        # Writing the DataFrame to an xlsx file
        df.to_excel(new_file_path, index=False, engine='openpyxl')
        return new_file_path

    except Exception as e:
        print(f"Error converting {file_path} to xlsx: {str(e)}")
        return None

def process_file(file_name):
    _, file_extension = os.path.splitext(file_name)
    
    try:
        # If file is .xls, assume it might be mislabeled HTML and try to convert
        if file_extension.lower() == '.xls':
            file_name = convert_html_to_xlsx(file_name)
            if file_name is None:
                return None, None, None
        
        df = pd.read_excel(file_name, engine='openpyxl')
        lot = df.iloc[0:2, 1]
        lot.name = lot.iloc[0]
        lot = lot.drop(lot.index[0]).reset_index(drop=True)

        df = df.iloc[7:, [2, 9]]
        df.columns = df.iloc[0]
        df = df.drop(df.index[0])
        grouped = df.groupby('SIZE')['ORIGINAL QTY'].sum().reset_index()

        s_count = grouped.loc[grouped['SIZE'] == 'S', 'ORIGINAL QTY'].sum()
        m_count = grouped.loc[grouped['SIZE'] == 'M', 'ORIGINAL QTY'].sum()

        return lot.iloc[0], s_count, m_count

    except Exception as e:
        print(f"Error processing {file_name}: {str(e)}")
        return None, None, None







def process_files():
    # Lists to store the results
    lots = []
    s_counts = []
    m_counts = []

    for file_path in file_paths:
        lot, s_count, m_count = process_file(file_path)
        
        # Append only if values are not None (indicating an error hasn't occurred)
        if lot is not None:
            lots.append(lot)
            s_counts.append(s_count)
            m_counts.append(m_count)

    # Construct the final DataFrame
    result_df = pd.DataFrame({
        'lot': lots,
        'S': s_counts,
        'M': m_counts
    })

    # Get the user's selected file path through GUI
    output_path = filedialog.asksaveasfilename(defaultextension=".xlsx", 
                                                filetypes=[("Excel files", "*.xlsx")],
                                                title="Choose filename and location to save")

    # Check if user canceled the save file dialog
    if output_path == "":
        print("File save was canceled")
        return

    # Export to Excel
    try:
        result_df.to_excel(output_path, index=False)
        print(f"File saved successfully to {output_path}")
    except Exception as e:
        print(f"Error while saving the file: {str(e)}")

def upload_action():
    global file_paths
    new_file_paths = filedialog.askopenfilenames(title="Choose Excel files", 
                                                 filetypes=[("Excel files", "*.xlsx"),
                                                            ("Legacy Excel files", "*.xls")])
    file_paths.extend(new_file_paths)  # Add new selected files to the existing list
    label.config(text=f"{len(file_paths)} files selected")



def clear_action():
    global file_paths
    file_paths = []  # Clear the list of file paths
    label.config(text="No files selected")  # Update the label


# Initial setup
file_paths = []

# GUI setup
root = tk.Tk()
root.title("Excel Processor")
root.geometry("500x400")

large_font = ('Verdana', 16)


upload_button = tk.Button(root, text="Upload Files", command=upload_action, font=large_font, width=20, height=2)
upload_button.pack(pady=10)

run_button = tk.Button(root, text="Run", command=process_files, font=large_font, width=20, height=2)
run_button.pack(pady=10)

clear_button = tk.Button(root, text="Clear", command=clear_action, font=large_font, width=20, height=2)
clear_button.pack(pady=10)

label = tk.Label(root, text="No files selected", font=large_font)
label.pack(pady=20)

root.mainloop()