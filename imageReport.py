import tkinter as tk
from tkinter import filedialog
import pandas as pd
import openpyxl
import os
# pip install pyinstaller
# pyinstaller --onefile --icon=bart.ico imageReport.py

#################################################################################################################################


def process_excel():
    global df
    global file_path
    file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
    if file_path:
        df = pd.read_excel(file_path)
        #print(df)
        return df

def select_folder():
    global folder_path
    folder_path = filedialog.askdirectory()
    #print(folder_path)
    return folder_path

def checking_files(folder_path, main_df):
    df_len = len(main_df)

    for i in range(df_len):
        folder_path = folder_path.replace("/", "\\")
        curBrand = main_df['Brands'].iloc[i]
        curSKU = main_df['SKU'].iloc[i]
        uploadedDir = folder_path + '\\' + curBrand + '\\UPLOADED\\' + curSKU + '-4.jpg'

        #print(uploadedDir)

        file_exists1 = os.path.isfile(uploadedDir)
        if file_exists1:
            main_df.at[i, 'Uploaded'] = 'Exists'
            #print('exists')
        else:
            main_df.at[i, 'Uploaded'] = 'Not in folder'
            #print('does not exists')

        boxDir = folder_path + '\\' + curBrand + '\\SHOE BOX\\' + curSKU + '.jpg'
        file_exists2 = os.path.isfile(boxDir)
        if file_exists2:
            main_df.at[i, 'Box image'] = 'Exists'
        else:
            main_df.at[i, 'Box image'] = 'Not in folder'

    
    main_df.to_excel(file_path, index=False)





#################################################################################################################################


window = tk.Tk()
window.geometry("300x150")

window.title("Photo Dept.")

# Body
label = tk.Label(window, text="Select an Excel file:")
label.pack()

# Create an input field
select_dir = tk.Button(window, text="Picture folder", command=select_folder)
select_dir.pack()

select_file = tk.Button(window, text="Select Excel File", command=process_excel)
select_file.pack()

# Create a button
button = tk.Button(window, text="Run Script", command=lambda : checking_files(folder_path, df))
button.pack()

# Initialize variables
folder_path = ''
file_path = ''
df = pd.DataFrame()

window.mainloop()