#### Purpose: This script is for plotting testing data from excel, beware of the layout of data.
#### Author: Hsin-Yun Hung
#### Inition version: 2025/02/07

import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter.constants import *
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import numpy as np
import os

def run_app():
    ## 建立窗口
    root = tk.Tk()
    root.title("Test results plotting tool")
    root.geometry('380x410')
    root.iconbitmap('EG logo.ico')
    root.resizable(False, False)
   
    ## 輸入框: Batch
    label_frame = tk.LabelFrame(root, width=380, height=100, text="Step 1", bg="light blue", bd=10, relief='groove')
    label_frame.pack(padx=10, pady=5, fill="x")
    label_frame.pack_propagate(False)
    tk.Label(label_frame, text="Maximmum # of Batch ?", font=('Arial', 13), bg="light blue").grid(row=0, column=0, padx=60, pady=10)
    batch_num = tk.StringVar()
    batch_entry = tk.Spinbox(label_frame, from_=1, to=100, textvariable=batch_num, font=("Arial", 12))
    batch_entry.grid(row=1, column=0, padx=70, pady=15)

    ## 文件選擇按鈕
    def select_file():
        global file_path
        file_path = filedialog.askopenfilename(title="Select file", filetypes=[("All Files", "*.*")])
        if not file_path:
            return
        else:
            print(f"Successfully selecting: {file_path}")
    label_frame_2 = tk.LabelFrame(root, width=380, height=100, text="Step 2", bg="DarkOliveGreen2", bd=10, relief='groove')
    label_frame_2.pack(padx=10, pady=5, fill="x")
    tk.Label(label_frame_2, text="Beware of the datasets layout!", font=('Arial', 12), bg="DarkOliveGreen2").grid(row=0, column=0, padx=66, pady=5)
    file_path = None
    tk.Button(label_frame_2, text="Select file", command=select_file).grid(row=1, column=0, padx=45, pady=15)
    
     ## 文件導入按鈕
    def load_sheet():
        global file_path
        batch_value = batch_num.get()
        batch_value = int(batch_value)
        wb = load_workbook(file_path)
        sheet_name = wb.sheetnames
        folder_path = filedialog.askdirectory(title="選擇儲存路徑")
        for sheet in sheet_name:
            sheet = wb[sheet]
            data_total = []
            for data in sheet.iter_rows(min_row=1, max_row=batch_value+1, min_col=1, max_col=10, values_only=True):
                data_total.append(data)
            print(data_total)
                
            ### 定義各項變數
            condition = data_total[1][0]
            test_item = data_total[0][0]
            num = 10
            value_limit, lower_limit, upper_limit = None, None, None
            if data_total[0][2] == "±":
                lower_limit = data_total[0][1]-data_total[0][3]
                upper_limit = data_total[0][1]+data_total[0][3]
                value_limit = np.linspace(lower_limit, upper_limit, num=num)
            elif data_total[0][2] == "<" or data_total[0][2] == "≦":
                upper_limit = data_total[0][3]
                distance = upper_limit*0.1
                value_limit = np.linspace(upper_limit-distance*(num-1), upper_limit, num=num)
            elif data_total[0][2] == ">" or data_total[0][2] == "≧":
                lower_limit = data_total[0][3]
                distance = lower_limit*0.1
                value_limit = np.linspace(lower_limit, lower_limit+distance*(num-1), num=num)
            else:
                print("Report data")
            decimal = len(str(value_limit[0]).split(".")[1])
            value_limit = [round(float(value), decimal) for value in value_limit]
            
            ### 製折線圖
            chart_title = f"{condition}-{test_item}"
            datasets = data_total[2:]  # 數據本身
            x_axis = data_total[1][1:]
            fig, ax = plt.subplots(1, 1, sharex='col', figsize=(10, 6))
            for row in datasets:
                label = row[0]  # 每一行的標籤
                values = row[1:]  # 每一行的數據（跳過標籤）
                ax.plot(x_axis, values, marker='o', linestyle='-', linewidth=2, alpha = 0.6, label=label)
                
            ### 在圖上設定規格上下限
            if lower_limit is not None:
                ax.axhline(y=lower_limit, color='#8B0000', linestyle='--', linewidth=1.5)
            if upper_limit is not None:
                ax.axhline(y=upper_limit, color='#8B0000', linestyle='--', linewidth=1.5)
            ax.set_title(chart_title, fontsize=18, fontweight='bold')
            
            ### 製作表格
            data_rows = [row[1:] for row in datasets]  # 每一行的數據
            bbox_height = 0.3 + (batch_value * 0.02)
            bbox_y = -0.5 - (batch_value * 0.01)
            ax.table(cellText=data_rows, colLabels=x_axis, loc='bottom', cellLoc='center', colLoc='center', bbox=[0, bbox_y, 1, bbox_height])
            
            ax.set_xlabel("Time point (months)")
            ax.set_ylabel(test_item, fontsize=15)
            ax.set_yticks(value_limit)
            ax.set_yticklabels(ax.get_yticks(), fontsize=12)
            ax.grid(True, linestyle='--', alpha=0.6)
            ax.legend(loc='upper right', bbox_to_anchor=(1.05, 1), fontsize=12)
            ax.grid(True)
            plt.tight_layout()
            plt.savefig(f"{folder_path}/{chart_title}.png", dpi=300)
        wb.close()
        if messagebox.askyesno("Plotting complete", "All charts have been successfully created. Do you want to exit?"):
            root.quit()
    label_frame_3 = tk.LabelFrame(root, width=380, height=100, text="Step 2", bg="khaki1", bd=10, relief='groove')
    label_frame_3.pack(padx=10, pady=5, fill="x")
    tk.Button(label_frame_3, text="Load file", command=load_sheet).grid(row=0, column=0, padx=55, pady=10)
    tk.Label(label_frame_3, text="Selecting folder to save chart after loading the file.", font=('Arial', 11), bg="khaki1").grid(row=1, column=0, padx=6, pady=10)

    
    root.mainloop()
    


if __name__ == "__main__":
    run_app()
