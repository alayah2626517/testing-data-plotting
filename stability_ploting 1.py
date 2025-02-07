import tkinter as tk
from tkinter import filedialog
from tkinter.constants import *
import matplotlib.pyplot as plt
from openpyxl import load_workbook
import numpy as np
import os

# 擺好位置

def run_app():
    # 建立窗口
    root = tk.Tk()
    root.title("Stability data plotting tool")
    root.geometry('380x400')
    root.resizable(False, False)
   
    # 輸入框: Batch
    tk.Label(root, text="Maximmum Batch?", font=('Arial', 14, 'bold')).pack(pady=5)
    batch_num = tk.StringVar()
    tk.Entry(root, textvariable=batch_num).pack(pady=10)


    # 文件選擇按鈕
    file_path = None
    def select_file():
        global file_path
        file_path = filedialog.askopenfilename(title="Select file", filetypes=[("All Files", "*.*")])
        if not file_path:
            return
        else:
            print(f"Successfully selecting: {file_path}")
    tk.Button(root, text="Select file", command=select_file).pack(pady=10)

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
            # print(data_total)
            
            # 定義各項變數
            condition = data_total[1][0]
            test_item = data_total[0][0]
            value_limit = None
            if data_total[0][2] == "±":
                low_limit = data_total[0][1]-data_total[0][3]
                upper_limit = data_total[0][1]+data_total[0][3]
                value_limit = np.linspace(low_limit, upper_limit, num=10)
            else:
                print("No specification defined.")            

            #製圖
            title = f"{condition}-{test_item}"
            datasets = data_total[2:]  # 數據本身
            plt.figure(figsize=(10, 6))
            for row in datasets:
                label = row[0]  # 每一行的標籤
                values = row[1:]  # 每一行的數據（跳過標籤）
                # x_axis = tuple(x for x in data_total[1][1:] if x is not None)
                x_axis = data_total[1][1:]
                plt.plot(x_axis, values, marker='o', linestyle='-', linewidth=2, alpha = 0.6, label=label)
            plt.title(title, fontsize=18, fontweight='bold')
            plt.xlabel("Time point (months)")
            plt.ylabel(test_item, fontsize=15)
            plt.yticks(value_limit, fontsize=12)
            plt.grid(True, linestyle='--', alpha=0.6)
            plt.legend()
            plt.grid(True)
            plt.savefig(f"{folder_path}/{title}.png", dpi=300)
        wb.close()
    tk.Button(root, text="Load file", command=load_sheet).pack(pady=10)

    root.mainloop()

if __name__ == "__main__":
    run_app()
