#python-pptx == 0.6.23

from pptx import Presentation
from pptx.util import Pt, Inches

import tkinter as tk
from tkinter import filedialog

import os

def add_slide_page(pptx_file_path, first_count, localte_per_from_left, localte_per_from_top, save_file_path):
    presentation = Presentation(pptx_file_path)

    rate = float(Inches(1))
    
    left = Inches(localte_per_from_left*(presentation.slide_width)/(100*rate))
    top = Inches(localte_per_from_top*(presentation.slide_height)/(100*rate))
    width = Inches((100 - localte_per_from_left)*(presentation.slide_width)/(100*rate))
    height = Inches((100 - localte_per_from_top)*(presentation.slide_height)/(100*rate))
    
    
    n = first_count
    all_sliders_num = len(presentation.slides) + n - 1

    for slide in presentation.slides:
        textbox = slide.shapes.add_textbox(left, top, width, height)
        textbox.text = f"{n}/{all_sliders_num}"
        n += 1

    presentation.save(save_file_path)


def main_run():
    pptx_file_path = filedialog.askopenfilename(filetypes=[("テキストファイル", "*.pptx")])
    directory, filename = os.path.split(pptx_file_path)
    first_count = entry_first_num.get()
    
    localte_per_from_left = entry_left_local.get()
    localte_per_from_top = entry_top_local.get()
    try:
        file_name, extension = os.path.splitext(filename)
        save_file = f"{file_name}_AddPageNumber{extension}"
        save_file_path = os.path.join(directory, save_file)
        ##############################################################################################
        i = 1
        while os.path.exists(save_file_path):
            save_file = f"{file_name}_AddPageNumber({i}){extension}"
            save_file_path = os.path.join(directory, save_file)
            i += 1
        ##############################################################################################
        add_slide_page(pptx_file_path, int(first_count), float(localte_per_from_left), float(localte_per_from_top), save_file_path)
        message = f"正常に動作しました．\n{save_file_path}\nとして保存されました．"
    except Exception as e:
        error_message = str(e)
        if f"Package not found at '{pptx_file_path}'" in error_message:
            message = "対象のパワーポイントファイルが見つかりませんでした．\n対象のパワーポイントファイルを開いている場合は閉じてください．"
        else:
            message = "予期せぬエラーが発生しました．"
    run_label.config(text=message)

    

if __name__ == "__main__":
    root = tk.Tk()
    root.title("Add page number")

    lavel_first_num = tk.Label(root, text="\n開始する数字を入力してください")
    lavel_first_num.grid()
    entry_first_num = tk.Entry(root)
    entry_first_num.grid()

    
    lavel_left_local = tk.Label(root, text="\n左端からの位置をパーセンテージで入力してください")
    lavel_left_local.grid()
    entry_left_local = tk.Entry(root)
    entry_left_local.grid()

    lavel_top_local = tk.Label(root, text="\n上端からの位置をパーセンテージで入力してください")
    lavel_top_local.grid()
    entry_top_local = tk.Entry(root)
    entry_top_local.grid()
    

    run_button = tk.Button(root, text = "実行", command = main_run)
    run_button.grid()

    run_label = tk.Label(root, text="")
    run_label.grid()

    root.mainloop()
