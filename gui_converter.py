import tkinter as tk
from tkinter import filedialog, messagebox
import os
from convert_e_notice import (
    convert_excel_to_text,
)  # asumsi kamu simpan fungsi ini di file terpisah


def select_file():
    filepath = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx")])
    if filepath:
        output_path = os.path.splitext(filepath)[0] + ".txt"
        convert_excel_to_text(filepath, output_path)
        messagebox.showinfo("Sukses", f"File berhasil dikonversi:\n{output_path}")


# Buat UI
root = tk.Tk()
root.title("Konversi E-notice HF")

frame = tk.Frame(root, padx=20, pady=20)
frame.pack()

label = tk.Label(frame, text="Klik tombol di bawah untuk memilih file Excel:")
label.pack(pady=(0, 10))

button = tk.Button(frame, text="Pilih File Excel", command=select_file)
button.pack()

root.mainloop()
