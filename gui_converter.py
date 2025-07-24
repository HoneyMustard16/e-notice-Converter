import tkinter as tk
from tkinter import filedialog, messagebox
from convert_e_notice import convert_excel_to_text


class ConverterGUI(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("e-notice converter")
        self.geometry("450x200")

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()

        # Input file selection
        tk.Label(self, text="Select Excel File:").pack(pady=(15, 5))
        input_frame = tk.Frame(self)
        input_frame.pack()
        tk.Entry(input_frame, textvariable=self.input_path, width=50).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        tk.Button(input_frame, text="Browse", command=self.browse_input).pack(
            side=tk.LEFT
        )

        # Output file selection
        tk.Label(self, text="Select Output File Location:").pack(pady=(15, 5))
        output_frame = tk.Frame(self)
        output_frame.pack()
        tk.Entry(output_frame, textvariable=self.output_path, width=50).pack(
            side=tk.LEFT, padx=(0, 5)
        )
        tk.Button(output_frame, text="Browse", command=self.browse_output).pack(
            side=tk.LEFT
        )

        # Convert button
        tk.Button(self, text="Convert", command=self.convert).pack(pady=20)

    def browse_input(self):
        file_path = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            self.input_path.set(file_path)

    def browse_output(self):
        file_path = filedialog.asksaveasfilename(
            defaultextension=".txt", filetypes=[("Text files", "*.txt")]
        )
        if file_path:
            self.output_path.set(file_path)

    def convert(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()
        if not input_file or not output_file:
            messagebox.showerror("Error", "Please select both input and output files.")
            return
        try:
            convert_excel_to_text(input_file, output_file)
            messagebox.showinfo(
                "Success", f"File converted successfully:\n{output_file}"
            )
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")


if __name__ == "__main__":
    app = ConverterGUI()
    app.mainloop()
