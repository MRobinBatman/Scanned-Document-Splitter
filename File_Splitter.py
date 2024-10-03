# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import filedialog, messagebox, Menu
import os
import fitz  # PyMuPDF
from PIL import Image, ImageTk
import csv

class PDFSplitterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Splitter")
        self.root.state('zoomed')
        self.pdf_path = None
        self.pdf_document = None
        self.page_images = []
        self.page_names = []
        self.year_prefix = tk.StringVar()
        self.custom_year = tk.StringVar()
        
        self.setup_ui()

        # Bind mouse wheel event to the root window
        self.root.bind_all("<MouseWheel>", self.on_mouse_wheel)
        
        # Bind the Enter key press event to submit names for all relevant entry fields
        self.bind_enter_key(self.new_name_entry)
        for entry in self.page_names:
            self.bind_enter_key(entry)


    def bind_enter_key(self, entry):
        entry.bind("<Return>", lambda event: self.update_name())
        
    def setup_ui(self):
        menubar = tk.Menu(self.root)
    
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Load", command=self.load_page_names)
        file_menu.add_command(label="Save", command=self.save_page_names)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.root.quit)
        menubar.add_cascade(label="File", menu=file_menu)
    
        self.root.config(menu=menubar)
    
        top_frame = tk.Frame(self.root)
        top_frame.pack(pady=10)
    
        self.file_button = tk.Button(top_frame, text="Select PDF", command=self.select_pdf)
        self.file_button.pack(side="left", padx=5)
    
        self.process_button = tk.Button(top_frame, text="Process", command=self.process_document)
        self.process_button.pack(side="left", padx=5)
    
        self.clear_button = tk.Button(top_frame, text="Clear", command=self.clear)
        self.clear_button.pack(side="left", padx=5)
    
        self.year_button_2022 = tk.Button(top_frame, text="2022", command=lambda: self.set_year_prefix("2022"))
        self.year_button_2022.pack(side="left", padx=5)
    
        self.year_button_2023 = tk.Button(top_frame, text="2023", command=lambda: self.set_year_prefix("2023"))
        self.year_button_2023.pack(side="left", padx=5)
    
        self.year_button_2024 = tk.Button(top_frame, text="2024", command=lambda: self.set_year_prefix("2024"))
        self.year_button_2024.pack(side="left", padx=5)
    
        self.year_button_2025 = tk.Button(top_frame, text="2025", command=lambda: self.set_year_prefix("2025"))
        self.year_button_2025.pack(side="left", padx=5)
    
        self.custom_year_entry = tk.Entry(top_frame, textvariable=self.custom_year, width=10)
        self.custom_year_entry.pack(side="left", padx=5)
    
        self.custom_year_button = tk.Button(top_frame, text="Set Custom Year", command=self.set_custom_year_prefix)
        self.custom_year_button.pack(side="left", padx=5)
    
        self.info_label = tk.Label(self.root, text="")
        self.info_label.pack(pady=10)
    
        self.frame = tk.Frame(self.root)
        self.frame.pack(fill="both", expand=True)
    
        self.canvas = tk.Canvas(self.frame)
        self.scroll_y = tk.Scrollbar(self.frame, orient="vertical", command=self.canvas.yview)
        self.scroll_y.pack(side="right", fill="y")
    
        self.canvas.config(yscrollcommand=self.scroll_y.set)
        self.canvas.pack(side="left", fill="both", expand=True)
    
        self.inner_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
    
        self.canvas.bind("<MouseWheel>", self.on_mouse_wheel)
    
        self.submitted_names_listbox = tk.Listbox(self.frame, width=30)
        self.submitted_names_listbox.pack(side="right", fill="y", pady=5, padx=(0, 10))
    
        self.new_name_entry = tk.Entry(self.frame)
        self.new_name_entry.pack(side="right", fill="x", pady=5, padx=(0, 10))
    
        self.update_name_button = tk.Button(self.frame, text="Update Name", command=self.update_name)
        self.update_name_button.pack(side="right", pady=5, padx=(0, 10))
    
        self.name_count_label = tk.Label(self.frame, text="Count: 0")
        self.name_count_label.pack(side="right", pady=5, padx=(0, 10))
    
            

    def set_year_prefix(self, year):
        self.year_prefix.set(year)

    def set_custom_year_prefix(self):
        self.year_prefix.set(self.custom_year.get())

    def select_pdf(self):
        self.pdf_path = filedialog.askopenfilename(filetypes=[("PDF Files", "*.pdf")])
        if self.pdf_path:
            self.load_pdf()

    def load_pdf(self):
        self.pdf_document = fitz.open(self.pdf_path)
        self.page_images = [self.convert_page_to_image(page) for page in self.pdf_document]

        self.info_label.config(text=f"Document: {os.path.basename(self.pdf_path)}\n"
                                    f"Pages: {len(self.pdf_document)}\n"
                                    f"Size: {os.path.getsize(self.pdf_path) / 1024:.2f} KB")

        self.display_pages()

    def convert_page_to_image(self, page):
        pix = page.get_pixmap()
        img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
        return img


    def display_pages(self):
        # Destroy any existing widgets in the inner frame
        for widget in self.inner_frame.winfo_children():
            widget.destroy()
    
        self.page_names = [tk.StringVar() for _ in self.page_images]
    
        total_height = 0  # Variable to store the total height of all images
    
        for i, (page_num, image) in enumerate(zip(range(1, len(self.page_images) + 1), self.page_images), start=1):
            img = ImageTk.PhotoImage(image)
            panel = tk.Label(self.inner_frame, image=img, bg="#FFFFFF")
            panel.image = img
            panel.grid(row=i, column=0, padx=(10, 0), pady=(10, 5), sticky="w")
    
            page_label = tk.Label(self.inner_frame, text=f"Page {page_num}:", bg="#FFFFFF")
            page_label.grid(row=i, column=1, padx=(10, 0), pady=(10, 5), sticky="w")
    
            entry = tk.Entry(self.inner_frame, textvariable=self.page_names[i - 1], width=50)
            entry.grid(row=i, column=2, padx=(0, 10), pady=(10, 5), sticky="w")
    
            # Bind the Enter key press event to the submit_name method
            entry.bind("<Return>", lambda event, idx=i-1: self.submit_name(idx))
    
            submit_button = tk.Button(self.inner_frame, text="Submit", command=lambda idx=i-1: self.submit_name(idx),
                                      bg="#007BFF", fg="white", relief="flat")
            submit_button.grid(row=i, column=3, padx=(0, 10), pady=(10, 5), sticky="w")
    
            total_height += image.height + 20
    
    
        
        print(f"Debug: Total height required for all pages: {total_height}")  # Debug statement for total height
    
        # Set the size of the inner frame to accommodate all pages
        self.inner_frame.config(width=self.canvas.winfo_width(), height=total_height)
    
        # Update the canvas scroll region
        self.canvas.config(scrollregion=self.canvas.bbox("all"))
    
        # Set the canvas height to match the total height required for all pages
        self.canvas.config(height=total_height)
    
        print("Debug: After setting canvas size -")
        print(f"Debug: Canvas Size - Width: {self.canvas.winfo_width()}, Height: {self.canvas.winfo_height()}")
    

    def process_document(self):
        names = [name.get() for name in self.page_names]
        if any(name == "" for name in names):
            messagebox.showerror("Invalid Input", "Please enter a name for each page.")
            return

        if len(names) != len(set(names)):
            messagebox.showerror("Invalid Input", "Page names must be unique.")
            return

        output_folder = os.path.splitext(os.path.basename(self.pdf_path))[0]
        pdf_directory = os.path.dirname(self.pdf_path)
        output_folder_path = os.path.join(pdf_directory, output_folder)
        os.makedirs(output_folder_path, exist_ok=True)

        year_prefix = self.year_prefix.get()
        for i, name in enumerate(names):
            output_pdf = fitz.open()
            output_pdf.insert_pdf(self.pdf_document, from_page=i, to_page=i)

            if year_prefix:
                prefixed_name = f"ATCH_{year_prefix}_{name}.pdf"
            else:
                prefixed_name = f"ATCH_{name}.pdf"
            
            output_path = os.path.join(output_folder_path, prefixed_name)
            output_pdf.save(output_path)

        self.save_page_names()
        messagebox.showinfo("Success", f"PDF split successfully into folder: {output_folder_path}")

    def submit_name(self, index):
        name = self.page_names[index].get()
        if name == "":
            messagebox.showerror("Invalid Input", "Please enter a name for the page.")
        else:
            self.submitted_names_listbox.delete(0, tk.END)  # Clear the list
            for i, name in enumerate(self.page_names):
                display_name = f"ATCH_{self.year_prefix.get()}_{name.get()}" if self.year_prefix.get() else f"ATCH_{name.get()}"
                self.submitted_names_listbox.insert(tk.END, display_name)  # Update list with new names
            messagebox.showinfo("Success", f"Name for page {index + 1} submitted: {name}")

    def update_name(self):
        selected_index = self.submitted_names_listbox.curselection()
        if selected_index:
            new_name = self.new_name_entry.get()
            if new_name == "":
                messagebox.showerror("Invalid Input", "Please enter a new name.")
                return
            year_prefix = self.year_prefix.get()
            if year_prefix:
                prefixed_name = f"ATCH_{year_prefix}_{new_name}"
            else:
                prefixed_name = f"ATCH_{new_name}"
            self.submitted_names_listbox.delete(selected_index)
            self.submitted_names_listbox.insert(selected_index, prefixed_name)
            self.page_names[selected_index[0]].set(new_name)
            self.update_name_count()
            messagebox.showinfo("Success", f"Name updated to: {prefixed_name}")

    def update_name_count(self):
        count = self.submitted_names_listbox.size()
        self.name_count_label.config(text=f"Count: {count}")

    def save_page_names(self):
        if not self.page_names:
            return
    
        pdf_directory = os.path.dirname(self.pdf_path)
    
        csv_path = os.path.join(pdf_directory, os.path.splitext(os.path.basename(self.pdf_path))[0] + ".csv")
    
        with open(csv_path, 'w', newline='') as file:
            writer = csv.writer(file)
            writer.writerow(["Page Number", "Name", "Filename"])  # Include filename column
            for i, name_var in enumerate(self.page_names):
                year_prefix = self.year_prefix.get()
                filename = f"ATCH_{year_prefix}_{name_var.get()}.pdf" if year_prefix else f"ATCH_{name_var.get()}.pdf"
                writer.writerow([i + 1, name_var.get(), filename])


    def load_page_names(self):
        if not self.pdf_path:
            messagebox.showerror("Error", "Please load a PDF file first.")
            return
    
        pdf_directory = os.path.dirname(self.pdf_path)
        csv_path = filedialog.askopenfilename(initialdir=pdf_directory, filetypes=[("CSV Files", "*.csv")])
    
        if csv_path:
            with open(csv_path, 'r') as file:
                reader = csv.reader(file)
                next(reader)  # Skip header
                for i, row in enumerate(reader):
                    if i < len(self.page_names):
                        self.page_names[i].set(row[1])
                        display_name = f"ATCH_{self.year_prefix.get()}_{row[1]}" if self.year_prefix.get() else f"ATCH_{row[1]}"
                        self.submitted_names_listbox.insert(i, display_name)
            self.update_name_count()


    def clear(self):
        self.pdf_path = None
        self.pdf_document = None
        self.page_images = []
        self.page_names = []
        self.year_prefix.set("")
        self.custom_year.set("")
        self.info_label.config(text="")
        self.submitted_names_listbox.delete(0, tk.END)
        self.new_name_entry.delete(0, tk.END)
        self.name_count_label.config(text="Count: 0")
        self.inner_frame.destroy()
        self.inner_frame = tk.Frame(self.canvas)
        self.canvas.create_window((0, 0), window=self.inner_frame, anchor="nw")
        self.canvas.config(scrollregion=self.canvas.bbox("all"))


    def on_mouse_wheel(self, event):
        delta = event.delta
        if event.num == 5 or event.delta == -120:
            self.canvas.yview_scroll(1, "units")
        elif event.num == 4 or event.delta == 120:
            self.canvas.yview_scroll(-1, "units")


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFSplitterApp(root)
    root.mainloop()

           
