import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from tkinter import Listbox
from PyPDF2 import PdfMerger
import webbrowser
import win32com.client as win32

# Function to merge PDFs based on user input for sorting criteria
def merge_pdfs(pdf_files, output_folder, sorting_criteria, progress_var):
    pdf_groups = {}

    if sorting_criteria == "Merge Without Sorting":
        output_filename = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                       filetypes=[("PDF files", "*.pdf")],
                                                       title="Save Merged PDF As",
                                                       initialdir=output_folder)
        if not output_filename:
            messagebox.showerror("Error", "No file name provided.")
            return
        
        merger = PdfMerger()
        for pdf_file in pdf_files:
            merger.append(pdf_file)
        
        merger.write(output_filename)
        merger.close()
        
        messagebox.showinfo("Success", f"PDFs merged successfully into {output_filename}")
        merged_files_listbox.delete(0, tk.END)
        merged_files_listbox.insert(tk.END, os.path.basename(output_filename))
        if messagebox.askyesno("Open Folder", "Do you want to open the output folder?"):
            webbrowser.open(output_folder)

        # Always ask if the user wants to send the email
        if messagebox.askyesno("Send Email", "Do you want to send the merged PDF via email?"):
            send_email_with_attachment(output_filename)

        return [output_filename]
    
    for pdf_file in pdf_files:
        filename = os.path.basename(pdf_file).lower()

        if sorting_criteria == "Common Start":
            match = re.match(r"([a-zA-Z]+)_.*\.pdf", filename)
        elif sorting_criteria == "Contains Numbers":
            match = re.match(r".*?(\d+).*\.pdf", filename)
        elif sorting_criteria == "Contains Specific Word":
            keyword = user_input_entry.get().lower()
            if keyword in filename:
                match = re.match(rf".*{keyword}.*\.pdf", filename)
            else:
                match = None
        else:
            match = None
        
        if match:
            common_name = match.group(1) if sorting_criteria == "Common Start" else keyword
            if common_name not in pdf_groups:
                pdf_groups[common_name] = []
            pdf_groups[common_name].append(pdf_file)
    
    if not pdf_groups:
        messagebox.showerror("Error", "No files matched the given criteria.")
        return

    total_groups = len(pdf_groups)
    progress_var.set(0)
    progress_bar['maximum'] = total_groups

    merged_files = []
    for i, (common_name, pdf_list) in enumerate(pdf_groups.items(), 1):
        merger = PdfMerger()
        for pdf_file in pdf_list:
            merger.append(pdf_file)
        
        output_filename = filedialog.asksaveasfilename(defaultextension=".pdf",
                                                       filetypes=[("PDF files", "*.pdf")],
                                                       title=f"Save Merged Group '{common_name}' As",
                                                       initialdir=output_folder,
                                                       initialfile=f"{common_name}_merged.pdf")
        if not output_filename:
            messagebox.showerror("Error", "No file name provided.")
            return

        merger.write(output_filename)
        merger.close()
        
        merged_files.append(output_filename)
        progress_var.set(i)
        root.update_idletasks()

    messagebox.showinfo("Success", "PDFs merged successfully!")
    
    merged_files_listbox.delete(0, tk.END)
    for merged_file in merged_files:
        merged_files_listbox.insert(tk.END, os.path.basename(merged_file))
    
    if messagebox.askyesno("Open Folder", "Do you want to open the output folder?"):
        webbrowser.open(output_folder)

    # Ask if the user wants to send the merged files via email
    if messagebox.askyesno("Send Email", "Do you want to send the merged PDF via email?"):
        for merged_file in merged_files:
            send_email_with_attachment(merged_file)

    return merged_files

# Function to send the merged PDF as an attachment via Outlook
def send_email_with_attachment(file_path):
    try:
        outlook = win32.Dispatch('outlook.application')
        mail = outlook.CreateItem(0)
        mail.Attachments.Add(file_path)
        mail.Subject = "Merged PDF File"
        mail.Body = "Please find the merged PDF attached."
        mail.To = ""  # Add recipient here
        mail.Display(True)
    except Exception as e:
        messagebox.showerror("Error", f"Failed to open Outlook: {str(e)}")

# Function to select PDF files
def select_pdf_files():
    pdf_files = filedialog.askopenfilenames(title="Select PDF Files", filetypes=[("PDF Files", "*.pdf")])
    selected_files_listbox.delete(0, tk.END)
    for pdf_file in pdf_files:
        selected_files_listbox.insert(tk.END, os.path.basename(pdf_file))
    return pdf_files

# Function to select output folder
def select_output_folder():
    folder_selected = filedialog.askdirectory(title="Select Output Folder")
    return folder_selected

# Function to start the merging process
def start_merging():
    pdf_files = select_pdf_files()
    if not pdf_files:
        messagebox.showerror("Error", "No PDF files selected.")
        return
    
    output_folder = select_output_folder()
    if not output_folder:
        messagebox.showerror("Error", "No output folder selected.")
        return

    sorting_criteria = sorting_criteria_var.get()

    if sorting_criteria == "Contains Specific Word" and not user_input_entry.get().strip():
        messagebox.showerror("Error", "Please enter a specific word to search for in file names.")
        return

    messagebox.showinfo("Processing", "Merging process started, please wait.")
    merge_pdfs(pdf_files, output_folder, sorting_criteria, progress_var)

# Set up the GUI
root = tk.Tk()
root.title("PDF Merger")

# Make the GUI auto-adjust to screen
root.state('zoomed')  # Auto-adjust to full screen

root.configure(bg="#F0F0F0")

# Styling
style = ttk.Style(root)
style.theme_use("clam")
style.configure("TButton", font=("Helvetica", 12), padding=6)
style.configure("TLabel", font=("Helvetica", 12))
style.configure("TCombobox", font=("Helvetica", 12))
style.configure("TProgressbar", thickness=20, troughcolor='white', background='green')  # Green progress bar

# Frames
top_frame = ttk.Frame(root, padding=20)
top_frame.pack(fill=tk.BOTH, expand=True)

middle_frame = ttk.Frame(root, padding=20)
middle_frame.pack(fill=tk.BOTH, expand=True)

bottom_frame = ttk.Frame(root, padding=20)
bottom_frame.pack(fill=tk.BOTH, expand=True)

# Centering frames
for frame in [top_frame, middle_frame, bottom_frame]:
    frame.grid_rowconfigure(0, weight=1)
    frame.grid_rowconfigure(1, weight=1)
    frame.grid_columnconfigure(0, weight=1)
    frame.grid_columnconfigure(1, weight=1)

# Listbox for selected files
selected_files_label = ttk.Label(top_frame, text="Selected PDF Files:")
selected_files_label.grid(row=0, column=0, columnspan=2, pady=5, sticky="ew")
selected_files_listbox = Listbox(top_frame, width=50, height=5, font=("Helvetica", 10))
selected_files_listbox.grid(row=1, column=0, columnspan=2, pady=10, padx=10)

# Listbox for merged files
merged_files_label = ttk.Label(middle_frame, text="Merged PDF Files:")
merged_files_label.grid(row=0, column=0, columnspan=2, pady=5, sticky="ew")
merged_files_listbox = Listbox(middle_frame, width=50, height=3, font=("Helvetica", 10))
merged_files_listbox.grid(row=1, column=0, columnspan=2, pady=10, padx=10)

# Sorting criteria
sorting_label = ttk.Label(bottom_frame, text="Sorting Criteria:")
sorting_label.grid(row=0, column=0, padx=10, pady=5, sticky="w")
sorting_criteria_var = tk.StringVar(value="Merge Without Sorting")
sorting_dropdown = ttk.Combobox(bottom_frame, textvariable=sorting_criteria_var, values=["Merge Without Sorting", "Common Start", "Contains Numbers", "Contains Specific Word"])
sorting_dropdown.grid(row=0, column=1, padx=10, pady=5, sticky="ew")

# Entry for user input (for specific word search)
user_input_label = ttk.Label(bottom_frame, text="Specific Word:")
user_input_label.grid(row=1, column=0, padx=10, pady=5, sticky="w")
user_input_entry = ttk.Entry(bottom_frame)
user_input_entry.grid(row=1, column=1, padx=10, pady=5, sticky="ew")

# Progress bar
progress_var = tk.DoubleVar()
progress_bar = ttk.Progressbar(bottom_frame, variable=progress_var, mode="determinate")
progress_bar.grid(row=2, column=0, columnspan=2, padx=10, pady=10, sticky="ew")

# Merge button
merge_button = ttk.Button(bottom_frame, text="Start Merging", command=start_merging)
merge_button.grid(row=3, column=0, columnspan=2, pady=10)

root.mainloop()
