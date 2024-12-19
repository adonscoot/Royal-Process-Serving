import tkinter as tk
from tkinter import messagebox
from tkinter import ttk
from tkcalendar import DateEntry
from docx import Document
import os
from datetime import datetime

# Load previously saved addresses from the file
def load_addresses():
    try:
        with open("addresses.txt", "r") as file:
            return [line.strip() for line in file.readlines()]
    except FileNotFoundError:
        return []

# Save the address to the file, ensuring there are no duplicates
def save_address(address):
    try:
        with open("addresses.txt", "r") as file:
            existing_addresses = [line.strip() for line in file.readlines()]
    except FileNotFoundError:
        existing_addresses = []

    # Check if the address is already in the list of saved addresses
    if address not in existing_addresses:
        with open("addresses.txt", "a") as file:
            file.write(f"{address}\n")
    else:
        messagebox.showinfo("Duplicate Address", "This address is already saved.")

# Update the Combobox with filtered addresses as the user types
def update_address_dropdown(event):
    typed_address = address.get()
    matching_addresses = [addr for addr in saved_addresses if typed_address.lower() in addr.lower()]
    address['values'] = matching_addresses

def save_affidavit():
    case_number_value = case_number.get()
    party_value = party.get()
    address_value = address.get()
    time_served_value = time_served.get()
    date_served_value = date_served.get()
    paper_type_value = paper_type.get()
    served_by_value = served_by.get()  # Get who served the paper

    # Data validation - check if all fields are filled
    if not case_number_value or not party_value or not address_value or not time_served_value or not date_served_value or not paper_type_value or not served_by_value:
        messagebox.showerror("Error", "All fields must be filled out!")
        return
    
    # Check if a file with the same name already exists
    filename = f"Affidavit_{case_number_value}.docx"
    if os.path.exists(filename):
        response = messagebox.askyesno("File Exists", f"The file {filename} already exists. Do you want to overwrite it?")
        if not response:
            return  # If the user chooses not to overwrite, exit the function
    
    # Load the template Word document
    template_path = "AffidavitTemplate.docx"  # Path to the Word template file
    if not os.path.exists(template_path):
        messagebox.showerror("Error", "Template file not found!")
        return

    template = Document(template_path)
    
    # Replace placeholders with actual data
    for paragraph in template.paragraphs:
        if "<<CaseNumber>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<CaseNumber>>", case_number_value)
        if "<<Party>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<Party>>", party_value)
        if "<<Address>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<Address>>", address_value)
        if "<<TimeServed>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<TimeServed>>", time_served_value)
        if "<<DateServed>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<DateServed>>", date_served_value)
        if "<<PaperType>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<PaperType>>", paper_type_value)
        if "<<ServedBy>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<ServedBy>>", served_by_value)
        
        # Add e-signature and date
        if "<<Signature>>" in paragraph.text:
            paragraph.text = paragraph.text.replace("<<Signature>>", served_by_value)
        
        if "<<DateSigned>>" in paragraph.text:
            current_date = datetime.now().strftime("%m/%d/%Y")
            paragraph.text = paragraph.text.replace("<<DateSigned>>", current_date)

    # Save the filled affidavit as a new Word document
    template.save(filename)

    # Confirmation message
    messagebox.showinfo("Success", f"Affidavit saved as {filename}")

    # Automatically open the Word document (works on Windows)
    os.startfile(filename)

    # Enable the "Create New Affidavit" button after saving
    create_new_button.grid(row=8, column=0, columnspan=2, pady=10)

# Function to create a new affidavit (clear fields and refresh the address dropdown)
def create_new_affidavit():
    case_number.delete(0, tk.END)
    party.delete(0, tk.END)
    address.set("")  # Clear the address field
    time_served.delete(0, tk.END)
    date_served.set_date("")  # Reset the date entry
    paper_type.set("")  # Reset the paper type combobox
    served_by.set("John Doe")  # Default served by value
    # Refresh the address dropdown to load the updated addresses
    saved_addresses = load_addresses()
    address['values'] = saved_addresses

# Set up the main application window
root = tk.Tk()
root.title("Affidavit Creator")

# Increase the window size
root.geometry("800x600")  # Set a bigger default window size

# Load previously saved addresses
saved_addresses = load_addresses()

# Create the notebook (tabs) widget
notebook = ttk.Notebook(root)

# Create frames for each tab
tab_affidavit = ttk.Frame(notebook)
tab_manage_affidavits = ttk.Frame(notebook)

# Add tabs to the notebook
notebook.add(tab_affidavit, text="Create Affidavit")
notebook.add(tab_manage_affidavits, text="Manage Affidavits")

# Grid layout for the notebook (tabs)
notebook.grid(row=0, column=0, padx=10, pady=10)

# Configure rows and columns to expand with window size
root.grid_rowconfigure(0, weight=1, uniform="equal")  # Make row 0 stretchable
root.grid_columnconfigure(0, weight=1, uniform="equal")  # Make column 0 stretchable

# Configure grid for the tab_affidavit frame
tab_affidavit.grid_rowconfigure(0, weight=1)
tab_affidavit.grid_rowconfigure(1, weight=1)
tab_affidavit.grid_rowconfigure(2, weight=1)
tab_affidavit.grid_rowconfigure(3, weight=1)
tab_affidavit.grid_rowconfigure(4, weight=1)
tab_affidavit.grid_rowconfigure(5, weight=1)
tab_affidavit.grid_rowconfigure(6, weight=1)
tab_affidavit.grid_rowconfigure(7, weight=1)
tab_affidavit.grid_columnconfigure(0, weight=1)
tab_affidavit.grid_columnconfigure(1, weight=3)

# Create a larger font for readability
large_font = ('Arial', 16)  # Font size 16 for better readability

# Tab 1 - Affidavit Details (Single Tab for Affidavit Creation)
tk.Label(tab_affidavit, text="Case Number:", font=large_font).grid(row=0, column=0, sticky="w", padx=10, pady=10)
case_number = tk.Entry(tab_affidavit, font=large_font)
case_number.grid(row=0, column=1, sticky="ew", padx=10, pady=10)

tk.Label(tab_affidavit, text="Party to be Served:", font=large_font).grid(row=1, column=0, sticky="w", padx=10, pady=10)
party = tk.Entry(tab_affidavit, font=large_font)
party.grid(row=1, column=1, sticky="ew", padx=10, pady=10)

tk.Label(tab_affidavit, text="Address:", font=large_font).grid(row=2, column=0, sticky="w", padx=10, pady=10)
address = ttk.Combobox(tab_affidavit, font=large_font)
address.grid(row=2, column=1, sticky="ew", padx=10, pady=10)
address['values'] = saved_addresses  # Load previous addresses into the ComboBox

# Bind the event to update the dropdown as user types
address.bind("<KeyRelease>", update_address_dropdown)

tk.Label(tab_affidavit, text="Time Served:", font=large_font).grid(row=3, column=0, sticky="w", padx=10, pady=10)
time_served = tk.Entry(tab_affidavit, font=large_font)
time_served.grid(row=3, column=1, sticky="ew", padx=10, pady=10)

tk.Label(tab_affidavit, text="Date Served:", font=large_font).grid(row=4, column=0, sticky="w", padx=10, pady=10)

# Add the DateEntry widget for calendar selection
date_served = DateEntry(tab_affidavit, font=large_font, date_pattern="mm/dd/yyyy")
date_served.grid(row=4, column=1, sticky="ew", padx=10, pady=10)

tk.Label(tab_affidavit, text="Paper Type:", font=large_font).grid(row=5, column=0, sticky="w", padx=10, pady=10)
paper_type = ttk.Combobox(tab_affidavit, font=large_font, values=["Summons", "Subpoena", "Writ", "Other"])
paper_type.grid(row=5, column=1, sticky="ew", padx=10, pady=10)

tk.Label(tab_affidavit, text="Served By:", font=large_font).grid(row=6, column=0, sticky="w", padx=10, pady=10)
served_by = ttk.Combobox(tab_affidavit, font=large_font, values=["John Doe", "Jane Smith", "Bob Johnson"])
served_by.grid(row=6, column=1, sticky="ew", padx=10, pady=10)

# Button to save affidavit
save_button = ttk.Button(tab_affidavit, text="Save Affidavit", command=save_affidavit, style='TButton', width=20, padding=10)
save_button.grid(row=7, column=0, columnspan=2, pady=10)

# Button to create a new affidavit
create_new_button = ttk.Button(tab_affidavit, text="Create New Affidavit", command=create_new_affidavit, style='TButton', width=20, padding=10)
create_new_button.grid(row=8, column=0, columnspan=2, pady=10)

root.mainloop()
