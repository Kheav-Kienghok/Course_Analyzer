import customtkinter as ctk
from tkinter import filedialog
import tkinter as tk
from PIL import Image, ImageTk
import os
import sys

from src.file_checker import get_valid_excel_files
from src.data_processing import process_excel_files, save_to_excel

# Main Application Setup
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

app = ctk.CTk()
app.title("Course Analyzer")

# Get the correct path for the icon
if getattr(sys, 'frozen', False):  # If running as an executable
    script_dir = sys._MEIPASS  # Temporary folder for bundled files
else:
    script_dir = os.path.dirname(os.path.abspath(__file__))

icon_path = os.path.join(script_dir, "images", "excel_icon.ico")

try: 
    app.iconbitmap(icon_path)
except Exception as e:
    print(f"Error loading icon: {e}")  # Print error for debugging

# Centering the window on the screen
width, height = 550, 520
screen_width, screen_height = app.winfo_screenwidth(), app.winfo_screenheight()
x, y = (screen_width - width) // 2, (screen_height - height) // 2
app.geometry(f'{width}x{height}+{x}+{y}')
app.resizable(False, False)

app.configure(bg="#ECF2FF")

# ============================== FUNCTIONS ===============================

# Global variable to hold the output directory
excel_files, output_directory = None, None


def select_folder():

    global excel_files  # Declare the variable as global to be accessed in button_callback

    # Get the folder path where the files are located
    folder_path = filedialog.askdirectory()

    if folder_path:
        excel_files = get_valid_excel_files(folder_path)
        file_count_label.configure(text=f"Files Selected: {len(excel_files)}")
        print(excel_files)

        if excel_files:
            show_files_button.configure(state="normal")  # Enable "Show Files" button
        else:
            status_label.configure(text="Status: No valid files found", text_color="red")

def select_output_location():

    global output_directory

    # Get the folder path where the files are located
    output_directory = filedialog.askdirectory()


def generate_result():
    # Check if there are any selected files
    if not excel_files:
        status_label.configure(text="Status: No files to process", text_color="red")
        return

    # Process the Excel files
    final_result = process_excel_files(excel_files)

    # Save the final result to an Excel file
    save_to_excel(final_result, selected_path=output_directory)

    status_label.configure(text="Status: Processing Completed ✅", text_color="green")


def show_files():
    pass

def set_icon(window):
    try:
        window.iconbitmap("images/excel_icon.ico")
    except Exception as e:
        print(f"Error setting icon: {e}")


# Function to load and display photo
def capture_photo():
    image_path = "images/images.png"
    if os.path.exists(image_path):
        img = Image.open(image_path)
        img = img.resize((45, 45), Image.Resampling.LANCZOS)
        img = img.convert("RGBA")
        return ImageTk.PhotoImage(img)
    else:
        img = Image.new('RGBA', (45, 45), color=(200, 200, 200, 255))
        return ImageTk.PhotoImage(img)


def show_files():
    """Function to show a new mini-frame with the file names in a new window."""

    customer_window = ctk.CTkToplevel(app)
    customer_window.title("Course Analyzer")

    # Ensure the main window updates before getting its dimensions
    app.update_idletasks()

    # Get the main window position and size
    main_x = app.winfo_x()
    main_y = app.winfo_y()
    main_width = app.winfo_width()
    main_height = app.winfo_height()

    # Dimensions of the toplevel window
    width, height = 380, 380

    # Calculate the position to center the toplevel inside the main window
    x = main_x + (main_width - width) // 2
    y = main_y + (main_height - height) // 2

    customer_window.geometry(f"{width}x{height}+{x}+{y}")
    customer_window.resizable(False, False)

    # Ensure the toplevel window is visible on top
    customer_window.transient(app)
    customer_window.grab_set()

    # Delay setting the icon for 1000 ms (1 second)
    customer_window.after(500, set_icon, customer_window)


    # Label for selecting files
    select_label = ctk.CTkLabel(customer_window, text="List of selected files:", font=("Arial", 14, "bold"))
    select_label.pack(pady=10)

    # Create a scrollable frame
    scrollable_frame = ctk.CTkScrollableFrame(customer_window)
    scrollable_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.8, relheight=0.7)

    # Get list of .xlsx files
    data = excel_files
    

    # Add labels and photos for each file
    for file in data:
        file = os.path.basename(file) 
        frame = ctk.CTkFrame(scrollable_frame, corner_radius=10)
        frame.pack(fill="x", pady=5, padx=5)
        
        img = capture_photo()
        img_label = tk.Label(frame, image=img, bg="white")
        img_label.image = img
        img_label.pack(side="left", padx=5, pady=5)
        
        file_label = ctk.CTkLabel(frame, text=file, anchor="w", font=("Arial", 14))
        file_label.pack(side="left", padx=10)


# ============================== UI COMPONENTS ===============================

# Title Label
app_title = ctk.CTkLabel(app, text=" Course Analyzer App ", font=("Helvetica", 20, "bold", "underline"), text_color="#2C3E50", bg_color="#ECF2FF")
app_title.pack(pady=8)

# Main Frame Setup
main_frame = ctk.CTkFrame(app, corner_radius=10, fg_color="#D6EAF8")
main_frame.pack(padx=30, pady=15, fill="both", expand=True)

# Upload Section
title_label = ctk.CTkLabel(main_frame, text="Upload Files", font=("Times New Roman", 17, "bold", "underline"), text_color="#154360")
title_label.pack(pady=5)

drop_frame = ctk.CTkFrame(main_frame, width=250, height=150, corner_radius=10, fg_color="#B3E5FC")
drop_frame.pack(pady=10, padx=20, fill="x")
upload_icon = ctk.CTkLabel(drop_frame, text="📁", font=("Arial", 60))
upload_icon.pack(pady=10)
upload_text = ctk.CTkLabel(drop_frame, text="Drag and drop your files here", font=("Arial", 14), text_color="#0D47A1")
upload_text.pack(pady=5)



# File Count Section
file_count_frame = ctk.CTkFrame(main_frame, fg_color="#D6EAF8")
file_count_frame.pack(pady=5, fill="x")
file_count_label = ctk.CTkLabel(file_count_frame, text="Files Selected: 0", font=("Arial", 12), text_color="black")
file_count_label.pack(pady=5)

# Button Section
button_frame = ctk.CTkFrame(main_frame, corner_radius=10, fg_color="#D6EAF8")
button_frame.pack(pady=5, fill="x")

# Select Folder Button
button_select_folder = ctk.CTkButton(
    button_frame, 
    text="🗂 Select Folder", 
    command=select_folder, 
    corner_radius=12, 
    width=180, 
    height=40, 
    font=("Times New Roman", 14, "bold"), 
    fg_color="#2E86C1", 
    hover_color="#1A5276", 
    text_color="white"
)
button_select_folder.pack(side="left", padx=15, pady=7, expand=True)

# Select Output Location Button
button_select_output = ctk.CTkButton(
    button_frame, 
    text="📁 Select Output", 
    command=select_output_location, 
    corner_radius=12, 
    width=180, 
    height=40, 
    font=("Times New Roman", 14, "bold"), 
    fg_color="#2E86C1", 
    hover_color="#1A5276", 
    text_color="white"
)
button_select_output.pack(side="right", padx=15, pady=7, expand=True)

# Generate Button
button_generate = ctk.CTkButton(
    main_frame, 
    text="🚀 Generate Result", 
    command=generate_result, 
    corner_radius=12, 
    width=250, 
    height=45, 
    font=("Arial", 16, "bold"), 
    fg_color="#2196F3", 
    hover_color="#1565C0", 
    text_color="white"
)
button_generate.pack(pady=5, padx=30, fill="x")

# Show Files Button
show_files_button = ctk.CTkButton(
    main_frame, 
    text="📄 Show Files", 
    command=show_files, 
    corner_radius=12, 
    width=250, 
    height=45, 
    font=("Arial", 16, "bold"), 
    fg_color="#2196F3", 
    hover_color="#1565C0", 
    text_color="white",
    state="disabled"  # Disabled until a folder is selected
)
show_files_button.pack(pady=5, padx=30, fill="x")

# Status Label
status_label = ctk.CTkLabel(main_frame, text="Status: Ready to Start ✅", font=("Arial", 12), text_color="green")
status_label.pack(pady=10)

app.mainloop()