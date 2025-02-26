import customtkinter
import os
from tkinterdnd2 import DND_FILES, TkinterDnD
from tkinter import filedialog
from src.file_checker import get_valid_excel_files
from src.data_processing import process_excel_files, save_to_excel

folder_path = None

# Function Callbacks
def button_callback():
    status_label.configure(text="Processing...", text_color="orange")
    print("Generate result clicked")
    
    # Process files if a folder is selected
    if not folder_path:
        status_label.configure(text="Error: No folder selected.", text_color="red")
        return
    
    try:
        # Get the list of valid Excel files from the folder
        excel_files = get_valid_excel_files(folder_path)

        if not excel_files:
            status_label.configure(text="No valid Excel files found.", text_color="red")
            return

        # Process the Excel files
        final_result = process_excel_files(excel_files)

        # Save the final result to an Excel file
        save_to_excel(final_result)

        status_label.configure(text="Processing completed. File saved!", text_color="green")
        print("Processing completed.")
    except (FileNotFoundError, ValueError) as e:
        status_label.configure(text=f"Error: {e}", text_color="red")

def select_folder():

    global folder_path
    folder_path = filedialog.askdirectory()

    valid_file = get_valid_excel_files(folder_path)

    file_count_label.configure(text=f"Files Selected: {len(valid_file)}")

    if folder_path:
        status_label.configure(text=f"Folder Selected: {os.path.basename(folder_path)}", text_color="blue")
        print("Folder selected:", folder_path)
        show_files_button.configure(state="normal")  # Enable the show files button once folder is selected

def select_output_location():
    output_selected = filedialog.askdirectory()
    if output_selected:
        status_label.configure(text=f"Output Folder: {os.path.basename(output_selected)}", text_color="blue")
        print("Output location selected:", output_selected)

# def on_drop(event):
#     file_path = event.data.strip("{}")  # Remove curly braces around file paths
#     update_file_count()

#     # Call the function to check if the dropped file is valid
#     valid_files = get_valid_excel_files(file_list)  # Update the list of valid files
    
#     if valid_files:
#         status_label.configure(text=f"File Added: {os.path.basename(file_path)} - Valid", text_color="green")
#     else:
#         status_label.configure(text=f"File Added: {os.path.basename(file_path)} - Invalid", text_color="red")
    
#     print("File added:", file_path)

def show_files():
    """Function to show a new mini-frame with the file names in a new window."""
    if not folder_path:
        status_label.configure(text="Error: No folder selected.", text_color="red")
        return  # Ensure the folder is selected before proceeding

    # Create a new top-level window to show files
    file_list_window = customtkinter.CTkToplevel(app)
    file_list_window.title("Files in Selected Folder")
    file_list_window.geometry("400x300")  # Set size for the window
    file_list_window.configure(bg="#D6EAF8")

    # Create a title label in the new window
    title_label = customtkinter.CTkLabel(file_list_window, text="Files in Selected Folder", font=("Times New Roman", 17, "bold", "underline"), text_color="#154360")
    title_label.pack(pady=5)

    # Get the valid Excel files
    valid_files = get_valid_excel_files(folder_path)
    print(f"Valid files: {valid_files}")  # Check if valid files are returned

    if not valid_files:
        status_label.configure(text="No valid Excel files found.", text_color="red")
        return

    # Display each valid file name in a label
    for file in valid_files:
        file_name = os.path.basename(file)  # Extract the file name
        file_label = customtkinter.CTkLabel(file_list_window, text=file_name, font=("Arial", 12), text_color="black")
        file_label.pack(pady=3)
    
    # Fix: Don't call block_update_dimensions_event, and remove any update() calls
    file_list_window.update_idletasks()  # This will ensure proper layout


# Main Application Setup
customtkinter.set_appearance_mode("light")
customtkinter.set_default_color_theme("blue")

app = TkinterDnD.Tk()
app.title("Excel Course Analyzer")
app.iconbitmap("./images/excel_icon.ico")

# Centering the window on the screen
width, height = 550, 520
screen_width, screen_height = app.winfo_screenwidth(), app.winfo_screenheight()
x, y = (screen_width - width) // 2, (screen_height - height) // 2
app.geometry(f'{width}x{height}+{x}+{y}')
app.resizable(False, False)

app.configure(bg="#ECF2FF")

# Title Label
app_title = customtkinter.CTkLabel(app, text="Excel Course Analyzer", font=("Helvetica", 20, "bold", "underline"), text_color="#2C3E50", bg_color="#ECF2FF")
app_title.pack(pady=8)

# Main Frame Setup
main_frame = customtkinter.CTkFrame(app, corner_radius=10, fg_color="#D6EAF8")
main_frame.pack(padx=30, pady=15, fill="both", expand=True)

# Upload Section
title_label = customtkinter.CTkLabel(main_frame, text="Upload Files", font=("Times New Roman", 17, "bold", "underline"), text_color="#154360")
title_label.pack(pady=5)

drop_frame = customtkinter.CTkFrame(main_frame, width=250, height=150, corner_radius=10, fg_color="#B3E5FC")
drop_frame.pack(pady=10, padx=20, fill="x")
upload_icon = customtkinter.CTkLabel(drop_frame, text="📁", font=("Arial", 60))
upload_icon.pack(pady=10)
upload_text = customtkinter.CTkLabel(drop_frame, text="Drag and drop your files here", font=("Arial", 14), text_color="#0D47A1")
upload_text.pack(pady=5)

drop_frame.drop_target_register(DND_FILES)
drop_frame.dnd_bind('<<Drop>>')
# drop_frame.dnd_bind('<<Drop>>', on_drop)

# File Count Section
file_count_frame = customtkinter.CTkFrame(main_frame, fg_color="#D6EAF8")
file_count_frame.pack(pady=5, fill="x")
file_count_label = customtkinter.CTkLabel(file_count_frame, text="Files Selected: 0", font=("Arial", 12), text_color="black")
file_count_label.pack(pady=5)

# Button Section
button_frame = customtkinter.CTkFrame(main_frame, corner_radius=10, fg_color="#D6EAF8")
button_frame.pack(pady=5, fill="x")

# Select Folder Button
button_select_folder = customtkinter.CTkButton(
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
button_select_output = customtkinter.CTkButton(
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
button_generate = customtkinter.CTkButton(
    main_frame, 
    text="🚀 Generate Result", 
    command=button_callback, 
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
show_files_button = customtkinter.CTkButton(
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
status_label = customtkinter.CTkLabel(main_frame, text="Status: Ready to Start ✅", font=("Arial", 12), text_color="green")
status_label.pack(pady=10)

app.mainloop()
