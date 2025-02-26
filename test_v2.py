import customtkinter as ctk
import tkinter as tk
from PIL import Image, ImageTk, ImageOps
import os

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

# Function to list .xlsx files in the directory
def get_xlsx_files(directory="."):
    return [f for f in os.listdir(directory) if f.endswith(".xlsx")]


# Main Application Setup
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")


# Create the main window
root = ctk.CTk()
root.title("Course Analyzer")
root.geometry("400x400")

# Label for selecting files
select_label = ctk.CTkLabel(root, text="Select files:", font=("Arial", 14, "bold"))
select_label.pack(pady=10)

# Create a scrollable frame
scrollable_frame = ctk.CTkScrollableFrame(root)
scrollable_frame.place(relx=0.5, rely=0.5, anchor="center", relwidth=0.8, relheight=0.7)

# Get list of .xlsx files
data = get_xlsx_files()

# Add labels and photos for each file
for file in data:
    frame = ctk.CTkFrame(scrollable_frame, corner_radius=10)
    frame.pack(fill="x", pady=5, padx=5)
    
    img = capture_photo()
    img_label = tk.Label(frame, image=img, bg="white")
    img_label.image = img
    img_label.pack(side="left", padx=5, pady=5)
    
    file_label = ctk.CTkLabel(frame, text=file, anchor="w", font=("Arial", 12))
    file_label.pack(side="left", padx=10)

# Run the main loop
root.mainloop()