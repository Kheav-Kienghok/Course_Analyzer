import os
import tkinter as tk

# Get the directory of the currently running script
script_dir = os.path.dirname(os.path.realpath(__file__))

# Form the path to the icon file relative to the script
icon_path = os.path.join(script_dir, "images", "excel_icon.ico")

# Print the icon path for debugging purposes
print(f"Using icon path: {icon_path}")

# Create the Tkinter window and set the icon
app = tk.Tk()
app.iconbitmap(icon_path)

app.mainloop()
