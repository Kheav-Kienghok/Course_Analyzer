# 📘 Course Analyzer Application - Documentation

## 📌 Project Overview  
The **Course Analyzer Application** automates student course data analysis from Excel files, streamlining the identification of missing courses per student and generating structured reports. The application provides both **command-line** and **GUI-based** execution for ease of use.

### ✨ Key Features  
- **Automated Excel Processing** – Reads and processes `.xlsx` files efficiently.  
- **Student Course Analysis** – Identifies courses that students have not taken per major.  
- **Report Generation** – Outputs structured results in `result.xlsx`.  
- **User-Friendly GUI** – A graphical interface for simplified user interaction.  

---

## 🚀 Installation & Setup  

### 🔹 Prerequisites  
- **Python** `3.13.2`  
- Required dependencies listed in `requirements.txt`  

### 🔹 Installation Steps  

#### 1️⃣ Clone the Repository  
```bash
git clone https://github.com/Kheav-Kienghok/Course_Analyzer.git
cd Course_Analyzer
```

#### 2️⃣ Create a Virtual Environment *(Recommended)*  
```bash
python -m venv venv
source venv/bin/activate  # macOS/Linux
venv\Scripts\activate     # Windows
```

#### 3️⃣ Install Dependencies  
```bash
pip install -r requirements.txt
```

#### 4️⃣ Run the Application  

- **Command-Line Mode:**  
  ```bash
  python main.py
  ```

- **Graphical Interface (GUI) Mode:**  
  ```bash
  python GUI_v1.py
  ```

---

## 📂 Project Structure  

```bash
📂 Course_Analyzer/
├── 📁 images/                 # (Optional) Stores images for GUI or documentation
├── 📁 output/                 # Stores generated reports (e.g., result.xlsx)
├── 📁 src/                    # Core application logic
│   ├── 📜 __init__.py         # Marks src as a package
│   ├── 📜 data_processing.py  # Handles data analysis & report generation
│   ├── 📜 file_checker.py     # Validates .xlsx file format
│   ├── 📜 file_reader.py      # Extracts student data from Excel files
├── 📜 .gitignore              # Excludes unnecessary files from version control
├── 🎨 GUI_v1.py               # GUI version of the application
├── 🚀 main.py                 # Application entry point
├── 📖 README.md               # Documentation & usage instructions
├── 📜 requirements.txt        # List of dependencies
```

---

## 📊 Sample Output (`result.xlsx`)  

The generated **Excel report** (`result.xlsx`) provides a structured overview of students who have not taken specific courses per major.

| Course Name  | Major: CS | Major: IT | Major: EE | Total  |
|-------------|----------|----------|----------|-------|
| Math 101    | 5        | 3        | 7        | 15    |
| CS 201      | 2        | 4        | 3        | 9     |
| Physics 102 | 6        | 2        | 5        | 13    |
| IT 305      | 1        | 3        | 6        | 10    |

The **"Total"** column represents the total number of students across all majors who haven't taken the respective course.

---

## 🛠️ Application Components  

### 📂 `src/` - Core Logic  
This folder contains the core logic of the application.  

- **`file_checker.py`**  
  - Verifies whether the selected file is a valid `.xlsx` file.  
  - Returns the file path if valid or prompts an error otherwise.  

- **`file_reader.py`**  
  - Reads and extracts student course data from the Excel file.  
  - Filters out unnecessary headers (e.g., "Report of").  

- **`data_processing.py`**  
  - Orchestrates data processing for course analysis.  
  - Identifies missing courses per student based on their major.  
  - Generates `result.xlsx` with multiple worksheets:
    - **All Courses Worksheet** – Summary of all missing courses categorized by major.  
    - **General Education Worksheet** – List of missing general education courses per student.  
    - **Major Core & Electives Worksheet** – Identifies missing major core and elective courses per student.  

---

### 📂 `output/` - Processed Reports  
- Stores generated `result.xlsx` reports.  
- Each report contains structured student course completion data.

---

### 🎨 GUI (`GUI_v1.py`)  
- Provides a **graphical user interface** for non-technical users.  
- Allows file selection without command-line interaction.  
- Presents results interactively and user-friendly.  

---

### 🚀 `main.py` - Application Entry Point  
- Orchestrates the full workflow from file validation to data processing.  
- Executes via the command line using:  
  ```bash
  python main.py
  ```

---

## 🐞 Known Issues & Solutions  

### 1️⃣ Excel File Format Requirements  
The application expects the input **Excel file** (`.xlsx`) to follow this structure:

| **Report of** |      |      |      |      |      |
|--------------|------|------|------|------|------|
| **Srl No**  | **Name** | **Math101** | **CS201** | **IT305** | **EE202** |
| 1           | John  | A    | RG   | B+   | C    |
| 2           | Alice | C    | B    |      | B+   |
| 3           | Bob   | B+   | A    | C    |      |

🔹 **Issue:** If the file structure differs from this format, the program may fail to process the data.  
✅ **Solution:** Ensure that:
- The first row contains `"Report of"`.
- Data begins from the second row.
- Course columns contain student grades or blank entries for missing courses.

---

## 📜 License  
This project is licensed under the **MIT License**. See [`LICENSE.md`](./LICENSE.md) for full details.

