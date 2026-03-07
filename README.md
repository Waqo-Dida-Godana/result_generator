# Kenyan CBC Student Results System

A desktop application built with Python and Tkinter for managing student marks under the Competency Based Curriculum (CBC).

## 🚀 Features
- **Secure Login System**: Separate access for Admin and Teachers.
- **Student Management**: Add, delete, and view student records.
- **CBC Grading**: Automatic calculation of Total, Average, and CBC Levels (EE, ME, AE, BE).
- **Automatic Ranking**: Students are automatically ranked based on performance.
- **Data Export**:
  - Export full class results to **Excel** (`.xlsx`).
  - Export full class results to **PDF**.
  - Generate individual **Student Report Cards** in PDF format.
- **Database Support**: Integration with MySQL (graceful fallback to in-memory mode if MySQL is offline).
- **Desktop Ready**: Can be converted to a standalone `.exe` file.

## 🛠 Prerequisites
Ensure you have Python installed. You will also need the following libraries:
```bash
pip install pandas reportlab fpdf openpyxl mysql-connector-python pyinstaller
```

## 📖 How to Run
1. Navigate to the project folder.
2. Run the application:
   ```bash
   python cbc_student_app.py
   ```

## 🔐 Login Credentials
| Role | Username | Password |
| :--- | :--- | :--- |
| **Admin** | `admin` | `1234` |
| **Teacher** | `teacher` | `1111` |

## 📦 Building the Desktop App (.EXE)
To create a standalone executable for Windows:
1. Run the provided `build_exe.bat` file.
2. Alternatively, run this command in your terminal:
   ```bash
   pyinstaller --onefile --windowed cbc_student_app.py
   ```
3. Your software will be available in the `dist` folder.

## 📊 CBC Grading Scale
- **EE (Exceeding Expectations)**: 80 - 100
- **ME (Meeting Expectations)**: 65 - 79
- **AE (Approaching Expectations)**: 50 - 64
- **BE (Below Expectations)**: 0 - 49
