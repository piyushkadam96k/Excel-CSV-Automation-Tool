# ✨ Excel / CSV Automation Tool ✨

> A small GUI utility to merge, clean and generate PDF reports from Excel/CSV files 📊📄

-------------------------------------------------------------


Excel / CSV Automation Tool v1.0     
🧰 Built with Python & CustomTkinter 


## ✨ Features

- 🔗 **Merge multiple Excel/CSV files** (adds `Source_File` column)
- 🧹 **Clean data** (trim whitespace, coerce numeric columns)
- 📋 **Remove duplicates** and produce smart summaries
- 📊 **Generate beautiful charts** (bar charts, value distributions)
- 🖨️ **PDF report generation** with statistics and visualizations
- 🖥️ **Intuitive GUI** built with `customtkinter`
- ⚡ **Real-time progress tracking** with visual feedback
- 📁 **Organized output** in timestamped folders

🚀 Quick Start

### 📦 Installation

- **Install dependencies:**

```powershell
pip install -r requirements.txt
```

### ▶️ Running the App

**Option 1: Direct Python (Always Works)** 🐍
```powershell
python app.py
```

**Option 2: Batch File Launcher** 💨
```powershell
Excel-CSV-Tool.bat
```

**Option 3: Desktop Shortcut** 🖱️ (Recommended)
- Look for `Excel-CSV-Tool.lnk` on your Desktop
- Double-click to launch instantly!

---

## 📌 Create Desktop Shortcut (Easy Methods)

### Method 1️⃣: Quick PowerShell Command (Fastest) ⚡

Copy and paste this in **PowerShell**:

```powershell
$DesktopPath = [Environment]::GetFolderPath('Desktop')
$WshShell = New-Object -ComObject WScript.Shell
$Shortcut = $WshShell.CreateShortcut("$DesktopPath\Excel-CSV-Tool.lnk")
$Shortcut.TargetPath = "C:\Users\***\AppData\Local\Programs\Python\Python311\python.exe"
$Shortcut.Arguments = "`"c:\Users\***\OneDrive\Desktop\working projects by me\Excel csv work\app.py`""
$Shortcut.WorkingDirectory = "c:\Users\***\OneDrive\Desktop\***\Excel csv work"
$Shortcut.IconLocation = "C:\Users\***\AppData\Local\Programs\Python\Python311\python.exe,0"
$Shortcut.Save()
Write-Host "✅ Desktop shortcut created successfully!"
```

### Method 2️⃣: Manual Shortcut Creation (Windows GUI) 🖱️

1. **Right-click** on your Desktop → **New** → **Shortcut**
2. **Paste this in the location field:**
   ```
   C:\Users\kadam\AppData\Local\Programs\Python\Python311\python.exe "c:\Users\kadam\OneDrive\Desktop\working projects by me\Excel csv work\app.py"
   ```
3. **Click Next** ➡️
4. **Name it:** `Excel-CSV-Tool` 📝
5. **Click Finish** ✅
6. *(Optional)* Right-click shortcut → **Properties** → **Advanced** → Check **Run as administrator** (if needed)

### Method 3️⃣: Use Batch File 🔧

Already created for you: `Excel-CSV-Tool.bat`
- Right-click → **Send to** → **Desktop (create shortcut)**
- Or just double-click the `.bat` file to run immediately!

---

## 🎯 Desktop Shortcut Tips

| 💡 Tip | Details |
|--------|---------|
| **Pin to Taskbar** | Right-click shortcut → Pin to Taskbar for quick access |
| **Change Icon** | Right-click → Properties → Change Icon (choose an icon from `python.exe` or custom `.ico`) |
| **Run Minimized Console** | Right-click → Properties → Advanced → Check "Run with reduced window" |
| **Keyboard Shortcut** | Right-click → Properties → Shortcut tab → **Shortcut key** (e.g., `Ctrl+Alt+E`) |

---

## 📂 How It Works

1. **Select Files** 📁 → Choose your Excel/CSV files
2. **Process** ⚙️ → Data gets cleaned & merged
3. **Review** 👀 → See previews in real-time
4. **Generate Reports** 📊 → Automatic PDF + Excel exports

The app creates an `output/` folder with timestamped subfolders for each run.

---

## 🔧 Troubleshooting

| ❌ Issue | ✅ Solution |
|---------|-----------|
| **`customtkinter` import error** | Run: `pip install customtkinter` |
| **Excel file won't open** | Install: `pip install openpyxl xlrd` |
| **PDF generation fails** | Install: `pip install fpdf2` |
| **Shortcut won't work** | Check Python path: `python --version` in PowerShell |
| **"No readable data found"** | Ensure files have proper headers and data |
| **Charts not displaying** | Install: `pip install matplotlib pillow` |

---

## 📦 Dependencies

All required packages are in `requirements.txt`:

```
pandas              # 📊 Data manipulation
matplotlib          # 📈 Chart generation
fpdf2               # 🖨️ PDF creation
customtkinter       # 🖥️ GUI framework
openpyxl            # 📁 Excel support
xlrd                # 📄 Legacy Excel reader
Pillow              # 🖼️ Image handling
```

Install all at once:
```powershell
pip install -r requirements.txt
```

---

## ⚖️ License & Redistribution

This project is licensed under **Creative Commons Attribution-NonCommercial 4.0 (CC BY-NC 4.0)** 🔐

- ✅ **You can:** Use, modify, and redistribute for **personal/educational purposes**
- ❌ **You cannot:** Use for commercial purposes without permission
- 📝 **You must:** Give appropriate credit to the original author

See embedded license in `app.py` for full details.

---

## 🌟 Quick Reference

| Command | Purpose |
|---------|---------|
| `python app.py` | 🚀 Launch the GUI |
| `Excel-CSV-Tool.bat` | 💨 Quick launcher (no terminal) |
| Double-click `Excel-CSV-Tool.lnk` | 🖱️ Desktop shortcut launch |
| `pip install -r requirements.txt` | 📦 Install dependencies |

---

## 💡 Pro Tips

🔹 **Batch Processing:** Select multiple CSV/Excel files at once for faster merging  
🔹 **Large Files:** The tool handles thousands of rows efficiently  
🔹 **Custom Output:** All reports are saved in organized timestamped folders  
🔹 **Reuse Sessions:** Previous runs are accessible in the `output/` folder  
🔹 **Keyboard Shortcuts:** Set one up for lightning-fast access!

---

## 👨‍💻 Made with ❤️

Created by **Amit Kadam** 🎯

*Enjoy automating your data workflows!* 🚀✨
