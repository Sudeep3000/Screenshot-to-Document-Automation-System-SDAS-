# 🧠 EDAM – E-Learning Documentation Automation Module

### Part of the SmartDoc Automation Suite (Powered by SDAS)

---

## 🎯 Objective
EDAM automates capturing screenshots from e-learning or digital book platforms and compiles them into structured Word or PDF documents with minimal user interaction.

---

## ⚙️ Key Features
- Region-based screenshot capture.
- Idle-time detection for automatic page capture.
- Word/PDF document creation and management.
- Real-time notifications via Windows Toast.
- Automated navigation for continuous capture.

---

## 🧩 Usage
1. Configure your file path, region, and idle timeout.
2. Run the module:
   ```bash
   python EDAM_main.py
   ```
3. The system captures screenshots automatically and appends them to your document.

---

## 🧾 Output
- A `.docx` or `.pdf` document containing all screenshots.
- Notifications for every key step (region setup, capture, page turn, and completion).

---

## 📚 Dependencies
```bash
pip install pyautogui pillow python-docx winotify pynput docx2pdf
```

---

## 👨‍💻 Author
Developed by **Sudeep N**
