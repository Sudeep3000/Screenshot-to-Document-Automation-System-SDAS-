# 🧾 LDSAM – Lab Document Screenshot Automation Module

### Part of the SmartDoc Automation Suite (Powered by SDAS)

---

## 🎯 Objective
LDSAM automates the process of inserting screenshots into pre-designed lab report templates by replacing specified placeholders with captured images.

---

## ⚙️ Key Features
- Placeholder-based image insertion.
- Idle-time detection for automated screenshot timing.
- Region-based capture.
- Error handling and process notifications.
- Automatic document saving with screenshots embedded.

---

## 🧩 Usage
1. Ensure placeholders in your Word file match those in the code.
2. Run the module:
   ```bash
   python Module2.py
   ```
3. When idle time elapses, a screenshot is captured and inserted at the corresponding placeholder.

---

## 🧾 Output
- A Word or PDF document with screenshots inserted precisely where placeholders existed.

---

## 📚 Dependencies
```bash
pip install pyautogui pillow python-docx winotify pynput docx2pdf
```

---

## 👨‍💻 Author
Developed by **Sudeep N**
