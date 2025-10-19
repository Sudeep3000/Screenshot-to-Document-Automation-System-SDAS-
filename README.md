# 🧠 SmartDoc Automation Suite (Powered by SDAS)
### Screenshot-to-Document Automation System (SDAS)

A modular Python-based automation suite designed to **capture, organize, and insert screenshots into Word documents** with minimal user intervention.  
It is built as part of the **SDAS – Screenshot-to-Document Automation System**, featuring two primary modules for educational and laboratory use cases.

---

## 📘 Overview

**SmartDoc Automation Suite (SDAS)** is a comprehensive system that automates manual documentation tasks like:
- Capturing screenshots from online or lab environments.
- Organizing them into structured Word/PDF reports.
- Replacing predefined placeholders in templates with screenshots automatically.
- Notifying the user with progress updates via Windows Toast notifications.

This project combines **GUI automation, document processing, and event-driven logic** into a seamless user experience.

---

## 🧩 System Modules

| Module | Name | Description |
|--------|------|-------------|
| **Module 1** | 🧠 **EDAM – E-Learning Documentation Automation Module** | Automates the process of taking region-based screenshots from e-learning content and compiling them into structured Word or PDF documents. |
| **Module 2** | 🧾 **LDSAM – Lab Document Screenshot Automation Module** | Inserts screenshots into specified positions or placeholders in existing lab report templates. Ideal for academic and research documentation. |

---

## ⚙️ Tech Stack

- **Python 3.x**
- **Libraries:** `pyautogui`, `pynput`, `python-docx`, `Pillow`, `winotify`, `docx2pdf`
- **Platform:** Windows (WinToast notifications)

---

## 🚀 Features

✅ Intelligent idle-time detection and auto-screenshot capture  
✅ Region-based screen capture and navigation automation  
✅ Automatic Word/PDF generation and file safety checks  
✅ Windows toast notifications for progress updates  
✅ Modular architecture (EDAM & LDSAM under SDAS)

---

## 🧭 Setup & Installation

### Step 1: Clone the Repository
```bash
git clone https://github.com/<yourusername>/Screenshot_to_Document_Automation_System.git
cd Screenshot_to_Document_Automation_System
