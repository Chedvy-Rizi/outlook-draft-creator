# 📧 Desktop-Bridge Email Automator

### Bridging the gap between Web-Sandbox and Local Desktop Automation

---

## 🎯 Project Overview
This project solves a common challenge in HR and recruitment systems: **The Browser Sandbox Limitation**. Web browsers are restricted from interacting with local desktop applications for security reasons.

I developed a **Local Agent Architecture** using a Flask-based Bridge that enables a web interface to securely trigger and control **Microsoft Outlook** on the user's machine, allowing for mass-creation of personalized email drafts with physical attachments.

---

## 📸 Demo & UI
> **[כאן מומלץ להוסיף את ה-GIF שלך]**
> 
> ![Application Demo](https://via.placeholder.com/800x400?text=Insert+Your+Demo+GIF+Here)

---

## 🛠️ Key Technical Challenges & Solutions

### 1. The Sandbox Bridge (Architecture)
* **Challenge:** Web browsers cannot access the local file system or launch `Outlook.exe` directly.
* **Solution:** Implemented a Local Python Agent that listens to secure REST requests and communicates with the **Win32COM API** to bridge the gap.

### 2. High-Performance Concurrency
* **Challenge:** Launching multiple Outlook instances can be slow and might freeze the Web UI.
* **Solution:** Utilized **Multi-threading** to offload the Outlook processing to a background thread. The server returns an **HTTP 202 Accepted** status immediately, ensuring a smooth User Experience.

### 3. Security & Resource Integrity
* **Sanitization:** Integrated `Bleach` to sanitize HTML inputs, preventing XSS attacks within generated emails.
* **Path Traversal Protection:** Used `secure_filename` to ensure uploaded attachments are stored safely.
* **Automated Cleanup:** Developed a resource management routine that securely deletes temporary files after processing to maintain a zero-footprint on the host machine.

---

## 🚀 Getting Started

### Prerequisites
* Windows OS (Required for Outlook COM interaction)
* Microsoft Outlook installed and configured
* Python 3.9+

### Installation

**1. Clone the Repository:**
```bash
git clone [https://github.com/](https://github.com/)[Chedvy-Rizi]/outlook-draft-creator.git
cd outlook-draft-creator
```

**2. Install Dependencies:**
```bash
pip install flask flask-cors pywin32 bleach werkzeug
```

**3. Run the Agent:**
```bash
python server.py
```

**4. Launch the UI:**
Open index.html in any modern browser.


##🧬 System Architecture
Developer: Chedvy Rizi
