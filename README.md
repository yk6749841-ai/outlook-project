# 📧 Outlook Automation Bridge (Flask + COM Interop)

A specialized tool that bridges the gap between a Web UI and Microsoft Outlook, enabling automated creation of email drafts with attachments by bypassing browser security limitations.

## 🛠️ The Technology
This project uses a **Flask** backend to act as a local server, communicating with the **Windows COM interface** via `pywin32`. This architecture allows the application to perform native operations that are impossible for standard web apps, such as accessing the local file system and controlling the Outlook Desktop client.

## ✨ Key Technical Features
- **COM Interface Integration:** Utilizes `win32com.client` to programmatically control Outlook objects (`Outlook.Application`).
- **Dynamic Attachment Handling:** Manages temporary file storage and cleanup (`os.remove`) to attach local files to email drafts securely.
- **Multithreading Management:** Implements `pythoncom.CoInitialize()` to ensure stable COM communication within a Flask environment.
- **Batch Processing:** Parses comma-separated recipient lists to generate multiple personalized drafts in one request.

## 🚀 How to Run
1. Install dependencies: `pip install flask flask-cors pywin32`
2. Run the server: `python app.py`
3. Access the UI via `http://127.0.0.1:5000`
4. Fill in the form, attach a file, and watch Outlook open your drafts automatically.

## 📁 Project Structure
- `app.py`: Flask server and Outlook COM logic.
- `templates/index.html`: Web form for data input.
- `static/`: Frontend styling and client-side logic.
