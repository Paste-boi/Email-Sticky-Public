
# 🧠 AI-Email-Sticky  
**Automated Gmail-to-To-Do Sticky Note with GPT-5-Mini Summarization**

---

### 📋 Overview
**AI-Email-Sticky** automatically turns Gmail messages into actionable sticky-note style to-dos.  
It reads new emails using Gmail’s IMAP API, summarizes them with OpenAI’s GPT-5-Mini model, and displays each actionable item as a persistent note with a checkbox and color-coded status.  

The app is designed to **stay running quietly**, refreshing itself automatically, and survive PC restarts while keeping all tasks saved locally.

---

### 🚀 Features
✅ **AI Summaries:** Uses GPT-5-Mini to condense email bodies into a single actionable line.  
✅ **Smart Filtering:** Ignores marketing, “FYI,” and other low-value emails.  
✅ **De-Duplication:** Prevents duplicate sticky notes from the same message UID.  
✅ **Cutoff Date:** Ignores emails received before a specified date (configurable).  
✅ **Leave Unread Option:** IMAP toggle lets you leave emails unread after polling.  
✅ **Manual Controls:**  
 • **Poll Now** → Instantly checks for new emails  
 • **Clear Completed** → Archives finished tasks immediately  
✅ **Visual Cues:**  
 • Checked tasks turn **subtle green (#6A9955)** for instant recognition  
 • Colorful text per field (From = blue, Date = red, Summary = white/black)  
✅ **Persistent Data:** Tasks saved in a local SQLite database (`tasks.db`)  
✅ **Configurable Themes:** Dark & Light modes (VS Code-matched colors)  
✅ **Logging:** Rolling log (`ai_sticky.log`) for debugging or audit trails  
✅ **Help Menu:** Self-test GPT connection, view last error, open log folder  

---

### 🧩 Current UI
- 🟦 **Header:**  
 → Title, **Poll Now**, **Clear Completed**, live status indicator  
- 📋 **Main List:**  
 → One line per email (From, Received, Summary)  
 → Large checkbox (toggles green text on completion)  
 → ❌ button removes note immediately  
- ⚙️ **Menu Bar:**  
 → *View > Dark Mode / Colorful Text*  
 → *Help > AI Self-Test / View Logs*  

---

### ⚙️ Installation

1. **Clone or extract** this project to a local folder, e.g.  
   ```
   C:\Users\<YourName>\Desktop\AI-Email-Sticky
   ```

2. **Install requirements:**
   ```bash
   pip install openai tk
   ```

3. **Create an OpenAI API key:**
   - Go to https://platform.openai.com/api-keys  
   - Copy the key and set it as an environment variable:
     ```powershell
     setx OPENAI_API_KEY "sk-xxxx"
     ```

4. **Enable Gmail IMAP and create an App Password:**
   - Visit: https://myaccount.google.com/apppasswords  
   - Choose “Mail” → “Windows Computer” → Copy the password  
   - Paste into your `config.ini` under `[imap] password=`  

5. **Run the app:**
   ```powershell
   python Ai-Email-Sticky.py
   ```

---

### 🧠 AI Behavior
- **Model:** GPT-5-Mini (fast + cost-efficient)  
- **Temperature:** Fixed to 1 (required by model)  
- **Summarization:** Converts body text → concise to-do line  
- **Classification (optional):** Drops “marketing” or “FYI” emails before adding  

---

### 🗂 Configuration File (`config.ini`)
The first run creates this file automatically. Example:

```ini
[imap]
server = imap.gmail.com
username = you@gmail.com
password = your_app_password
folder = INBOX
ssl = true
cutoff_date = 2025-10-20
mark_as_read = false

[ai]
enabled = true
model = gpt-5-mini
temperature = 1
api_key =
base_url =
classify_before_add = true
drop_labels = marketing, fyi

[app]
retention_hours = 12
poll_seconds = 300
ui_refresh_seconds = 30

[ui]
font_size = 10
always_on_top = true
theme = dark
colorful_text = true
```

---

### 💾 Data Storage
All files are saved under:  
```
C:\Users\<YourName>\.ai_email_sticky\
│
├── tasks.db          # SQLite database of all tasks
├── ai_sticky.log     # Rotating log file
└── config.ini        # Configuration
```

---

### 🎨 Theme Reference (sampled from VS Code)
| Mode | Element | Color |
|------|----------|--------|
| **Dark** | Border | `#2D2D2D` |
| | Background | `#1E1E1E` |
| | Text | `#D4D4D4` |
| | Completed | `#6A9955` |
| **Light** | Border | `#E9E9E9` |
| | Background | `#FFFFFF` |
| | Text | `#1F1F1F` |
| | Completed | `#6A9955` |

---

### 🔍 Hotkeys
| Shortcut | Action |
|-----------|---------|
| **Ctrl + R** | Manual poll now |
| **Alt + F4 / X** | Save geometry & quit |

---

### 🧾 Logs and Troubleshooting
- View log: *Help → View Latest Log*  
- Error messages like `insufficient_quota` or `unsupported_value` mean:  
  - Check your OpenAI account billing  
  - Ensure `temperature=1` for GPT-5-Mini  

---

### 🧱 Future Enhancements
- System tray icon + background startup  
- Windows Task Scheduler integration  
- Auto-archive / sync with task managers (ClickUp, Todoist, etc.)  
- Custom rule filters (regex, domain-based)  

---

### 👨‍💻 Credits
Developed collaboratively by **Peyton Strippelhoff** and **ChatGPT (GPT-5)**  
Focused on **productivity, automation, and information clarity**.
