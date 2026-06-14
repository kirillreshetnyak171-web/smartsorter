# SmartSorter — AI Assistant for Excel & CSV

**Desktop app · Python + PyQt6 · Pet project / demo**

Tool for **small teams and offices**: upload a spreadsheet, get an **AI draft**, review it, then save. Main demo scenario — **classify customer complaints** and route files to department folders. Interface in **German, English, and Russian**.

---

## What it does

| Scenario | What happens |
|----------|--------------|
| **Complaint routing** | Load `complaints_example.csv` → AI suggests category & department → you approve → files go to department folders |
| **Spreadsheet cleanup** | Templates: duplicates, dates, sorting, trim spaces — or describe the task in plain language |
| **Human in the loop** | AI never saves silently — you check the draft first |
| **Privacy (BYOK)** | Your Anthropic API key stays on your Mac in `config.txt` (not in this repo) |

> **Note:** Demo / portfolio project. Includes sample data (`complaints_example.csv`, `company_rules.example.txt`) — no real customer data.

---

## Tech stack

Python 3, PyQt6, Pandas, OpenPyXL, Anthropic Claude API.

---

## Quick start (Mac)

### Install once

```bash
git clone https://github.com/kirillreshetnyak171-web/smartsorter.git
cd smartsorter
python3 -m pip install -r requirements.txt
```

### Run without Terminal

Double-click **`SmartSorter.command`** in the project folder.  
First time on macOS: **Right-click → Open → Open**.

### Run from Terminal

```bash
python3 app.py
```

### API key

1. Create a key at [console.anthropic.com](https://console.anthropic.com)  
2. Paste in the field top-left (`sk-ant-…`)  
3. Copy `config.example.txt` → `config.txt` if starting fresh  

Data is sent to Anthropic **only when you click “Get draft”**.

---

## Demo: complaint classification

1. **File** → open `complaints_example.csv`  
2. **Task** → template **«Complaint classification»**  
3. **Attach rules** → `company_rules.example.txt`  
4. **Get draft** → review report & table  
5. **To departments** → files land in folders (see `department_folders.demo.txt`)

Create department folders on your Desktop first, or edit paths in `department_folders.demo.txt`.

---

## Main files

| File | Purpose |
|------|---------|
| `app.py` | Application |
| `SmartSorter.command` | Double-click launcher (Mac) |
| `config.example.txt` | API key template (copy to `config.txt`) |
| `company_rules.example.txt` | Sample company rules + departments |
| `complaints_example.csv` | Sample complaints |
| `department_folders.demo.txt` | Where to deliver files per department |

---

## Other tasks

Templates for duplicates, dates, sorting, whitespace cleanup, and more.  
Open any `.xlsx` / `.csv` → pick a template or describe the task → draft → **Save**.

---

## Troubleshooting

| Message | Fix |
|---------|-----|
| No API key | Paste key top-left |
| No rules file | Attach `company_rules.example.txt` for complaints template |
| No «Department» column | Run classification first |
| Department folders missing | Check `department_folders.demo.txt` |

---

## License

Pet / demo project. Use at your own discretion.
