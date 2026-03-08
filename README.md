# 🐱 Si-Som DOCX Skill (si-som-docx)

> **"Professional Educational Document Generator for Don (Krudony)"**

This is a specialized Gemini CLI skill designed to create high-quality Word documents (.docx) following the **Thai Educational Standard** (specifically tailored for Ban Mae Sai School).

---

## 🚀 Features

- 📑 **Exam Mode**: Create 2-column A4 exams with TH SarabunPSK 16pt.
- 📂 **Project Mode**: Generate school project reports with precise spacing and margins (1 inch).
- 🖋️ **Auto-Formatting**: Integrated with `docx_engine.py` and `project_engine.py` for automated layout management.
- 🏫 **School Branding**: Pre-configured headers for Ban Mae Sai School.

## 🛠️ Requirements

To use this skill effectively, the target machine MUST have:

1. **Python 3.x**: The core engine runs on Python.
   - Install dependencies: `pip install python-docx docxtpl`
2. **TH SarabunPSK Font**: Standard Thai government font. **[CRITICAL]**
3. **Node.js**: Required if using the Obsidian Bridge script.

## 📦 Installation (Gemini CLI)

```bash
gemini skills install https://github.com/Krudony/si-som-docx-skill.git
```

## 📖 Usage

Activate the skill in your Gemini session:
`"activate_skill si-som-docx"`

**Example Commands:**
- "ช่วยสร้างข้อสอบวิชาคอมพิวเตอร์ ป.6 ให้หน่อย 30 ข้อ"
- "ทำสรุปโครงการพานักเรียนไปทัศนศึกษาให้ที"

## 📂 Structure

- `scripts/`: Python engines (`docx_engine.py`, `project_engine.py`).
- `references/`: Formatting rules and markdown templates.
- `examples/`: Sample scripts and usage demonstrations.

---
*Created with ❤️ by **Si-Som 🐱** for **Don (Krudony)***
*Repository: [Krudony/si-som-docx-skill](https://github.com/Krudony/si-som-docx-skill)*
