
---

## 🏢 Circular 29 Auto-Generator
A one-click tool to generate **Circular 29** reports from a real estate project's **Form 3 Excel** — designed for developers, RERA consultants, and compliance teams.

<img width="1918" height="911" alt="image" src="https://github.com/user-attachments/assets/40b36db3-66ca-41e9-b101-b86fae224047" />

---

## 🔧 Overview

Preparing Circular 29 manually is repetitive and error-prone. This app automates it by extracting key details from Form 3 (as per MahaRERA) and producing a structured Excel file for reporting.

---

## ✨ Features

- 📄 **Reads Form 3 Excel** with complex multi-sheet structures
- 🏢 Extracts **Project Name** & **RERA Registration Number**
- 📅 Captures **As-on Date** from `Table B`
- 🧮 Parses unit data for sections like:
  - Sold / Booked
  - Unsold / Not Booked
  - Mortgaged
  - Not for Sale
- 📊 Outputs a **clean Circular 29 Excel sheet** with formatted data
- 🔁 Handles minor inconsistencies in header naming and spacing
- 🚫 Skips blank or invalid rows without crashing

---

## 🖥️ Live App

👉 **[Launch Now](https://circular29-autogenerator.streamlit.app/)** – No setup required

---

## Output Snapshot

<img width="1919" height="912" alt="image" src="https://github.com/user-attachments/assets/03b85327-9837-47ee-aea4-c01b60254c48" />

---

## 📁 Usage

### 1. Prepare Input
Use the Excel Form 3 received from your CA or project team — it should include sheets like `Table A`, `Table B`, etc.

### 2. Upload
Go to the app and upload the `.xlsx` file.

### 3. Download
Once processed, the Circular 29 file will be ready to download instantly.

---

## 🛠 Tech Stack

- **Streamlit** – interactive web UI
- **Pandas** – data parsing & transformation
- **Regex** – for extracting project metadata
- **OpenPyXL / XlsxWriter** – Excel generation
- **Python Logging** – for clean debugging

---

## 🧠 Internals

- Dynamically locates the **unit data** by scanning for `Sr. No`, `Flat No.`, `Carpet Area`, etc.
- Stops reading a section **only when a new section header is detected** (e.g., from Sold ➝ Unsold)
- Handles varied formatting across Excel files by tolerating blank rows and minor header changes
- Extracts metadata either from the sheet **or** as a fallback from the file name

---

## 🚀 Deployment

Hosted via [Streamlit Cloud](https://streamlit.io/cloud). To run locally:

---

## 👨‍💻 Author

**Aryan Parte**  
🔗 [GitHub](https://github.com/aryanparte03)  
📫 Business Intelligence & Analytics | Web + Data Automation | MBA  

```bash
git clone https://github.com/aryanparte03/test
cd test
pip install -r requirements.txt
streamlit run full_streamlit_app.py
