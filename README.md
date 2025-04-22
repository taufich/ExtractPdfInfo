# 📄 Transcript Extractor & Highlighter

A powerful and user-friendly web application built with **Streamlit** for parsing academic transcript PDFs and exporting structured, color-coded Excel files. This tool is ideal for academic administrators, registrars, or students looking to extract and analyze transcript data efficiently.

---

## 🚀 Features

- ✅ Upload academic transcripts in PDF format
- 🔍 Extract and display:
  - Student information (Name, Reg. Number, Major, Program)
  - Detailed course records by academic year and semester
  - Distribution of credit and overall averages
- 🎨 Auto-highlight:
  - Excellent grades (>= 14.0) in green
  - Low/failing grades (< 12.0 or EC flagged) in red
- 📊 Export organized data into an Excel file with multiple sheets:
  - Student Info
  - Courses (with highlights)
  - Credit Distribution
  - Averages

---

## 🛠️ Technologies Used

- [Streamlit](https://streamlit.io/) - for building the web interface
- [PyMuPDF (fitz)](https://pymupdf.readthedocs.io/) - for PDF text extraction
- [Pandas](https://pandas.pydata.org/) - for data manipulation
- [OpenPyXL](https://openpyxl.readthedocs.io/) - for Excel file creation and styling
- [Regular Expressions (re)](https://docs.python.org/3/library/re.html) - for parsing transcript text

---

## 📦 Installation

1. **Clone the repository:**
   ```bash
   git clone https://github.com/yourusername/transcript-extractor.git
   cd transcript-extractor
   
2. **Create and activate a virtual environment (optional but recommended):**
   ```bash
   python -m venv venv
   source venv/bin/activate  # On Windows: venv\Scripts\activate
   
3. **Install dependencies:**
   ```bash
   pip install -r requirements.txt
   
4. **Run the application:**
   ```bash
   streamlit run main.py


