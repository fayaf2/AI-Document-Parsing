# 📄 AI Document Parsing Web App

**AI-Document-Parsing** is a web-based application that intelligently extracts, analyzes, and processes content from uploaded `.doc` and `.docx` files using AI/NLP techniques. This tool is designed to reduce manual effort in document review by extracting key data and optionally allowing further enhancements like summarization, question answering, or image parsing.

---

## ⚙️ Features

- 📂 Upload `.doc` or `.docx` files via a web interface
- 🤖 Automatically extract text using NLP methods
- 🧠 AI-based processing for:
  - Extracting structured information
  - Summarizing content (optional)
  - Identifying keywords or highlights
- 📊 Progress tracking with `progress.txt`

AI-Document-Parsing/
├── app.py # Flask app to handle file uploads and processing
├── main.py # Core document processing logic
├── your_script.py # Additional processing script
├── templates/ # HTML files (upload form, results display)
├── static/ # Static files (CSS, JS, images)
├── uploads/ # Uploaded document files
├── progress.txt # Progress tracker or logging
├── scope of work test.doc # Sample input file
├── ss.docx # Another test document



---

## 🚀 How to Run the Application

### 1. Clone the Repository

```bash
git clone https://github.com/fayaf2/AI-Document-Parsing.git
cd AI-Document-Parsing
