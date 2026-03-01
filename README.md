# 🎓 Certificate Generator

A powerful web application for batch-generating personalized certificates from Excel data and DOCX templates. Built with React and TypeScript.

![Certificate Generator Demo](https://via.placeholder.com/800x400?text=Certificate+Generator+Demo)

## ✨ Features

### 📄 Template Management
- **Upload DOCX templates** with placeholders like `{name}`, `{position}`, `{date}`
- **Automatic placeholder detection** - finds all placeholders in your template
- **Save templates permanently** in your browser's IndexedDB
- **Load saved templates** instantly - no need to re-upload
- **Delete unwanted templates** with one click

### 📊 Excel Data Handling
- **Upload Excel files** (.xlsx, .xls) as your data source
- **Edit data inline** with a full spreadsheet editor
  - Add/delete rows
  - Add/delete columns
  - Edit any cell
  - Bulk paste CSV data
  - Import additional data from files
- **Smart date conversion** - automatically converts Excel dates to readable format
- **Persistent storage** - data saves in browser database

### 🔍 Advanced Filtering
- **Multi-column filtering** with AND logic
- **Real-time preview** of filtered results
- **Value suggestions** based on actual column data
- **Active filter indicators** with one-click clear
- **Download filtered certificates** only

### 📥 Certificate Generation
- **Preview certificates** before downloading
- **Navigate through records** with Previous/Next buttons
- **Download options:**
  - Current certificate
  - Range of certificates
  - All certificates at once
- **Print certificates** individually or in batch
- **Sequential numbering** with `{CERTver-DATE_ISO}` placeholder (YYYY-MM-01 format)

### 💾 File System Integration
- **Direct file saving** using File System Access API (Chrome/Edge)
- **Fallback to download** for other browsers
- **Explicit user permission** for file modifications
- **Clear status indicators** for write access

## 🚀 Getting Started

### Prerequisites
- Node.js (v14 or higher)
- npm or yarn

### Installation

1. **Clone the repository**
```bash
git clone https://github.com/jefefefef/Certifateapp.git
cd Certifateapp.git

2. **Install dependencies**
```bash
npm install

3. **Start the development server**
npm run dev

📖 How to Use
Step 1: Upload a DOCX Template
Click the "Upload DOCX Template" area

Select a Word document with placeholders like {name}, {position}, {date}

Enter a name for your template when prompted

The template is saved permanently in your browser

Step 2: Upload Excel Data
Click the "Upload Excel Data" area

Select an Excel file with columns matching your placeholders

Data is loaded and saved for the session

Step 3: Generate Certificates
Preview certificates using the navigation buttons

Choose your download option:

Download Current - only the visible certificate

Download Range - specify a range of records

Download All - all records at once

🎯 Advanced Features
Editing Excel Data
Click "Edit Excel Data" button

Modify cells, add rows/columns, or paste CSV data

Enable Direct Save (Chrome/Edge) to save back to original file, or Download Updated File for a new copy

Filtering Records
Click "Filter Records" button

Add filter conditions (column + value)

See real-time preview of matching records

Apply filters or download filtered results

Using Sequential Numbers
Add any placeholder ending with -DATE_ISO in your template for automatically incrementing numbers:

{CERTver-DATE_ISO} → 2026-03-01

{CERTlang-DATE_ISO} → 2026-03-01

{FOO-DATE_ISO} → 2026-03-01

All -DATE_ISO placeholders share the same counter per template per month:

First download this month: 2026-03-01

Second download: 2026-03-02

Resets each month

Placeholder Modifiers
Uppercase: Add _UPPER suffix (e.g., {name_UPPER}) to convert values to uppercase

🏗️ Technical Architecture
Technologies Used
React - UI framework

TypeScript - Type safety

Vite - Build tool

Tailwind CSS - Styling

IndexedDB - Local storage for templates and data

LocalStorage - Counter storage (CSV format)

File System Access API - Direct file saving

SheetJS (XLSX) - Excel file parsing

Docxtemplater - DOCX template processing

Mammoth.js - DOCX to HTML conversion

FileSaver.js - File download fallback

Lucide React - Icons

Data Flow
text
Upload DOCX → Store in IndexedDB ← Load saved template
     ↓
Extract placeholders from ALL XML files (including headers)
     ↓
Upload Excel → Parse with SheetJS
     ↓
Edit data → Update in memory + IndexedDB
     ↓
Filter records → Apply conditions
     ↓
Generate certificates → Docxtemplater with dynamic placeholders
     ↓
Download/Save → File System API / FileSaver
Key Components
Placeholder Detection
The app scans ALL XML files in the DOCX zip, including:

word/document.xml (main content)

word/header*.xml (headers)

word/footer*.xml (footers)

Counter System
Sequential numbers are stored in LocalStorage as CSV:

csv
templateId,templateName,monthKey,count
template_123,Training Certificate,2026-03,5
template_123,Training Certificate,2026-02,12
template_456,Completion Certificate,2026-03,3
🌐 Browser Support
Feature	Chrome	Edge	Firefox	Safari
Basic functionality	✅	✅	✅	✅
Direct file saving	✅	✅	❌	❌
IndexedDB	✅	✅	✅	✅
Print	✅	✅	✅	✅
📁 Project Structure
text
certificate-generator/
├── src/
│   ├── App.tsx              # Main application component
│   ├── main.tsx             # Entry point
│   ├── index.css            # Global styles
│   └── vite-env.d.ts        # Vite type definitions
├── public/                   # Static assets
├── index.html               # HTML template
├── package.json             # Dependencies
├── tsconfig.json            # TypeScript configuration
├── vite.config.ts           # Vite configuration
├── tailwind.config.js       # Tailwind CSS configuration
├── postcss.config.js        # PostCSS configuration
└── README.md                # This file
🔧 Configuration
Environment Variables
Create a .env file in the root directory:

env
VITE_APP_TITLE=Certificate Generator
VITE_APP_VERSION=1.0.0
Database Version
The app uses IndexedDB with version control. Update DB_VERSION in App.tsx when making structural changes:

typescript
const DB_VERSION = 2; // Increment when changing schema

📞 Support
For issues or questions:

Check the Issues page

Create a new issue with detailed description

Include browser console logs when reporting bugs

Specify your browser and operating system

🚧 Known Limitations
Direct file saving only works in Chromium-based browsers (Chrome, Edge, Opera)

Large Excel files (>10MB) may cause performance issues

DOCX templates with complex formatting may not preview perfectly

Counters reset when clearing browser data

File handles are lost when closing the browser tab (need to re-enable direct save)

🔜 Roadmap
Planned Features
Dark mode support

Multiple Excel sheets support

Custom placeholder functions

Cloud storage integration (Google Drive, Dropbox)

Batch template editing

PDF export option

Email certificates directly

User authentication

Template sharing

Audit log of downloads

QR code generation

Custom counter formats

In Progress
Enhanced error handling

Performance optimizations for large datasets

More placeholder modifiers

Template categories/tags

📊 Performance Tips
Keep Excel files under 10MB for optimal performance

Use simple DOCX templates without complex tables or images

Clear browser data periodically to remove old counters

Use filtering to work with subsets of large datasets

Chrome/Edge recommended for direct file saving feature

🎯 Use Cases
Educational institutions - Generate certificates for graduates

Corporate training - Issue completion certificates

Event management - Create participation certificates

HR departments - Generate employment verification letters

Any batch document generation from template + data

💡 Tips & Tricks
Template Design: Create your DOCX with placeholders like {name}, {date}, {course}

Excel Columns: Match column names exactly to placeholders (case-insensitive)

Sequential Numbers: Use {CERTver-DATE_ISO} anywhere in the document (even headers!)

Bulk Editing: Use the CSV paste feature for quick data entry

Filtering: Combine multiple conditions for precise record selection

Direct Save: Enable in Chrome/Edge for seamless workflow

Made with ❤️ for batch certificate generation