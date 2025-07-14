##Business Report Generator

This project automates the generation of professional business reports in **PDF**, **DOCX**, and **PPTX** formats using cutting-edge AI (Claude Sonnet 3.7), real-time web research via **Serper.dev**, and dynamic HTML rendering techniques. The reports are styled for executives and are grounded with **IEEE-style citations** sourced from live Google-like search results.

---

## 🚀 Features

- 🎯 **Business Report Generator**: Generates formal HTML reports with:
  - Clean semantic structure
  - Inline CSS styling for PDF/DOCX conversion
  - Executive-ready formatting
  - Live references and citations

- 📊 **AI Presentation Builder**:
  - Generates 8–10 slide decks in HTML
  - Converts them into PNG → PPTX using Selenium and ChromeDriver
  - Fully visual, brand-friendly design

- 🔍 **Live Web Citation Integration**:
  - Uses [Serper.dev](https://serper.dev) to fetch top search results
  - Embeds IEEE-style citations both inline and in a reference section

- ☁️ **Cloud Storage Ready**:
  - Automatically uploads generated reports to **AWS S3**
  - Public URLs for easy sharing and integration

---

## 🧩 Tech Stack

| Component           | Technology Used                                 |
|---------------------|--------------------------------------------------|
| AI Agent            | `strands.Agent` using Claude Sonnet 3.7         |
| Web Search API      | [Serper.dev](https://serper.dev) (Google Search)|
| HTML to PDF         | `pdfkit` + `wkhtmltopdf`                        |
| HTML to DOCX        | `pypandoc`                                      |
| PPTX Presentation   | `python-pptx`, `selenium`, `PIL`                |
| Cloud Integration   | `boto3` (AWS S3 SDK for Python)                 |

---

## 📂 Input Format

Create an `input.json` file:

```json
{
  "use_case_name": "Contract Intelligence Assistant",
  "description": "An AI system that reads legal contracts, extracts key clauses, summarizes terms, and recommends actions.",
  "project_id": "ci_project_001",
  "user_id": "john_doe"
}
