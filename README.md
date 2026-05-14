# FormatFlow - Premium Document Processing Platform

FormatFlow is a professional-grade document processing and generation platform designed to transform raw text and existing documents into beautifully branded, high-fidelity DOCX files. It combines a modern, conversational interface with a powerful document synthesis engine.

## 🚀 Key Features

*   **Chat-to-DOCX Generation**: A seamless, ChatGPT-inspired experience that converts rich text (TinyMCE) into professional documents with hierarchical heading styles and consistent alignment.
*   **PDF Re-branding Engine**: Leveraging the official **Adobe PDF Services SDK**, FormatFlow extracts semantic structures (headings, tables, lists) from PDFs and regenerates them within your corporate branding.
*   **Intelligent Styling**: The `StyleManager` ensures that every document follows a strict design baseline, fixing common "staircase" indentation bugs and enforcing consistent typography.
*   **Persistent Conversations**: Full support for chat history, pinned conversations, and "Untitled" session tracking.
*   **Document Preview & Export**: Generate documents in real-time, with support for nested tables, bulleted lists, and inline rich-text formatting (bold, italics, underline).
*   **High-Fidelity PDF Export**: Integrated `docx2pdf` support for creating native PDF versions of your branded documents.

## 🛠️ Technology Stack

*   **Backend**: Flask (Python 3.10+) with JWT-based authentication.
*   **Frontend**: Vanilla JS with TinyMCE 6 (Rich Text Editor).
*   **Document Engine**: `python-docx` for generation, `BeautifulSoup4` for HTML parsing.
*   **PDF Intelligence**: Adobe PDF Services API (v4+).
*   **Database**: SQLite/MySQL via SQLAlchemy.

## 🔧 Installation & Setup

1.  **Clone the repository**:
    ```bash
    git clone https://github.com/Codes-Technology/CODES-FormatFlow.git
    cd CODES-FormatFlow
    ```

2.  **Environment Configuration**:
    Create a `.env` file in the root directory:
    ```env
    # AI Title Generation
    GROQ_API=your_groq_api_key

    # Adobe PDF Services (Required for PDF processing)
    ADOBE_CLIENT_ID=your_adobe_id
    ADOBE_CLIENT_SECRET=your_adobe_secret

    # Database
    SQLALCHEMY_DATABASE_URI=sqlite:///formatflow.db
    ```

3.  **Install Dependencies**:
    ```bash
    pip install -r requirements.txt
    ```

4.  **Launch the App**:
    ```bash
    python app.py
    ```

## 📂 Core Components

| Component | Description |
| :--- | :--- |
| `document_processor.py` | The "brain" of the application. Handles HTML parsing and DOCX synthesis. |
| `utils/style_manager.py` | Manages document branding rules and baseline typography. |
| `utils/adobe_helper.py` | High-fidelity extraction via Adobe PDF Services SDK. |
| `route/process.py` | Orchestrates the document processing pipelines. |

## 💡 Troubleshooting

> [!TIP]
> If you encounter PDF conversion errors on Windows, ensure that **Microsoft Word** is installed and accessible, as `docx2pdf` relies on the Word COM interface for high-fidelity conversion.

---
*Created for premium document formatting.*
