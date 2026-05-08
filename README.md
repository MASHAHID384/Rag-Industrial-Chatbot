# AI Powered Multi Doc Conversational Chatbot Using RAG

## 🎯 Project Purpose
This project is built to make document understanding faster and easier through conversational AI. It allows users to upload multiple documents, retrieve relevant context, and get grounded answers in different response styles. The goal is to provide a practical, deployable RAG chatbot experience for real-world academic and professional use.

## 1) ✨ Features
- Multi-document chat over PDF, DOCX, TXT, CSV, JSON, and XML files.
- Auto file processing and chunk indexing for fast retrieval.
- Multiple answer styles: Detailed, Short, Bullet Points, Professional.
- Mobile-safe upload mode and URL-based document ingestion.
- Numeric dashboard with API, file, and query metrics.
- Public cloud deployment with secure API key through secrets.

## 2) 🔄 Pipeline
1. Ingest: Upload local files or add document URL.
2. Parse: Extract text from supported document formats.
3. Chunk: Split content into retrieval-ready chunks.
4. Index: Store chunks in session context for fast lookup.
5. Retrieve: Select top relevant chunks for each question.
6. Generate: Build prompt and generate response with Gemini.
7. Format: Clean, deduplicate, and style output response.

## 3) 🧰 Technology Used
- Python
- Streamlit
- Google Gemini API (`google-genai`)
- Pandas
- PyPDF
- python-docx
- JSON
- Git
- GitHub
- Streamlit Cloud

## 4) 🗂️ Project Structure
```text
chatbot-project/
  app.py
  requirements.txt
  .env.example
  rag_chatbot/
    __init__.py
    main.py
  docs/
    DEPLOYMENT.md
  .github/workflows/
    ci.yml
```

## 5) 🚀 Quickstart (How I Run This Project)
1. Open terminal in project root.
2. Activate virtual environment.
3. Set API key:
   `GEMINI_API_KEY=your_api_key_here`
4. Install dependencies:
   `pip install -r requirements.txt`
5. Run project:
   `streamlit run app.py`
