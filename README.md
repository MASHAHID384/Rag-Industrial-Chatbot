# AI Powered Multi Doc Conversational Chatbot Using RAG

## 🎯 Project Purpose & Objective
- Build a practical, production-ready RAG chatbot that helps users understand multiple documents through natural conversation.
- Reduce manual reading time by retrieving relevant context and generating grounded answers from uploaded files.
- Provide a clean, deployable solution that works across desktop and mobile with secure API-based inference.
## ✨ Features
- Multi-document chat over PDF, DOCX, TXT, CSV, JSON, and XML files.
- Auto file processing and chunk indexing for fast retrieval.
- Multiple answer styles: Detailed, Short, Bullet Points, Professional.
- Mobile-safe upload mode and URL-based document ingestion.
- Numeric dashboard with API, file, and query metrics.
- Public cloud deployment with secure API key through secrets.

## 🔄 Pipeline
1. Ingest: Upload local files or add document URL.
2. Parse: Extract text from supported document formats.
3. Chunk: Split content into retrieval-ready chunks.
4. Index: Store chunks in session context for fast lookup.
5. Retrieve: Select top relevant chunks for each question.
6. Generate: Build prompt and generate response with Gemini.
7. Format: Clean, deduplicate, and style output response.

## 🧰 Technology Used
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

## 🗂️ Project Structure
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

## 🚀 Quickstart
1. Open terminal in project root.
2. Activate virtual environment.
3. Set API key:
   `GEMINI_API_KEY=your_api_key_here`
4. Install dependencies:
   `pip install -r requirements.txt`
5. Run project:
   `streamlit run app.py`

