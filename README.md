# AI Powered Multi Doc Conversational Chatbot Using RAG

Production-ready Streamlit chatbot with multi-file RAG, auto-indexing, answer-style control, and numeric dashboard metrics.

## Project Structure
```text
chatbot-project/
  app.py                      # Thin Streamlit entrypoint
  requirements.txt            # Runtime dependencies
  .env.example                # Environment variable template
  rag_chatbot/
    __init__.py
    main.py                   # Core app logic (UI + RAG + API calls)
  data/                       # Optional local data files
  faiss_index/                # Optional index artifacts
  docs/
    DEPLOYMENT.md             # Deploy checklist and steps
  .github/workflows/
    ci.yml                    # CI: install + compile check
```

## Quick Start
1. Open terminal in project root.
2. Activate virtual environment.
3. Run:
```powershell
streamlit run app.py
```

## Environment Variable
Set this once in your environment:
```text
GEMINI_API_KEY=your_api_key_here
```

## Deploy Readiness
- `app.py` kept as stable entrypoint for local run and cloud deploy.
- Core logic moved to `rag_chatbot/main.py` for cleaner maintenance.
- CI workflow added for basic validation before pushing.
- Runtime files and secrets excluded via `.gitignore`.

## GitHub Push
```powershell
git add .
git commit -m "refactor: deploy-ready project structure"
git push
```
