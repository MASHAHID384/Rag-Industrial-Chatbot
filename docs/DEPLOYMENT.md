# Deployment Guide

## 1. Local verification
```powershell
streamlit run app.py
```

## 2. Required environment variable
- `GEMINI_API_KEY`

## 3. Streamlit Community Cloud
1. Push this repo to GitHub.
2. Create app in Streamlit Cloud.
3. Set `Main file path` = `app.py`.
4. Add `GEMINI_API_KEY` in app Secrets.
5. Deploy.

## 4. Optional: Render / VM
Use startup command:
```bash
streamlit run app.py --server.port $PORT --server.address 0.0.0.0
```

## 5. Pre-deploy checklist
- API key is valid.
- Uploaded file parsing works (pdf/docx/csv/json).
- Multi-file answering tested.
- Mobile layout tested.
