# RALCO Field Walkthrough — Export Backend

Flask backend that receives job data from the field walkthrough app and
returns a fully formatted EMA Sales Builder Excel file.

## Files

- `app.py` — Flask application
- `requirements.txt` — Python dependencies
- `render.yaml` — Render deployment config
- `template.xlsx` — RALCO EMA Sales Builder template (DO NOT RENAME)

## Local testing

```bash
pip install -r requirements.txt
python app.py
```

Then test the health check:
```
http://localhost:5000/health
```

## Deploying to Render

See deployment guide provided separately.
