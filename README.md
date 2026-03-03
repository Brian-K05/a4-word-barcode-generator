# A4 Word Barcode Generator

**Code 128 only.** Enter a barcode value → click **Generate & Download** → get an A4 Word file with the barcode in the footer.

Built by **Brian Kyle L. Salor**

---

## Run locally

```bash
pip install -r requirements.txt
python app.py
```

Open http://127.0.0.1:5000

---

## Deploy live (Render, Railway, etc.)

1. Push this repo to GitHub (see below).
2. On [Render](https://render.com): **New → Web Service**, connect your repo.
   - **Build command:** `pip install -r requirements.txt`
   - **Start command:** `gunicorn app:app`
   - Click **Create Web Service**. Your app will be live at `https://your-app-name.onrender.com`
3. On [Railway](https://railway.app): **New Project → Deploy from GitHub** → select repo. Railway auto-detects the app; if not, set start command to `gunicorn app:app`.

---

## Push to GitHub

From the project folder in terminal:

```bash
git init
git add .
git commit -m "Initial commit: A4 Word barcode generator"
```

Create a **new repository** on [GitHub](https://github.com/new) (e.g. `a4-word-barcode-generator`), then:

```bash
git remote add origin https://github.com/YOUR_USERNAME/a4-word-barcode-generator.git
git branch -M main
git push -u origin main
```

Replace `YOUR_USERNAME` with your GitHub username and the repo name if you chose a different one.
