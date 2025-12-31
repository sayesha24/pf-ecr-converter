# PF ECR Converter - Deployment Guide

## Features
- ✅ Convert ECR PDF to Excel
- ✅ Convert Excel to CSV  
- ✅ Supports both normal AND rotated/landscape PDFs (via OCR)
- ✅ Removes name columns automatically
- ✅ Cleans numeric formatting

---

## 🚀 Deploy to Render (FREE)

### Step 1: Create a GitHub Repository

1. Go to https://github.com and sign in
2. Click **"+"** (top right) → **"New repository"**
3. Name it: `pf-ecr-converter`
4. Keep it **Public**
5. Click **"Create repository"**

### Step 2: Upload All Files

1. On your repository page, click **"uploading an existing file"**
2. Drag and drop ALL 5 files from this folder:
   - `app.py`
   - `requirements.txt`  
   - `render.yaml`
   - `Aptfile` ← **Important! This installs Tesseract OCR**
   - `README.md` (optional)
3. Click **"Commit changes"**

### Step 3: Deploy on Render

1. Go to https://render.com
2. Sign up/Login with GitHub
3. Click **"New +"** → **"Web Service"**
4. Select your `pf-ecr-converter` repository
5. Settings should auto-fill from render.yaml
6. Click **"Create Web Service"**

### Step 4: Wait for Build (3-5 minutes)

Wait for the status to show **"Live"**.

### Step 5: Get Your URL! 🎉

Your app will be available at:
```
https://pf-ecr-converter.onrender.com
```

---

## 📁 Files Explained

| File | Purpose |
|------|---------|
| `app.py` | Main Flask application |
| `requirements.txt` | Python packages |
| `render.yaml` | Render deployment config |
| `Aptfile` | **System packages (Tesseract OCR)** |

---

## ⚠️ IMPORTANT: The `Aptfile`

The `Aptfile` tells Render to install Tesseract OCR on the server. Without it, rotated PDFs won't work!

Contents of Aptfile:
```
tesseract-ocr
tesseract-ocr-eng
```

---

## 🔧 Troubleshooting

### "tesseract is not installed" error
- Make sure `Aptfile` is uploaded to your GitHub repo
- The file must be named exactly `Aptfile` (capital A, no extension)
- Redeploy after adding the file

### Build takes too long
- First build with Aptfile may take 5+ minutes
- Subsequent builds are faster

---

## ⚠️ Free Tier Notes

- App sleeps after 15 min of inactivity
- First request after sleep takes 30-60 seconds
