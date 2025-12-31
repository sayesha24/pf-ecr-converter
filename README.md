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
2. Drag and drop ALL 4 files from this folder:
   - `app.py`
   - `requirements.txt`  
   - `Dockerfile`
   - `render.yaml`
3. Click **"Commit changes"**

### Step 3: Deploy on Render

1. Go to https://render.com
2. Sign up/Login with GitHub
3. Click **"New +"** → **"Web Service"**
4. Select your `pf-ecr-converter` repository
5. Render will detect it's a Docker app automatically
6. Click **"Create Web Service"**

### Step 4: Wait for Build (5-10 minutes)

Docker builds take longer than Python builds. Wait for the status to show **"Live"**.

### Step 5: Get Your URL! 🎉

Your app will be available at:
```
https://pf-ecr-converter.onrender.com
```

---

## 📱 Local Development

### Prerequisites
```bash
# Mac
brew install tesseract

# Ubuntu/Debian  
sudo apt-get install tesseract-ocr

# Windows
# Download from: https://github.com/UB-Mannheim/tesseract/wiki
```

### Install & Run
```bash
pip install -r requirements.txt
python app.py
```

Open: http://localhost:5000

---

## 🔧 Troubleshooting

### "Build Failed" on Render
- Make sure ALL 4 files are uploaded (especially Dockerfile)
- Check that file names are exact (case-sensitive)

### "Application Error" after deployment
- Wait 2-3 minutes for the app to fully start
- Check Render logs for specific errors

### OCR not working locally
- Make sure Tesseract is installed: `tesseract --version`
- On Mac: `brew install tesseract`
- On Ubuntu: `sudo apt-get install tesseract-ocr`

### Rotated PDF not converting properly
- The OCR may not capture all records perfectly
- Try re-downloading the PDF from EPFO portal
- Some heavily compressed PDFs may have issues

---

## 📁 Files

| File | Purpose |
|------|---------|
| `app.py` | Main Flask application |
| `requirements.txt` | Python dependencies |
| `Dockerfile` | Docker configuration with Tesseract OCR |
| `render.yaml` | Render deployment settings |

---

## ⚠️ Free Tier Notes

- App sleeps after 15 min of inactivity
- First request after sleep takes 30-60 seconds
- Docker builds use more free tier minutes
