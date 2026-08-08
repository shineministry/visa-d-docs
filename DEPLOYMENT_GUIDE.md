# Visa D Application Form — Deployment Guide

## Files

| File | Purpose |
|------|---------|
| `index.html` | The web form (Google Forms-style UI) |
| `Code.gs` | Google Apps Script backend (email sending) |
| `visa_page_1.png` | Original application form page 1 |
| `visa_page_2.png` | Original application form page 2 |

---

## Step-by-Step Setup

### 1. Upload Images to Google Drive

1. Go to [drive.google.com](https://drive.google.com)
2. Create a new folder (e.g. "Visa D Form Assets")
3. Upload `visa_page_1.png` and `visa_page_2.png` into it
4. Open the folder — copy the **Folder ID** from the URL:
   ```
   https://drive.google.com/drive/folders/1aBcDeFgHiJkLmNoPqRsTuVwXyZ
                                                       ^^^^^^^^^^^^^^^^^^^
                                                       This is the Folder ID
   ```

### 2. Create a New Google Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Click **"+ New Project"**

### 3. Add the Code Files

1. **Code.gs** — Delete the default code, paste the contents of `Code.gs`
2. **index.html** — Click **"+"** next to "Files" → **"HTML"** → name it `index` → paste the contents of `index.html`

### 4. Configure

In `Code.gs`, update these two values:

```javascript
const ADMIN_COPY_EMAIL = 'your-email@gmail.com';   // your email for copies (or '' to skip)
const VISA_FORM_FOLDER_ID = '1aBcDeFgHiJkLmNoPqRsTuVwXyZ';  // paste your Folder ID
```

### 5. Deploy

1. Click **"Deploy"** → **"New deployment"**
2. Gear icon → **"Web app"**
3. Set: Execute as **Me**, Who has access **Anyone**
4. Click **"Deploy"** → copy the **Web App URL**

### 6. Connect the URL

1. In `index.html`, find: `const WEB_APP_URL = 'YOUR_GOOGLE_APPS_SCRIPT_WEB_APP_URL';`
2. Replace with your URL
3. Save → **Deploy** → **Manage deployments** → edit → **New version** → **Deploy**

### 7. Test

1. Open your Web App URL
2. Fill in the form, upload photo & signature
3. Submit — check your email

---

## How It Works

1. User opens the form, fills all fields, uploads photo & signature
2. On submit, data goes to Apps Script
3. Script pulls `visa_page_1.png` and `visa_page_2.png` from your Google Drive folder
4. Builds an HTML email: answers on top, form images on bottom
5. Sends email to the applicant's email address
6. Optionally sends a copy to you
