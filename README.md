# PPT Slide Cleanup — Setup Guide

## What this is
A PowerPoint task pane add-in with a single "Cleanup Slide" button.
It reads your slide, duplicates it, then uses Claude AI to reformat the
duplicate to match your slide master (fonts, colours, sizes, positions).

---

## Step 1 — Put this on GitHub (free)

1. Go to https://github.com/new
2. Create a new **public** repository called `ppt-cleanup`
3. Upload all files from this folder into it (drag & drop in the GitHub UI)

---

## Step 2 — Deploy to Vercel (free)

1. Go to https://vercel.com and sign up (use your GitHub account)
2. Click **Add New → Project**
3. Import your `ppt-cleanup` GitHub repo
4. Vercel auto-detects Create React App — just click **Deploy**
5. After ~1 minute you get a URL like: `https://ppt-cleanup-abc123.vercel.app`

---

## Step 3 — Update the manifest

Open `manifest.xml` and replace every occurrence of:
```
https://YOUR-APP.vercel.app
```
with your actual Vercel URL (e.g. `https://ppt-cleanup-abc123.vercel.app`).

Save the file.

---

## Step 4 — Sideload in PowerPoint (Windows)

1. Open PowerPoint
2. Go to **Insert → Get Add-ins → Upload My Add-in** (or via the gear icon)
3. Browse to and select your `manifest.xml` file
4. Click **Upload**
5. A new **"Cleanup Slide"** button will appear in the **Home** ribbon under **AI Tools**

---

## Step 5 — Use it

1. Open any presentation
2. Click on a slide in the panel to select it
3. Click **Cleanup Slide** in the ribbon (or **Insert → My Add-ins → Slide Cleanup**)
4. Hit the big **✦ Cleanup Slide** button in the task pane
5. Watch the log — it duplicates the slide and applies AI fixes

---

## Notes

- The original slide is **always preserved** — cleanup works on a duplicate
- Requires Office 2016+ or Microsoft 365 (for PowerPoint JS API support)
- The add-in calls `api.anthropic.com` — make sure your network allows it
- Icons (icon-16.png, icon-32.png, icon-80.png) can be any small PNG you host in `/public`

---

## Customising the master theme rules

The formatting rules are in `src/App.js` inside `buildPrompt()`.
Edit the "SLIDE MASTER RULES" section to match your actual theme:
- Font names
- Colours (hex)
- Font sizes
- Position values (in inches)
