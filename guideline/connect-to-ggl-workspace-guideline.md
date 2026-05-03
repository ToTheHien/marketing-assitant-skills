# Google Workspace CLI — Setup Guide

## What This Is

The Google Workspace CLI (`gws`) gives Claude Code direct access to your Google apps — Gmail, Drive, Docs, Sheets, Calendar, and more. One tool, one setup, and Claude Code becomes a personal assistant that can actually manage your Google life.

---

## Step 1: Use Your Main Google Account

Use your main Google account for all steps below. Proceed to Step 2.

---

## Step 2: Install the CLI

Open your terminal and check Node.js is installed:

```bash
node --version
```

If you don't have Node.js, install it from https://nodejs.org first.

Install the Workspace CLI:

```bash
npm install -g @googleworkspace/cli
```

Verify the installation:

```bash
gws --version
```

---

## Step 3: Create a Google Cloud Project

1. Go to https://console.cloud.google.com
2. Sign in with your Google account
3. Click the project dropdown at the top → **New Project**
4. Name it `AI Assistant CLI` (or anything you want) → **Create**
5. Make sure that project is selected in the top dropdown
6. **Save your Project ID** — this is different from the project name. Click the project dropdown and you'll see both a Name column and an ID column. The ID looks like `ai-assistant-cli-438219` (lowercase with a number). Copy it somewhere for later.

---

## Step 4: Configure the OAuth Consent Screen

1. In the left sidebar → **APIs & Services** → **OAuth consent screen**
2. Select **External** → **Create**
3. Fill in the required fields:
   - **App name:** `GWS CLI` (or anything)
   - **User support email:** your email
   - **Developer contact email:** your email
   - Everything else: leave blank / skip
4. Click through the **Scopes** page — skip it
5. Click **Audience** in the left sidebar → **Publish App** → confirm

> **Important:** If you leave the app in Testing mode, your login will expire every 7 days and you'll have to re-authenticate weekly. Publishing the app prevents that.
> You'll see a "unverified app" warning during login — that's normal, it's your own app on your own account.

---

## Step 5: Create OAuth Credentials

1. In the left sidebar → **APIs & Services** → **Credentials**
2. Click **+ Create Credentials** → **OAuth client ID**
3. **Application type:** Desktop app
4. **Name:** `GWS CLI` (or anything)
5. Click **Create**
6. Click **Download JSON** on the popup
   - If you accidentally closed it, go back to Credentials, click on your client, and the download button is next to the Client Secret field
7. The downloaded file will have a long name — rename it to exactly `client_secret.json`
8. Save it to:
   - **Mac/Linux:** `~/.config/gws/client_secret.json`
   - **Windows:** `C:\Users\YourName\.config\gws\client_secret.json`
   - You may need to create the `gws` folder

---

## Step 6: Enable Billing & Google APIs

### Enable Billing

1. In the left sidebar → **Billing** (or go to https://console.cloud.google.com/billing)
2. Click **Link a billing account** → add a credit card

> The APIs used here are free or have generous free tiers. Google requires billing to be enabled before some APIs will work, but personal use costs nothing.

### Enable APIs

1. In the left sidebar → **APIs & Services** → **Library**
2. Search for and enable each of the following (click into each → click **Enable** → go back):
   - Gmail API
   - Google Drive API
   - Google Calendar API
   - Google Docs API
   - Google Sheets API
   - Google Slides API *(optional)*

---

## Step 7: Log In

In your terminal:

```bash
gws auth login
```

Your browser will open. Sign in with your account and approve the permissions.

You'll see a **"Google hasn't verified this app"** screen. Click **Advanced** → **Go to [app name] (unsafe)**. This is normal — it's your own app.

Test that it works:

**Mac/Linux:**
```bash
gws gmail users messages list --params '{"userId": "me", "maxResults": 5}'
```

**Windows (PowerShell):**
```powershell
gws gmail users messages list --params '{\"userId\": \"me\", \"maxResults\": 5}'
```

You should see your recent emails. If you do, you're all set.

---

## Step 7b: If You Get a Decryption Error

If you see an error like `Failed to decrypt credentials`, make sure you're on the latest version. Earlier versions (0.9.1 and below) had a bug where the encryption key was lost after login — fixed in version 0.10.0.

Update and retry:

```bash
npm install -g @googleworkspace/cli@latest
gws auth login
```

---

## You're Done

Claude Code now has access to your Google Workspace. You can use the `gws-*` skills in this project to interact with Gmail, Drive, Docs, Sheets, and Calendar.
