# Schedule Email Worker

Cloudflare Email Worker that processes Excel schedule files sent to `schedule@statisticalprograms.org` and uploads them to Firebase.

## Features

- Receives emails with Excel attachments (.xlsx, .xls)
- Auto-detects Nevada vs Utah schedules
- Parses schedule data (DATE, TECH, TEST, ZIP, IOCS, RT columns)
- Uploads to Firebase Realtime Database
- Replaces old state data with new data

## Setup

### 1. Install Dependencies

```bash
npm install
```

### 2. Configure Cloudflare

This worker is connected to Cloudflare via GitHub integration. Any push to the `main` branch will automatically deploy.

### 3. Email Routing Setup

In Cloudflare Dashboard:
- Email Routing → statisticalprograms.org
- Routing Rules → schedule@statisticalprograms.org → Send to Worker → schedule-email-worker

## Usage

Send an Excel file to `schedule@statisticalprograms.org` with the following columns:
- DATE (required)
- TECH (required)
- TEST (optional)
- ZIP (optional)
- IOCS (optional)
- RT (optional)

The worker will:
1. Detect state from filename or data
2. Parse Excel data
3. Delete old data for that state
4. Upload new schedule entries to Firebase

## Firebase Structure
