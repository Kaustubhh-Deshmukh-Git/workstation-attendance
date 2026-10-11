# 🏢 Workstation Attendance System (QR-Based Check-in)

A complete, production-grade Workstation Attendance System powered by **Google Apps Script** bound to a **Google Sheet** with a mobile-responsive **HTML5 QR Scanner Web App**.

Includes full support for Member profiles, organizational structure, and constraints:
- 🏛️ **College / Institution**
- 🔬 **Department / Branch**
- 💼 **Position in Workstation** (e.g. *Lab Lead, Research Fellow, Hardware Specialist, Intern*)

---

## 📌 Architecture & Workflow

```mermaid
flowchart TD
    subgraph Admin ["⚙️ Administration"]
        A["Admin adds Member (College, Dept, Position)"] --> B["setup() / Token Generator"]
        B -->|HMAC-SHA256 Secret| C["Permanent QR: MemberID | Token"]
        C --> D["Custom Menu: 'Send QR Codes'"]
        D -->|MailApp with Inline & Attachment| E["Member Receives QR ID Badge with Position & College"]
    end

    subgraph Entrance ["🚪 Entrance Kiosk / Smartphone"]
        E --> F["Member shows QR Badge"]
        G["Scanner Web App (Scanner.html)"] -->|Camera / USB Scanner| H["Scan QR Code"]
        F -.-> H
        H -->|HTTP POST| I["markAttendance(qrData)"]
    end

    subgraph Server ["☁️ Google Apps Script Backend (Code.gs)"]
        I --> J{"1. LockService Acquired?"}
        J -->|Yes| K{"2. HMAC Signature Valid?"}
        K -->|Yes| L{"3. Within Check-in Window?"}
        L -->|Yes| M{"4. Already Scanned Today?"}
        M -->|No| N["Append to 'Attendance_Log' with College, Dept & Position"]
        N --> O["Return SUCCESS (Green) with Position & College Tags"]
        M -->|Yes| P["Return ALREADY_MARKED (Amber)"]
        K -->|No| Q["Return INVALID / COUNTERFEIT (Red)"]
        L -->|No| R["Return OUTSIDE_HOURS (Red)"]
    end

    subgraph Reporting ["📊 End-of-Day Reporting"]
        S["Time-driven Trigger (e.g. 21:00)"] --> T["generateDailyReport()"]
        T --> U["Create Sheet: Report_YYYY-MM-DD with 8 Columns"]
        T --> V["Email KPI Summary & Link to Admin"]
    end
```

---

## 📂 Project Structure

```
d:\Workstation Attendance system\
├── appsscript.json        # Manifest file with Asia/Kolkata timezone & scopes
├── Code.gs                # Complete backend logic (HMAC, Sheets, MailApp, Reports, Member Constraints)
├── Scanner.html           # Modern glassmorphism Web App with HTML5-QRCode, Live Meta Pills & Web Audio
├── sample_members.csv     # Sample CSV with College, Department, and Position constraints
└── README.md              # Complete setup and operational guide
```

---

## 🏛️ Member Schema & Constraints

| Column | Field Name | Description | Example |
|---|---|---|---|
| **A** | `Member ID` | Unique permanent identifier | `WS-001` |
| **B** | `Name` | Full name of the member | `Aarav Sharma` |
| **C** | `Email` | Recipient email address | `aarav.sharma@example.com` |
| **D** | `College` | College, University, or Institute | `College of Engineering Pune` |
| **E** | `Department` | Academic or technical department | `Computer Engineering` |
| **F** | `Position in Workstation` | Role / designation inside the workstation | `Lab Lead` |
| **G** | `Token` | Cryptographic HMAC-SHA256 signature | *(Auto-generated)* |
| **H** | `QR Link` | Direct link to high-res QR code image | *(Auto-generated)* |
| **I** | `Email Sent` | Status flag (`Yes` or `No`) | `No` |

---

## 🚀 Step-by-Step Setup Instructions

### Step 1: Create a Google Spreadsheet
1. Open [Google Sheets](https://sheets.new) in your browser.
2. Name your spreadsheet: **Workstation Attendance System**.

---

### Step 2: Open Apps Script & Add Files
1. In your Google Sheet, go to **Extensions** &rarr; **Apps Script**.
2. Rename the project at top-left to: **Workstation Attendance System**.
3. **Enable Manifest File**:
   - Click the **Gear icon** (Project Settings) on the left sidebar.
   - Check the box: `Show "appsscript.json" manifest file in editor`.
4. **Copy `appsscript.json`**:
   - Return to the **Editor** (`<>` icon).
   - Click `appsscript.json` and replace all its contents with the code from [`appsscript.json`](file:///d:/Workstation%20Attendance%20system/appsscript.json).
5. **Copy `Code.gs`**:
   - Open `Code.gs` in the editor.
   - Replace its entire contents with the code from [`Code.gs`](file:///d:/Workstation%20Attendance%20system/Code.gs).
6. **Add `Scanner.html`**:
   - Click the **`+`** icon next to *Files* &rarr; Select **HTML**.
   - Name the file `Scanner` (Apps Script automatically appends `.html`).
   - Paste the code from [`Scanner.html`](file:///d:/Workstation%20Attendance%20system/Scanner.html).
7. Press `Ctrl + S` (or `Cmd + S`) to save all files.

---

### Step 3: Run System Initialization (`setup`)
1. In the Apps Script toolbar, ensure the function dropdown is set to **`setup`**.
2. Click **Run**.
3. **Grant Google Permissions**:
   - When the *"Authorization required"* modal appears, click **Review permissions**.
   - Choose your Google account.
   - Click **Advanced** &rarr; Click **Go to Workstation Attendance System (unsafe)**.
   - Click **Allow**.
4. Once `setup` finishes execution:
   - Return to your Google Spreadsheet tab and refresh the page.
   - You will see 3 sheets created and styled:
     1. **`Members`**: Pre-filled with sample members having `College`, `Department`, and `Position in Workstation`.
     2. **`Attendance_Log`**: Ready to record check-in events with 9 columns.
     3. **`Config`**: Pre-filled with workstation settings.
   - A new custom menu **🏢 Workstation Attendance** will appear in the top menu bar.

---

### Step 4: Deploy the Scanner Web App
1. In the top-right corner of Apps Script, click **Deploy** &rarr; **New deployment**.
2. Click the gear icon next to *Select type* &rarr; Choose **Web app**.
3. Configure the deployment:
   - **Description**: `Workstation Scanner v1.0`
   - **Execute as**: `Me (<your-email@gmail.com>)`
   - **Who has access**: `Anyone` *(or "Anyone within your organization" if using Google Workspace)*.
4. Click **Deploy**.
5. Copy the **Web app URL** (ends in `/exec`).
   > 💡 **Tip:** Bookmark this URL on your tablet or smartphone kiosk at the workstation entrance.

---

## 📋 Features & How to Use

### 1. Sending QR Badges to Members (`Feature 1`)
1. Open the Google Sheet &rarr; Click menu **🏢 Workstation Attendance** &rarr; **📧 Send QR Codes to Members**.
2. **Badge Content**: Emailed passes display the member's **Position**, **Department**, **College**, and permanent QR badge.
3. **Email Provider Setup (Gmail / Brevo / Resend)**:
   - **Gmail (Default)**: Free Google accounts send up to 100 emails/day.
   - **Brevo API (300 Free/Day)** or **Resend (3,000 Free/Month)**:
     - Click **🏢 Workstation Attendance** &rarr; **⚡ Email Provider Setup (Brevo / Resend / Gmail)**.
     - Paste your free API key, and click **Save Settings** to bypass Gmail quota limits.
4. **Alternative: Export All QR Badges to Google Drive**:
   - Click **🏢 Workstation Attendance** &rarr; **📁 Export All QR Badges to Google Drive**.
   - Generates individual PNG badges (`WS-001_Aarav_Sharma.png`) inside a new Google Drive folder.

### 2. Adding New Members with Constraints
1. Go to the **`Members`** sheet.
2. Add a new row:
   - Column A: `WS-011`
   - Column B: `Member Name`
   - Column C: `member.email@example.com`
   - Column D: `College / University Name`
   - Column E: `Department`
   - Column F: `Position in Workstation` (e.g. *Research Intern*)
   - Leave `Token` and `QR Link` blank; set `Email Sent` to `No`.
3. Click **🏢 Workstation Attendance** &rarr; **🚀 Initialize / Refresh System Setup**.
4. The system will automatically compute the cryptographic HMAC token and QR code URL for the new member!
5. Click **📧 Send QR Codes to Members** to dispatch their pass.

### 3. Using the Scanner Web App (`Feature 2 & 3`)
- Open the Web App URL on any smartphone, iPad/tablet, or laptop webcam.
- **Camera Scanning**: Present the QR code on the phone screen or printed card within the reticle.
- **Instant Visual Identification**:
  - 🟢 **Green ("Welcome, [Name]!")**: Check-in logged with timestamp, position badge (`💼 Lab Lead`), and academic tag (`🔬 Computer Eng • COEP`).
  - 🟠 **Amber ("[Name], already marked present today")**: Duplicate scan prevented.
  - 🔴 **Red ("Outside Check-in Window" / "Invalid QR")**: Access denied.
- **Auto-Reset**: Resets back to continuous scanning in 1.8 seconds.
- **USB Scanner / Manual Fallback**: Connect a USB handheld barcode reader or type `WS-001|<token>` into the input box below the camera.

### 4. End-of-Day Daily Attendance Report (`Feature 4`)
- **Automated Scheduling**:
  1. Click **🏢 Workstation Attendance** &rarr; **⏰ Setup Daily Report Trigger**.
  2. The system reads the `Report Generation Time` from `Config` (e.g., `21:00`) and automatically runs at that hour every evening.
  3. It generates a new sheet **`Report_YYYY-MM-DD`** with all 8 columns (`Member ID`, `Name`, `Email`, `College`, `Department`, `Position`, `Status`, `Check-in Time`), color-codes Present (Green) vs Absent (Red), and emails the summary KPI to the `Admin Email`.
- **Manual Generation**:
  - Click **🏢 Workstation Attendance** &rarr; **📊 Generate Today's Report** anytime to produce today's report on demand.

---

## ⚙️ Configuration Options (`Config` Sheet)

| Setting | Default | Description |
|---|---|---|
| **Workstation Name** | `Main Innovation Lab` | Displayed in Web App header, emails, and reports |
| **Check-in Start Time** | `08:00` | Earliest allowed scan time (HH:mm 24-hr format) |
| **Check-in End Time** | `20:00` | Latest allowed scan time (HH:mm 24-hr format) |
| **Report Generation Time** | `21:00` | Trigger hour for automated daily report & email |
| **Admin Email** | `admin@example.com` | Receives end-of-day summary reports |
| **Allowed Re-scan** | `No` | Set to `Yes` if members need to scan multiple times/day |

---

## 🔒 Security & Anti-Counterfeiting Details

- **Cryptographic Signing**: Each token is calculated using `HMAC-SHA256(MemberID, SecretKey)`.
- **Protected Secret**: The `HMAC_SECRET_KEY` is stored in Google Apps Script `ScriptProperties` (never exposed to client-side code or spreadsheet viewers).
- **Concurrency Protection**: Check-in transactions utilize `LockService.getScriptLock()` to prevent race conditions when multiple members scan simultaneously.
