# C64 Facilities and Maintenance Report Form

A custom, fully-featured Google Apps Script Web Application designed to streamline facilities and maintenance reporting for Anglo Singapore International School. This application replaces third-party form builders with a native, robust, and cost-effective solution integrated directly with Google Workspace.

## 🌟 Key Features

*   **Modern User Interface:** A clean, responsive, and mobile-friendly form featuring the official school branding and intuitive layout.
*   **Comprehensive Location Selection:** Includes a complete, dynamic dropdown of over 160 specific school locations across multiple buildings and floors.
*   **Direct Google Drive Integration:** Users can upload images of maintenance issues directly. Images are securely and automatically saved to a designated Google Drive folder with smart sharing permissions to ensure accessibility for officers.
*   **Google Sheets Database:** All submissions are instantly logged into a Google Sheet for permanent, easily accessible record-keeping.
*   **Task Assignment:** Officers can assign specific requests to team members by entering their email address directly in the dashboard.
*   **Analytics & Reports Dashboard:** A powerful, built-in analytics page (`?page=analytics`) powered by Chart.js. It features dynamic filtering (daily/weekly/monthly), visual charts (status, priority, location), and a complete export suite (CSV, Excel, PDF, Print) without relying on external third-party BI tools.
*   **Automated Email Notifications:**
    *   **Submission Confirmations:** Senders receive immediate confirmation emails.
    *   **Officer Alerts:** The Facilities Team is instantly notified of new requests with direct links to review.
    *   **Status Updates & Assignments:** When an officer updates a status or assigns a task, an automated email is sent to the requester detailing the update. Assigned team members are automatically CC'd.

## 🛠️ Technology Stack

*   **Backend:** Google Apps Script (`Code.gs`)
*   **Frontend:** HTML5, CSS3, JavaScript (`Index.html`, `Dashboard.html`, `Analytics.html`)
*   **Libraries:** jQuery, DataTables (with export extensions), Chart.js, SweetAlert2
*   **Data Storage:** Google Sheets, Google Drive
*   **Email Delivery:** Google MailApp

## 🚀 Setup & Installation

1.  **Create a Google Apps Script Project:**
    *   Open your destination Google Sheet.
    *   Go to `Extensions` > `Apps Script`.
2.  **Copy Files:**
    *   Copy the contents of `Code.gs`, `Index.html`, `Dashboard.html`, and `Analytics.html` from this repository into your Apps Script project.
3.  **Sheet Preparation:**
    *   Ensure your destination Google Sheet has a header named `Assigned To` in Column M (the 13th column) to support the Task Assignment feature.
3.  **Configuration:**
    *   In `Code.gs`, update the configuration variables at the top of the file:
        *   `SHEET_ID`: The ID of your Google Sheet.
        *   `DRIVE_FOLDER_ID`: The ID of the Google Drive folder for image uploads.
        *   `OFFICER_EMAIL`: The email address of the facilities officer.
4.  **Authorization (CRITICAL):**
    *   Select the `setupPermissions` function from the dropdown menu at the top of the Apps Script editor.
    *   Click **Run**.
    *   Follow the prompts to grant the necessary permissions (Drive, Sheets, Email). *This step guarantees that the script has the required Drive access to save uploaded files.*
5.  **Deployment:**
    *   Click **Deploy** > **New deployment**.
    *   Select **Web app**.
    *   Set **Execute as:** `Me`.
    *   Set **Who has access:** `Anyone` (or restrict to your Google Workspace domain as needed).
    *   Click **Deploy** to generate your live form URL.

## 🔒 Security Notes

This application handles Google Drive permissions gracefully. If deployed within a strict Google Workspace domain (like a school) that prevents public link sharing, the script is designed to automatically fall back to sharing the file with anyone inside the domain (`DOMAIN_WITH_LINK`), preventing submission crashes.

---
*Built for Anglo Singapore International School*
