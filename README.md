# wedding-invitation

## Google Sheets RSVP Setup

1. Create a Google Sheet for RSVP submissions.
2. Open Apps Script from that sheet and paste in [google-sheets-backend/Code.gs](google-sheets-backend/Code.gs).
3. Replace `PASTE_YOUR_SHEET_ID_HERE` with your Google Sheet ID.
4. Deploy the Apps Script as a Web App:
	- Execute as: `Me`
	- Who has access: `Anyone`
5. Copy the deployed Web App URL.
6. Paste that URL into `googleScriptUrl` in both invitation files:
	- [cuttack/index.html](cuttack/index.html)
	- [bhilai/index.html](bhilai/index.html)

If `googleScriptUrl` is left blank, the pages will fall back to the current manual WhatsApp or email RSVP flow.