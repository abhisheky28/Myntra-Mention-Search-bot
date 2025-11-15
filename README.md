# Unlinked Mention Finder – Google Apps Script ⚙️
A fully automated Google Apps Script that finds brand mentions across the web, detects whether they are linked or unlinked, and logs opportunities into a Google Sheets dashboard.
This tool helps SEO teams quickly discover unlinked mentions and convert them into high-value backlink opportunities.

# 📌 Features
1. Uses Google Custom Search Engine (CSE) API to find brand mentions.
2. Scrapes pages safely using a Googlebot-style user agent.
3. Detects unlinked mentions by scanning page content.

# Automatically logs results in a Google Sheet:
New opportunities
Context snippets
Date found
Status tracking
Moves processed entries to Archive automatically.
Includes a daily scheduled trigger for full automation.
Multi-page search (up to 30 results per query) for deeper discovery.

# 📁 Sheet Structure
Your spreadsheet must include these sheets:
Sheet Name	Purpose
Dashboard & Controls	Status, trigger info, excluded domains, counters
Queries	List of search queries + last checked time
Results – New	Fresh unlinked mentions
Archive	Previously reviewed or processed mentions

# ⚙️ Setup Requirements
Before running, update the script with:
API_KEY → Your Google CSE API key
CSE_ID → Your Custom Search Engine ID
BRAND_NAME → Brand to detect in content
BRAND_DOMAIN → Domain used to check for links
Also ensure that the Queries sheet has queries in column A.

#🚀 How to Use?
Paste the script into Apps Script.
Update your API key, CSE ID, brand name, and domain.
Reload the Sheet → Menu Mention Finder will appear.
Click ▶️ Find New Mentions to run manually.
(Optional) Click Setup Daily Trigger to automate daily scans.

# 📌 What the Script Does Internally
Builds paginated CSE requests (up to 30 results/query).

Extracts URLs and checks:
Excluded domains
Duplicates (Results + Archive)
Brand mention without a backlink
Fetches page HTML using a safer Googlebot UA.
Extracts a clean context snippet around the brand name.
Writes the result into the sheet with status "New".

🗂️ Automation
The script includes automatic:
Daily trigger creation
Row movement from Results → Archive when status changes
Daily query counter in the dashboard
