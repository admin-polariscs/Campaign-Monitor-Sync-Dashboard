  _____  _                  _     _____      _            _     
 |  __ \| |                | |   |  __ \    | |          (_)    
 | |__) | | __ _ _ __   ___| |_  | |__) |__ | | __ _ _ __ _ ___ 
 |  ___/| |/ _` | '_ \ / _ \ __| |  ___/ _ \| |/ _` | '__| / __|
 | |    | | (_| | | | |  __/ |_  | |  | (_) | | (_| | |  | \__ \
 |_|    |_|\__,_|_| |_|\___|\__| |_|   \___/|_|\__,_|_|  |_|___/
                                                                
                                                                
                              GPL 3.0 LICENSE



Campaign Monitor Sync Dashboard with Flask + Pandas

Managing multiple subscriber lists across multiple databases is a nightmare when done manually. Each month, subscriber spreadsheets had to be uploaded one by one into Campaign Monitor, unsubscribes handled separately, and invalid addresses manually identified.

To solve this, I built a Flask-based dashboard that connects local Excel databases directly to Campaign Monitor via the API, with the following features:

✨ Features

One-click sync per database – or run all at once.

Real-time progress log in the browser (via SSE streaming).

Automatic unsubscribe handling – removes subscribers not present in the Excel database.

Invalid email detection – Campaign Monitor validation results are logged.

Export invalids to Excel – one workbook with a sheet per database containing invalid emails only.

Debug logging – shows active/removed subscribers in Campaign Monitor during unsubscribe checks.

🛠 Tech stack

Python 3.13

Flask (for the lightweight dashboard & streaming logs)

Pandas (Excel reading/writing, CSV/invalid exports)

Requests (API calls to Campaign Monitor)

XlsxWriter (multi-sheet Excel export)

⚙️ Requirements

Python ≥ 3.9

Installed packages:

pip install flask pandas requests xlsxwriter

📂 Data structure requirements

Each database must meet the following conditions for the sync to work:

File location is fixed – paths are hardcoded in the configuration. Moving/renaming files breaks the sync.

Excel format – .xlsx files with at least one column containing emails.

Email column – must contain “mail” in its name (e.g. Email, E-mail, Mail).

Optional columns – Name and Surname are mapped directly, all other columns are passed as custom fields to Campaign Monitor.

Unique list IDs – each database is linked to a specific Campaign Monitor list ID.

🚀 Usage

Configure your databases in the databases list inside app.py:

{
    "name": "name of the DB",
    "listId": "qdsgqdsfgqdsgqdf",
    "file": r"G:\My Drive\Detcetc\namedb.xlsx"
}


Run the app:

python app.py


Open in your browser:

http://127.0.0.1:5000


Choose a database (or “sync all”), optionally tick the unsubscribe option, and watch results stream live.

After sync, download invalid addresses with one click (invalid_emails.xlsx with a sheet per database).

📊 Example output (browser log)
Batch sync started: nameofdb.xlsx
Columns found: ['Name', 'Surname', 'E-mail', ...]
Total subscribers to sync: 3491
Batch 1: new=0, existing=997, duplicates=2, invalid=1
...
=== Final report ===
Total new: 0
Total existing: 2989
Total duplicates: 3
Total invalid: 8
--- Unsubscribe check ---
[DEBUG] Active subscribers in CM: 3476
[DEBUG] Example: ['jondo@allwaysdfive.com.au', 'johndoe@voyagfeslapara.com']
Unsubscribed: 0

💡 Why it matters

Before: manual updates across 30+  databases, high risk of errors, wasted hours.
Now: automated, auditable, and exportable sync in minutes.
