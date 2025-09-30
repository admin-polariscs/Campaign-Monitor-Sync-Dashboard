"""
_____  _                  _     _____      _            _     
 |  __ \| |                | |   |  __ \    | |          (_)    
 | |__) | | __ _ _ __   ___| |_  | |__) |__ | | __ _ _ __ _ ___ 
 |  ___/| |/ _` | '_ \ / _ \ __| |  ___/ _ \| |/ _` | '__| / __|
 | |    | | (_| | | | |  __/ |_  | |  | (_) | | (_| | |  | \__ \
 |_|    |_|\__,_|_| |_|\___|\__| |_|   \___/|_|\__,_|_|  |_|___/

 PLANET POLARIS (app.py)
Licensed under GPL 3.0
                                                                
"""                                                          

import os
import pandas as pd
import requests
import time
from flask import Flask, render_template_string, request, Response, send_file

app = Flask(__name__)

# === SETTINGS ===
API_KEY = "YOUR_API_KEY_HERE"   # <-- Put your Campaign Monitor API key here
API_BASE = "https://api.createsend.com/api/v3.2/subscribers/"
AUTH = (API_KEY, "x")

# Example mapping: local Excel file → Campaign Monitor listId
# Replace with your own
databases = [
    {
        "name": "Example-List-1",
        "listId": "LIST_ID_HERE",
        "file": r"C:\path\to\Example1.xlsx"
    },
    {
        "name": "Example-List-2",
        "listId": "LIST_ID_HERE",
        "file": r"C:\path\to\Example2.xlsx"
    }
]

# === Globals ===
progress_log = []
invalids_per_db = {}

def log(msg):
    print(msg)
    progress_log.append(msg)

# === HTML Template ===
html_template = """
<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <title>Campaign Monitor Sync</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 700px; margin: 40px auto; }
    #output { background: #f7f7f7; padding: 10px; border: 1px solid #ccc;
              white-space: pre-line; height: 400px; overflow-y: auto; }
    button { margin: 6px; padding: 8px 14px; }
  </style>
</head>
<body>
  <h2>Campaign Monitor Sync Dashboard</h2>

  <label for="dbSelect">Choose database:</label>
  <select id="dbSelect">
    {% for db in databases %}
    <option value="{{ loop.index0 }}">{{ db.name }}</option>
    {% endfor %}
  </select>

  <br><br>
  <label>
    <input type="checkbox" id="doUnsub">
    Also unsubscribe missing addresses
  </label>
  <br>

  <button onclick="syncSelected()">Sync selected</button>
  <button onclick="syncAll()">Sync all</button>
  <a href="/download_invalids"><button>Download invalids</button></a>

  <h3>Results (live)</h3>
  <div id="output"></div>

  <script>
    let evtSource;
    function startStream() {
      const output = document.getElementById("output");
      output.innerText = ""; 
      if (evtSource) evtSource.close();
      evtSource = new EventSource("/stream");
      evtSource.onmessage = function(e) {
        output.innerText += e.data + "\\n";
        output.scrollTop = output.scrollHeight;
      };
    }

    function syncSelected() {
      const idx = document.getElementById("dbSelect").value;
      const unsub = document.getElementById("doUnsub").checked ? "1" : "0";
      startStream();
      fetch("/sync/" + idx + "?unsub=" + unsub);
    }

    function syncAll() {
      const unsub = document.getElementById("doUnsub").checked ? "1" : "0";
      startStream();
      fetch("/sync_all?unsub=" + unsub);
    }
  </script>
</body>
</html>
"""

# === Helpers ===
def clean(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    if s.lower() == "nan":
        return ""
    return s

def detect_email_column(df):
    """Detects the email column automatically"""
    possible_cols = [c for c in df.columns if "mail" in c.lower()]
    return possible_cols[0] if possible_cols else None

def get_active_subscribers(list_id):
    """Fetch all active subscribers from Campaign Monitor"""
    emails = set()
    page = 1
    while True:
        url = f"https://api.createsend.com/api/v3.2/lists/{list_id}/active.json?pagesize=1000&page={page}"
        r = requests.get(url, auth=AUTH)
        data = r.json()
        results = data.get("Results", [])
        for sub in results:
            emails.add(sub["EmailAddress"].lower().strip())
        if page >= data.get("NumberOfPages", 1):
            break
        page += 1
    return emails

def unsubscribe_missing(list_id, drive_emails):
    cm_emails = get_active_subscribers(list_id)
    drive_set = set(e.lower().strip() for e in drive_emails if e)
    to_unsub = cm_emails - drive_set
    log(f"Unsubscribing {len(to_unsub)} addresses")
    for email in to_unsub:
        payload = {"EmailAddress": email}
        url = f"{API_BASE}{list_id}/unsubscribe.json"
        requests.post(url, auth=AUTH, json=payload)
        log(f"⛔ Unsubscribed: {email}")

def sync_file(filepath, list_id, do_unsub=False):
    if not os.path.exists(filepath):
        log(f"⚠️ File not found: {filepath}")
        return

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        log(f"⚠️ Could not read file {filepath}, error: {e}")
        return

    log(f"Batch sync started: {os.path.basename(filepath)}")
    log(f"Columns found: {list(df.columns)}")

    email_col = detect_email_column(df)
    if not email_col:
        log("⚠️ No email column found (look for 'Email', 'E-mail', 'Mail'...)")
        return

    subscribers = []
    drive_emails = []
    for _, row in df.iterrows():
        email = clean(row.get(email_col, ""))
        if not email:
            continue
        drive_emails.append(email)
        firstname = clean(row.get("Name", ""))
        lastname = clean(row.get("Surname", ""))

        custom_fields = []
        for col in df.columns:
            if col not in [email_col, "Name", "Surname"]:
                value = clean(row.get(col, ""))
                if value:
                    custom_fields.append({"Key": col, "Value": value})

        subscribers.append({
            "EmailAddress": email,
            "Name": f"{firstname} {lastname}".strip(),
            "Resubscribe": True,
            "ConsentToTrack": "Yes",
            "CustomFields": custom_fields
        })

    url = f"{API_BASE}{list_id}/import.json"
    total_new = total_existing = total_duplicates = total_failures = 0
    invalids_per_db[list_id] = []

    chunk_size = 1000
    for i in range(0, len(subscribers), chunk_size):
        batch = subscribers[i:i + chunk_size]
        payload = {"Subscribers": batch, "Resubscribe": True}
        r = requests.post(url, auth=AUTH, json=payload)
        data = r.json()
        rd = data.get("ResultData", {})
        new = rd.get("TotalNewSubscribers", 0)
        existing = rd.get("TotalExistingSubscribers", 0)
        dups = len(rd.get("DuplicateEmailsInSubmission", []))
        fails = len(rd.get("FailureDetails", []))
        invalids = rd.get("FailureDetails", [])

        total_new += new
        total_existing += existing
        total_duplicates += dups
        total_failures += fails

        for f in invalids:
            invalids_per_db[list_id].append(f)

        log(f"Batch: new={new}, existing={existing}, duplicates={dups}, invalid={fails}")

    log("=== Final report ===")
    log(f"New: {total_new}, Existing: {total_existing}, Duplicates: {total_duplicates}, Invalid: {total_failures}")

    if do_unsub:
        unsubscribe_missing(list_id, drive_emails)

def export_invalids_to_excel():
    if not invalids_per_db:
        log("⚠️ No invalid addresses to export.")
        return None

    filename = "invalid_emails.xlsx"
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        for db_name, rows in invalids_per_db.items():
            if rows:
                df = pd.DataFrame(rows)
                df.to_excel(writer, sheet_name=str(db_name)[:30], index=False)
    log(f"✅ Invalid addresses exported to {filename}")
    return filename

# === Routes ===
@app.route("/")
def index():
    return render_template_string(html_template, databases=databases)

@app.route("/stream")
def stream():
    def event_stream():
        last_index = 0
        while True:
            if last_index < len(progress_log):
                for msg in progress_log[last_index:]:
                    yield f"data: {msg}\n\n"
                last_index = len(progress_log)
            time.sleep(1)
    return Response(event_stream(), mimetype="text/event-stream")

@app.route("/sync/<int:idx>")
def sync_selected(idx):
    global progress_log
    progress_log = []
    db = databases[idx]
    do_unsub = request.args.get("unsub") == "1"
    sync_file(db["file"], db["listId"], do_unsub)
    return "OK"

@app.route("/sync_all")
def sync_all():
    global progress_log
    progress_log = []
    do_unsub = request.args.get("unsub") == "1"
    for db in databases:
        log(f"--- {db['name']} ---")
        sync_file(db["file"], db["listId"], do_unsub)
    return "OK"

@app.route("/download_invalids")
def download_invalids():
    filename = export_invalids_to_excel()
    if not filename or not os.path.exists(filename):
        return "⚠️ No export file found. Run a sync first."
    return send_file(filename, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True, threaded=True)
