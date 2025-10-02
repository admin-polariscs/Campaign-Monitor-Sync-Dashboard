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

# === GLOBALS ===
progress_log = []
invalids_per_db = {}   # { db_name: [ {"Email":..., "Reason":...}, ... ] }

def log(msg):
    print(msg)
    progress_log.append(msg)

# === HTML TEMPLATE ===
html_template = """
<!DOCTYPE html>
<html lang="nl">
<head>
  <meta charset="UTF-8">
  <title>Mailing Sofitel Legend The Grand Sync</title>
  <style>
    body { font-family: Arial, sans-serif; max-width: 800px; margin: 40px auto; }
    h2 { margin-bottom: 10px; }
    #output { background: #f7f7f7; padding: 10px; border: 1px solid #ccc;
              white-space: pre-line; height: 400px; overflow-y: auto; }
    button { margin: 6px; padding: 8px 14px; }
  </style>
</head>
<body>
  <h2>Mailing Sofitel Legend The Grand Sync</h2>

  <label for="dbSelect">Kies database:</label>
  <select id="dbSelect">
    {% for db in databases %}
    <option value="{{ loop.index0 }}">{{ db.name }}</option>
    {% endfor %}
  </select>

  <br><br>
  <label>
    <input type="checkbox" id="doUnsub">
    Ook unsubscriben wat niet in Excel staat
  </label>
  <br>

  <button onclick="syncSelected()">Sync geselecteerde</button>
  <button onclick="syncAll()">Sync alle</button>
  <a href="/download_invalids"><button>Download ongeldige adressen</button></a>

  <h3>Resultaten (live)</h3>
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

# === HELPERS ===
def clean(value):
    if pd.isna(value):
        return ""
    s = str(value).strip()
    return "" if s.lower() == "nan" else s

def detect_email_column(df):
    possible_cols = [c for c in df.columns if "mail" in c.lower()]
    return possible_cols[0] if possible_cols else None

def get_active_subscribers(list_id):
    emails = set()
    page = 1
    while True:
        url = f"https://api.createsend.com/api/v3.2/lists/{list_id}/active.json?pagesize=1000&page={page}"
        r = requests.get(url, auth=AUTH)
        data = r.json()
        results = data.get("Results", [])
        for sub in results:
            emails.add(sub["EmailAddress"].lower().strip())
        total_pages = data.get("NumberOfPages", 1)
        if page >= total_pages:
            break
        page += 1
    return emails

def get_unsubscribed_subscribers(list_id):
    """Fetch unsubscribed emails from Campaign Monitor"""
    emails = set()
    page = 1
    while True:
        url = f"https://api.createsend.com/api/v3.2/lists/{list_id}/unsubscribed.json?pagesize=1000&page={page}"
        r = requests.get(url, auth=AUTH)
        data = r.json()
        results = data.get("Results", [])
        for sub in results:
            emails.add(sub["EmailAddress"].lower().strip())
        total_pages = data.get("NumberOfPages", 1)
        if page >= total_pages:
            break
        page += 1
    return emails

def unsubscribe_missing(list_id, drive_emails):
    cm_emails = get_active_subscribers(list_id)
    drive_set = set(e.lower().strip() for e in drive_emails if e)
    to_unsub = cm_emails - drive_set
    log(f"Aantal in Campaign Monitor (active): {len(cm_emails)}")
    log(f"Aantal in Excel: {len(drive_set)}")
    log(f"Te unsubscriben: {len(to_unsub)} adressen")
    for email in to_unsub:
        payload = {"EmailAddress": email}
        url = f"{API_BASE}{list_id}/unsubscribe.json"
        r = requests.post(url, auth=AUTH, json=payload)
        if r.status_code in (200, 201):
            log(f"⛔ Unsubscribed: {email}")
        else:
            log(f"⚠️ Unsub error {email}: {r.status_code} {r.text}")

def export_invalids_to_excel():
    if not invalids_per_db:
        log("⚠️ Geen ongeldige adressen om te exporteren.")
        return None
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    filename = f"invalid_emails_{timestamp}.xlsx"
    with pd.ExcelWriter(filename, engine="xlsxwriter") as writer:
        for db_name, rows in invalids_per_db.items():
            if rows:
                df = pd.DataFrame(rows)
                safe_name = db_name[:30]
                df.to_excel(writer, sheet_name=safe_name, index=False)
    log(f"✅ Ongeldige adressen geëxporteerd naar {filename}")
    return filename

def sync_file(filepath, list_id, db_name, do_unsub=False):
    if not os.path.exists(filepath):
        log(f"⚠️ Bestand niet gevonden: {filepath}")
        return

    try:
        df = pd.read_excel(filepath)
    except Exception as e:
        log(f"⚠️ Kon bestand niet lezen: {filepath}, fout: {e}")
        return

    log(f"Batch sync gestart: {os.path.basename(filepath)}")
    log(f"Kolommen gevonden: {list(df.columns)}")

    email_col = detect_email_column(df)
    if not email_col:
        log("⚠️ Geen kolom met e-mail gevonden")
        return

    # 🚫 Get unsubscribed list and skip them
    unsubscribed = get_unsubscribed_subscribers(list_id)
    log(f"[DEBUG] {len(unsubscribed)} unsubscribed addresses will be skipped")

    subscribers = []
    for _, row in df.iterrows():
        email = clean(row.get(email_col, ""))
        if not email:
            continue
        if email.lower() in unsubscribed:
            log(f"⏩ Skipping unsubscribed {email}")
            continue

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
            "Resubscribe": False,
            "ConsentToTrack": "Yes",
            "CustomFields": custom_fields
        })

    if not subscribers:
        log("⚠️ Geen geldige subscribers gevonden in dit bestand.")
        return

    url = f"{API_BASE}{list_id}/import.json"
    chunk_size = 1000
    total = len(subscribers)
    log(f"Totaal subscribers te syncen: {total}")

    total_new = total_existing = total_duplicates = total_failures = 0
    drive_emails = [s["EmailAddress"] for s in subscribers]
    invalids_per_db[db_name] = []

    for i in range(0, total, chunk_size):
        batch = subscribers[i:i + chunk_size]
        payload = {"Subscribers": batch, "Resubscribe": False}
        r = requests.post(url, auth=AUTH, json=payload)
        batch_nr = i // chunk_size + 1
        try:
            data = r.json()
        except:
            data = {}

        if r.status_code in (200, 201) or r.status_code == 400:
            rd = data.get("ResultData", {})
            new = rd.get("TotalNewSubscribers", 0)
            existing = rd.get("TotalExistingSubscribers", 0)
            dups = len(rd.get("DuplicateEmailsInSubmission", []))
            fails = len(rd.get("FailureDetails", []))
            total_new += new
            total_existing += existing
            total_duplicates += dups
            total_failures += fails

            for f in rd.get("FailureDetails", []):
                invalids_per_db[db_name].append({
                    "Email": f.get("EmailAddress"),
                    "Reason": f.get("Message")
                })

            log(f"Batch {batch_nr}: nieuwe={new}, bestaande={existing}, duplicaten={dups}, ongeldig={fails}")
        else:
            log(f"Batch {batch_nr}: fout {r.status_code} {r.text}")
        time.sleep(0.2)

    log("=== Eindrapport ===")
    log(f"Totaal nieuwe: {total_new}")
    log(f"Totaal bestaande: {total_existing}")
    log(f"Totaal duplicaten: {total_duplicates}")
    log(f"Totaal ongeldig: {total_failures}")

    if do_unsub:
        log("--- Unsubscribe check ---")
        unsubscribe_missing(list_id, drive_emails)
    else:
        log("--- Unsubscribe overslaan (niet aangevinkt) ---")

    export_invalids_to_excel()

# === ROUTES ===
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
    sync_file(db["file"], db["listId"], db["name"], do_unsub)
    return "OK"

@app.route("/sync_all")
def sync_all():
    global progress_log
    progress_log = []
    do_unsub = request.args.get("unsub") == "1"
    for db in databases:
        log(f"--- {db['name']} ---")
        sync_file(db["file"], db["listId"], db["name"], do_unsub)
    return "OK"

@app.route("/download_invalids")
def download_invalids():
    filename = export_invalids_to_excel()
    if filename and os.path.exists(filename):
        return send_file(filename, as_attachment=True)
    return "⚠️ Geen exportbestand gevonden. Run eerst een sync."

if __name__ == "__main__":
    app.run(debug=True, threaded=True)
