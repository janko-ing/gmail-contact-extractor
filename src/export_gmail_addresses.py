from __future__ import print_function
import os, re, time, sys
from datetime import datetime
import csv
from email.utils import getaddresses
import threading
import tkinter as tk
from tkinter import scrolledtext, messagebox

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build

SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# --------------------- Gmail API ---------------------
def get_service():
    creds = None
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return build('gmail', 'v1', credentials=creds)

def extract_name_email(headers):
    pairs = []
    for h in headers:
        if h['name'] in ['From', 'To', 'Cc', 'Bcc']:
            pairs.extend(getaddresses([h['value']]))
    return pairs

def fetch_batch(service, query):
    records = set()
    page_token = None
    quota_used = 0
    start_time = time.time()
    
    while True:
        # Check if we need to wait to respect rate limits
        if quota_used >= 14000:  # Leave some buffer
            elapsed = time.time() - start_time
            if elapsed < 60:  # If less than a minute has passed
                wait_time = 60 - elapsed + 1  # Wait until next minute + 1 sec buffer
                print(f"Rate limit approaching, waiting {wait_time:.1f} seconds...")
                time.sleep(wait_time)
                quota_used = 0
                start_time = time.time()
        
        # Fetch message list (5 quota units)
        results = service.users().messages().list(
            userId='me', q=query, maxResults=100, pageToken=page_token
        ).execute()
        quota_used += 5
        
        messages = results.get('messages', [])
        
        # Process messages in smaller batches to avoid hitting limits
        batch_size = min(50, (14000 - quota_used) // 5)  # Reserve quota for message.get calls
        
        for i in range(0, len(messages), batch_size):
            batch_messages = messages[i:i + batch_size]
            
            for msg in batch_messages:
                # Check quota before each message.get call
                if quota_used >= 14000:
                    elapsed = time.time() - start_time
                    if elapsed < 60:
                        wait_time = 60 - elapsed + 1
                        print(f"Rate limit approaching, waiting {wait_time:.1f} seconds...")
                        time.sleep(wait_time)
                        quota_used = 0
                        start_time = time.time()
                
                msg_data = service.users().messages().get(
                    userId='me', id=msg['id'], format='metadata',
                    metadataHeaders=['From','To','Cc','Bcc']
                ).execute()
                quota_used += 5
                
                headers = msg_data['payload']['headers']
                for name, email in extract_name_email(headers):
                    if email:
                        records.add((name.strip(), email.strip().lower()))
            
            # Small delay between batches to be gentle on the API
            time.sleep(0.2)
        
        page_token = results.get('nextPageToken')
        if not page_token:
            break
    
    return records

# --------------------- GUI ---------------------
class GmailExporterApp:
    def __init__(self, root):
        self.root = root
        root.title("Gmail Email Exporter")
        root.geometry("600x400")

        # Create main frame
        self.frame = tk.Frame(root)
        self.frame.pack(expand=True, fill='both', padx=10, pady=10)

        # Start button at the top
        self.start_btn = tk.Button(self.frame, text="Start Export", command=self.start_export)
        self.start_btn.pack(pady=5)

        # ScrolledText for log
        self.log = scrolledtext.ScrolledText(self.frame, state='disabled', wrap='word')
        self.log.pack(expand=True, fill='both')

    def log_message(self, msg):
        self.log.config(state='normal')
        self.log.insert(tk.END, msg + "\n")
        self.log.see(tk.END)
        self.log.config(state='disabled')
        self.root.update()

    def start_export(self):
        self.start_btn.config(state='disabled')
        threading.Thread(target=self.run_export).start()

    def run_export(self):
        try:
            self.log_message("Initializing Gmail API...")
            service = get_service()
            all_records = set()

            current_year = datetime.now().year
            for year in range(2005, current_year + 1):
                query = f"after:{year}/01/01 before:{year+1}/01/01"
                self.log_message(f"Fetching {year} emails...")
                batch = fetch_batch(service, query)
                self.log_message(f"  Found {len(batch)} addresses in {year}")
                all_records.update(batch)

            output_file = os.path.join(os.path.dirname(sys.executable if getattr(sys, 'frozen', False) else __file__), "emails.csv")
            with open(output_file, "w", newline="", encoding="utf-8") as f:
                writer = csv.writer(f)
                writer.writerow(["Name", "Email"])
                for name, email in sorted(all_records, key=lambda x: x[1]):
                    writer.writerow([name, email])

            self.log_message(f"✅ Done! Exported {len(all_records)} unique email addresses.")
            self.log_message(f"Saved to: {output_file}")
            messagebox.showinfo("Export Complete", f"Export finished!\nSaved to:\n{output_file}")

        except Exception as e:
            self.log_message(f"❌ Error: {e}")
            messagebox.showerror("Error", str(e))
        finally:
            self.start_btn.config(state='normal')

# --------------------- Main ---------------------
if __name__ == "__main__":
    root = tk.Tk()
    app = GmailExporterApp(root)
    root.mainloop()
