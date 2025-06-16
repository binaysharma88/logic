# ✅ Updated SignNowAutomationGUI with Multi-Token Support
import requests
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, scrolledtext
import pandas as pd
from pathlib import Path
import random
from edit_sendplan_gui import edit_sendplan_gui

class SignNowAutomationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("SignNow Automation")
        self.root.geometry("900x600")
        self.root.configure(bg="white")

        self.tokens = []
        

        self.recipients_file = Path("recipients.xlsx")
        self.sendplan_file = Path("sendplan.xlsx")
        self.token_input_file = Path("signnow_accounts.xlsx")
        self.token_output_file = Path("tokens.xlsx")
        self.files_dir = Path("files")
        self.auto_refresh_id = None

        self.ensure_files_exist()
        self.create_widgets()
        self.load_tokens_from_file()

    def load_tokens_from_file(self):
        path = self.token_output_file
        if not path.exists():
            df = pd.DataFrame(columns=["Token", "Limit", "Email"])
            df.to_excel(path, index=False)
        df = pd.read_excel(path)
        self.tokens = []
        for _, row in df.iterrows():
            self.tokens.append({
                "token": row["Token"],
                "limit": int(row["Limit"]),
                "used": 0,
                "email": row["Email"]
            })

    def fetch_tokens_from_excel(self):
        if not self.token_input_file.exists():
            messagebox.showerror("Missing File", f"{self.token_input_file.name} not found.")
            return

        try:
            df = pd.read_excel(self.token_input_file)
            if not all(col in df.columns for col in ["Email", "Password", "ClientID", "ClientSecret"]):
                messagebox.showerror("Invalid Format", "Excel must contain: Email, Password, ClientID, ClientSecret.")
                return

            tokens = []
            for _, row in df.iterrows():
                payload = {
                    "username": row["Email"],
                    "password": row["Password"],
                    "grant_type": "password",
                    "client_id": row["ClientID"],
                    "client_secret": row["ClientSecret"]
                }
                headers = {"Content-Type": "application/x-www-form-urlencoded"}
                try:
                    res = requests.post("https://api.signnow.com/oauth2/token", data=payload, headers=headers)
                    res.raise_for_status()
                    access_token = res.json().get("access_token")
                    if access_token:
                        tokens.append({"Token": access_token, "Limit": 10, "Email": row["Email"]})
                except Exception as e:
                    print(f"❌ Failed for {row['Email']}: {e}")

            if tokens:
                existing_df = pd.read_excel(self.token_output_file) if self.token_output_file.exists() else pd.DataFrame(columns=["Token", "Limit", "Email"])
                updated_df = pd.concat([existing_df, pd.DataFrame(tokens)], ignore_index=True).drop_duplicates(subset="Token")
                updated_df.to_excel(self.token_output_file, index=False)
                messagebox.showinfo("✅ Done", f"Fetched and saved {len(tokens)} token(s).")
            else:
                messagebox.showwarning("No Tokens", "No valid tokens were generated.")

        except Exception as e:
            messagebox.showerror("Error", f"Failed to read or process Excel: {e}")

    def get_next_token(self):
        for token_obj in self.tokens:
            if token_obj["used"] < token_obj["limit"]:
                return token_obj
        return None

    def ensure_files_exist(self):
        if not self.recipients_file.exists():
            df = pd.DataFrame(columns=["Name", "Email"])
            df.to_excel(self.recipients_file, index=False)
            messagebox.showinfo("Created", "recipients.xlsx created. Please add recipient data.")

        if not self.sendplan_file.exists():
            df = pd.DataFrame(columns=["Subject", "Body", "AttachmentType", "AttachmentPath", "HTMLTemplate", "TFNA"])
            df.to_excel(self.sendplan_file, sheet_name="SendPlan", index=False)
            messagebox.showinfo("Created", "sendplan.xlsx created. Please add your send plans.")

        self.files_dir.mkdir(exist_ok=True)
        for i in range(1, 3):
            pdf_path = self.files_dir / f"sample{i}.pdf"
            if not pdf_path.exists():
                with open(pdf_path, "wb") as f:
                    f.write(b"%PDF-1.4\n% Dummy PDF content\n%%EOF")

        if not self.token_input_file.exists():
            df = pd.DataFrame(columns=["Email", "Password", "ClientID", "ClientSecret"])
            df.to_excel(self.token_input_file, index=False)
            messagebox.showinfo("Created", f"{self.token_input_file.name} created. Please fill in your SignNow credentials.")            

    def create_widgets(self):
        top_frame = tk.Frame(self.root, bg="white", pady=10)
        top_frame.pack(fill="x")

        tk.Button(top_frame, text="Manage Recipients", command=self.open_text_manage_window_recipients).pack(side="left", padx=10)
        tk.Button(top_frame, text="Manage SendPlan", command=lambda: edit_sendplan_gui(self.root, ".", None)).pack(side="left", padx=10)
        tk.Button(top_frame, text="Reload Tokens", command=self.load_tokens_from_file).pack(side="left", padx=10)
        tk.Button(top_frame, text="Fetch Tokens", command=self.fetch_tokens_from_excel).pack(side="right", padx=10)

        #self.log_box = scrolledtext.ScrolledText(self.root, height=15, state='disabled')
        #self.log_box.pack(fill="both", expand=True, padx=10, pady=5)
        self.log_text_widget = scrolledtext.ScrolledText(self.root, height=12, width=100, bg="black", fg="white", state="disabled")
        self.log_text_widget.pack(padx=10, pady=10, fill="both", expand=True)
    

        self.start_btn = tk.Button(self.root, text="Start Automation", command=self.start_automation)
        self.start_btn.pack(pady=10)
        

    def log(self, message):
        self.log_text_widget.config(state='normal')
        self.log_text_widget.insert("end", message + "\n")
        self.log_text_widget.see("end")
        self.log_text_widget.config(state='disabled')
        self.root.update_idletasks()
        

    def start_automation(self):
        try:
            recipients_df = pd.read_excel(self.recipients_file)
            sendplan_df = pd.read_excel(self.sendplan_file)
        except Exception as e:
            self.log(f"❌ Failed to read Excel files: {e}")
            return

        if recipients_df.empty:
            self.log("❌ Recipients list is empty.")
            return
        if sendplan_df.empty:
            self.log("❌ SendPlan is empty.")
            return

        plan = sendplan_df.iloc[0]
        subject = plan['Subject']
        message = plan['Body']
        pdf_path = plan['AttachmentPath']

        if not os.path.isfile(pdf_path):
            self.log(f"❌ File not found: {pdf_path}")
            return
        if os.path.getsize(pdf_path) == 0:
            self.log(f"❌ File is empty: {pdf_path}")
            return

        sender_email = recipients_df.iloc[0]['Email'] if 'Email' in recipients_df.columns else ""

        for i, row in recipients_df.iterrows():
            token_obj = self.get_next_token()
            if not token_obj:
                self.log("🚫 All tokens exhausted. Stopping automation.")
                return

            token = token_obj["token"]
            token_obj["used"] += 1

            # Upload document per recipient
            try:
                #self.log(f"📤 Uploading file for: {row['Email']} using token: {token[:6]}...")
                with open(pdf_path, 'rb') as f:
                    files = {'file': ('document.pdf', f, 'application/pdf')}
                    headers = {'Authorization': f'Bearer {token}'}
                    response = requests.post("https://api.signnow.com/document", headers=headers, files=files)
                    response.raise_for_status()
                    doc_id = response.json().get("id")
                    self.log(f"✅ Upload successful. Doc ID: {doc_id}")
            except Exception as e:
                self.log(f"❌ Upload failed for token {token[:6]}: {e}")
                continue

            # Add signature field
            try:
                headers = {
                    'Authorization': f'Bearer {token}',
                    'Content-Type': 'application/json'
                }
                role = f"Signer {i+1}"
                fields_payload = {
                    "roles": [{"name": role, "signing_order": 1}],
                    "fields": [{
                        "x": 100,
                        "y": 150,
                        "width": 200,
                        "height": 40,
                        "page_number": 0,
                        "role": role,
                        "required": True,
                        "type": "signature"
                    }]
                }
                resp = requests.put(f"https://api.signnow.com/document/{doc_id}", json=fields_payload, headers=headers)
                #self.log(f"🖋️ Field response: {resp.status_code} - {resp.text}")
                resp.raise_for_status()
            except Exception as e:
                self.log(f"❌ Failed to set fields: {e}")
                continue

            # Send invite
            try:
                invite_data = {
                    "to": [{
                        "email": row["Email"],
                        "role": role,
                        "order": 1
                    
                    }],
                    "from": sender_email,
                    "subject": subject,
                    "message": message
                }
                invite_url = f"https://api.signnow.com/document/{doc_id}/invite"
                invite_response = requests.post(invite_url, json=invite_data, headers=headers)
                #self.log(f"📬 Invite response: {invite_response.status_code} - {invite_response.text}")
                invite_response.raise_for_status()
                self.log(f"✅ Sent to: {row['Email']}")
            except Exception as e:
                self.log(f"❌ Failed to send invite to {row['Email']}: {e}")



    def open_text_manage_window_recipients(self):
        RANDOM_NAMES = ["Alice", "Bob", "Charlie", "Daisy", "Ethan", "Fiona", "George", "Hannah"]

        win = tk.Toplevel(self.root)
        win.title("Manage Recipients (Text Editor)")
        win.geometry("500x500")
        win.configure(bg="lightblue")

        name_mode_var = tk.StringVar(value="use_existing")

        control_frame = tk.Frame(win, bg="lightblue")
        control_frame.pack(pady=5)

        tk.Label(control_frame, text="Name Handling Mode:", bg="lightblue").grid(row=0, column=0, padx=5)
        mode_menu = tk.OptionMenu(control_frame, name_mode_var, "use_existing", "fixed_name", "random_name")
        mode_menu.grid(row=0, column=1, padx=5)

        fixed_name_label = tk.Label(control_frame, text="Fixed Name:", bg="lightblue")
        fixed_name_entry = tk.Entry(control_frame)
        fixed_name_entry.insert(0, "John Doe")

        def update_name_input_visibility(*args):
            if name_mode_var.get() == "fixed_name":
                fixed_name_label.grid(row=1, column=0, pady=5)
                fixed_name_entry.grid(row=1, column=1, pady=5)
            else:
                fixed_name_label.grid_remove()
                fixed_name_entry.grid_remove()

        name_mode_var.trace_add("write", update_name_input_visibility)
        update_name_input_visibility()

        text_area = scrolledtext.ScrolledText(win, height=20)
        text_area.pack(fill="both", expand=True, padx=10, pady=10)

        def load():
            try:
                df = pd.read_excel(self.recipients_file)
                text_area.delete("1.0", tk.END)
                for _, row in df.iterrows():
                    text_area.insert(tk.END, f"{row['Name']} {row['Email']}\n")
                count_label.config(text=f"Total recipients: {len(df)}")
            except Exception as e:
                messagebox.showerror("Error", str(e))

        def save():
            try:
                lines = text_area.get("1.0", tk.END).strip().splitlines()
                data = []
                mode = name_mode_var.get()
                fixed_name = fixed_name_entry.get().strip()

                for line in lines:
                    parts = line.strip().split()
                    if len(parts) >= 2:
                        name = " ".join(parts[:-1])
                        email = parts[-1]
                    elif len(parts) == 1:
                        email = parts[0]
                        name = ""
                    else:
                        continue

                    if mode == "fixed_name":
                        name = fixed_name
                    elif mode == "random_name":
                        name = random.choice(RANDOM_NAMES)

                    data.append((name, email))

                df = pd.DataFrame(data, columns=["Name", "Email"])
                df.to_excel(self.recipients_file, index=False)
                messagebox.showinfo("✅ Success", "Recipients updated!")
                count_label.config(text=f"Total recipients: {len(df)}")
                win.lift()
                win.attributes('-topmost', True)
                win.after(100, lambda: win.attributes('-topmost', False))

            except Exception as e:
                messagebox.showerror("Error", str(e))

        btn_frame = tk.Frame(win, bg="lightblue")
        btn_frame.pack(pady=10)
        count_label = tk.Label(btn_frame, text="Total recipients: 0", bg="lightblue")
        count_label.pack()

        tk.Button(btn_frame, text="💾 Update / Save", command=save).pack(side="left", padx=10)
        tk.Button(btn_frame, text="🔄 Load", command=load).pack(side="left", padx=10)

    def on_close(self):

        if self.auto_refresh_id:
           self.root.after_cancel(self.auto_refresh_id)

        if messagebox.askokcancel("Quit", "Do you really want to stop the automation and exit?"):
            
            self.root.destroy()    

def start_gui():
    #prevent_multiple_gui()
    root = tk.Tk()
    app = SignNowAutomationGUI(root)
    root.protocol("WM_DELETE_WINDOW", app.on_close)
    root.mainloop()


if __name__ == "__main__":
    start_gui()    
