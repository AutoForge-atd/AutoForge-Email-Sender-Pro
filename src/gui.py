import csv
import json
import os
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

import pandas as pd

from email_sender import EmailSender
from utils import ensure_output_folder, build_log_path, safe_get


class EmailSenderProGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("AutoForge – Email Sender Pro")
        self.root.geometry("1060x820")
        self.root.configure(bg="white")

        self.config_data = self.load_config()
        self.contacts_df = None
        self.csv_path = None
        self.attachment_path = None

        self.build_ui()

    def load_config(self):
        default_config = {
            "company_name": "AutoForge",
            "tagline": "Automated Tools Development",
            "default_output_folder": "output",
            "smtp_server": "smtp.gmail.com",
            "smtp_port": 587,
            "sender_email": "",
            "sender_name": "AutoForge",
            "subject": ""
        }

        if not os.path.exists("config.json"):
            return default_config

        try:
            with open("config.json", "r", encoding="utf-8") as file:
                config = json.load(file)

            for key, value in default_config.items():
                if key not in config:
                    config[key] = value

            return config
        except Exception:
            return default_config

    def save_config(self):
        with open("config.json", "w", encoding="utf-8") as file:
            json.dump(self.config_data, file, indent=4)

    def build_ui(self):
        self.canvas = tk.Canvas(
            self.root,
            width=1060,
            height=820,
            bg="#0b1f5c",
            highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True)

        # Header
        self.canvas.create_text(
            530, 15,
            text="Email Sender Pro",
            fill="#ffffff",
            font=("Arial", 20, "bold")
        )

        self.canvas.create_text(
            530, 32,
            text="Send personalized bulk emails in seconds",
            fill="#7dd3fc",
            font=("Arial", 10)
        )

        self.canvas.create_text(
            530, 45,
            text=f"{self.config_data['company_name']} • {self.config_data['tagline']}",
            fill="#cbd5f5",
            font=("Arial", 10)
        )

        
        self.canvas.tag_raise("all")

        # Top button grid
        top_buttons = tk.Frame(self.root, bg="#0b1f5c")
        top_buttons.place(relx=0.5, rely=0.21, anchor="center")

        button_specs = [
            ("Load CSV", self.load_csv),
            ("Preview Contacts", self.preview_contacts),
            ("Attach File", self.select_attachment),
            ("Send Emails", self.send_emails),
            ("Export Log", self.export_log),
            ("Open Output Folder", self.open_output_folder),
            ("Settings", self.open_settings_window),
            ("Clear Activity Log", self.clear_log)
        ]

        for index, (label, command) in enumerate(button_specs):
            row = index // 2
            col = index % 2
            button = tk.Button(
                top_buttons,
                text=label,
                command=command,
                width=24,
                height=2,
                bg="#2563eb",
                fg="white",
                activebackground="#1d4ed8",
                activeforeground="white",
                relief="flat",
                bd=0,
                font=("Arial", 9, "bold"),
                cursor="hand2"
            )
            button.grid(row=row, column=col, padx=10, pady=8)

        # Status labels
        self.csv_label = tk.Label(
            self.root,
            text="CSV file: None selected",
            bg="#0b1f5c",
            fg="white",
            font=("Arial", 9),
            wraplength=800,
            justify="center"
        )
        self.csv_label.place(relx=0.5, rely=0.36, anchor="center")

        self.attachment_label = tk.Label(
            self.root,
            text="Attachment: None selected",
            bg="#0b1f5c",
            fg="white",
            font=("Arial", 9),
            wraplength=800,
            justify="center"
        )
        self.attachment_label.place(relx=0.5, rely=0.38, anchor="center")

        # Credentials frame
        credentials_frame = tk.Frame(
            self.root,
            bg="#102869",
            bd=0,
            highlightthickness=1,
            highlightbackground="#2a4db8"
        )
        credentials_frame.place(relx=0.5, rely=0.47, anchor="center")

        tk.Label(
            credentials_frame,
            text="Sender Email",
            bg="#102869",
            fg="white",
            font=("Arial", 9, "bold")
        ).grid(row=0, column=0, padx=12, pady=8, sticky="w")

        tk.Label(
            credentials_frame,
            text="App Password",
            bg="#102869",
            fg="white",
            font=("Arial", 10, "bold")
        ).grid(row=1, column=0, padx=12, pady=8, sticky="w")

        tk.Label(
            credentials_frame,
            text="Subject",
            bg="#102869",
            fg="white",
            font=("Arial", 10, "bold")
        ).grid(row=2, column=0, padx=12, pady=8, sticky="w")

        self.sender_email_entry = tk.Entry(credentials_frame, width=42, font=("Arial", 9))
        self.app_password_entry = tk.Entry(credentials_frame, width=42, show="*", font=("Arial", 9))
        self.subject_entry = tk.Entry(credentials_frame, width=42, font=("Arial", 9))

        self.sender_email_entry.grid(row=0, column=1, padx=12, pady=8)
        self.app_password_entry.grid(row=1, column=1, padx=12, pady=8)
        self.subject_entry.grid(row=2, column=1, padx=12, pady=8)

        self.sender_email_entry.insert(0, self.config_data.get("sender_email", ""))
        self.subject_entry.insert(0, self.config_data.get("subject", ""))

        # Message editor frame
        editor_frame = tk.Frame(
            self.root,
            bg="#102869",
            bd=0,
            highlightthickness=1,
            highlightbackground="#2a4db8"
        )
        editor_frame.place(relx=0.5, rely=0.67, anchor="center")

        tk.Label(
            editor_frame,
            text="Message Body (use {name}, {email}, {company})",
            bg="#102869",
            fg="white",
            font=("Arial", 9, "bold")
        ).pack(pady=(10, 6))

        self.message_box = tk.Text(
            editor_frame,
            width=92,
            height=7,
            bg="white",
            fg="black",
            relief="flat",
            bd=0,
            font=("Arial", 9)
        )
        self.message_box.pack(padx=12, pady=(0, 12))
        self.message_box.insert(
            "1.0",
            "Hi {name},\n\n"
            "I’m reaching out from AutoForge.\n\n"
            "This is a personalized email for {company}.\n\n"
            "Best regards,\n"
            "AutoForge"
        )

        # Activity log
        log_title = tk.Label(
            self.root,
            text="Activity Log",
            bg="#0b1f5c",
            fg="white",
            font=("Arial", 9, "bold")
        )
        log_title.place(relx=0.5, rely=0.79, anchor="center")

        self.log_box = tk.Text(
            self.root,
            width=112,
            height=8,
            bg="white",
            fg="black",
            relief="flat",
            bd=0,
            font=("Consolas", 9)
        )
        self.log_box.place(relx=0.5, rely=0.88, anchor="center")

        self.latest_log_rows = []

    def log_message(self, message):
        from datetime import datetime
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.log_box.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_box.see(tk.END)

    def clear_log(self):
        self.log_box.delete("1.0", tk.END)

    def shorten_path(self, path_text, max_length=85):
        if not path_text:
            return ""
        if len(path_text) <= max_length:
            return path_text
        return "..." + path_text[-(max_length - 3):]

    def load_csv(self):
        path = filedialog.askopenfilename(
            title="Select Contact CSV",
            filetypes=[("CSV files", "*.csv")]
        )

        if not path:
            return

        try:
            df = pd.read_csv(path, dtype=str).fillna("")
            required_columns = {"email", "name", "company"}
            actual_columns = {col.strip().lower() for col in df.columns}

            if not required_columns.issubset(actual_columns):
                raise ValueError("CSV must contain columns: email, name, company")

            normalized_columns = {col: col.strip().lower() for col in df.columns}
            df = df.rename(columns=normalized_columns)

            self.contacts_df = df
            self.csv_path = path
            self.csv_label.config(text=f"CSV file: {self.shorten_path(path)}")
            self.log_message(f"Loaded CSV with {len(df)} contact(s).")
        except Exception as exc:
            messagebox.showerror("CSV Error", str(exc))

    def preview_contacts(self):
        if self.contacts_df is None:
            messagebox.showwarning("No CSV", "Please load a CSV file first.")
            return

        preview_window = tk.Toplevel(self.root)
        preview_window.title("Contact Preview")
        preview_window.geometry("760x420")
        preview_window.configure(bg="white")

        tree = ttk.Treeview(
            preview_window,
            columns=("email", "name", "company"),
            show="headings",
            height=15
        )

        tree.heading("email", text="Email")
        tree.heading("name", text="Name")
        tree.heading("company", text="Company")

        tree.column("email", width=280)
        tree.column("name", width=180)
        tree.column("company", width=220)

        for _, row in self.contacts_df.head(50).iterrows():
            tree.insert(
                "",
                tk.END,
                values=(
                    safe_get(row.get("email")),
                    safe_get(row.get("name")),
                    safe_get(row.get("company"))
                )
            )

        tree.pack(fill="both", expand=True, padx=12, pady=12)

    def select_attachment(self):
        path = filedialog.askopenfilename(
            title="Select Optional Attachment",
            filetypes=[("All files", "*.*")]
        )

        if path:
            self.attachment_path = path
            self.attachment_label.config(text=f"Attachment: {self.shorten_path(path)}")
            self.log_message(f"Attachment selected: {path}")

    def render_message(self, template, row):
        rendered = template
        replacements = {
            "{name}": safe_get(row.get("name")),
            "{email}": safe_get(row.get("email")),
            "{company}": safe_get(row.get("company"))
        }

        for placeholder, value in replacements.items():
            rendered = rendered.replace(placeholder, value)

        return rendered

    def validate_before_send(self):
        if self.contacts_df is None:
            raise ValueError("Please load a CSV file first.")

        sender_email = self.sender_email_entry.get().strip()
        app_password = self.app_password_entry.get().strip()
        subject = self.subject_entry.get().strip()
        message_template = self.message_box.get("1.0", tk.END).strip()

        if not sender_email:
            raise ValueError("Please enter your sender email.")
        if not app_password:
            raise ValueError("Please enter your app password.")
        if not subject:
            raise ValueError("Please enter an email subject.")
        if not message_template:
            raise ValueError("Please enter a message body.")

        return sender_email, app_password, subject, message_template

    def send_emails(self):
        try:
            sender_email, app_password, subject, message_template = self.validate_before_send()

            self.config_data["sender_email"] = sender_email
            self.config_data["subject"] = subject
            self.save_config()

            sender = EmailSender(
                smtp_server=self.config_data.get("smtp_server", "smtp.gmail.com"),
                smtp_port=self.config_data.get("smtp_port", 587),
                sender_email=sender_email,
                app_password=app_password,
                sender_name=self.config_data.get("sender_name", "AutoForge")
            )

            self.latest_log_rows = []
            success_count = 0
            fail_count = 0

            for _, row in self.contacts_df.iterrows():
                recipient_email = safe_get(row.get("email"))
                recipient_name = safe_get(row.get("name"))
                recipient_company = safe_get(row.get("company"))

                personalized_body = self.render_message(message_template, row)

                try:
                    msg = sender.create_message(
                        recipient_email=recipient_email,
                        subject=subject,
                        body=personalized_body,
                        attachment_path=self.attachment_path
                    )
                    sender.send_message(msg)

                    self.log_message(f"Sent: {recipient_email}")
                    self.latest_log_rows.append({
                        "email": recipient_email,
                        "name": recipient_name,
                        "company": recipient_company,
                        "status": "Sent",
                        "error": ""
                    })
                    success_count += 1
                except Exception as exc:
                    self.log_message(f"Failed: {recipient_email} | {exc}")
                    self.latest_log_rows.append({
                        "email": recipient_email,
                        "name": recipient_name,
                        "company": recipient_company,
                        "status": "Failed",
                        "error": str(exc)
                    })
                    fail_count += 1

                self.root.update_idletasks()

            self.write_latest_log_to_file()
            messagebox.showinfo(
                "Sending Complete",
                f"Emails complete.\n\nSent: {success_count}\nFailed: {fail_count}"
            )
        except Exception as exc:
            messagebox.showerror("Send Error", str(exc))

    def write_latest_log_to_file(self):
        output_folder = ensure_output_folder(
            self.config_data.get("default_output_folder", "output")
        )
        log_path = build_log_path(output_folder)

        with open(log_path, "w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(
                file,
                fieldnames=["email", "name", "company", "status", "error"]
            )
            writer.writeheader()
            writer.writerows(self.latest_log_rows)

        self.log_message(f"Delivery log saved: {log_path}")

    def export_log(self):
        if not self.latest_log_rows:
            messagebox.showwarning("No Log Data", "No email log is available yet.")
            return

        save_path = filedialog.asksaveasfilename(
            title="Save Delivery Log",
            defaultextension=".csv",
            initialdir=self.config_data.get("default_output_folder", "output"),
            initialfile="delivery_log.csv",
            filetypes=[("CSV files", "*.csv")]
        )

        if not save_path:
            return

        with open(save_path, "w", newline="", encoding="utf-8") as file:
            writer = csv.DictWriter(
                file,
                fieldnames=["email", "name", "company", "status", "error"]
            )
            writer.writeheader()
            writer.writerows(self.latest_log_rows)

        self.log_message(f"Exported log: {save_path}")
        messagebox.showinfo("Exported", f"Delivery log saved to:\n{save_path}")

    def open_output_folder(self):
        output_folder = os.path.abspath(
            ensure_output_folder(self.config_data.get("default_output_folder", "output"))
        )
        try:
            os.startfile(output_folder)
        except AttributeError:
            messagebox.showinfo("Output Folder", output_folder)

    def open_settings_window(self):
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Settings")
        settings_window.geometry("560x320")
        settings_window.configure(bg="white")
        settings_window.resizable(False, False)

        title_label = tk.Label(
            settings_window,
            text="App Settings",
            font=("Arial", 18, "bold"),
            bg="white",
            fg="#0b1f5c"
        )
        title_label.pack(pady=15)

        form_frame = tk.Frame(settings_window, bg="white")
        form_frame.pack(padx=20, pady=10, fill="both", expand=True)

        tk.Label(form_frame, text="Company Name", bg="white").grid(row=0, column=0, sticky="w", pady=8)
        tk.Label(form_frame, text="Tagline", bg="white").grid(row=1, column=0, sticky="w", pady=8)
        tk.Label(form_frame, text="Sender Name", bg="white").grid(row=2, column=0, sticky="w", pady=8)
        tk.Label(form_frame, text="Output Folder", bg="white").grid(row=3, column=0, sticky="w", pady=8)

        company_name_entry = tk.Entry(form_frame, width=40)
        tagline_entry = tk.Entry(form_frame, width=40)
        sender_name_entry = tk.Entry(form_frame, width=40)
        output_folder_entry = tk.Entry(form_frame, width=40)

        company_name_entry.grid(row=0, column=1, padx=8, pady=8)
        tagline_entry.grid(row=1, column=1, padx=8, pady=8)
        sender_name_entry.grid(row=2, column=1, padx=8, pady=8)
        output_folder_entry.grid(row=3, column=1, padx=8, pady=8)

        company_name_entry.insert(0, self.config_data.get("company_name", ""))
        tagline_entry.insert(0, self.config_data.get("tagline", ""))
        sender_name_entry.insert(0, self.config_data.get("sender_name", "AutoForge"))
        output_folder_entry.insert(0, self.config_data.get("default_output_folder", "output"))

        def save_settings():
            self.config_data["company_name"] = company_name_entry.get().strip()
            self.config_data["tagline"] = tagline_entry.get().strip()
            self.config_data["sender_name"] = sender_name_entry.get().strip() or "AutoForge"
            self.config_data["default_output_folder"] = output_folder_entry.get().strip() or "output"
            self.save_config()
            messagebox.showinfo("Saved", "Settings saved successfully.")
            settings_window.destroy()

        save_button = tk.Button(
            settings_window,
            text="Save Settings",
            command=save_settings,
            width=20,
            height=2,
            bg="#2563eb",
            fg="white",
            activebackground="#1d4ed8",
            activeforeground="white",
            relief="flat",
            bd=0,
            font=("Arial", 11, "bold"),
            cursor="hand2"
        )
        save_button.pack(pady=18)