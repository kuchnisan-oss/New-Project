import csv
import json
import re
import threading
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from tkinter import Tk, ttk, StringVar, messagebox

import msal
import requests

GRAPH_SCOPE = ["Mail.Read"]
GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0"
MAX_EMAILS_PER_RUN = 500
PHONE_REGEX = re.compile(
    r"(?:(?:\+?\d{1,3}[\s.-]?)?(?:\(\d{2,4}\)[\s.-]?)?\d{3,4}[\s.-]?\d{4})"
)


@dataclass
class AppConfig:
    tenant_id: str
    client_id: str
    mailbox: str

    @property
    def authority(self) -> str:
        return f"https://login.microsoftonline.com/{self.tenant_id.strip()}"


class InboxExtractorApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.root.title("Outlook Inbox Contact Extractor")
        self.tenant_id_var = StringVar()
        self.client_id_var = StringVar()
        self.mailbox_var = StringVar()
        self.status_var = StringVar(value="Ready.")

        self._build_ui()

    def _build_ui(self) -> None:
        main_frame = ttk.Frame(self.root, padding=16)
        main_frame.grid(row=0, column=0, sticky="nsew")

        ttk.Label(main_frame, text="Tenant ID").grid(row=0, column=0, sticky="w")
        ttk.Entry(main_frame, textvariable=self.tenant_id_var, width=48).grid(
            row=0, column=1, sticky="ew", pady=4
        )

        ttk.Label(main_frame, text="Client ID").grid(row=1, column=0, sticky="w")
        ttk.Entry(main_frame, textvariable=self.client_id_var, width=48).grid(
            row=1, column=1, sticky="ew", pady=4
        )

        ttk.Label(main_frame, text="Mailbox Email").grid(row=2, column=0, sticky="w")
        ttk.Entry(main_frame, textvariable=self.mailbox_var, width=48).grid(
            row=2, column=1, sticky="ew", pady=4
        )

        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, pady=(12, 4))

        ttk.Button(button_frame, text="Start", command=self.start).grid(
            row=0, column=0, padx=4
        )
        ttk.Button(button_frame, text="Reset Cursor", command=self.reset_cursor).grid(
            row=0, column=1, padx=4
        )

        ttk.Label(main_frame, textvariable=self.status_var, wraplength=420).grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(8, 0)
        )

        self.root.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

    def start(self) -> None:
        config = self._get_config()
        if not config:
            return
        self.status_var.set("Authenticating with Microsoft Graph...")
        threading.Thread(target=self._run_extraction, args=(config,), daemon=True).start()

    def reset_cursor(self) -> None:
        cursor_path = self._cursor_path()
        if cursor_path.exists():
            cursor_path.unlink()
            self.status_var.set("Cursor reset. Next run will start from newest email.")
        else:
            self.status_var.set("No cursor to reset.")

    def _get_config(self) -> AppConfig | None:
        tenant_id = self.tenant_id_var.get().strip()
        client_id = self.client_id_var.get().strip()
        mailbox = self.mailbox_var.get().strip()

        if not tenant_id or not client_id or not mailbox:
            messagebox.showerror("Missing Info", "Please fill in all fields.")
            return None
        return AppConfig(tenant_id=tenant_id, client_id=client_id, mailbox=mailbox)

    def _run_extraction(self, config: AppConfig) -> None:
        try:
            token = self._authenticate(config)
            if not token:
                self._set_status("Authentication failed. Please try again.")
                return

            self._set_status("Fetching emails and extracting contacts...")
            emails = self._fetch_emails(token, config.mailbox)
            if emails is None:
                return

            contacts = self._extract_contacts(emails)
            if not contacts:
                self._set_status("No new contacts found in this batch.")
                return

            saved_count = self._save_contacts(contacts)
            self._set_status(
                f"Saved {saved_count} new contacts. Run again to continue."
            )
        except Exception as exc:  # noqa: BLE001
            self._set_status(f"Error: {exc}")

    def _authenticate(self, config: AppConfig) -> str | None:
        app = msal.PublicClientApplication(
            client_id=config.client_id, authority=config.authority
        )
        flow = app.initiate_device_flow(scopes=GRAPH_SCOPE)
        if "user_code" not in flow:
            self._set_status("Unable to start device code flow.")
            return None

        messagebox.showinfo(
            "Device Login",
            f"Use the following code to sign in: {flow['user_code']}\n\n"
            f"Go to {flow['verification_uri']} to authenticate.",
        )
        result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            self._set_status(result.get("error_description", "Login failed."))
            return None
        return result["access_token"]

    def _fetch_emails(self, token: str, mailbox: str) -> list[dict] | None:
        cursor = self._load_cursor()
        if cursor and cursor.get("next_link"):
            url = cursor["next_link"]
        else:
            url = (
                f"{GRAPH_ENDPOINT}/users/{mailbox}/messages"
                f"?$top={MAX_EMAILS_PER_RUN}"
                "&$orderby=receivedDateTime desc"
                "&$select=sender,body"
            )

        headers = {"Authorization": f"Bearer {token}"}
        response = requests.get(url, headers=headers, timeout=30)
        if response.status_code != 200:
            self._set_status(
                f"Graph API error {response.status_code}: {response.text}"
            )
            return None

        data = response.json()
        next_link = data.get("@odata.nextLink")
        self._save_cursor(next_link)
        return data.get("value", [])

    def _extract_contacts(self, messages: list[dict]) -> list[dict]:
        contacts = []
        for message in messages:
            sender = message.get("sender", {}).get("emailAddress", {})
            email = sender.get("address")
            name = sender.get("name")
            if not email:
                continue
            domain = email.split("@")[-1] if "@" in email else ""
            company = domain.split(".")[0] if domain else ""
            body_content = message.get("body", {}).get("content", "")
            phone = self._extract_phone(body_content)
            contacts.append(
                {
                    "name": name or "",
                    "email": email,
                    "company": company,
                    "phone": phone or "",
                    "extracted_at": datetime.utcnow().isoformat(),
                }
            )
        return contacts

    def _extract_phone(self, text: str) -> str | None:
        match = PHONE_REGEX.search(text)
        if match:
            return match.group(0).strip()
        return None

    def _save_contacts(self, contacts: list[dict]) -> int:
        csv_path = self._csv_path()
        existing_emails = self._load_existing_emails(csv_path)
        new_contacts = [c for c in contacts if c["email"] not in existing_emails]

        if not new_contacts:
            return 0

        csv_path.parent.mkdir(parents=True, exist_ok=True)
        file_exists = csv_path.exists()
        with csv_path.open("a", newline="", encoding="utf-8") as handle:
            writer = csv.DictWriter(
                handle, fieldnames=["name", "email", "company", "phone", "extracted_at"]
            )
            if not file_exists:
                writer.writeheader()
            writer.writerows(new_contacts)
        return len(new_contacts)

    def _load_existing_emails(self, csv_path: Path) -> set[str]:
        if not csv_path.exists():
            return set()
        with csv_path.open("r", newline="", encoding="utf-8") as handle:
            reader = csv.DictReader(handle)
            return {row.get("email", "").strip() for row in reader if row.get("email")}

    def _cursor_path(self) -> Path:
        return Path.home() / "Desktop" / "outlook_contact_cursor.json"

    def _csv_path(self) -> Path:
        return Path.home() / "Desktop" / "outlook_contacts.csv"

    def _load_cursor(self) -> dict | None:
        cursor_path = self._cursor_path()
        if not cursor_path.exists():
            return None
        with cursor_path.open("r", encoding="utf-8") as handle:
            return json.load(handle)

    def _save_cursor(self, next_link: str | None) -> None:
        cursor_path = self._cursor_path()
        cursor_path.parent.mkdir(parents=True, exist_ok=True)
        payload = {"next_link": next_link}
        with cursor_path.open("w", encoding="utf-8") as handle:
            json.dump(payload, handle, indent=2)

    def _set_status(self, message: str) -> None:
        self.status_var.set(message)


if __name__ == "__main__":
    root = Tk()
    app = InboxExtractorApp(root)
    root.mainloop()
