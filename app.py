import sys
import os
import tkinter as tk
from tkinter import messagebox
from datetime import datetime
import openpyxl


def get_app_dir():
    """Return the directory containing the app, whether frozen or running as script."""
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


def save_submission(text):
    app_dir = get_app_dir()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    # --- submissions.txt ---
    txt_path = os.path.join(app_dir, "submissions.txt")
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write(f"{timestamp}\t{text}\n")

    # --- submissions.xlsx ---
    xlsx_path = os.path.join(app_dir, "submissions.xlsx")
    if os.path.exists(xlsx_path):
        wb = openpyxl.load_workbook(xlsx_path)
        ws = wb.active
    else:
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "Submissions"
        ws.append(["Timestamp", "Entry"])

    ws.append([timestamp, text])

    # Write to a temp file first so a crash mid-write can't corrupt the xlsx.
    tmp_path = xlsx_path + ".tmp"
    wb.save(tmp_path)
    os.replace(tmp_path, xlsx_path)


class SubmissionApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Submission App")
        self.resizable(False, False)

        tk.Label(self, text="Enter text:", font=("Helvetica", 12)).pack(
            padx=20, pady=(20, 4)
        )

        self.text_box = tk.Text(self, width=50, height=6, font=("Helvetica", 11))
        self.text_box.pack(padx=20)

        tk.Button(
            self,
            text="Submit",
            font=("Helvetica", 11),
            width=14,
            command=self.on_submit,
        ).pack(pady=(12, 6))

        self.status_label = tk.Label(
            self, text="", font=("Helvetica", 10), fg="green"
        )
        self.status_label.pack(pady=(0, 16))

    def on_submit(self):
        text = self.text_box.get("1.0", tk.END).strip()
        if not text:
            self.status_label.config(text="Nothing to save — text box is empty.", fg="red")
            return
        try:
            save_submission(text)
            self.text_box.delete("1.0", tk.END)
            self.status_label.config(text="Saved!", fg="green")
        except Exception as exc:
            messagebox.showerror("Save error", str(exc))
            self.status_label.config(text="Error saving — see dialog.", fg="red")


if __name__ == "__main__":
    app = SubmissionApp()
    app.mainloop()
