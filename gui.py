from __future__ import annotations

import hashlib
import json
import logging
import os
import subprocess
import sys
import threading
import tkinter as tk
import uuid
from pathlib import Path
from tkinter import filedialog, messagebox, ttk

from config import APP_DATA_DIR, APP_NAME, DEFAULT_CONFIG, DEFAULT_TEMPLATE, load_json_config
from excel_reader import list_sheets, load_source_file
from generator import generate_workbook
from mapping import detect_columns, normalize_data
from models import Record, SourceTable


LOGGER = logging.getLogger(__name__)
REQUIRED_FIELDS = ("asp", "module", "module_type", "channel", "datapoint", "description", "source")
FIELD_LABELS = {
    "asp": "ASP", "module": "Modul", "module_type": "Modultyp", "channel": "Kanal",
    "datapoint": "Datenpunkt", "description": "Beschreibung", "source": "Quelle",
}


class MappingDialog(tk.Toplevel):
    def __init__(self, parent: tk.Misc, headers: list[str], detected: dict[str, str]) -> None:
        super().__init__(parent)
        self.title("Spaltenzuordnung")
        self.resizable(False, False)
        self.result: dict[str, str] | None = None
        self.save_profile = tk.BooleanVar(value=False)
        self.variables: dict[str, tk.StringVar] = {}
        ttk.Label(self, text="Ordnen Sie die Rohdaten-Spalten den internen Feldern zu.").grid(row=0, column=0, columnspan=2, padx=12, pady=12)
        choices = ["(nicht verwenden)", *headers]
        for row, field in enumerate(REQUIRED_FIELDS, start=1):
            ttk.Label(self, text=FIELD_LABELS[field]).grid(row=row, column=0, sticky="w", padx=12, pady=3)
            variable = tk.StringVar(value=detected.get(field, choices[0]))
            self.variables[field] = variable
            ttk.Combobox(self, textvariable=variable, values=choices, state="readonly", width=35).grid(row=row, column=1, padx=12, pady=3)
        ttk.Checkbutton(self, text="Für diese Spaltenstruktur speichern", variable=self.save_profile).grid(
            row=len(REQUIRED_FIELDS) + 1, column=0, columnspan=2, sticky="w", padx=12, pady=6
        )
        buttons = ttk.Frame(self)
        buttons.grid(row=len(REQUIRED_FIELDS) + 2, column=0, columnspan=2, pady=12)
        ttk.Button(buttons, text="Übernehmen", command=self._accept).pack(side="left", padx=5)
        ttk.Button(buttons, text="Abbrechen", command=self.destroy).pack(side="left", padx=5)
        self.transient(parent)
        self.grab_set()

    def _accept(self) -> None:
        result = {key: value.get() for key, value in self.variables.items() if value.get() != "(nicht verwenden)"}
        if "channel" not in result or "datapoint" not in result:
            messagebox.showerror("Fehler", "Kanal und Datenpunkt müssen zugeordnet sein.", parent=self)
            return
        self.result = result
        self.destroy()


class SmartStruxureApp(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=12)
        self.master = master
        self.source_var = tk.StringVar()
        self.template_var = tk.StringVar(value=str(DEFAULT_TEMPLATE) if DEFAULT_TEMPLATE.exists() else "")
        self.output_var = tk.StringVar()
        self.sheet_var = tk.StringVar()
        self.status_var = tk.StringVar(value="Bitte Rohdatei und Vorlage auswählen.")
        self.progress_var = tk.IntVar()
        self.source_table: SourceTable | None = None
        self.column_mapping: dict[str, str] = {}
        self.profiles_path = APP_DATA_DIR / "mapping_profiles.json"
        self._build()

    def _build(self) -> None:
        self.pack(fill="both", expand=True)
        self.master.title(APP_NAME)
        self.master.geometry("980x650")
        self.master.minsize(800, 520)
        for row, (label, variable, command, button) in enumerate([
            ("Rohdatei", self.source_var, self.choose_source, "Rohdatei auswählen"),
            ("Master-Vorlage", self.template_var, self.choose_template, "Vorlage auswählen"),
            ("Ausgabedatei", self.output_var, self.choose_output, "Speicherort auswählen"),
        ]):
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=4)
            ttk.Entry(self, textvariable=variable).grid(row=row, column=1, sticky="ew", padx=8)
            ttk.Button(self, text=button, command=command).grid(row=row, column=2, sticky="ew")
        ttk.Label(self, text="Tabellenblatt").grid(row=3, column=0, sticky="w", pady=4)
        self.sheet_box = ttk.Combobox(self, textvariable=self.sheet_var, state="readonly")
        self.sheet_box.grid(row=3, column=1, sticky="ew", padx=8)
        self.sheet_box.bind("<<ComboboxSelected>>", lambda _: self.load_source())
        ttk.Button(self, text="Zuordnung (nur Fremdformat)", command=self.configure_mapping).grid(row=3, column=2, sticky="ew")
        self.preview = ttk.Treeview(self, show="headings", height=16)
        self.preview.grid(row=4, column=0, columnspan=3, sticky="nsew", pady=(12, 6))
        preview_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.preview.xview)
        preview_scroll.grid(row=5, column=0, columnspan=3, sticky="ew")
        self.preview.configure(xscrollcommand=preview_scroll.set)
        self.generate_button = ttk.Button(self, text="Beschriftung generieren", command=self.start_generation)
        self.generate_button.grid(row=6, column=0, columnspan=3, sticky="ew", pady=12, ipady=8)
        ttk.Progressbar(self, variable=self.progress_var, maximum=100).grid(row=7, column=0, columnspan=3, sticky="ew")
        ttk.Label(self, textvariable=self.status_var, wraplength=900).grid(row=8, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Button(self, text="Ordner öffnen", command=self.open_output_folder).grid(row=8, column=2, sticky="e")
        self.columnconfigure(1, weight=1)
        self.rowconfigure(4, weight=1)

    def choose_source(self) -> None:
        value = filedialog.askopenfilename(filetypes=[("Rohdaten", "*.xlsx *.xlsm *.csv"), ("Alle Dateien", "*.*")])
        if value:
            self.source_var.set(value)
            default_output = Path(value).with_name(f"{Path(value).stem}_Beschriftung.xlsx")
            self.output_var.set(str(default_output))
            try:
                sheets = list_sheets(Path(value))
                self.sheet_box["values"] = sheets
                preferred = next((name for name in sheets if "_DP" in name.upper()), sheets[0] if sheets else "")
                self.sheet_var.set(preferred)
                self.load_source()
            except Exception as exc:
                self._error(exc)

    def choose_template(self) -> None:
        value = filedialog.askopenfilename(filetypes=[("Excel-Vorlage", "*.xlsx *.xlsm")])
        if value:
            self.template_var.set(value)

    def choose_output(self) -> None:
        value = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel-Datei", "*.xlsx")])
        if value:
            self.output_var.set(value)

    def load_source(self) -> None:
        self.source_table = load_source_file(Path(self.source_var.get()), self.sheet_var.get() or None)
        self.column_mapping = self._load_mapping_profile(self.source_table.headers) or detect_columns(self.source_table.headers)
        self._show_preview(self.source_table)
        missing = [FIELD_LABELS[field] for field in REQUIRED_FIELDS if field not in self.column_mapping]
        suffix = f" Nicht erkannt: {', '.join(missing)}." if missing else " Spalten automatisch erkannt."
        self.status_var.set(f"Datei geladen. {len(self.source_table.rows)} Datensätze erkannt.{suffix}")

    def _show_preview(self, table: SourceTable) -> None:
        self.preview.delete(*self.preview.get_children())
        columns = table.headers[:12]
        self.preview["columns"] = columns
        for column in columns:
            self.preview.heading(column, text=column)
            self.preview.column(column, width=135, stretch=True)
        for row in table.rows[:20]:
            self.preview.insert("", "end", values=[row.get(column, "") for column in columns])

    def configure_mapping(self) -> None:
        if not self.source_table:
            messagebox.showinfo("Hinweis", "Bitte zuerst eine Rohdatei auswählen.")
            return
        dialog = MappingDialog(self, self.source_table.headers, self.column_mapping)
        self.wait_window(dialog)
        if dialog.result is not None:
            self.column_mapping = dialog.result
            if dialog.save_profile.get():
                self._save_mapping_profile(self.source_table.headers, dialog.result)
            self.status_var.set("Spaltenzuordnung übernommen.")

    @staticmethod
    def _profile_key(headers: list[str]) -> str:
        return hashlib.sha256("\0".join(headers).encode("utf-8")).hexdigest()

    def _load_mapping_profile(self, headers: list[str]) -> dict[str, str] | None:
        if not self.profiles_path.exists():
            return None
        try:
            profiles = json.loads(self.profiles_path.read_text(encoding="utf-8"))
            profile = profiles.get(self._profile_key(headers))
            if isinstance(profile, dict) and all(value in headers for value in profile.values()):
                return {str(key): str(value) for key, value in profile.items()}
        except (OSError, ValueError, TypeError):
            LOGGER.warning("Mapping-Profile konnten nicht gelesen werden.", exc_info=True)
        return None

    def _save_mapping_profile(self, headers: list[str], mapping: dict[str, str]) -> None:
        try:
            profiles = json.loads(self.profiles_path.read_text(encoding="utf-8")) if self.profiles_path.exists() else {}
            profiles[self._profile_key(headers)] = mapping
            self.profiles_path.write_text(json.dumps(profiles, ensure_ascii=False, indent=2), encoding="utf-8")
        except (OSError, ValueError, TypeError):
            LOGGER.warning("Mapping-Profil konnte nicht gespeichert werden.", exc_info=True)

    def start_generation(self) -> None:
        if not self.source_table:
            messagebox.showerror("Fehler", "Bitte zuerst eine Rohdatei laden.")
            return
        missing = [field for field in ("channel", "datapoint") if field not in self.column_mapping]
        if missing:
            self.configure_mapping()
            if any(field not in self.column_mapping for field in missing):
                return
        master = Path(self.template_var.get())
        output = Path(self.output_var.get())
        if not self.template_var.get().strip():
            messagebox.showerror("Fehler", "Bitte eine Master-Vorlage auswählen.")
            return
        if not self.output_var.get().strip():
            messagebox.showerror("Fehler", "Bitte einen Speicherort auswählen.")
            return
        replace_existing = False
        try:
            if master.resolve() == output.resolve():
                messagebox.showerror("Fehler", "Die Master-Vorlage darf nicht überschrieben werden.")
                return
            if output.exists():
                overwrite = messagebox.askyesno(
                    "Datei bereits vorhanden",
                    f"Die Ausgabedatei existiert bereits:\n\n{output}\n\nSoll sie ersetzt werden?",
                )
                if not overwrite:
                    self.status_var.set("Generierung abgebrochen. Bitte einen anderen Speicherort auswählen.")
                    return
                replace_existing = True
        except OSError as exc:
            self._error(exc)
            return
        records = normalize_data(self.source_table.rows, self.column_mapping)
        self.generate_button.configure(state="disabled")
        self.progress_var.set(0)
        threading.Thread(
            target=self._generate,
            args=(master, output, records, replace_existing),
            daemon=True,
        ).start()

    def _generate(self, master: Path, output: Path, records: list[Record], replace_existing: bool) -> None:
        generated_path = output
        if replace_existing:
            generated_path = output.with_name(f".{output.stem}.{uuid.uuid4().hex}.tmp.xlsx")
        try:
            config = load_json_config(DEFAULT_CONFIG)
            result = generate_workbook(master, generated_path, records, config, self._progress)
            if replace_existing:
                os.replace(generated_path, output)
                result.output_path = output.resolve()
            details = f"Fertig. {result.boxes_created} Kästchen, {result.records_processed} Datenpunkte.\nAusgabe: {result.output_path}"
            if result.warnings:
                details += "\n" + "\n".join(result.warnings)
            self.master.after(0, lambda: self._finish(details))
        except Exception as exc:
            LOGGER.exception("Generierung fehlgeschlagen")
            if generated_path != output:
                try:
                    generated_path.unlink(missing_ok=True)
                except OSError:
                    LOGGER.warning("Temporäre Ausgabedatei konnte nicht entfernt werden: %s", generated_path)
            # Python clears the exception variable after leaving an except
            # block. Bind it as a default argument for the delayed callback.
            self.master.after(0, lambda error=exc: self._finish_error(error))

    def _progress(self, value: int, text: str) -> None:
        self.master.after(0, lambda: (self.progress_var.set(value), self.status_var.set(text)))

    def _finish(self, text: str) -> None:
        self.generate_button.configure(state="normal")
        self.progress_var.set(100)
        self.status_var.set(text)
        messagebox.showinfo("Fertig", text)

    def _finish_error(self, exc: Exception) -> None:
        self.generate_button.configure(state="normal")
        self._error(exc)

    def _error(self, exc: Exception) -> None:
        self.status_var.set(f"Fehler: {exc}")
        messagebox.showerror("Fehler", str(exc))

    def open_output_folder(self) -> None:
        path = Path(self.output_var.get()).parent
        if not path.exists():
            return
        if sys.platform == "win32":
            os.startfile(path)  # type: ignore[attr-defined]
        elif sys.platform == "darwin":
            subprocess.Popen(["open", str(path)])
        else:
            subprocess.Popen(["xdg-open", str(path)])


def run_gui() -> None:
    root = tk.Tk()
    SmartStruxureApp(root)
    root.mainloop()
