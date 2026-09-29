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

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD
except ImportError:  # GUI bleibt auch ohne optionale Drag-and-drop-Bibliothek startbar.
    DND_FILES = None
    TkinterDnD = None

from config import APP_DATA_DIR, APP_NAME, DEFAULT_CONFIG, DEFAULT_TEMPLATE, load_json_config
from excel_reader import list_sheets, load_source_file
from generator import generate_multi_sheet_workbook
from mapping import detect_columns, normalize_data
from models import Record, SourceTable


LOGGER = logging.getLogger(__name__)
REQUIRED_FIELDS = ("asp", "module", "module_type", "channel", "datapoint", "description", "source")
FIELD_LABELS = {
    "asp": "ASP", "module": "Modul", "module_type": "Modultyp", "channel": "Kanal",
    "datapoint": "Datenpunkt", "description": "Beschreibung", "source": "Quelle",
}
HEADER_FIELDS = [
    ("DOKUMENTTITEL", "Dokumenttitel"),
    ("NETZTEIL_MODUL_1", "Netzteil – Modul 1"),
    ("NETZTEIL_TEXT_1", "Netzteil – Text 1"),
    ("NETZTEIL_MODUL_2", "Netzteil – Modul 2"),
    ("NETZTEIL_TEXT_2", "Netzteil – Text 2"),
    ("AS_MODUL", "Automation Server – Modul"),
    ("ANLAGE", "Anlage"),
    ("AUTOMATION_SERVER", "Automation Server"),
    ("IP_ADRESSE", "IP-Adresse"),
    ("SUBNETZMASKE", "Subnetzmaske"),
    ("AS_VERSION", "AS-Version"),
    ("BENUTZER", "Benutzer"),
    ("PASSWORT", "Passwort"),
]


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


class HeaderDataDialog(tk.Toplevel):
    def __init__(
        self,
        parent: tk.Misc,
        initial_values: dict[str, str] | None = None,
        initial_image: Path | None = None,
    ) -> None:
        super().__init__(parent)
        self.title("Kopfdaten und Bild")
        self.resizable(False, False)
        self.result: tuple[dict[str, str], Path | None] | None = None
        self.variables: dict[str, tk.StringVar] = {}
        self.image_var = tk.StringVar(value=str(initial_image or ""))
        ttk.Label(
            self,
            text="Diese Angaben werden für alle ausgewählten Tabellenblätter verwendet.\n"
                 "Leere Felder bleiben als {{PLATZHALTER}} in Excel erhalten.",
            justify="left",
        ).grid(row=0, column=0, columnspan=3, sticky="w", padx=14, pady=(14, 10))
        values = initial_values or {}
        for row, (field, label) in enumerate(HEADER_FIELDS, start=1):
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", padx=(14, 8), pady=3)
            variable = tk.StringVar(value=values.get(field, ""))
            self.variables[field] = variable
            ttk.Entry(self, textvariable=variable, width=52, show="*" if field == "PASSWORT" else "").grid(
                row=row, column=1, columnspan=2, sticky="ew", padx=(0, 14), pady=3
            )
        image_row = len(HEADER_FIELDS) + 1
        ttk.Label(self, text="Bild/Logo").grid(row=image_row, column=0, sticky="w", padx=(14, 8), pady=(10, 3))
        ttk.Entry(self, textvariable=self.image_var, width=42).grid(row=image_row, column=1, sticky="ew", pady=(10, 3))
        ttk.Button(self, text="Bild auswählen", command=self._choose_image).grid(
            row=image_row, column=2, padx=(8, 14), pady=(10, 3)
        )
        buttons = ttk.Frame(self)
        buttons.grid(row=image_row + 1, column=0, columnspan=3, pady=14)
        ttk.Button(buttons, text="Übernehmen", command=self._accept).pack(side="left", padx=5)
        ttk.Button(buttons, text="Mit Platzhaltern fortfahren", command=self._skip).pack(side="left", padx=5)
        ttk.Button(buttons, text="Abbrechen", command=self.destroy).pack(side="left", padx=5)
        self.transient(parent)
        self.grab_set()

    def _choose_image(self) -> None:
        value = filedialog.askopenfilename(
            parent=self,
            filetypes=[("Bilddateien", "*.png *.jpg *.jpeg *.bmp *.gif"), ("Alle Dateien", "*.*")],
        )
        if value:
            self.image_var.set(value)

    def _accept(self) -> None:
        image_value = self.image_var.get().strip()
        image_path = Path(image_value) if image_value else None
        if image_path is not None and not image_path.is_file():
            messagebox.showerror("Fehler", "Die ausgewählte Bilddatei wurde nicht gefunden.", parent=self)
            return
        values = {field: variable.get().strip() for field, variable in self.variables.items() if variable.get().strip()}
        self.result = (values, image_path)
        self.destroy()

    def _skip(self) -> None:
        self.result = ({}, None)
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
        self.header_values: dict[str, str] = {}
        self.header_image: Path | None = None
        self.profiles_path = APP_DATA_DIR / "mapping_profiles.json"
        self._build()

    def _build(self) -> None:
        self.pack(fill="both", expand=True)
        self.master.title(APP_NAME)
        self.master.geometry("980x720")
        self.master.minsize(800, 520)
        for row, (label, variable, command, button) in enumerate([
            ("Rohdatei", self.source_var, self.choose_source, "Rohdatei auswählen"),
            ("Master-Vorlage", self.template_var, self.choose_template, "Vorlage auswählen"),
            ("Ausgabedatei", self.output_var, self.choose_output, "Speicherort auswählen"),
        ]):
            ttk.Label(self, text=label).grid(row=row, column=0, sticky="w", pady=4)
            entry = ttk.Entry(self, textvariable=variable)
            entry.grid(row=row, column=1, sticky="ew", padx=8)
            if row == 0:
                self.source_entry = entry
            ttk.Button(self, text=button, command=command).grid(row=row, column=2, sticky="ew")
        self.drop_label = tk.Label(
            self,
            text="Excel- oder CSV-Rohdatei hierher ziehen",
            relief="ridge",
            borderwidth=2,
            padx=12,
            pady=10,
            bg="#f3f6fa",
            fg="#334155",
        )
        self.drop_label.grid(row=3, column=0, columnspan=3, sticky="ew", pady=(8, 10))
        self._configure_drag_drop()
        ttk.Label(self, text="Tabellenblätter").grid(row=4, column=0, sticky="nw", pady=4)
        sheet_frame = ttk.Frame(self)
        sheet_frame.grid(row=4, column=1, sticky="ew", padx=8)
        self.sheet_list = tk.Listbox(sheet_frame, selectmode="extended", exportselection=False, height=4)
        self.sheet_list.pack(side="left", fill="x", expand=True)
        sheet_scroll = ttk.Scrollbar(sheet_frame, orient="vertical", command=self.sheet_list.yview)
        sheet_scroll.pack(side="right", fill="y")
        self.sheet_list.configure(yscrollcommand=sheet_scroll.set)
        self.sheet_list.bind("<<ListboxSelect>>", lambda _event: self.load_selected_preview())
        sheet_buttons = ttk.Frame(self)
        sheet_buttons.grid(row=4, column=2, sticky="nsew")
        ttk.Button(sheet_buttons, text="Alle auswählen", command=self._select_all_sheets).pack(fill="x", pady=(0, 3))
        ttk.Button(sheet_buttons, text="Nur *_DP", command=self._select_dp_sheets).pack(fill="x", pady=3)
        ttk.Button(sheet_buttons, text="Zuordnung (Fremdformat)", command=self.configure_mapping).pack(fill="x", pady=(3, 0))
        self.preview = ttk.Treeview(self, show="headings", height=16)
        self.preview.grid(row=5, column=0, columnspan=3, sticky="nsew", pady=(12, 6))
        preview_scroll = ttk.Scrollbar(self, orient="horizontal", command=self.preview.xview)
        preview_scroll.grid(row=6, column=0, columnspan=3, sticky="ew")
        self.preview.configure(xscrollcommand=preview_scroll.set)
        self.generate_button = ttk.Button(self, text="Beschriftung generieren", command=self.start_generation)
        self.generate_button.grid(row=7, column=0, columnspan=3, sticky="ew", pady=12, ipady=8)
        ttk.Progressbar(self, variable=self.progress_var, maximum=100).grid(row=8, column=0, columnspan=3, sticky="ew")
        ttk.Label(self, textvariable=self.status_var, wraplength=900).grid(row=9, column=0, columnspan=2, sticky="w", pady=8)
        ttk.Button(self, text="Ordner öffnen", command=self.open_output_folder).grid(row=9, column=2, sticky="e")
        self.columnconfigure(1, weight=1)
        self.rowconfigure(5, weight=1)

    def _configure_drag_drop(self) -> None:
        if DND_FILES is None or not hasattr(self.drop_label, "drop_target_register"):
            self.drop_label.configure(text="Drag-and-drop nicht verfügbar – bitte Rohdatei auswählen")
            return
        for widget in (self.drop_label, self.source_entry):
            widget.drop_target_register(DND_FILES)
            widget.dnd_bind("<<Drop>>", self._on_source_drop)
        self.drop_label.dnd_bind("<<DragEnter>>", lambda _event: self.drop_label.configure(bg="#dbeafe"))
        self.drop_label.dnd_bind("<<DragLeave>>", lambda _event: self.drop_label.configure(bg="#f3f6fa"))

    def _on_source_drop(self, event: object) -> str:
        self.drop_label.configure(bg="#f3f6fa")
        data = getattr(event, "data", "")
        try:
            paths = [Path(value) for value in self.master.tk.splitlist(data)]
        except tk.TclError:
            paths = []
        if len(paths) != 1:
            messagebox.showerror("Fehler", "Bitte genau eine Rohdatei ablegen.")
            return "break"
        self._set_source_file(paths[0])
        return "copy"

    def choose_source(self) -> None:
        value = filedialog.askopenfilename(filetypes=[("Rohdaten", "*.xlsx *.xlsm *.csv"), ("Alle Dateien", "*.*")])
        if value:
            self._set_source_file(Path(value))

    def _set_source_file(self, path: Path) -> None:
        if not path.is_file():
            messagebox.showerror("Fehler", f"Die abgelegte Datei wurde nicht gefunden:\n{path}")
            return
        if path.suffix.lower() not in {".xlsx", ".xlsm", ".csv"}:
            messagebox.showerror("Fehler", "Unterstützt werden nur .xlsx-, .xlsm- und .csv-Dateien.")
            return
        self.source_var.set(str(path))
        self.output_var.set(str(path.with_name(f"{path.stem}_Beschriftung.xlsx")))
        try:
            sheets = list_sheets(path)
            self._set_sheet_choices(sheets)
            self.load_selected_preview()
            self.drop_label.configure(text=f"Geladen: {path.name}")
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

    def _set_sheet_choices(self, sheets: list[str]) -> None:
        self.sheet_list.delete(0, "end")
        if not sheets:
            self.sheet_list.insert("end", "CSV-Datei")
            self.sheet_list.selection_set(0)
            return
        for sheet in sheets:
            self.sheet_list.insert("end", sheet)
        self._select_dp_sheets()

    def _select_all_sheets(self) -> None:
        self.sheet_list.selection_set(0, "end")
        self.load_selected_preview()

    def _select_dp_sheets(self) -> None:
        self.sheet_list.selection_clear(0, "end")
        matches = [index for index in range(self.sheet_list.size()) if "_DP" in self.sheet_list.get(index).upper()]
        for index in matches or ([0] if self.sheet_list.size() else []):
            self.sheet_list.selection_set(index)
        self.load_selected_preview()

    def _selected_sheets(self) -> list[str | None]:
        if Path(self.source_var.get()).suffix.lower() == ".csv":
            return [None]
        return [self.sheet_list.get(index) for index in self.sheet_list.curselection()]

    def load_selected_preview(self) -> None:
        selected = self._selected_sheets()
        if not selected or not self.source_var.get():
            return
        sheet_name = selected[0]
        self.sheet_var.set(sheet_name or "")
        self.source_table = load_source_file(Path(self.source_var.get()), sheet_name)
        self.column_mapping = self._load_mapping_profile(self.source_table.headers) or detect_columns(self.source_table.headers)
        self._show_preview(self.source_table)
        missing = [FIELD_LABELS[field] for field in REQUIRED_FIELDS if field not in self.column_mapping]
        suffix = f" Nicht erkannt: {', '.join(missing)}." if missing else " Spalten automatisch erkannt."
        self.status_var.set(
            f"{len(selected)} Tabellenblatt/-blätter ausgewählt. "
            f"Vorschau: {len(self.source_table.rows)} Datensätze.{suffix}"
        )

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
        selected_sheets = self._selected_sheets()
        if not selected_sheets:
            messagebox.showerror("Fehler", "Bitte mindestens ein Tabellenblatt auswählen.")
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
        datasets: list[tuple[str, list[Record]]] = []
        try:
            for sheet_name in selected_sheets:
                table = load_source_file(Path(self.source_var.get()), sheet_name)
                if self.source_table is not None and table.headers == self.source_table.headers:
                    mapping = dict(self.column_mapping)
                else:
                    mapping = self._load_mapping_profile(table.headers) or detect_columns(table.headers)
                missing = [field for field in ("module", "module_type", "channel", "description") if field not in mapping]
                if missing:
                    shown_name = sheet_name or "CSV-Datei"
                    raise ValueError(
                        f"Tabellenblatt '{shown_name}': Felder nicht automatisch erkannt: {', '.join(missing)}"
                    )
                datasets.append((sheet_name or Path(self.source_var.get()).stem, normalize_data(table.rows, mapping)))
        except Exception as exc:
            self._error(exc)
            return
        header_dialog = HeaderDataDialog(self, self.header_values, self.header_image)
        self.wait_window(header_dialog)
        if header_dialog.result is None:
            return
        self.header_values, self.header_image = header_dialog.result
        self.generate_button.configure(state="disabled")
        self.progress_var.set(0)
        threading.Thread(
            target=self._generate,
            args=(master, output, datasets, replace_existing, self.header_values, self.header_image),
            daemon=True,
        ).start()

    def _generate(
        self,
        master: Path,
        output: Path,
        datasets: list[tuple[str, list[Record]]],
        replace_existing: bool,
        header_values: dict[str, str],
        header_image: Path | None,
    ) -> None:
        generated_path = output
        if replace_existing:
            generated_path = output.with_name(f".{output.stem}.{uuid.uuid4().hex}.tmp.xlsx")
        try:
            config = load_json_config(DEFAULT_CONFIG)
            result = generate_multi_sheet_workbook(
                master,
                generated_path,
                datasets,
                config,
                header_values=header_values,
                header_image=header_image,
                progress=self._progress,
            )
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
    root = TkinterDnD.Tk() if TkinterDnD is not None else tk.Tk()
    SmartStruxureApp(root)
    root.mainloop()
