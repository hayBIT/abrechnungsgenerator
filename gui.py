#!/usr/bin/env python3
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, scrolledtext

from provision import run_provision


class FilePickerRow(tk.Frame):
    def __init__(self, master: tk.Misc, label: str, multiple: bool = False) -> None:
        super().__init__(master)
        self._multiple = multiple
        tk.Label(self, text=label, width=20, anchor="w").grid(row=0, column=0, sticky="w")
        self.entry = tk.Entry(self, width=60)
        self.entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        tk.Button(self, text="Auswählen…", command=self._select).grid(row=0, column=2)
        if multiple:
            tk.Button(self, text="Leeren", command=self.clear).grid(row=0, column=3, padx=(6, 0))
        self.grid_columnconfigure(1, weight=1)

    def _select(self) -> None:
        if self._multiple:
            paths = filedialog.askopenfilenames()
            if paths:
                self.entry.delete(0, tk.END)
                self.entry.insert(0, "; ".join(paths))
        else:
            path = filedialog.askopenfilename()
            if path:
                self.entry.delete(0, tk.END)
                self.entry.insert(0, path)

    def clear(self) -> None:
        self.entry.delete(0, tk.END)

    def values(self) -> list[Path]:
        raw = self.entry.get().strip()
        if not raw:
            return []
        parts = [part.strip() for part in raw.split(";") if part.strip()]
        return [Path(part) for part in parts]

    def value(self) -> Path | None:
        raw = self.entry.get().strip()
        return Path(raw) if raw else None


class ProvisionApp(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self.title("Provisionsabrechnungssystem")
        self.geometry("900x720")

        container = tk.Frame(self, padx=12, pady=12)
        container.pack(fill="both", expand=True)

        header = tk.Label(
            container,
            text="Provisionsabrechnungssystem",
            font=("Arial", 16, "bold"),
        )
        header.pack(anchor="w", pady=(0, 12))

        self.ameise_row = FilePickerRow(container, "Ameise-CSV")
        self.ameise_row.pack(fill="x", pady=4)

        self.kravag_row = FilePickerRow(container, "KRAVAG-CSV(s)", multiple=True)
        self.kravag_row.pack(fill="x", pady=4)

        self.rv_row = FilePickerRow(container, "R+V-CSV(s)", multiple=True)
        self.rv_row.pack(fill="x", pady=4)

        self.vema_row = FilePickerRow(container, "VEMA-CSV(s)", multiple=True)
        self.vema_row.pack(fill="x", pady=4)

        self.ff_row = FilePickerRow(container, "Fonds Finanz XLSX", multiple=True)
        self.ff_row.pack(fill="x", pady=4)

        self.vermittler_row = FilePickerRow(container, "Vermittlerliste (optional)")
        self.vermittler_row.pack(fill="x", pady=4)

        output_frame = tk.Frame(container)
        output_frame.pack(fill="x", pady=8)
        tk.Label(output_frame, text="Output-Verzeichnis", width=20, anchor="w").grid(
            row=0, column=0, sticky="w"
        )
        self.output_entry = tk.Entry(output_frame, width=60)
        self.output_entry.insert(0, str(Path("output").resolve()))
        self.output_entry.grid(row=0, column=1, sticky="ew", padx=(0, 6))
        tk.Button(output_frame, text="Auswählen…", command=self._select_output).grid(
            row=0, column=2
        )
        output_frame.grid_columnconfigure(1, weight=1)

        self.run_button = tk.Button(
            container, text="Abrechnung starten", command=self._start_run, height=2
        )
        self.run_button.pack(fill="x", pady=(12, 8))

        self.log = scrolledtext.ScrolledText(container, height=16, state="disabled")
        self.log.pack(fill="both", expand=True)

    def _select_output(self) -> None:
        path = filedialog.askdirectory()
        if path:
            self.output_entry.delete(0, tk.END)
            self.output_entry.insert(0, path)

    def _append_log(self, message: str) -> None:
        self.log.configure(state="normal")
        self.log.insert(tk.END, message + "\n")
        self.log.see(tk.END)
        self.log.configure(state="disabled")

    def _start_run(self) -> None:
        ameise = self.ameise_row.value()
        if not ameise:
            messagebox.showerror("Fehlende Datei", "Bitte eine Ameise-CSV auswählen.")
            return

        kravag = self.kravag_row.values()
        rv = self.rv_row.values()
        if not kravag and not rv:
            messagebox.showerror(
                "Fehlende Abrechnung",
                "Bitte mindestens eine KRAVAG- oder R+V-Abrechnung hinzufügen.",
            )
            return

        output_dir = self.output_entry.get().strip() or "output"
        vermittlerliste = self.vermittler_row.value()

        self.run_button.configure(state="disabled")
        self._append_log("Starte Abrechnung …")

        thread = threading.Thread(
            target=self._run_provision,
            args=(
                ameise,
                kravag,
                rv,
                self.vema_row.values(),
                self.ff_row.values(),
                vermittlerliste,
                Path(output_dir),
            ),
            daemon=True,
        )
        thread.start()

    def _run_provision(
        self,
        ameise: Path,
        kravag: list[Path],
        rv: list[Path],
        vema: list[Path],
        ff: list[Path],
        vermittlerliste: Path | None,
        output_dir: Path,
    ) -> None:
        try:
            run_provision(
                ameise=ameise,
                kravag=kravag,
                rv=rv,
                vema=vema,
                ff=ff,
                vermittlerliste=vermittlerliste,
                output_dir=output_dir,
            )
        except Exception as exc:  # noqa: BLE001 - show user-friendly error dialog
            self.after(0, lambda: self._handle_error(exc))
            return
        self.after(0, self._handle_success)

    def _handle_error(self, exc: Exception) -> None:
        self._append_log(f"Fehler: {exc}")
        messagebox.showerror("Abrechnung fehlgeschlagen", str(exc))
        self.run_button.configure(state="normal")

    def _handle_success(self) -> None:
        self._append_log("Abrechnung abgeschlossen. Dateien wurden erstellt.")
        messagebox.showinfo("Abrechnung abgeschlossen", "Die Ausgaben wurden erstellt.")
        self.run_button.configure(state="normal")


def main() -> None:
    app = ProvisionApp()
    app.mainloop()


if __name__ == "__main__":
    main()
