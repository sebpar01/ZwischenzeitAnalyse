"""
Rodelrennen – Zwischenzeitenanalyse

Funktionen:
- Excel-Datei auswählen
- Lauf auswählen (z.B. 1. Lauf / 2. Lauf)
- Athleten auswählen, die verglichen werden sollen
- Einen der ausgewählten Athleten als Referenz bestimmen
- Zwischenzeiten gegen die Referenz berechnen
- x-Achse mit echten Zeitabständen darstellen
- schneller als Referenz = grün
- langsamer als Referenz = rot
- Referenz = y=0
- Diagramm als PNG speichern

Benötigte Pakete:
    pip install pandas matplotlib openpyxl
"""

import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt


# ------------------------------------------------------------
# Excel einlesen
# ------------------------------------------------------------

def load_excel(path):
    sheets = pd.ExcelFile(path).sheet_names
    if not sheets:
        raise ValueError("Die Excel-Datei enthält keine Tabellenblätter.")

    df = pd.read_excel(path, sheet_name=sheets[0])

    # Spaltennamen bereinigen
    df.columns = [str(c).strip() for c in df.columns]

    required = {"Name", "Lauf"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Folgende benötigte Spalten fehlen: {', '.join(missing)}")

    # Mögliche Schreibweisen für Nation und Disziplin
    nation_candidates = ["Nation", "NATION", "Nat", "NAT", "Country"]
    discipline_candidates = [
        "Disziplin", "Discipline", "Event", "Wettkampf", "Competition"
    ]

    nation_col = next((c for c in nation_candidates if c in df.columns), None)
    discipline_col = next((c for c in discipline_candidates if c in df.columns), None)

    # Zwischenzeit-Spalten automatisch erkennen
    split_cols = [
        c for c in df.columns
        if str(c).lower().startswith("zwischenzeit")
    ]

    if not split_cols:
        raise ValueError("Keine Spalten 'Zwischenzeit ...' gefunden.")

    # Nur numerische Zwischenzeiten verwenden
    for c in split_cols:
        df[c] = pd.to_numeric(df[c], errors="coerce")

    df["Name"] = df["Name"].astype(str).str.strip()
    df["Lauf"] = df["Lauf"].astype(str).str.strip()

    if nation_col:
        df[nation_col] = df[nation_col].fillna("").astype(str).str.strip()

    if discipline_col:
        df[discipline_col] = df[discipline_col].fillna("").astype(str).str.strip()

    # "Name" aus einer eventuellen Kopfzeile entfernen
    df = df[df["Name"].str.lower() != "name"].copy()

    return df, split_cols, nation_col, discipline_col


# ------------------------------------------------------------
# Analyse
# ------------------------------------------------------------

def find_finish_col(df):
    """Findet die Laufzeit-/Zielzeit-Spalte."""
    candidates = [
        "Laufzeit", "Laufzeit [s]", "Laufzeit_s",
        "Zielzeit", "Zielzeit [s]", "Finish", "Finishzeit",
        "Gesamtzeit", "Total Time"
    ]
    for c in candidates:
        if c in df.columns:
            return c
    return None


def calculate_analysis(
    df,
    split_cols,
    run_name,
    selected_names,
    reference_name,
):
    run_df = df[df["Lauf"] == run_name].copy()

    if run_df.empty:
        raise ValueError(f"Keine Daten für '{run_name}' gefunden.")

    run_df = run_df.drop_duplicates(subset=["Name"], keep="first")

    reference_rows = run_df[run_df["Name"] == reference_name]
    if reference_rows.empty:
        raise ValueError(
            f"Referenzathlet '{reference_name}' wurde im Lauf nicht gefunden."
        )

    reference = reference_rows.iloc[0]

    # Die Laufzeit in der Excel ist die Zielzeit und wird als
    # letzter Punkt "Ziel" ergänzt.
    finish_col = find_finish_col(df)

    if finish_col:
        # Laufzeit/Zielzeit kann in Excel als Zahl oder als Text
        # vorliegen. Für die Berechnung immer explizit in Sekunden
        # umwandeln.
        run_df[finish_col] = pd.to_numeric(
            run_df[finish_col],
            errors="coerce"
        )
        reference_finish = pd.to_numeric(
            reference[finish_col],
            errors="coerce"
        )
    else:
        reference_finish = np.nan

    # x-Positionen: tatsächliche kumulierte Zeiten des Referenzathleten.
    x = [0.0]
    labels = ["Start"]

    valid_cols = []

    for col in split_cols:
        value = reference[col]
        if pd.notna(value):
            valid_cols.append(col)
            x.append(float(value))
            labels.append(str(col))

    # Ziel/Laufzeit als letzter Punkt.
    if pd.notna(reference_finish):
        x.append(float(reference_finish))
        labels.append("Ziel")

    result = {}

    for name in selected_names:
        athlete_rows = run_df[run_df["Name"] == name]
        if athlete_rows.empty:
            continue

        athlete = athlete_rows.iloc[0]
        y = [0.0]

        for col in valid_cols:
            athlete_time = athlete[col]
            ref_time = reference[col]

            if pd.isna(athlete_time) or pd.isna(ref_time):
                y.append(np.nan)
            else:
                # positiv = langsamer
                # negativ = schneller
                y.append(float(athlete_time - ref_time))

        # Zielzeit ebenfalls relativ zur Referenz.
        if pd.notna(reference_finish):
            athlete_finish = pd.to_numeric(
                athlete[finish_col],
                errors="coerce"
            )

            if pd.isna(athlete_finish):
                y.append(np.nan)
            else:
                y.append(float(athlete_finish - reference_finish))

        result[name] = y

    return np.array(x), labels, result


def get_event_title(df, run_name, discipline_col):
    """
    Erstellt den Titel aus Lauf + Disziplin.

    Beispiel:
        2nd Run / WC Men's Singles Lake Placid
    """
    run_part = str(run_name).strip()

    if discipline_col and discipline_col in df.columns:
        values = (
            df[discipline_col]
            .dropna()
            .astype(str)
            .str.strip()
        )
        values = values[values != ""]

        if not values.empty:
            return f"{run_part} / {values.iloc[0]}"

    return run_part


# ------------------------------------------------------------
# Diagramm
# ------------------------------------------------------------

def plot_analysis(x, labels, result, reference_name, title, output_path=None):
    fig, ax = plt.subplots(figsize=(14, 7))

    # Referenz = 0
    ax.axhline(
        0,
        color="black",
        linewidth=1.5,
        linestyle="-",
        zorder=4
    )

    colors = plt.cm.tab10(np.linspace(0, 1, max(len(result), 1)))

    # --------------------------------------------------------
    # WICHTIG:
    #
    # Die Farbe eines Sektors beschreibt NICHT, ob der Athlet
    # insgesamt im Rückstand/Vorsprung ist.
    #
    # Sie beschreibt die Veränderung des Rückstands zwischen
    # zwei Messpunkten:
    #
    # y[i]   = Rückstand/Vorsprung bei ZZ_i
    # y[i+1] = Rückstand/Vorsprung bei ZZ_i+1
    #
    # delta = y[i+1] - y[i]
    #
    # delta < 0 -> Rückstand wird kleiner / Vorsprung wird größer
    #              => Athlet war in diesem Abschnitt SCHNELLER
    #              => GRÜN
    #
    # delta > 0 -> Rückstand wird größer / Vorsprung wird kleiner
    #              => Athlet war in diesem Abschnitt LANGSAMER
    #              => ROT
    #
    # Beispiel:
    # ZZ2 = +0.123 s
    # ZZ3 = +0.023 s
    # delta = 0.023 - 0.123 = -0.100 s
    # => 0.100 s aufgeholt => GRÜN
    # --------------------------------------------------------

    for name, y in result.items():
        if name == reference_name:
            continue

        y = np.asarray(y, dtype=float)

        for i in range(len(x) - 1):
            if np.isnan(y[i]) or np.isnan(y[i + 1]):
                continue

            x0, x1 = x[i], x[i + 1]
            y0, y1 = y[i], y[i + 1]

            delta = y1 - y0

            if delta < 0:
                # Rückstand wird kleiner / Vorsprung wird größer
                # => schneller in diesem Sektor
                fill_color = "green"
            elif delta > 0:
                # Rückstand wird größer / Vorsprung wird kleiner
                # => langsamer in diesem Sektor
                fill_color = "red"
            else:
                # exakt gleich schnell wie Referenz
                continue

            # Die Fläche wird zwischen der tatsächlichen Athletenlinie
            # und der Referenzlinie dargestellt. Die Farbe kommt aber
            # ausschließlich vom Sektor-Zeitgewinn/-verlust.
            ax.fill_between(
                [x0, x1],
                [y0, y1],
                [0, 0],
                color=fill_color,
                alpha=0.10,
                linewidth=0,
                zorder=1
            )

    # --------------------------------------------------------
    # Gestrichelte Linien von jedem Messpunkt zur Referenz.
    # Diese zeigen den kumulierten Rückstand/Vorsprung.
    # --------------------------------------------------------
    color_index = 0

    for name, y in result.items():
        if name == reference_name:
            continue

        y = np.asarray(y, dtype=float)
        line_color = colors[color_index % len(colors)]
        color_index += 1

        for xi, yi in zip(x, y):
            if np.isnan(yi):
                continue

            ax.plot(
                [xi, xi],
                [0, yi],
                color=line_color,
                linestyle="--",
                linewidth=0.9,
                alpha=0.35,
                zorder=3
            )

        # Athletenlinie
        ax.plot(
            x,
            y,
            marker="o",
            linewidth=2,
            markersize=5,
            label=name,
            color=line_color,
            zorder=6
        )

    # --------------------------------------------------------
    # Referenzlinie
    # --------------------------------------------------------
    ax.plot(
        x,
        np.zeros(len(x)),
        color="black",
        linewidth=2.5,
        label=f"Referenz: {reference_name}",
        zorder=7
    )

    # --------------------------------------------------------
    # Sektorgrenzen
    # --------------------------------------------------------
    for xi in x:
        ax.axvline(
            xi,
            color="gray",
            linewidth=0.6,
            linestyle=":",
            alpha=0.22,
            zorder=0
        )

    ax.set_xticks(x)
    ax.set_xticklabels(labels, rotation=35, ha="right")

    ax.set_xlabel(
        "Zeit im Lauf [s] – Abstände proportional zu den tatsächlichen Zwischenzeiten"
    )

    ax.set_ylabel(
        "Kumulierte Zeitdifferenz zur Referenz [s]\n"
        "negativ = Vorsprung | positiv = Rückstand"
    )

    ax.set_title(title)

    ax.grid(True, axis="y", alpha=0.25)

    handles, legend_labels = ax.get_legend_handles_labels()

    ref_idx = next(
        (
            i for i, label in enumerate(legend_labels)
            if label == f"Referenz: {reference_name}"
        ),
        None
    )

    if ref_idx is not None:
        ref_handle = handles.pop(ref_idx)
        ref_label = legend_labels.pop(ref_idx)
        handles.insert(0, ref_handle)
        legend_labels.insert(0, ref_label)

    ax.legend(handles, legend_labels, loc="best", fontsize=9)

    if len(x) > 1:
        margin = max((x[-1] - x[0]) * 0.02, 0.1)
        ax.set_xlim(x[0] - margin, x[-1] + margin)

    fig.tight_layout()

    if output_path:
        fig.savefig(output_path, dpi=200, bbox_inches="tight")

    plt.show()


# ------------------------------------------------------------
# GUI
# ------------------------------------------------------------

class AnalysisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Rodelrennen – Zwischenzeitenanalyse")
        self.root.geometry("900x700")

        self.df = None
        self.split_cols = None
        self.nation_col = None
        self.discipline_col = None
        self.file_path = None

        self.run_var = tk.StringVar()
        self.reference_var = tk.StringVar()
        self.display_to_name = {}
        self.selected_names = []

        self.build_gui()

    def build_gui(self):
        main = ttk.Frame(self.root, padding=15)
        main.pack(fill="both", expand=True)

        # Datei
        file_frame = ttk.LabelFrame(main, text="1. Excel-Datei", padding=10)
        file_frame.pack(fill="x", pady=(0, 10))

        self.file_label = ttk.Label(
            file_frame,
            text="Noch keine Datei ausgewählt."
        )
        self.file_label.pack(side="left", fill="x", expand=True)

        ttk.Button(
            file_frame,
            text="Excel auswählen",
            command=self.select_file
        ).pack(side="right")

        # Lauf
        run_frame = ttk.LabelFrame(main, text="2. Lauf auswählen", padding=10)
        run_frame.pack(fill="x", pady=(0, 10))

        self.run_combo = ttk.Combobox(
            run_frame,
            textvariable=self.run_var,
            state="readonly",
            width=30
        )
        self.run_combo.pack(side="left")
        self.run_combo.bind("<<ComboboxSelected>>", self.update_athletes)

        # Referenz
        ref_frame = ttk.LabelFrame(
            main,
            text="3. Referenzathlet",
            padding=10
        )
        ref_frame.pack(fill="x", pady=(0, 10))

        ttk.Label(
            ref_frame,
            text="Die Referenz liegt bei y = 0:"
        ).pack(side="left", padx=(0, 10))

        self.reference_combo = ttk.Combobox(
            ref_frame,
            textvariable=self.reference_var,
            state="readonly",
            width=40
        )
        self.reference_combo.pack(side="left")


        # Athleten
        athlete_frame = ttk.LabelFrame(
            main,
            text="4. Athleten auswählen",
            padding=10
        )
        athlete_frame.pack(fill="both", expand=True, pady=(0, 10))

        info = ttk.Label(
            athlete_frame,
            text="Strg gedrückt halten, um mehrere Athleten auszuwählen."
        )
        info.pack(anchor="w", pady=(0, 5))

        list_frame = ttk.Frame(athlete_frame)
        list_frame.pack(fill="both", expand=True)

        scrollbar = ttk.Scrollbar(
            list_frame,
            orient="vertical"
        )
        scrollbar.pack(side="right", fill="y")

        self.athlete_list = tk.Listbox(
            list_frame,
            selectmode=tk.EXTENDED,
            yscrollcommand=scrollbar.set,
            font=("Segoe UI", 10)
        )
        self.athlete_list.pack(side="left", fill="both", expand=True)

        scrollbar.config(command=self.athlete_list.yview)

        button_frame = ttk.Frame(athlete_frame)
        button_frame.pack(fill="x", pady=(8, 0))

        ttk.Button(
            button_frame,
            text="Alle auswählen",
            command=self.select_all
        ).pack(side="left", padx=(0, 5))

        ttk.Button(
            button_frame,
            text="Auswahl löschen",
            command=self.clear_selection
        ).pack(side="left")

        # Analyse
        ttk.Button(
            main,
            text="ANALYSE ERSTELLEN",
            command=self.run_analysis
        ).pack(fill="x", ipady=8)

    def select_file(self):
        path = filedialog.askopenfilename(
            title="Excel-Datei auswählen",
            filetypes=[
                ("Excel-Dateien", "*.xlsx *.xls"),
                ("Alle Dateien", "*.*")
            ]
        )

        if not path:
            return

        try:
            self.df, self.split_cols, self.nation_col, self.discipline_col = load_excel(path)
            self.file_path = path

            self.file_label.config(
                text=os.path.basename(path)
            )

            runs = [
                r for r in self.df["Lauf"].dropna().unique()
                if str(r).strip()
            ]

            self.run_combo["values"] = runs

            if runs:
                self.run_combo.current(0)
                self.run_var.set(runs[0])
                self.update_athletes()

        except Exception as e:
            messagebox.showerror(
                "Fehler beim Einlesen",
                str(e)
            )

    def update_athletes(self, event=None):
        if self.df is None:
            return

        # Aktuelle Auswahl sichern, bevor die Listbox neu aufgebaut wird.
        current_selected = []
        for i in self.athlete_list.curselection():
            display = self.athlete_list.get(i)
            name = self.display_to_name.get(display, display)
            current_selected.append(name)

        # Falls bereits eine Auswahl gespeichert wurde, diese ebenfalls
        # berücksichtigen.
        if current_selected:
            self.selected_names = current_selected

        run_name = self.run_var.get()
        run_df = self.df[self.df["Lauf"] == run_name].copy()

        names = []

        for _, row in run_df.drop_duplicates(subset=["Name"]).iterrows():
            name = str(row["Name"]).strip()

            if self.nation_col and self.nation_col in row.index:
                nation = str(row[self.nation_col]).strip()
            else:
                nation = ""

            display = f"{name} ({nation})" if nation else name
            names.append((name, display))

        names.sort(key=lambda item: item[0])

        self.athlete_list.delete(0, tk.END)
        self.display_to_name = {}

        for name, display in names:
            self.athlete_list.insert(tk.END, display)
            self.display_to_name[display] = name

        # Bereits ausgewählte Athleten wieder markieren.
        available_names = set(self.display_to_name.values())

        for i in range(self.athlete_list.size()):
            display = self.athlete_list.get(i)
            name = self.display_to_name.get(display)

            if name in self.selected_names and name in available_names:
                self.athlete_list.selection_set(i)

        self.reference_combo["values"] = [
            display for _, display in names
        ]

        # Referenz beibehalten, wenn sie im neuen Lauf vorhanden ist.
        current_reference = self.reference_var.get()

        if current_reference in self.reference_combo["values"]:
            self.reference_combo.set(current_reference)
        elif names:
            self.reference_combo.current(0)

    def select_all(self):
        self.athlete_list.select_set(0, tk.END)
        self.selected_names = [
            self.display_to_name.get(self.athlete_list.get(i), self.athlete_list.get(i))
            for i in range(self.athlete_list.size())
        ]

    def clear_selection(self):
        self.athlete_list.selection_clear(0, tk.END)
        self.selected_names = []

    def run_analysis(self):
        if self.df is None:
            messagebox.showwarning(
                "Keine Datei",
                "Bitte zuerst eine Excel-Datei auswählen."
            )
            return

        run_name = self.run_var.get()
        reference_display = self.reference_var.get()
        reference_name = self.display_to_name.get(
            reference_display,
            reference_display
        )

        selected_indices = self.athlete_list.curselection()

        # Auswahl aus Schritt 3 explizit speichern.
        if selected_indices:
            self.selected_names = [
                self.display_to_name.get(
                    self.athlete_list.get(i),
                    self.athlete_list.get(i)
                )
                for i in selected_indices
            ]

        selected_names = list(self.selected_names)

        # Der Referenzathlet wird automatisch in die Analyse aufgenommen,
        # auch wenn er in Schritt 4 nicht nochmals ausgewählt wurde.
        if reference_name not in selected_names:
            selected_names.insert(0, reference_name)

        if not selected_names:
            messagebox.showwarning(
                "Keine Athleten",
                "Bitte mindestens einen Athleten auswählen oder eine Referenz festlegen."
            )
            return

        try:
            x, labels, result = calculate_analysis(
                self.df,
                self.split_cols,
                run_name,
                selected_names,
                reference_name
            )

            # Speichern neben der Excel-Datei
            base = os.path.splitext(self.file_path)[0]
            output_path = (
                f"{base}_Analyse_{run_name.replace(' ', '_')}.png"
            )

            event_title = get_event_title(
                self.df,
                run_name,
                self.discipline_col
            )

            title = (
                f"{event_title}\n"
                f"Zeitdifferenzen | Referenz: {reference_name}"
            )

            plot_analysis(
                x,
                labels,
                result,
                reference_name,
                title,
                output_path
            )

            messagebox.showinfo(
                "Analyse fertig",
                f"Diagramm wurde gespeichert:\n\n{output_path}"
            )

        except Exception as e:
            messagebox.showerror(
                "Fehler bei der Analyse",
                str(e)
            )


# ------------------------------------------------------------
# Start
# ------------------------------------------------------------

if __name__ == "__main__":
    root = tk.Tk()
    app = AnalysisApp(root)
    root.mainloop()
