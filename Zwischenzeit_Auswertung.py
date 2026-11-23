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
import sys
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages


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



# ------------------------------------------------------------
# Zusätzliche Sektoren-Tabelle
# ------------------------------------------------------------

def find_startnr_col(df):
    """Findet die Startnummer-Spalte (Bib)."""
    candidates = ["Startnr", "Startnummer", "Bib", "Bib Number", "Start-Nr"]
    return next((c for c in candidates if c in df.columns), None)


def calculate_sector_table_data(df, split_cols, run_name, selected_names):
    """Berechnet Sektorzeiten, virtuelle Bestzeiten und Platzierungen."""
    run_df = df[df["Lauf"] == run_name].copy()
    run_df = run_df.drop_duplicates(subset=["Name"], keep="first")

    finish_col = find_finish_col(df)
    if not finish_col:
        raise ValueError(
            "Für die Sektoren-Tabelle wurde keine Laufzeit-/Zielzeit-Spalte gefunden. "
            "Erwartet wird z. B. 'Laufzeit'."
        )

    startnr_col = find_startnr_col(df)
    if not startnr_col:
        raise ValueError(
            "Die Startnummer-Spalte wurde nicht gefunden. "
            "Bitte prüfe, ob sie in Excel 'Startnr' heißt."
        )

    nation_candidates = ["Nation", "NATION", "Nat", "NAT", "Country"]
    nation_col = next((c for c in nation_candidates if c in df.columns), None)
    if nation_col is None:
        raise ValueError("Die Nation-Spalte wurde nicht gefunden. Erwartet wird 'Nation'.")

    # Nur Zwischenzeiten berücksichtigen, für die zumindest ein
    # ausgewählter Athlet einen Wert besitzt.
    athletes = []
    for name in selected_names:
        rows = run_df[run_df["Name"] == name]
        if rows.empty:
            continue
        row = rows.iloc[0]
        cumulative = [0.0]
        for col in split_cols:
            value = pd.to_numeric(row[col], errors="coerce")
            cumulative.append(float(value) if pd.notna(value) else np.nan)

        finish = pd.to_numeric(row[finish_col], errors="coerce")
        if pd.isna(finish):
            # Ohne Zielzeit kann weder Gesamtplatz noch letzter Sektor
            # zuverlässig berechnet werden.
            continue

        cumulative.append(float(finish))
        sector_times = []
        for i in range(len(cumulative) - 1):
            start_time = cumulative[i]
            end_time = cumulative[i + 1]
            if pd.isna(start_time) or pd.isna(end_time):
                sector_times.append(np.nan)
            else:
                sector_times.append(end_time - start_time)

        athletes.append({
            "name": name,
            "bib": row[startnr_col],
            "nation": str(row[nation_col]).strip() if pd.notna(row[nation_col]) else "",
            "finish": float(finish),
            "sector_times": sector_times,
        })

    if not athletes:
        raise ValueError(
            "Für die ausgewählten Athleten konnten keine vollständigen Laufzeiten gefunden werden."
        )

    # Bei unvollständigen Zwischenzeiten kann die Anzahl der Sektoren
    # unterschiedlich sein. Für die Tabelle werden die vorhandenen
    # Zwischenzeit-Spalten gemeinsam verwendet; fehlende Werte bleiben leer.
    n_sectors = len(split_cols) + 1
    for athlete in athletes:
        athlete["sector_times"] = (athlete["sector_times"] + [np.nan] * n_sectors)[:n_sectors]

    sector_bests = []
    sector_ranks = {a["name"]: [] for a in athletes}
    sector_gaps = {a["name"]: [] for a in athletes}

    for sector_idx in range(n_sectors):
        values = [a["sector_times"][sector_idx] for a in athletes]
        valid = [v for v in values if pd.notna(v)]
        best = min(valid) if valid else np.nan
        sector_bests.append(best)

        for athlete in athletes:
            value = athlete["sector_times"][sector_idx]
            if pd.isna(value) or pd.isna(best):
                sector_gaps[athlete["name"]].append(np.nan)
                sector_ranks[athlete["name"]].append(np.nan)
            else:
                sector_gaps[athlete["name"]].append(value - best)
                # Competition ranking: 1, 2, 2, 4
                rank = 1 + sum(v < value for v in valid)
                sector_ranks[athlete["name"]].append(rank)

    virtual_total = float(np.nansum(sector_bests))

    # Gesamtplatzierung nach tatsächlicher Ziel-/Laufzeit.
    athletes.sort(key=lambda a: a["finish"])
    previous_time = None
    current_rank = 0
    for index, athlete in enumerate(athletes, start=1):
        if previous_time is None or not np.isclose(athlete["finish"], previous_time):
            current_rank = index
        athlete["overall_rank"] = current_rank
        previous_time = athlete["finish"]

    return {
        "athletes": athletes,
        "sector_bests": sector_bests,
        "sector_gaps": sector_gaps,
        "sector_ranks": sector_ranks,
        "virtual_total": virtual_total,
        "sector_labels": [f"Sektor {i + 1}" for i in range(n_sectors)],
    }


def format_seconds(value, total=False):
    if pd.isna(value):
        return "–"
    value = float(value)
    if value >= 60 or total:
        minutes = int(value // 60)
        seconds = value - minutes * 60
        return f"{minutes}:{seconds:06.3f}"
    return f"{value:.3f}"


def plot_sector_table(table_data, title, output_path, show=True):
    """Erstellt eine PNG-Tabelle im Stil der bereitgestellten Ergebnisübersicht."""
    athletes = table_data["athletes"]
    sector_bests = table_data["sector_bests"]
    sector_gaps = table_data["sector_gaps"]
    sector_ranks = table_data["sector_ranks"]
    n_sectors = len(sector_bests)

    # Spaltenbreiten: Platz, Bib, Name, Nation, je Sektor Zeit/Rückst./Rang, Gesamt.
    widths = [0.55, 0.62, 2.15, 0.72]
    for _ in range(n_sectors):
        widths.extend([0.86, 0.90, 0.58])
    widths.append(0.95)
    x_edges = np.cumsum([0] + widths)
    total_width = x_edges[-1]

    header_h = 0.48
    subheader_h = 0.43
    row_h = 0.48
    n_rows = len(athletes) + 1  # virtuelle Bestzeit + Athleten
    total_height = header_h + subheader_h + n_rows * row_h

    fig_width = max(12, total_width * 0.82)
    fig_height = max(3.5, total_height * 0.52 + 0.9)
    fig, ax = plt.subplots(figsize=(fig_width, fig_height))
    ax.set_xlim(0, total_width)
    ax.set_ylim(0, total_height)
    ax.axis("off")

    # Dezente Tabellenfarben ähnlich dem Beispiel.
    header_color = "#b7d3ec"
    group_header_color = "#d8e8f5"
    virtual_color = "#e5f3e9"
    grid_color = "#b8b8b8"

    def cell(x0, y0, w, h, text, facecolor="white", color="black",
             fontsize=8.5, weight="normal", ha="center", pad=0.05):
        rect = plt.Rectangle((x0, y0), w, h, facecolor=facecolor,
                             edgecolor=grid_color, linewidth=0.45)
        ax.add_patch(rect)
        tx = x0 + (pad if ha == "left" else w / 2)
        ax.text(tx, y0 + h / 2, str(text), ha=ha, va="center",
                fontsize=fontsize, color=color, fontweight=weight,
                clip_on=True)

    # Titel
    ax.text(total_width / 2, total_height + 0.12, title,
            ha="center", va="bottom", fontsize=11, fontweight="bold")

    # Kopfzeile mit zusammengefassten Sektorüberschriften.
    y_group = total_height - header_h
    for i, label in enumerate(["Platz", "Bib", "Name", "Nation"]):
        cell(x_edges[i], y_group - subheader_h, widths[i], header_h + subheader_h,
             label, facecolor=header_color, weight="bold", fontsize=8.5)

    for sector_idx in range(n_sectors):
        start_col = 4 + sector_idx * 3
        x0 = x_edges[start_col]
        w = sum(widths[start_col:start_col + 3])
        cell(x0, y_group, w, header_h, f"Sektor {sector_idx + 1}",
             facecolor=group_header_color, weight="bold", fontsize=9)
        for offset, sublabel in enumerate(["Zeit", "Rückst.", "Rang"]):
            col = start_col + offset
            cell(x_edges[col], y_group - subheader_h, widths[col], subheader_h,
                 sublabel, facecolor=header_color, weight="bold", fontsize=8)

    total_col = len(widths) - 1
    cell(x_edges[total_col], y_group - subheader_h, widths[total_col],
         header_h + subheader_h, "Gesamt", facecolor=header_color,
         weight="bold", fontsize=8.5)

    # Virtuelle Bestzeit-Zeile.
    y = y_group - subheader_h - row_h
    cell(x_edges[0], y, widths[0], row_h, "", facecolor=virtual_color)
    cell(x_edges[1], y, widths[1], row_h, "", facecolor=virtual_color)
    cell(x_edges[2], y, widths[2], row_h, "Virtuelle Bestzeit",
         facecolor=virtual_color, ha="left", weight="bold", fontsize=8)
    cell(x_edges[3], y, widths[3], row_h, "", facecolor=virtual_color)

    for sector_idx, best in enumerate(sector_bests):
        start_col = 4 + sector_idx * 3
        cell(x_edges[start_col], y, widths[start_col], row_h,
             format_seconds(best), facecolor=virtual_color,
             color="#00843d", weight="bold")
        # Rückstand und Rang bleiben bei der virtuellen Bestzeit leer.

    cell(x_edges[total_col], y, widths[total_col], row_h,
         format_seconds(table_data["virtual_total"], total=True),
         facecolor="#16833b", color="white", weight="bold")

    # Athleten-Zeilen.
    max_gap = max(
        [g for values in sector_gaps.values() for g in values if pd.notna(g)] + [0.01]
    )
    max_total_gap = max([a["finish"] - min(x["finish"] for x in athletes) for a in athletes] + [0.01])

    for row_idx, athlete in enumerate(athletes, start=1):
        y = y_group - subheader_h - (row_idx + 1) * row_h
        base_bg = "#f0f0f0" if row_idx % 2 == 0 else "#ffffff"
        cell(x_edges[0], y, widths[0], row_h, athlete["overall_rank"], facecolor=base_bg)
        bib = athlete["bib"]
        if pd.isna(bib):
            bib_text = "–"
        elif isinstance(bib, (float, np.floating)) and float(bib).is_integer():
            bib_text = str(int(bib))
        else:
            bib_text = str(bib)
        cell(x_edges[1], y, widths[1], row_h, bib_text, facecolor=base_bg)
        cell(x_edges[2], y, widths[2], row_h, athlete["name"],
             facecolor=base_bg, ha="left", fontsize=8)
        cell(x_edges[3], y, widths[3], row_h, athlete["nation"],
             facecolor=base_bg, fontsize=8)

        for sector_idx in range(n_sectors):
            start_col = 4 + sector_idx * 3
            sector_time = athlete["sector_times"][sector_idx]
            gap = sector_gaps[athlete["name"]][sector_idx]
            rank = sector_ranks[athlete["name"]][sector_idx]

            is_best = pd.notna(gap) and np.isclose(gap, 0.0)
            time_bg = "#e3f2e7" if is_best else base_bg
            time_color = "#00843d" if is_best else "black"
            cell(x_edges[start_col], y, widths[start_col], row_h,
                 format_seconds(sector_time), facecolor=time_bg, color=time_color)

            if pd.isna(gap):
                gap_bg = base_bg
                gap_text = "–"
            else:
                intensity = min(float(gap) / max_gap, 1.0)
                # Weiß bis kräftiges Rot.
                gap_bg = (1.0, 1.0 - 0.48 * intensity, 1.0 - 0.48 * intensity)
                gap_text = f"{gap:.3f}"
            cell(x_edges[start_col + 1], y, widths[start_col + 1], row_h,
                 gap_text, facecolor=gap_bg)

            rank_bg = "#e3f2e7" if pd.notna(rank) and rank == 1 else base_bg
            rank_color = "#00843d" if pd.notna(rank) and rank == 1 else "black"
            rank_text = "–" if pd.isna(rank) else int(rank)
            cell(x_edges[start_col + 2], y, widths[start_col + 2], row_h,
                 rank_text, facecolor=rank_bg, color=rank_color)

        total_gap = athlete["finish"] - min(a["finish"] for a in athletes)
        intensity = min(total_gap / max_total_gap, 1.0) if max_total_gap > 0 else 0
        total_bg = (1.0, 1.0 - 0.55 * intensity, 1.0 - 0.55 * intensity)
        cell(x_edges[total_col], y, widths[total_col], row_h,
             format_seconds(athlete["finish"], total=True),
             facecolor=total_bg, weight="bold")

    # Markante schwarze Trennlinien zwischen den Sektoren.
    bottom_y = 0
    # Linien laufen über die gesamte Tabellenhöhe, einschließlich der
    # oberen Sektor-Überschriften. Dadurch werden die Sektoren bereits
    # in der obersten Kopfzeile klar voneinander abgegrenzt.
    top_y = total_height
    for boundary_col in [4 + i * 3 for i in range(n_sectors)] + [total_col]:
        x = x_edges[boundary_col]
        ax.plot([x, x], [bottom_y, top_y], color="black", linewidth=1.8, zorder=10)
    ax.plot([0, total_width], [y_group - subheader_h - row_h,
                               y_group - subheader_h - row_h],
            color="black", linewidth=1.5, zorder=10)

    fig.tight_layout()
    fig.savefig(output_path, dpi=220, bbox_inches="tight")
    if show:
        plt.show()
    return fig

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

def plot_analysis(x, labels, result, reference_name, title, output_path=None, show=True):
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

    if show:
        plt.show()
    return fig


# ------------------------------------------------------------
# PDF-Export
# ------------------------------------------------------------

def create_analysis_pdf(line_fig, sector_fig, pdf_path, event_title, reference_name):
    """Erstellt einen zweiseitigen A4-Querformatbericht mit echten Vektorgrafiken.

    Die Matplotlib-Figuren werden direkt über den PDF-Backend gespeichert.
    Linien, Flächen, Tabellenzellen, Texte und Achsen bleiben dadurch
    Vektorelemente und bleiben auch beim starken Zoomen scharf.
    """
    page_size = (11.69, 8.27)  # A4 quer in Zoll
    pages = [
        (line_fig, "Zeitdifferenzen", f"Referenz: {reference_name}"),
        (sector_fig, "Sektorenanalyse", "Sektorzeiten, Rückstände und Ränge"),
    ]

    with PdfPages(pdf_path) as pdf:
        info = pdf.infodict()
        info["Title"] = str(event_title)
        info["Subject"] = "Zwischenzeitenanalyse – Zeitdifferenzen und Sektoren"
        info["Author"] = "Rodel Austria"

        for chart_fig, section_title, subtitle in pages:
            old_size = chart_fig.get_size_inches().copy()
            old_positions = [(ax, ax.get_position().frozen()) for ax in chart_fig.axes]
            old_text_sizes = []
            header_artists = []
            try:
                # Feste A4-Querformatseite; die ursprünglichen PNG-Exporte
                # wurden bereits vorher erstellt und bleiben unverändert.
                chart_fig.set_size_inches(*page_size, forward=True)

                # Platz für einen dezenten, einheitlichen PDF-Kopfbereich.
                for ax, _ in old_positions:
                    if section_title == "Sektorenanalyse":
                        # Die Tabelle behält ihre ursprünglichen Proportionen.
                        # So werden Zellen und Texte nicht horizontal gestaucht.
                        ax.set_position([0.04, 0.18, 0.92, 0.68])
                        ax.set_aspect("equal", adjustable="box")
                        # Die Tabelle wird proportional auf die Seitenbreite
                        # skaliert. Schriftgrößen werden im gleichen Verhältnis
                        # verkleinert, damit alle Werte in ihren Zellen bleiben.
                        text_scale = min(1.0, (11.69 * 0.92) / old_size[0])
                        for text_artist in ax.texts:
                            old_text_sizes.append((text_artist, text_artist.get_fontsize()))
                            text_artist.set_fontsize(text_artist.get_fontsize() * text_scale)
                            if text_artist.get_position()[1] > ax.get_ylim()[1]:
                                text_artist.set_visible(False)
                    else:
                        # Linienanalyse minimal verkleinert und etwas nach rechts
                        # gesetzt. Dadurch erhält die y-Achsenbeschriftung mehr
                        # Abstand zum linken Seitenrand, ohne das Layout merklich
                        # zu verändern.
                        ax.set_position([0.09, 0.16, 0.84, 0.70])
                        # Der PDF-Kopf enthält bereits den Veranstaltungstitel.
                        ax.set_title("")

                header_artists.append(chart_fig.text(
                    0.045, 0.965, str(event_title),
                    ha="left", va="top", fontsize=13,
                    fontweight="bold", color="#171717"
                ))
                header_artists.append(chart_fig.text(
                    0.955, 0.965, section_title,
                    ha="right", va="top", fontsize=10,
                    fontweight="bold", color="#333333"
                ))
                header_artists.append(chart_fig.text(
                    0.955, 0.925, subtitle,
                    ha="right", va="top", fontsize=8.5, color="#555555"
                ))
                header_artists.append(chart_fig.add_artist(plt.Line2D(
                    [0.045, 0.955], [0.895, 0.895],
                    transform=chart_fig.transFigure,
                    color="#222222", linewidth=0.8
                )))

                # Direkter PDF-Export der Figure: kein PNG-Zwischenschritt.
                # Matplotlib schreibt die Grafikobjekte als PDF-Vektoren.
                pdf.savefig(chart_fig, bbox_inches=None)
            finally:
                for artist in header_artists:
                    try:
                        artist.remove()
                    except Exception:
                        pass
                chart_fig.set_size_inches(*old_size, forward=True)
                for ax, position in old_positions:
                    ax.set_position(position)
                for text_artist, fontsize in old_text_sizes:
                    text_artist.set_fontsize(fontsize)
                    text_artist.set_visible(True)


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

            line_fig = plot_analysis(
                x,
                labels,
                result,
                reference_name,
                title,
                output_path,
                show=False
            )

            # Zweite Auswertung: Sektoren-Tabelle mit virtueller Bestzeit.
            sector_data = calculate_sector_table_data(
                self.df,
                self.split_cols,
                run_name,
                selected_names
            )
            sector_output_path = (
                f"{base}_Sektoren_{run_name.replace(' ', '_')}.png"
            )
            sector_fig = plot_sector_table(
                sector_data,
                event_title,
                sector_output_path,
                show=False
            )

            # PDF mit zwei Seiten: 1. Linienanalyse, 2. Sektoren-Tabelle.
            pdf_output_path = f"{base}_Analyse_{run_name.replace(' ', '_')}.pdf"
            create_analysis_pdf(
                line_fig, sector_fig, pdf_output_path, event_title, reference_name
            )

            # Grafiken wie bisher zusätzlich am Bildschirm anzeigen.
            plt.show()
            plt.close(line_fig)
            plt.close(sector_fig)

            messagebox.showinfo(
                "Analyse fertig",
                "Beide Auswertungen wurden gespeichert:\n\n"
                f"1. Diagramm:\n{output_path}\n\n"
                f"2. Sektoren-Tabelle:\n{sector_output_path}\n\n"
                f"3. PDF-Bericht (2 Seiten):\n{pdf_output_path}"
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
