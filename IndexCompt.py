#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""IndexCompt — suivi des relevés de compteur électrique."""

import json
import math
import os
import random
from collections import defaultdict
from datetime import date, datetime, timedelta

import tkinter as tk
from tkinter import messagebox, simpledialog, ttk

try:
    import pandas as pd
except ImportError:
    pd = None


BASE_DIR = os.path.dirname(os.path.abspath(__file__))
ENERGY_FILE = os.path.join(BASE_DIR, "energy_data.json")
EXCEL_FILE = os.path.join(BASE_DIR, "export_consommation.xlsx")

DATE_FMTS = (
    "%Y-%m-%d %H:%M:%S",
    "%Y-%m-%d",
    "%d/%m/%Y %H:%M:%S",
    "%d/%m/%Y",
    "%d-%m-%Y",
)


class TimeManager(tk.Frame):
    def __init__(self, master):
        super().__init__(master)
        self.master = master

        self.readings = []
        self.price_kwh = tk.DoubleVar(value=0.20)
        self.monthly_goal = tk.DoubleVar(value=250.0)
        self.is_dark = True

        self.current_power = 400
        self.simulation_running = False
        self.simulation_job = None
        self.graph_frame = None

        self.load_energy()
        self.apply_colors()

        master.title("IndexCompt")
        master.geometry("1180x760")
        master.minsize(900, 600)

        self.pack(fill="both", expand=True)
        self.build_menu()
        self.build_ui()
        self.refresh_table()

        master.protocol("WM_DELETE_WINDOW", self.on_close)

    @staticmethod
    def _parse_date(raw):
        raw = (raw or "").strip()
        if not raw:
            raise ValueError("La date est vide.")

        for fmt in DATE_FMTS:
            try:
                return datetime.strptime(raw, fmt).date()
            except ValueError:
                continue

        raise ValueError("Date invalide. Utilise JJ/MM/AAAA.")

    @staticmethod
    def _iso(value):
        if isinstance(value, date):
            return value.strftime("%Y-%m-%d")

        raw = str(value or "").strip()
        if not raw:
            raise ValueError("La date est vide.")

        for fmt in DATE_FMTS:
            try:
                return datetime.strptime(raw, fmt).strftime("%Y-%m-%d")
            except ValueError:
                continue

        raise ValueError("Date invalide.")

    @staticmethod
    def _display_date(raw):
        try:
            return datetime.strptime(str(raw)[:10], "%Y-%m-%d").strftime(
                "%d/%m/%Y"
            )
        except (ValueError, TypeError):
            return str(raw)[:10]

    @staticmethod
    def _date_key(raw):
        try:
            return datetime.strptime(str(raw)[:10], "%Y-%m-%d").date()
        except (ValueError, TypeError):
            return date.min

    def _sort_readings(self):
        self.readings.sort(key=lambda item: self._date_key(item.get("date", "")))

    def _normalize_readings(self):
        cleaned = []

        for reading in self.readings:
            try:
                reading_date = self._iso(reading.get("date", ""))
                index = float(
                    str(reading.get("index", "")).replace(",", ".")
                )
            except (ValueError, TypeError, AttributeError):
                continue

            if not math.isfinite(index) or index < 0:
                continue

            cleaned.append({"date": reading_date, "index": index})

        self.readings = cleaned
        self._sort_readings()

    def validate_reading(self, reading, exclude_index=None):
        """Valide un relevé et la cohérence des index voisins."""
        reading_date = self._parse_date(reading.get("date", ""))
        try:
            index = float(str(reading.get("index", "")).replace(",", "."))
        except (ValueError, TypeError):
            raise ValueError("L’index doit être un nombre valide.")

        if not math.isfinite(index) or index < 0:
            raise ValueError("L’index doit être un nombre positif valide.")

        for position, existing in enumerate(self.readings):
            if position == exclude_index:
                continue

            existing_date = self._date_key(existing.get("date", ""))
            existing_index = float(existing.get("index", 0))

            if existing_date == reading_date:
                raise ValueError("Un relevé existe déjà à cette date.")

            if existing_date < reading_date and existing_index > index:
                raise ValueError(
                    "L’index saisi est inférieur à celui d’un relevé antérieur."
                )

            if existing_date > reading_date and existing_index < index:
                raise ValueError(
                    "L’index saisi est supérieur à celui d’un relevé ultérieur."
                )

        return {
            "date": reading_date.strftime("%Y-%m-%d"),
            "index": index,
        }

    def apply_colors(self):
        if self.is_dark:
            self.bg_color = "#1e1e1e"
            self.fg_color = "#e8e8e8"
            self.accent = "#4A90E2"
            self.entry_bg = "#2d2d2d"
            self.panel_bg = "#252525"
            self.tree_bg = "#2a2a2a"
            self.tree_fg = "#e8e8e8"
            self.sel_bg = "#3a5a80"
        else:
            self.bg_color = "#f4f4f4"
            self.fg_color = "#222222"
            self.accent = "#2E6DA4"
            self.entry_bg = "#ffffff"
            self.panel_bg = "#ffffff"
            self.tree_bg = "#ffffff"
            self.tree_fg = "#222222"
            self.sel_bg = "#c5d8f0"

        self.configure(bg=self.bg_color)
        self.master.configure(bg=self.bg_color)

        style = ttk.Style()
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        style.configure(
            "Treeview",
            background=self.tree_bg,
            foreground=self.tree_fg,
            fieldbackground=self.tree_bg,
            rowheight=24,
            font=("Arial", 10),
        )
        style.configure(
            "Treeview.Heading",
            background=self.accent,
            foreground="white",
            font=("Arial", 10, "bold"),
        )
        style.map("Treeview", background=[("selected", self.sel_bg)])
        style.configure(
            "TProgressbar",
            troughcolor=self.entry_bg,
            background="#27AE60",
        )

    def toggle_theme(self):
        self.is_dark = not self.is_dark
        if not self.save_energy():
            self.is_dark = not self.is_dark
            return

        for widget in self.winfo_children():
            widget.destroy()

        self.apply_colors()
        self.build_ui()
        self.refresh_table()

    def build_menu(self):
        menubar = tk.Menu(self.master)
        self.master.config(menu=menubar)

        fichier = tk.Menu(menubar, tearoff=0)
        fichier.add_command(
            label="Nouveau relevé", command=self.add_new_reading
        )
        fichier.add_command(
            label="Exporter vers Excel", command=self.export_to_excel
        )
        fichier.add_separator()
        fichier.add_command(label="Quitter", command=self.on_close)
        menubar.add_cascade(label="Fichier", menu=fichier)

        affichage = tk.Menu(menubar, tearoff=0)
        affichage.add_command(label="Graphiques", command=self.show_graphs)
        affichage.add_command(
            label="Thème clair / sombre", command=self.toggle_theme
        )
        menubar.add_cascade(label="Affichage", menu=affichage)

        aide = tk.Menu(menubar, tearoff=0)
        aide.add_command(label="Aide & tutoriel", command=self.show_help)
        menubar.add_cascade(label="Aide", menu=aide)

    def build_ui(self):
        left = tk.Frame(self, bg=self.panel_bg, width=300)
        left.pack(side="left", fill="y", padx=(12, 6), pady=12)
        left.pack_propagate(False)

        right_frame = tk.Frame(self, bg=self.bg_color)
        right_frame.pack(
            side="right", fill="both", expand=True, padx=(6, 12), pady=12
        )

        tk.Label(
            left,
            text="IndexCompt",
            font=("Arial", 16, "bold"),
            bg=self.panel_bg,
            fg=self.accent,
        ).pack(pady=(10, 16))

        tk.Label(
            left,
            text="Prix du kWh (EUR)",
            bg=self.panel_bg,
            fg=self.fg_color,
        ).pack()
        tk.Entry(
            left,
            textvariable=self.price_kwh,
            width=14,
            bg=self.entry_bg,
            fg="#ffffff" if self.is_dark else "#000000",
            insertbackground=self.fg_color,
            justify="center",
        ).pack(pady=(2, 4))
        tk.Button(
            left,
            text="Appliquer le prix",
            bg=self.accent,
            fg="white",
            command=self.apply_price,
        ).pack(pady=(0, 12))

        tk.Label(
            left,
            text="Objectif mensuel (kWh)",
            bg=self.panel_bg,
            fg=self.fg_color,
        ).pack()
        tk.Entry(
            left,
            textvariable=self.monthly_goal,
            width=14,
            bg=self.entry_bg,
            fg="#ffffff" if self.is_dark else "#000000",
            insertbackground=self.fg_color,
            justify="center",
        ).pack(pady=(2, 4))
        tk.Button(
            left,
            text="Valider l’objectif",
            bg="#8E44AD",
            fg="white",
            command=self.apply_goal,
        ).pack(pady=(0, 10))

        self.progress_bar = ttk.Progressbar(
            left, length=240, mode="determinate", maximum=100
        )
        self.progress_bar.pack(pady=4)

        self.progress_label = tk.Label(
            left,
            text="Objectif : --",
            bg=self.panel_bg,
            fg=self.fg_color,
            font=("Arial", 10, "bold"),
        )
        self.progress_label.pack(pady=(0, 8))

        self.twelve_label = tk.Label(
            left,
            text="12 mois : --",
            bg=self.panel_bg,
            fg=self.fg_color,
            font=("Arial", 10),
            wraplength=260,
            justify="center",
        )
        self.twelve_label.pack(pady=(0, 12))

        self.power_label = tk.Label(
            left,
            text="Puissance simulée : %s W" % self.current_power,
            bg=self.panel_bg,
            fg=self.fg_color,
            font=("Arial", 10),
        )
        self.power_label.pack(pady=4)

        sim_bg = "#E74C3C" if self.simulation_running else "#27AE60"
        sim_txt = (
            "Arrêter la simulation"
            if self.simulation_running
            else "Démarrer la simulation"
        )
        self.sim_button = tk.Button(
            left,
            text=sim_txt,
            bg=sim_bg,
            fg="white",
            width=22,
            command=self.toggle_simulation,
        )
        self.sim_button.pack(pady=8)

        tk.Button(
            left,
            text="Nouveau relevé",
            bg="#27AE60",
            fg="white",
            width=22,
            command=self.add_new_reading,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Modifier",
            bg="#4A90E2",
            fg="white",
            width=22,
            command=self.modify_reading,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Supprimer",
            bg="#E67E22",
            fg="white",
            width=22,
            command=self.delete_reading,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Graphiques",
            bg="#16A085",
            fg="white",
            width=22,
            command=self.show_graphs,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Exporter Excel",
            bg="#2980B9",
            fg="white",
            width=22,
            command=self.export_to_excel,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Thème clair / sombre",
            bg="#7F8C8D",
            fg="white",
            width=22,
            command=self.toggle_theme,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Réinitialiser le suivi",
            bg="#C0392B",
            fg="white",
            width=22,
            command=self.reset_tracking,
        ).pack(pady=3)
        tk.Button(
            left,
            text="Tout effacer",
            bg="#922B21",
            fg="white",
            width=22,
            command=self.clear_all,
        ).pack(pady=3)

        scroll = ttk.Scrollbar(right_frame, orient="vertical")
        self.read_tree = ttk.Treeview(
            right_frame,
            columns=("date", "index", "conso", "cost"),
            show="headings",
            yscrollcommand=scroll.set,
        )
        scroll.config(command=self.read_tree.yview)

        self.read_tree.heading("date", text="Date")
        self.read_tree.heading("index", text="Index (kWh)")
        self.read_tree.heading("conso", text="Consommation")
        self.read_tree.heading("cost", text="Coût (EUR)")
        self.read_tree.column("date", width=220, anchor="center")
        self.read_tree.column("index", width=160, anchor="center")
        self.read_tree.column("conso", width=180, anchor="center")
        self.read_tree.column("cost", width=140, anchor="center")

        self.read_tree.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")

        self.read_tree.tag_configure(
            "year",
            background="#3d3d3d" if self.is_dark else "#dfe6ee",
            foreground=self.accent,
            font=("Arial", 10, "bold"),
        )

        self.total_label = tk.Label(
            self,
            text="",
            bg=self.bg_color,
            fg=self.fg_color,
            font=("Arial", 13, "bold"),
        )
        self.total_label.pack(side="bottom", pady=(0, 10))

    def show_graphs(self):
        if len(self.readings) < 2:
            messagebox.showinfo(
                "Information",
                "Il faut au moins 2 relevés pour afficher les graphiques.",
            )
            return

        try:
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
            import matplotlib.pyplot as plt  # noqa: F401
        except ImportError:
            messagebox.showerror(
                "Erreur",
                "Matplotlib n’est pas installé.\n"
                "Installe-le avec : pip install matplotlib",
            )
            return

        graph_win = tk.Toplevel(self)
        graph_win.title("Graphiques de consommation")
        graph_win.geometry("1100x750")
        graph_win.configure(bg=self.bg_color)

        tk.Label(
            graph_win,
            text="Graphiques de consommation",
            font=("Arial", 16, "bold"),
            bg=self.bg_color,
            fg=self.accent,
        ).pack(pady=15)

        btn_frame = tk.Frame(graph_win, bg=self.bg_color)
        btn_frame.pack(pady=10)

        tk.Button(
            btn_frame,
            text="Courbe cumulée",
            bg="#4A90E2",
            fg="white",
            width=22,
            command=lambda: self.plot_graph(graph_win, "cumul"),
        ).pack(side="left", padx=15)

        tk.Button(
            btn_frame,
            text="Barres mensuelles",
            bg="#27AE60",
            fg="white",
            width=22,
            command=lambda: self.plot_graph(graph_win, "monthly"),
        ).pack(side="left", padx=15)

        graph_win.graph_frame = tk.Frame(graph_win, bg=self.bg_color)
        graph_win.graph_frame.pack(
            fill="both", expand=True, padx=25, pady=15
        )
        self.graph_frame = graph_win.graph_frame
        self.plot_graph(graph_win, "cumul")

    def plot_graph(self, window, graph_type):
        frame = getattr(window, "graph_frame", None) or self.graph_frame
        if frame is None:
            return

        for widget in frame.winfo_children():
            widget.destroy()

        try:
            import matplotlib as mpl
            import matplotlib.pyplot as plt
            from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

            mpl.rcParams["font.family"] = "DejaVu Sans"
            mpl.rcParams["font.size"] = 9
            mpl.rcParams["axes.unicode_minus"] = False
        except Exception as error:
            messagebox.showerror(
                "Erreur", "Impossible de charger matplotlib :\n%s" % error
            )
            return

        try:
            fig = plt.Figure(
                figsize=(9.5, 6), dpi=90, facecolor=self.bg_color
            )
            ax = fig.add_subplot(111)
            ax.set_facecolor("#2a2a2a" if self.is_dark else "#f0f0f0")

            if graph_type == "cumul":
                labels = [
                    self._display_date(item["date"])
                    for item in self.readings
                ]
                indices = [item["index"] for item in self.readings]
                positions = list(range(len(labels)))

                ax.plot(
                    positions,
                    indices,
                    color=self.accent,
                    linewidth=2,
                    marker="o",
                    markersize=3,
                )
                step = max(1, len(positions) // 8)
                ticks = positions[::step]
                ax.set_xticks(ticks)
                ax.set_xticklabels(
                    [labels[pos] for pos in ticks],
                    rotation=30,
                    ha="right",
                )
                ax.set_title(
                    "Évolution de l’index cumulé", color=self.fg_color
                )
                ax.set_ylabel("Index (kWh)", color=self.fg_color)
                ax.grid(True, alpha=0.3)
            else:
                monthly = defaultdict(float)

                for position in range(1, len(self.readings)):
                    current = self.readings[position]
                    previous = self.readings[position - 1]
                    consumption = max(
                        current["index"] - previous["index"], 0
                    )
                    month = current["date"][:7]
                    monthly[month] += consumption

                months = sorted(monthly)
                values = [monthly[month] for month in months]

                ax.bar(range(len(months)), values, color="#4A90E2")
                ax.set_xticks(range(len(months)))
                ax.set_xticklabels(months, rotation=30, ha="right")
                ax.set_title(
                    "Consommation estimée par mois du relevé (kWh)",
                    color=self.fg_color,
                )
                ax.set_ylabel("kWh", color=self.fg_color)

            ax.tick_params(colors=self.fg_color, labelsize=8)
            for spine in ax.spines.values():
                spine.set_color(
                    "#555555" if self.is_dark else "#cccccc"
                )

            canvas = FigureCanvasTkAgg(fig, master=frame)
            canvas.draw()
            canvas.get_tk_widget().pack(fill="both", expand=True)
        except Exception as error:
            tk.Label(
                frame,
                text="Impossible d’afficher le graphique.\n%s"
                     % str(error)[:120],
                bg=self.bg_color,
                fg="#E74C3C",
                font=("Arial", 11),
                justify="center",
            ).pack(expand=True)

    def add_new_reading(self):
        win = tk.Toplevel(self)
        win.title("Nouveau relevé")
        win.geometry("420x280")
        win.configure(bg=self.bg_color)
        win.resizable(False, False)
        win.transient(self.master)

        tk.Label(
            win,
            text="Date (JJ/MM/AAAA) :",
            bg=self.bg_color,
            fg=self.fg_color,
            font=("Arial", 11),
        ).pack(pady=(20, 5))

        date_entry = tk.Entry(
            win,
            width=25,
            bg=self.entry_bg,
            fg="#ffffff" if self.is_dark else "#000000",
            font=("Arial", 11),
        )
        date_entry.insert(0, datetime.now().strftime("%d/%m/%Y"))
        date_entry.pack(pady=5)

        tk.Label(
            win,
            text="Index du compteur (kWh) :",
            bg=self.bg_color,
            fg=self.fg_color,
            font=("Arial", 11),
        ).pack(pady=(15, 5))

        index_entry = tk.Entry(
            win,
            width=25,
            bg=self.entry_bg,
            fg="#ffffff" if self.is_dark else "#000000",
            font=("Arial", 11),
        )
        index_entry.pack(pady=5)
        index_entry.focus_set()

        def save():
            try:
                reading = self.validate_reading(
                    {"date": date_entry.get(), "index": index_entry.get()}
                )
            except ValueError as error:
                messagebox.showerror("Erreur", str(error), parent=win)
                return

            self.readings.append(reading)
            self._sort_readings()

            if self.save_energy():
                win.destroy()
                messagebox.showinfo(
                    "Succès", "Relevé ajouté avec succès."
                )

        tk.Button(
            win,
            text="Ajouter le relevé",
            bg="#27AE60",
            fg="white",
            font=("Arial", 11, "bold"),
            command=save,
        ).pack(pady=25)

    def show_help(self):
        help_win = tk.Toplevel(self)
        help_win.title("Aide & tutoriel")
        help_win.geometry("900x700")
        help_win.configure(bg=self.bg_color)

        text = tk.Text(
            help_win,
            bg="#2a2a2a" if self.is_dark else "#ffffff",
            fg="#e0e0e0" if self.is_dark else "#333333",
            font=("Consolas", 11),
            wrap="word",
            padx=15,
            pady=15,
        )
        text.pack(fill="both", expand=True, padx=20, pady=20)

        help_text = (
            "IndexCompt — Tutoriel\n\n"
            "- Menu Fichier → Nouveau relevé\n"
            "- Boutons Modifier / Supprimer à gauche\n"
            "- La simulation est fictive : elle ne lit pas le compteur réel.\n"
            "- Le thème clair/sombre est sauvegardé automatiquement.\n"
            "- Les graphiques montrent l’index et la consommation estimée.\n"
            "- L’objectif concerne le mois civil en cours.\n"
            "- La consommation d’un intervalle est attribuée au mois "
            "du relevé le plus récent.\n"
            "- Export Excel : installer pandas et openpyxl.\n"
        )
        text.insert("1.0", help_text)
        text.config(state="disabled")

        tk.Button(
            help_win,
            text="Fermer",
            bg="#E74C3C",
            fg="white",
            command=help_win.destroy,
        ).pack(pady=10)

    def export_to_excel(self):
        if not self.readings:
            messagebox.showinfo(
                "Information", "Aucune donnée à exporter."
            )
            return

        if pd is None:
            messagebox.showerror(
                "Erreur",
                "pandas n’est pas installé.\n"
                "Installe-le avec : pip install pandas openpyxl",
            )
            return

        try:
            price = float(
                str(self.price_kwh.get()).replace(",", ".")
            )
            if not math.isfinite(price) or price < 0:
                raise ValueError("Prix invalide.")

            rows = []
            for position, reading in enumerate(self.readings):
                row = {
                    "Date": self._display_date(reading["date"]),
                    "Index (kWh)": round(reading["index"], 3),
                }

                if position > 0:
                    consumption = max(
                        reading["index"]
                        - self.readings[position - 1]["index"],
                        0,
                        )
                    row["Consommation estimée (kWh)"] = round(
                        consumption, 3
                    )
                    row["Coût estimé (EUR)"] = round(
                        consumption * price, 2
                    )
                else:
                    row["Consommation estimée (kWh)"] = 0
                    row["Coût estimé (EUR)"] = 0

                rows.append(row)

            pd.DataFrame(rows).to_excel(EXCEL_FILE, index=False)
            messagebox.showinfo(
                "Succès", "Exporté dans %s" % EXCEL_FILE
            )
        except (OSError, ValueError, TypeError, tk.TclError) as error:
            messagebox.showerror("Erreur", str(error))
        except Exception as error:
            messagebox.showerror(
                "Erreur",
                "Échec de l’export Excel : %s" % error,
                )

    def toggle_simulation(self):
        if self.simulation_running:
            self.simulation_running = False
            self.sim_button.config(
                text="Démarrer la simulation", bg="#27AE60"
            )
            if self.simulation_job is not None:
                try:
                    self.after_cancel(self.simulation_job)
                except tk.TclError:
                    pass
                self.simulation_job = None
            return

        self.simulation_running = True
        self.sim_button.config(
            text="Arrêter la simulation", bg="#E74C3C"
        )
        self.run_simulation()

    def run_simulation(self):
        if not self.simulation_running:
            return

        hour = datetime.now().hour
        if hour >= 22 or hour <= 6:
            target = random.randint(120, 450)
        else:
            target = random.randint(300, 2800)
            if random.random() < 0.25:
                target = random.randint(2200, 4200)

        self.current_power = int(
            self.current_power * 0.7 + target * 0.3
        )
        self.power_label.config(
            text="Puissance simulée : %s W" % self.current_power
        )

        today = datetime.now().strftime("%Y-%m-%d")
        added = self.current_power * (4 / 3_600_000.0)

        if self.readings and self.readings[-1]["date"] >= today:
            self.readings[-1]["index"] = round(
                self.readings[-1]["index"]
                + added
                + random.uniform(0.001, 0.018),
                3,
                )
        else:
            last_index = (
                self.readings[-1]["index"] if self.readings else 1200.0
            )
            self.readings.append(
                {
                    "date": today,
                    "index": round(
                        last_index
                        + added
                        + random.uniform(0.001, 0.018),
                        3,
                        ),
                }
            )

        self._sort_readings()
        if self.save_energy():
            self.refresh_table()

        self.simulation_job = self.after(4000, self.run_simulation)

    def load_energy(self):
        if not os.path.exists(ENERGY_FILE):
            return

        try:
            with open(ENERGY_FILE, "r", encoding="utf-8") as file:
                data = json.load(file)

            if not isinstance(data, dict):
                raise ValueError("Le contenu JSON n’est pas un objet.")

            readings = data.get("readings", [])
            if not isinstance(readings, list):
                raise ValueError("La liste des relevés est invalide.")

            self.readings = readings

            try:
                price = float(data.get("price_kwh", 0.20))
                if math.isfinite(price) and price >= 0:
                    self.price_kwh.set(price)
            except (ValueError, TypeError, tk.TclError):
                pass

            try:
                goal = float(data.get("monthly_goal", 250.0))
                if math.isfinite(goal) and goal >= 0:
                    self.monthly_goal.set(goal)
            except (ValueError, TypeError, tk.TclError):
                pass

            self.is_dark = bool(data.get("is_dark", True))
            self._normalize_readings()

        except (OSError, json.JSONDecodeError, ValueError, TypeError) as error:
            self.readings = []
            messagebox.showwarning(
                "Données illisibles",
                "Le fichier de données n’a pas pu être chargé :\n%s\n\n"
                "Le fichier original n’a pas été supprimé."
                % error,
                )

    def save_energy(self):
        try:
            price = float(
                str(self.price_kwh.get()).replace(",", ".")
            )
            goal = float(
                str(self.monthly_goal.get()).replace(",", ".")
            )

            if not math.isfinite(price) or price < 0:
                raise ValueError
            if not math.isfinite(goal) or goal < 0:
                raise ValueError
        except (tk.TclError, ValueError, TypeError):
            messagebox.showerror(
                "Erreur",
                "Le prix et l’objectif doivent être des nombres "
                "positifs valides.",
            )
            return False

        self._normalize_readings()

        data = {
            "readings": self.readings,
            "price_kwh": round(price, 4),
            "monthly_goal": goal,
            "is_dark": self.is_dark,
        }

        temporary_file = ENERGY_FILE + ".tmp"
        try:
            with open(temporary_file, "w", encoding="utf-8") as file:
                json.dump(data, file, ensure_ascii=False, indent=2)
            os.replace(temporary_file, ENERGY_FILE)
        except OSError as error:
            try:
                if os.path.exists(temporary_file):
                    os.remove(temporary_file)
            except OSError:
                pass

            messagebox.showerror(
                "Erreur", "Sauvegarde impossible :\n%s" % error
            )
            return False

        self.refresh_table()
        return True

    def apply_price(self):
        if self.save_energy():
            messagebox.showinfo(
                "Succès",
                "Prix du kWh mis à jour : %.4f EUR"
                % float(self.price_kwh.get()),
                )

    def apply_goal(self):
        if self.save_energy():
            messagebox.showinfo(
                "Succès",
                "Objectif mis à jour : %.1f kWh"
                % float(self.monthly_goal.get()),
                )

    def get_real_index(self, selected_item):
        item_id = str(selected_item)
        if item_id.startswith("r") and item_id[1:].isdigit():
            index = int(item_id[1:])
            if 0 <= index < len(self.readings):
                return index
        return None

    def modify_reading(self):
        selection = self.read_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Attention", "Sélectionne une ligne à modifier."
            )
            return

        real_index = self.get_real_index(selection[0])
        if real_index is None:
            messagebox.showwarning(
                "Attention",
                "Sélectionne un relevé, pas un en-tête d’année.",
            )
            return

        reading = self.readings[real_index]

        win = tk.Toplevel(self)
        win.title("Modifier le relevé")
        win.geometry("420x300")
        win.configure(bg=self.bg_color)
        win.transient(self.master)
        win.grab_set()

        tk.Label(
            win,
            text="Date (JJ/MM/AAAA) :",
            bg=self.bg_color,
            fg=self.fg_color,
        ).pack(pady=10)

        date_entry = tk.Entry(
            win,
            width=25,
            bg=self.entry_bg,
            fg="#ffffff" if self.is_dark else "#000000",
        )
        date_entry.insert(0, self._display_date(reading["date"]))
        date_entry.pack(pady=5)

        tk.Label(
            win,
            text="Index (kWh) :",
            bg=self.bg_color,
            fg=self.fg_color,
        ).pack(pady=10)

        index_entry = tk.Entry(
            win,
            width=25,
            bg=self.entry_bg,
            fg="#ffffff" if self.is_dark else "#000000",
        )
        index_entry.insert(0, "%.3f" % reading["index"])
        index_entry.pack(pady=5)
        index_entry.focus_set()
        index_entry.select_range(0, "end")

        def save_modification():
            try:
                updated = self.validate_reading(
                    {
                        "date": date_entry.get(),
                        "index": index_entry.get(),
                    },
                    exclude_index=real_index,
                )
            except ValueError as error:
                messagebox.showerror("Erreur", str(error), parent=win)
                return

            old_reading = self.readings[real_index]
            self.readings[real_index] = updated
            self._sort_readings()

            if self.save_energy():
                win.destroy()
                messagebox.showinfo("Succès", "Relevé modifié.")
            else:
                self.readings.append(old_reading)
                self._sort_readings()
                self.refresh_table()

        tk.Button(
            win,
            text="Enregistrer",
            bg="#4A90E2",
            fg="white",
            command=save_modification,
        ).pack(pady=20)
        win.bind("<Return>", lambda event: save_modification())

    def delete_reading(self):
        selection = self.read_tree.selection()
        if not selection:
            messagebox.showwarning(
                "Attention", "Sélectionne une ligne."
            )
            return

        real_index = self.get_real_index(selection[0])
        if real_index is None:
            messagebox.showwarning(
                "Attention",
                "Sélectionne un relevé, pas un en-tête d’année.",
            )
            return

        if not messagebox.askyesno(
                "Supprimer ?", "Supprimer ce relevé ?"
        ):
            return

        removed = self.readings.pop(real_index)
        if not self.save_energy():
            self.readings.append(removed)
            self._sort_readings()
            self.refresh_table()
            return

        messagebox.showinfo("Succès", "Relevé supprimé.")

    def clear_all(self):
        if not messagebox.askyesno(
                "Effacer tout ?", "Veux-tu vraiment effacer tous les relevés ?"
        ):
            return

        old_readings = self.readings[:]
        self.readings = []
        if not self.save_energy():
            self.readings = old_readings
            self.refresh_table()

    def reset_tracking(self):
        if not messagebox.askyesno(
                "Réinitialiser ?", "Réinitialiser le suivi ?"
        ):
            return

        start_index = simpledialog.askfloat(
            "Index de départ",
            "Nouvel index de départ :",
            initialvalue=1200.0,
            minvalue=0.0,
        )
        if start_index is None:
            return

        old_readings = self.readings[:]
        self.readings = [
            {
                "date": datetime.now().strftime("%Y-%m-%d"),
                "index": start_index,
            }
        ]

        if not self.save_energy():
            self.readings = old_readings
            self.refresh_table()

    def refresh_table(self):
        if not hasattr(self, "read_tree"):
            return

        for item in self.read_tree.get_children():
            self.read_tree.delete(item)

        try:
            price = float(
                str(self.price_kwh.get()).replace(",", ".")
            )
            if not math.isfinite(price) or price < 0:
                price = 0.0
        except (tk.TclError, ValueError, TypeError):
            price = 0.0

        try:
            goal = float(
                str(self.monthly_goal.get()).replace(",", ".")
            )
            if not math.isfinite(goal) or goal < 0:
                goal = 0.0
        except (tk.TclError, ValueError, TypeError):
            goal = 0.0

        total_consumption = 0.0
        total_cost = 0.0
        twelve_consumption = 0.0
        twelve_cost = 0.0
        month_consumption = 0.0

        now = datetime.now()
        cutoff = (now - timedelta(days=365)).date()
        current_month = now.strftime("%Y-%m")
        current_year = None

        for position, reading in enumerate(self.readings):
            year = reading["date"][:4]

            if year != current_year:
                current_year = year
                self.read_tree.insert(
                    "",
                    "end",
                    iid="y%s" % year,
                    values=("— Année %s —" % year, "", "", ""),
                    tags=("year",),
                )

            consumption_text = ""
            cost_text = ""

            if position > 0:
                consumption = max(
                    reading["index"]
                    - self.readings[position - 1]["index"],
                    0,
                    )
                cost = consumption * price

                consumption_text = "%.3f kWh" % consumption
                cost_text = "%.2f EUR" % cost
                total_consumption += consumption
                total_cost += cost

                # Estimation : l'intervalle est compté au mois du relevé courant.
                if reading["date"][:7] == current_month:
                    month_consumption += consumption

                if self._date_key(reading["date"]) >= cutoff:
                    twelve_consumption += consumption
                    twelve_cost += cost

            self.read_tree.insert(
                "",
                "end",
                iid="r%s" % position,
                values=(
                    self._display_date(reading["date"]),
                    "%.3f" % reading["index"],
                    consumption_text,
                    cost_text,
                ),
            )

        if goal > 0:
            percentage = min((month_consumption / goal) * 100, 100)
            self.progress_bar["value"] = percentage
            color = (
                "#27AE60"
                if percentage < 70
                else "#F39C12"
                if percentage < 90
                else "#E74C3C"
            )
            self.progress_label.config(
                text="Objectif du mois : %.1f%% (%.1f/%.1f)"
                     % (percentage, month_consumption, goal),
                fg=color,
            )
        else:
            self.progress_bar["value"] = 0
            self.progress_label.config(
                text="Objectif : --", fg=self.fg_color
            )

        self.total_label.config(
            text="Total estimé : %.3f kWh — %.2f EUR"
                 % (total_consumption, total_cost)
        )
        self.twelve_label.config(
            text="12 mois estimés : %.3f kWh — %.2f EUR"
                 % (twelve_consumption, twelve_cost)
        )

    def on_close(self):
        self.simulation_running = False

        if self.simulation_job is not None:
            try:
                self.after_cancel(self.simulation_job)
            except tk.TclError:
                pass
            self.simulation_job = None

        self.master.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = TimeManager(root)
    root.mainloop()