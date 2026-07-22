#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convertisseur_gui.py — Interface graphique du convertisseur CAO/DAO par lot.

Fenêtre tkinter (livré avec Python) pour :
  - choisir des fichiers via le gestionnaire de fichiers (multi-sélection) ;
  - cocher les formats voulus selon la nature (Pièces / Assemblages /
    Mises en plan), chaque nature ayant ses propres cases ;
  - ajouter un préfixe et/ou un suffixe aux noms de fichiers ;
  - choisir la destination (à côté des fichiers, ou un dossier personnalisé) ;
  - suivre l'avancement (barre + temps restant estimé) et le journal ;
  - obtenir un bilan en fin de conversion (et la liste des échecs).

La conversion tourne dans un thread séparé (avec CoInitialize pour COM).
"""

import os
import queue
import threading
import time

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import convertisseur_cao as engine


_ALL_EXTS = sorted(engine.ROUTES.keys())
_FILE_TYPES = [
    ("Fichiers CAO/DAO", " ".join("*" + e for e in _ALL_EXTS)),
    ("SolidWorks", "*.sldprt *.sldasm *.slddrw"),
    ("CATIA V5", "*.CATPart *.CATProduct *.CATDrawing"),
    ("Inventor", "*.ipt *.iam *.idw"),
    ("Creo", "*.prt *.asm *.drw"),
    ("Tous les fichiers", "*.*"),
]


def _fmt_duration(seconds):
    if seconds is None or seconds < 0:
        return "—"
    seconds = int(round(seconds))
    m, s = divmod(seconds, 60)
    if m >= 60:
        h, m = divmod(m, 60)
        return "%dh%02dm%02ds" % (h, m, s)
    if m:
        return "%dmin %02ds" % (m, s)
    return "%ds" % s


class App(ttk.Frame):
    def __init__(self, master, controller=None):
        super().__init__(master, padding=10)
        self.master = master
        self.controller = controller
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        self.files = []
        # format_vars[kind][target] -> BooleanVar (STEP de « Pièces » et de
        # « Assemblages » sont indépendants)
        self.format_vars = {k: {} for k in engine.KIND_ORDER}
        self.frames = {}
        self._queue = queue.Queue()
        self._worker = None
        self._stop = False
        self._start_time = None
        self._total = 0

        self._build()
        self._poll_queue()
        self._refresh_state()

    # ------------------------------------------------------------------ UI
    def _build(self):
        r = 0
        if self.controller is not None:
            header = ttk.Frame(self)
            header.grid(row=r, column=0, sticky="ew", pady=(0, 6))
            self.back_btn = ttk.Button(header, text="‹ Accueil",
                                       command=self._go_home)
            self.back_btn.pack(side="left")
            ttk.Label(header, text="   Conversion CAO/DAO",
                      font=("", 11, "bold")).pack(side="left")
            r += 1

        bar = ttk.Frame(self)
        bar.grid(row=r, column=0, sticky="ew")
        ttk.Button(bar, text="Ajouter des fichiers…",
                   command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Ajouter un dossier…",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Retirer la sélection",
                   command=self.remove_selected).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Vider la liste",
                   command=self.clear_files).pack(side="left", padx=(6, 0))
        r += 1

        ttk.Label(self, text="Fichiers à convertir :").grid(
            row=r, column=0, sticky="w", pady=(8, 2))
        r += 1

        lf = ttk.Frame(self)
        lf.grid(row=r, column=0, sticky="nsew")
        lf.columnconfigure(0, weight=1)
        self.rowconfigure(r, weight=1)
        self.listbox = tk.Listbox(lf, height=7, selectmode="extended",
                                  activestyle="none")
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(lf, orient="vertical", command=self.listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=sb.set)
        r += 1

        self.count_label = ttk.Label(self, text="")
        self.count_label.grid(row=r, column=0, sticky="w", pady=(2, 6))
        r += 1

        # --- Formats (3 cadres) --------------------------------------
        fmt = ttk.Frame(self)
        fmt.grid(row=r, column=0, sticky="ew", pady=(0, 6))
        for i, kind in enumerate(engine.KIND_ORDER):
            fmt.columnconfigure(i, weight=1)
            frame = ttk.LabelFrame(fmt, text=engine.KIND_LABELS[kind], padding=8)
            frame.grid(row=0, column=i, sticky="nsew",
                       padx=(0 if i == 0 else 4, 0))
            self.frames[kind] = frame
            for t in engine.TARGETS[kind]:
                var = tk.BooleanVar(value=(t != "pdf"))  # PDF décoché par défaut
                self.format_vars[kind][t] = var
                ttk.Checkbutton(frame, text=engine.TARGET_LABELS.get(t, t),
                                variable=var).pack(anchor="w")
        r += 1

        # --- Nommage (préfixe / suffixe) -----------------------------
        naming = ttk.LabelFrame(self, text="Nommage des fichiers", padding=8)
        naming.grid(row=r, column=0, sticky="ew", pady=(0, 6))
        naming.columnconfigure(1, weight=1)
        naming.columnconfigure(3, weight=1)
        ttk.Label(naming, text="Préfixe :").grid(row=0, column=0, sticky="w")
        self.prefix_var = tk.StringVar()
        self.prefix_var.trace_add("write", lambda *_: self._update_naming_preview())
        ttk.Entry(naming, textvariable=self.prefix_var, width=16).grid(
            row=0, column=1, sticky="ew", padx=(4, 10))
        ttk.Label(naming, text="Suffixe :").grid(row=0, column=2, sticky="w")
        self.suffix_var = tk.StringVar()
        self.suffix_var.trace_add("write", lambda *_: self._update_naming_preview())
        ttk.Entry(naming, textvariable=self.suffix_var, width=16).grid(
            row=0, column=3, sticky="ew", padx=(4, 0))
        self.naming_preview = ttk.Label(naming, text="", foreground="#555")
        self.naming_preview.grid(row=1, column=0, columnspan=4, sticky="w",
                                 pady=(4, 0))
        r += 1

        # --- Destination ---------------------------------------------
        dest = ttk.LabelFrame(self, text="Destination", padding=8)
        dest.grid(row=r, column=0, sticky="ew", pady=(0, 6))
        dest.columnconfigure(1, weight=1)
        self.dest_mode = tk.StringVar(value="beside")
        ttk.Radiobutton(
            dest, text="À côté de chaque fichier (sous-dossier « Export »)",
            variable=self.dest_mode, value="beside",
            command=self._refresh_state).grid(row=0, column=0, columnspan=3,
                                              sticky="w")
        ttk.Radiobutton(
            dest, text="Dossier personnalisé :", variable=self.dest_mode,
            value="custom", command=self._refresh_state).grid(
            row=1, column=0, sticky="w")
        self.dest_entry = ttk.Entry(dest)
        self.dest_entry.grid(row=1, column=1, sticky="ew", padx=6)
        self.dest_browse = ttk.Button(dest, text="Parcourir…",
                                      command=self.choose_dest)
        self.dest_browse.grid(row=1, column=2)
        self.subfolder_note = ttk.Label(
            dest, text="Plusieurs formats cochés seront rangés dans des "
                       "sous-dossiers (STEP/, STL/, PDF/…).",
            foreground="#555")
        self.subfolder_note.grid(row=2, column=0, columnspan=3, sticky="w",
                                 pady=(4, 0))
        r += 1

        # --- Progression + action ------------------------------------
        act = ttk.Frame(self)
        act.grid(row=r, column=0, sticky="ew", pady=(0, 6))
        act.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(act, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.convert_btn = ttk.Button(act, text="Convertir",
                                      command=self.start_conversion)
        self.convert_btn.grid(row=0, column=1)
        self.eta_label = ttk.Label(act, text="")
        self.eta_label.grid(row=1, column=0, columnspan=2, sticky="w",
                            pady=(3, 0))
        r += 1

        # --- Journal --------------------------------------------------
        ttk.Label(self, text="Journal :").grid(row=r, column=0, sticky="w")
        r += 1
        logf = ttk.Frame(self)
        logf.grid(row=r, column=0, sticky="nsew")
        logf.columnconfigure(0, weight=1)
        logf.rowconfigure(0, weight=1)
        self.rowconfigure(r, weight=2)
        self.log_text = tk.Text(logf, height=9, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        lsb = ttk.Scrollbar(logf, orient="vertical", command=self.log_text.yview)
        lsb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=lsb.set)

        self._update_naming_preview()

    def _go_home(self):
        busy = self._worker is not None and self._worker.is_alive()
        if busy:
            messagebox.showinfo("Conversion en cours",
                                "Attendez la fin de la conversion.")
            return
        if self.controller is not None:
            self.controller.show_home()

    # ------------------------------------------------------------- actions
    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="Choisir des fichiers CAO/DAO", filetypes=_FILE_TYPES)
        self._add_paths(paths)

    def add_folder(self):
        folder = filedialog.askdirectory(title="Choisir un dossier")
        if folder:
            self._add_paths(engine.collect_files([folder]))

    def _add_paths(self, paths):
        added = 0
        existing = set(self.files)
        for p in paths:
            p = os.path.abspath(p)
            if p in existing:
                continue
            route = engine.route_for(p)
            if route is None:
                self.log("Ignoré (format non géré) : %s" % os.path.basename(p))
                continue
            self.files.append(p)
            existing.add(p)
            soft, kind = route
            self.listbox.insert("end", "  [%s · %s]  %s" % (
                engine.SOFTWARE_LABELS.get(soft, soft),
                engine.KIND_LABELS[kind].rstrip("s").lower(), p))
            added += 1
        if added:
            self._refresh_state()

    def remove_selected(self):
        for i in reversed(self.listbox.curselection()):
            self.listbox.delete(i)
            del self.files[i]
        self._refresh_state()

    def clear_files(self):
        self.listbox.delete(0, "end")
        self.files = []
        self._refresh_state()

    def choose_dest(self):
        folder = filedialog.askdirectory(title="Dossier de destination")
        if folder:
            self.dest_entry.delete(0, "end")
            self.dest_entry.insert(0, folder)
            self.dest_mode.set("custom")
            self._refresh_state()

    # -------------------------------------------------------------- helpers
    def _kinds_present(self):
        present = set()
        for p in self.files:
            route = engine.route_for(p)
            if route:
                present.add(route[1])
        return present

    def _set_frame_state(self, frame, enabled):
        state = "normal" if enabled else "disabled"
        for child in frame.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass

    def _selection(self):
        """dict nature -> liste de formats cochés."""
        sel = {}
        for kind in engine.KIND_ORDER:
            sel[kind] = [t for t in engine.TARGETS[kind]
                         if self.format_vars[kind][t].get()]
        return sel

    def _distinct_selected_formats(self):
        present = self._kinds_present()
        sel = self._selection()
        distinct = set()
        for kind in present:
            distinct.update(sel.get(kind, []))
        return distinct

    def _update_naming_preview(self):
        pre = engine.sanitize_affix(self.prefix_var.get())
        suf = engine.sanitize_affix(self.suffix_var.get())
        self.naming_preview.configure(
            text="Exemple : %sma_piece%s.step" % (pre, suf))

    def _refresh_state(self):
        present = self._kinds_present()
        for kind in engine.KIND_ORDER:
            self._set_frame_state(self.frames[kind], kind in present)

        custom = self.dest_mode.get() == "custom"
        self.dest_entry.configure(state="normal" if custom else "disabled")
        self.dest_browse.configure(state="normal" if custom else "disabled")

        counts = {k: 0 for k in engine.KIND_ORDER}
        for p in self.files:
            route = engine.route_for(p)
            if route:
                counts[route[1]] += 1
        self.count_label.configure(
            text="%d fichier(s) : %d pièce(s), %d assemblage(s), %d mise(s) en plan"
            % (len(self.files), counts["part"], counts["assembly"],
               counts["drawing"]))

        busy = self._worker is not None and self._worker.is_alive()
        self.convert_btn.configure(
            state="disabled" if busy or not self.files else "normal")

    # ----------------------------------------------------------- conversion
    def start_conversion(self):
        if not self.files:
            return
        present = self._kinds_present()
        sel = self._selection()

        # nature présente mais sans aucun format coché -> avertir
        ignored = [engine.KIND_LABELS[k] for k in present if not sel.get(k)]
        if len(ignored) == len(present):
            messagebox.showwarning(
                "Aucun format", "Cochez au moins un format de sortie.")
            return
        if ignored:
            if not messagebox.askyesno(
                    "Formats manquants",
                    "Aucun format coché pour : %s.\n"
                    "Ces fichiers seront ignorés. Continuer ?"
                    % ", ".join(ignored)):
                return

        dest_dir = None
        if self.dest_mode.get() == "custom":
            dest_dir = self.dest_entry.get().strip()
            if not dest_dir:
                messagebox.showwarning("Destination",
                                       "Indiquez un dossier de destination.")
                return

        subfolders = len(self._distinct_selected_formats()) > 1
        opts = engine.ExportOptions(
            dest_dir=dest_dir, subfolders=subfolders,
            prefix=self.prefix_var.get(), suffix=self.suffix_var.get())

        groups, skipped = engine.group_by_software(self.files)
        for f in skipped:
            self.log("Ignoré (format non géré) : %s" % os.path.basename(f))
        if not groups:
            self.log("Aucun fichier CAO/DAO reconnu.")
            return

        self._total = engine.count_exports(groups, sel)
        if self._total == 0:
            messagebox.showwarning("Rien à faire",
                                   "Aucun export ne correspond aux formats cochés.")
            return

        self._stop = False
        self._start_time = time.time()
        self.progress.configure(maximum=self._total, value=0)
        self.eta_label.configure(text="Préparation…")
        self.convert_btn.configure(state="disabled")
        self.log("")
        self.log(">>> Démarrage de la conversion (%d export(s))…" % self._total)
        if subfolders:
            self.log("    Rangement par sous-dossiers de format activé.")

        def work():
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pythoncom = None
            try:
                res = engine.convert_groups(
                    groups, sel, opts,
                    log=lambda m: self._queue.put(("log", m)),
                    progress=lambda d, t: self._queue.put(("progress", (d, t))),
                    stop_flag=lambda: self._stop)
                self._queue.put(("result", res))
            except Exception as e:
                self._queue.put(("log", "ERREUR : %s" % e))
                self._queue.put(("result", (0, 0, [("", "", str(e))])))
            finally:
                if pythoncom is not None:
                    try:
                        pythoncom.CoUninitialize()
                    except Exception:
                        pass
                self._queue.put(("done", None))

        self._worker = threading.Thread(target=work, daemon=True)
        self._worker.start()
        self._refresh_state()

    # ------------------------------------------------------- boucle & log
    def _poll_queue(self):
        try:
            while True:
                kind, payload = self._queue.get_nowait()
                if kind == "log":
                    self.log(payload)
                elif kind == "progress":
                    self._on_progress(*payload)
                elif kind == "result":
                    self._last_result = payload
                elif kind == "done":
                    self._on_done()
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def _on_progress(self, done, total):
        self.progress.configure(value=done)
        elapsed = time.time() - (self._start_time or time.time())
        if done > 0:
            remaining = elapsed / done * (total - done)
            self.eta_label.configure(
                text="%d / %d exports  ·  temps restant estimé : %s"
                % (done, total, _fmt_duration(remaining)))
        else:
            self.eta_label.configure(text="%d / %d exports" % (done, total))

    def _on_done(self):
        self.progress.configure(value=self.progress["maximum"])
        self.eta_label.configure(text="Terminé.")
        self._refresh_state()
        res = getattr(self, "_last_result", None)
        self.log(">>> Conversion terminée.")
        if res is None:
            return
        ok, total, failures = res
        if not failures:
            messagebox.showinfo(
                "Conversion terminée",
                "Tout est OK : %d/%d export(s) réussi(s)." % (ok, total))
        else:
            lines = []
            for f, t, msg in failures:
                name = os.path.basename(f) if f else "(session)"
                lines.append("• %s  [%s]  — %s" % (name, t.upper() if t else "?", msg))
            detail = "\n".join(lines[:25])
            if len(lines) > 25:
                detail += "\n… et %d autre(s)." % (len(lines) - 25)
            messagebox.showwarning(
                "Conversion terminée avec des échecs",
                "%d/%d export(s) réussi(s).\n\n"
                "Fichiers/formats NON convertis :\n%s" % (ok, total, detail))

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


class HomeFrame(ttk.Frame):
    """Écran d'accueil : choix entre Conversion et Renommage."""
    def __init__(self, master, controller):
        super().__init__(master, padding=30)
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        ttk.Label(self, text="Outils CAO/DAO",
                  font=("", 18, "bold")).grid(row=0, column=0, pady=(10, 4))
        ttk.Label(self, text="Que voulez-vous faire ?",
                  foreground="#555").grid(row=1, column=0, pady=(0, 24))

        cards = ttk.Frame(self)
        cards.grid(row=2, column=0)

        conv = ttk.Frame(cards, padding=16, relief="ridge", borderwidth=1)
        conv.grid(row=0, column=0, padx=12)
        ttk.Label(conv, text="Conversion",
                  font=("", 13, "bold")).pack(pady=(0, 6))
        ttk.Label(conv, text="Convertir par lot en STEP, STL,\n"
                             "DWG, DXF, PDF ou fichier pièce.",
                  justify="center", foreground="#555").pack(pady=(0, 12))
        ttk.Button(conv, text="Ouvrir la conversion",
                   command=controller.show_conversion).pack()

        ren = ttk.Frame(cards, padding=16, relief="ridge", borderwidth=1)
        ren.grid(row=0, column=1, padx=12)
        ttk.Label(ren, text="Renommage",
                  font=("", 13, "bold")).pack(pady=(0, 6))
        ttk.Label(ren, text="Renommer des fichiers par lot.\n(À venir.)",
                  justify="center", foreground="#555").pack(pady=(0, 12))
        ttk.Button(ren, text="Ouvrir le renommage",
                   command=controller.show_rename).pack()


class RenameFrame(ttk.Frame):
    """Écran de renommage — réservé pour plus tard."""
    def __init__(self, master, controller):
        super().__init__(master, padding=30)
        self.controller = controller
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)

        header = ttk.Frame(self)
        header.grid(row=0, column=0, sticky="ew")
        ttk.Button(header, text="‹ Accueil",
                   command=controller.show_home).pack(side="left")
        ttk.Label(header, text="   Renommage",
                  font=("", 11, "bold")).pack(side="left")

        ttk.Label(self, text="Fonction de renommage à venir.",
                  font=("", 13)).grid(row=1, column=0, pady=40)
        ttk.Label(self, text="On en reparle : dis-moi les règles de "
                             "renommage souhaitées.",
                  foreground="#555").grid(row=2, column=0)


class Controller(object):
    """Aiguille entre l'accueil, la conversion et le renommage."""
    GEOMETRY = {"home": "560x360", "conversion": "820x820", "rename": "560x360"}

    def __init__(self, root):
        self.root = root
        self.show_home()

    def _clear(self):
        for w in self.root.winfo_children():
            w.destroy()

    def show_home(self):
        self._clear()
        self.root.geometry(self.GEOMETRY["home"])
        HomeFrame(self.root, self)

    def show_conversion(self):
        self._clear()
        self.root.geometry(self.GEOMETRY["conversion"])
        App(self.root, controller=self)

    def show_rename(self):
        self._clear()
        self.root.geometry(self.GEOMETRY["rename"])
        RenameFrame(self.root, self)


def main():
    root = tk.Tk()
    root.title("Outils CAO/DAO")
    try:
        ttk.Style().theme_use("vista")
    except tk.TclError:
        pass
    Controller(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    main()
