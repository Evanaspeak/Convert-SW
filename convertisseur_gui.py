#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convertisseur_gui.py — Interface graphique du convertisseur CAO/DAO par lot.

Fenêtre simple (tkinter, livré avec Python) pour :
  - choisir des fichiers via le gestionnaire de fichiers (multi-sélection) ;
  - cocher les formats voulus selon le type (pièce/assemblage ou mise en plan) ;
  - choisir la destination (à côté des fichiers, ou un dossier personnalisé) ;
  - lancer la conversion et suivre le journal en direct.

La conversion tourne dans un thread séparé pour ne pas figer la fenêtre ;
le pilotage COM y est initialisé (CoInitialize).
"""

import os
import queue
import threading

import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import convertisseur_cao as engine


# Extensions gérées, pour le filtre du sélecteur de fichiers
_ALL_EXTS = sorted(engine.ROUTES.keys())
_FILE_TYPES = [
    ("Fichiers CAO/DAO", " ".join("*" + e for e in _ALL_EXTS)),
    ("SolidWorks", "*.sldprt *.sldasm *.slddrw"),
    ("CATIA V5", "*.CATPart *.CATProduct *.CATDrawing"),
    ("Inventor", "*.ipt *.iam *.idw"),
    ("Creo", "*.prt *.asm *.drw"),
    ("Tous les fichiers", "*.*"),
]


class App(ttk.Frame):
    def __init__(self, master):
        super().__init__(master, padding=10)
        self.master = master
        self.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(6, weight=1)

        self.files = []                 # liste des chemins ajoutés
        self.format_vars = {}           # target -> tk.BooleanVar
        self._queue = queue.Queue()     # messages du thread de conversion
        self._worker = None
        self._stop = False

        self._build()
        self._poll_queue()
        self._refresh_state()

    # ------------------------------------------------------------------ UI
    def _build(self):
        row = 0

        # --- Fichiers -------------------------------------------------
        bar = ttk.Frame(self)
        bar.grid(row=row, column=0, sticky="ew")
        ttk.Button(bar, text="Ajouter des fichiers…",
                   command=self.add_files).pack(side="left")
        ttk.Button(bar, text="Ajouter un dossier…",
                   command=self.add_folder).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Retirer la sélection",
                   command=self.remove_selected).pack(side="left", padx=(6, 0))
        ttk.Button(bar, text="Vider la liste",
                   command=self.clear_files).pack(side="left", padx=(6, 0))
        row += 1

        ttk.Label(self, text="Fichiers à convertir :").grid(
            row=row, column=0, sticky="w", pady=(8, 2))
        row += 1

        list_frame = ttk.Frame(self)
        list_frame.grid(row=row, column=0, sticky="nsew")
        list_frame.columnconfigure(0, weight=1)
        self.rowconfigure(row, weight=1)
        self.listbox = tk.Listbox(list_frame, height=8, selectmode="extended",
                                  activestyle="none")
        self.listbox.grid(row=0, column=0, sticky="nsew")
        sb = ttk.Scrollbar(list_frame, orient="vertical",
                           command=self.listbox.yview)
        sb.grid(row=0, column=1, sticky="ns")
        self.listbox.configure(yscrollcommand=sb.set)
        row += 1

        self.count_label = ttk.Label(self, text="")
        self.count_label.grid(row=row, column=0, sticky="w", pady=(2, 6))
        row += 1

        # --- Formats --------------------------------------------------
        fmt = ttk.Frame(self)
        fmt.grid(row=row, column=0, sticky="ew", pady=(0, 6))
        fmt.columnconfigure(0, weight=1)
        fmt.columnconfigure(1, weight=1)

        self.frame_3d = ttk.LabelFrame(fmt, text="Pièces / assemblages",
                                       padding=8)
        self.frame_3d.grid(row=0, column=0, sticky="nsew", padx=(0, 4))
        for t in engine.TARGETS["3d"]:
            self._add_format_check(self.frame_3d, t, default=True)

        self.frame_2d = ttk.LabelFrame(fmt, text="Mises en plan", padding=8)
        self.frame_2d.grid(row=0, column=1, sticky="nsew", padx=(4, 0))
        for t in engine.TARGETS["2d"]:
            # PDF décoché par défaut, DWG/DXF cochés
            self._add_format_check(self.frame_2d, t, default=(t != "pdf"))
        row += 1

        # --- Destination ---------------------------------------------
        dest = ttk.LabelFrame(self, text="Destination", padding=8)
        dest.grid(row=row, column=0, sticky="ew", pady=(0, 6))
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
        row += 1

        # --- Action ---------------------------------------------------
        act = ttk.Frame(self)
        act.grid(row=row, column=0, sticky="ew", pady=(0, 6))
        act.columnconfigure(0, weight=1)
        self.progress = ttk.Progressbar(act, mode="indeterminate")
        self.progress.grid(row=0, column=0, sticky="ew", padx=(0, 8))
        self.convert_btn = ttk.Button(act, text="Convertir",
                                      command=self.start_conversion)
        self.convert_btn.grid(row=0, column=1)
        row += 1

        # --- Journal --------------------------------------------------
        ttk.Label(self, text="Journal :").grid(row=row, column=0, sticky="w")
        row += 1
        log_frame = ttk.Frame(self)
        log_frame.grid(row=row, column=0, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.rowconfigure(row, weight=2)
        self.log_text = tk.Text(log_frame, height=10, wrap="word", state="disabled")
        self.log_text.grid(row=0, column=0, sticky="nsew")
        lsb = ttk.Scrollbar(log_frame, orient="vertical",
                            command=self.log_text.yview)
        lsb.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=lsb.set)

    def _add_format_check(self, parent, target, default):
        var = tk.BooleanVar(value=default)
        self.format_vars[target] = var
        cb = ttk.Checkbutton(parent, text=engine.TARGET_LABELS.get(target, target),
                             variable=var)
        cb.pack(anchor="w")
        return cb

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
            if engine.route_for(p) is None:
                self.log("Ignoré (format non géré) : %s" % os.path.basename(p))
                continue
            self.files.append(p)
            existing.add(p)
            soft, kind = engine.route_for(p)
            label = "pièce/assemblage" if kind == "3d" else "mise en plan"
            self.listbox.insert("end", "  [%s · %s]  %s" % (
                engine.SOFTWARE_LABELS.get(soft, soft), label, p))
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
        has3d = has2d = False
        for p in self.files:
            route = engine.route_for(p)
            if not route:
                continue
            if route[1] == "3d":
                has3d = True
            else:
                has2d = True
        return has3d, has2d

    def _set_frame_state(self, frame, enabled):
        state = "normal" if enabled else "disabled"
        for child in frame.winfo_children():
            try:
                child.configure(state=state)
            except tk.TclError:
                pass

    def _refresh_state(self):
        has3d, has2d = self._kinds_present()
        self._set_frame_state(self.frame_3d, has3d)
        self._set_frame_state(self.frame_2d, has2d)

        custom = self.dest_mode.get() == "custom"
        self.dest_entry.configure(state="normal" if custom else "disabled")
        self.dest_browse.configure(state="normal" if custom else "disabled")

        n = len(self.files)
        n3 = sum(1 for p in self.files if (engine.route_for(p) or (None, None))[1] == "3d")
        n2 = n - n3
        self.count_label.configure(
            text="%d fichier(s) : %d pièce(s)/assemblage(s), %d mise(s) en plan"
            % (n, n3, n2))

        busy = self._worker is not None and self._worker.is_alive()
        self.convert_btn.configure(state="disabled" if busy or n == 0 else "normal")

    # ----------------------------------------------------------- conversion
    def _selected_formats(self, kind):
        return [t for t in engine.TARGETS[kind] if self.format_vars[t].get()]

    def start_conversion(self):
        if not self.files:
            return
        sel3d = self._selected_formats("3d")
        sel2d = self._selected_formats("2d")
        has3d, has2d = self._kinds_present()
        if (has3d and not sel3d) and (has2d and not sel2d):
            messagebox.showwarning("Aucun format",
                                   "Cochez au moins un format de sortie.")
            return
        if has3d and not sel3d:
            messagebox.showwarning(
                "Aucun format 3D",
                "Des pièces/assemblages sont sélectionnés mais aucun format "
                "(STEP/STL) n'est coché : ils seront ignorés.")
        if has2d and not sel2d:
            messagebox.showwarning(
                "Aucun format plan",
                "Des mises en plan sont sélectionnées mais aucun format "
                "(DWG/DXF/PDF) n'est coché : elles seront ignorées.")

        dest_dir = None
        if self.dest_mode.get() == "custom":
            dest_dir = self.dest_entry.get().strip()
            if not dest_dir:
                messagebox.showwarning("Destination",
                                       "Indiquez un dossier de destination.")
                return

        groups, skipped = engine.group_by_software(self.files)
        for f in skipped:
            self.log("Ignoré (format non géré) : %s" % os.path.basename(f))
        if not groups:
            self.log("Aucun fichier CAO/DAO reconnu.")
            return

        self._stop = False
        self.progress.start(12)
        self.convert_btn.configure(state="disabled")
        self.log("")
        self.log(">>> Démarrage de la conversion…")

        def work():
            # COM doit être initialisé dans ce thread
            try:
                import pythoncom
                pythoncom.CoInitialize()
            except Exception:
                pythoncom = None
            try:
                engine.convert_groups(
                    groups, sel3d, sel2d, dest_dir=dest_dir,
                    log=lambda m: self._queue.put(("log", m)),
                    stop_flag=lambda: self._stop)
            except Exception as e:
                self._queue.put(("log", "ERREUR : %s" % e))
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
                elif kind == "done":
                    self.progress.stop()
                    self.log(">>> Conversion terminée.")
                    self._refresh_state()
        except queue.Empty:
            pass
        self.after(120, self._poll_queue)

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")


def main():
    root = tk.Tk()
    root.title("Convertisseur CAO/DAO par lot")
    root.geometry("760x680")
    try:
        ttk.Style().theme_use("vista")   # thème natif Windows si dispo
    except tk.TclError:
        pass
    App(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    main()
