#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convertisseur_cao.py — Moteur de conversion CAO/DAO par lot (gratuit, local).

On ne convertit rien "à la main" : on PILOTE les logiciels de CAO déjà
installés via leur automation COM (pywin32). Chaque logiciel n'est lancé
qu'une seule fois, quel que soit le nombre de fichiers à traiter.

Formats de sortie possibles :
    - Pièces / assemblages  -> STEP (.step) et/ou STL (.stl)
    - Mises en plan         -> DWG (.dwg), DXF (.dxf) et/ou PDF (.pdf)

Ce fichier contient le MOTEUR (routage + handlers) et une petite interface
en ligne de commande. L'interface graphique est dans convertisseur_gui.py.

Sans argument, on ouvre l'interface graphique.

Usage ligne de commande :
    python convertisseur_cao.py fichier1.sldprt dossier\\ ...
    python convertisseur_cao.py --dry-run *.sldprt      (n'ouvre aucun logiciel)
    python convertisseur_cao.py --help

Ne fonctionne réellement que sous Windows avec pywin32 et les logiciels
concernés. Le mode --dry-run tourne partout (sert à vérifier le routage).
"""

import os
import re
import sys
import time
import traceback

# ---------------------------------------------------------------------------
# Configuration / routage
# ---------------------------------------------------------------------------

# ext (en minuscules) -> (logiciel, nature)   nature : "3d" ou "2d"
ROUTES = {
    ".sldprt":     ("solidworks", "3d"),
    ".sldasm":     ("solidworks", "3d"),
    ".slddrw":     ("solidworks", "2d"),

    ".catpart":    ("catia", "3d"),
    ".catproduct": ("catia", "3d"),
    ".catdrawing": ("catia", "2d"),

    ".ipt":        ("inventor", "3d"),
    ".iam":        ("inventor", "3d"),
    ".idw":        ("inventor", "2d"),

    ".prt":        ("creo", "3d"),
    ".asm":        ("creo", "3d"),
    ".drw":        ("creo", "2d"),
}

# Formats proposés par nature. Sert de valeurs par défaut (CLI) et de liste
# de référence pour l'interface graphique.
TARGETS = {
    "3d": ["step", "stl"],
    "2d": ["dwg", "dxf", "pdf"],
}

# Libellés affichés dans l'interface
TARGET_LABELS = {
    "step": "STEP (.step)",
    "stl":  "STL (.stl)",
    "dwg":  "DWG (.dwg)",
    "dxf":  "DXF (.dxf)",
    "pdf":  "PDF (.pdf)",
}

SOFTWARE_LABELS = {
    "solidworks": "SolidWorks",
    "catia": "CATIA V5",
    "inventor": "Inventor",
    "creo": "PTC Creo",
}

# Creo numérote ses fichiers : « carter.prt.3 ». On retire le suffixe « .N ».
_CREO_VERSIONED = re.compile(r"^(?P<stem>.*\.(?:prt|asm|drw))\.\d+$", re.IGNORECASE)


def route_for(path):
    """Retourne (logiciel, nature) pour un chemin, ou None si non géré."""
    name = os.path.basename(path).lower()
    m = _CREO_VERSIONED.match(name)
    if m:
        name = m.group("stem")
    _, ext = os.path.splitext(name)
    return ROUTES.get(ext)


def clean_stem(path):
    """Nom de base sans extension, en gérant le suffixe de version Creo."""
    name = os.path.basename(path)
    m = _CREO_VERSIONED.match(name.lower())
    if m:
        name = name[:len(m.group("stem"))]
    stem, _ = os.path.splitext(name)
    return stem


def export_dir_for(path):
    """Crée (au besoin) et renvoie le dossier « Export » à côté du fichier."""
    d = os.path.join(os.path.dirname(os.path.abspath(path)), "Export")
    os.makedirs(d, exist_ok=True)
    return d


def output_path(src, target, dest_dir=None):
    """Chemin de sortie pour un fichier source et un format cible.

    dest_dir=None -> sous-dossier « Export » à côté du fichier source.
    dest_dir=chemin -> ce dossier (créé au besoin).
    """
    if dest_dir:
        os.makedirs(dest_dir, exist_ok=True)
        folder = dest_dir
    else:
        folder = export_dir_for(src)
    return os.path.join(folder, clean_stem(src) + "." + target)


# ---------------------------------------------------------------------------
# Journalisation (par défaut : print ; l'interface graphique fournit la sienne)
# ---------------------------------------------------------------------------

def default_log(msg):
    print(time.strftime("[%H:%M:%S] ") + msg, flush=True)


# ---------------------------------------------------------------------------
# Utilitaire COM (importé paresseusement)
# ---------------------------------------------------------------------------

def _byref_long(value=0):
    """VARIANT entier passé par référence (VT_BYREF | VT_I4).

    C'est LE correctif du DISP_E_TYPEMISMATCH de SolidWorks : les arguments
    [out] « Errors » / « Warnings » de OpenDoc6 et de SaveAs doivent être
    passés comme des VARIANT byref, pas comme des entiers Python.
    """
    import pythoncom
    from win32com.client import VARIANT
    return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, value)


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

class Handler(object):
    name = "?"

    def __init__(self, log=default_log):
        self.app = None
        self.log = log

    def start(self):
        raise NotImplementedError

    def convert(self, src, targets, dest_dir=None):
        """Convertit un fichier vers chaque format de `targets`.

        Renvoie une liste de tuples (target, outpath, ok, message).
        """
        raise NotImplementedError

    def stop(self):
        pass


# ------------------------------- SolidWorks --------------------------------

class SolidWorksHandler(Handler):
    name = "SolidWorks"

    _DOCTYPE = {".sldprt": 1, ".sldasm": 2, ".slddrw": 3}  # swDocumentTypes_e
    _SILENT_OPEN = 1                                        # swOpenDocOptions_Silent
    _VER_CURRENT = 0                                        # swSaveAsCurrentVersion
    _SAVE_SILENT_COPY = 1 | 2                              # Silent | Copy

    def start(self):
        import win32com.client
        self.app = win32com.client.Dispatch("SldWorks.Application")
        self.app.Visible = True
        self._open_mode = None
        self._save_mode = None

    def _open(self, src, ext):
        import pythoncom
        dtype = self._DOCTYPE[ext]

        def m_variant():
            errs, warns = _byref_long(), _byref_long()
            return self.app.OpenDoc6(src, dtype, self._SILENT_OPEN, "", errs, warns)

        def m_int():
            return self.app.OpenDoc6(src, dtype, self._SILENT_OPEN, "", 0, 0)

        def m_open():
            return self.app.OpenDoc(src, dtype)

        modes = [("OpenDoc6+VARIANT", m_variant),
                 ("OpenDoc6+int", m_int),
                 ("OpenDoc", m_open)]
        if self._open_mode:
            modes.sort(key=lambda kv: kv[0] != self._open_mode)

        last_err = None
        for label, fn in modes:
            try:
                doc = fn()
                if doc is not None:
                    self._open_mode = label
                    return doc
                active = self.app.ActiveDoc
                if active is not None:
                    self._open_mode = label
                    return active
            except pythoncom.com_error as e:
                last_err = e
        if last_err:
            raise last_err
        raise RuntimeError("SolidWorks n'a pas pu ouvrir le fichier (doc None).")

    def _save(self, doc, outpath):
        import pythoncom
        ver, opt = self._VER_CURRENT, self._SAVE_SILENT_COPY
        ext = doc.Extension

        def m_ext_saveas():
            errs, warns = _byref_long(), _byref_long()
            return ext.SaveAs(outpath, ver, opt, None, errs, warns)

        def m_ext_saveas3():
            errs, warns = _byref_long(), _byref_long()
            return ext.SaveAs3(outpath, ver, opt, None, None, errs, warns)

        def m_doc_saveas4():
            errs, warns = _byref_long(), _byref_long()
            return doc.SaveAs4(outpath, ver, opt, errs, warns)

        def m_doc_saveas3():
            return doc.SaveAs3(outpath, ver, opt)

        modes = [("Ext.SaveAs", m_ext_saveas),
                 ("Ext.SaveAs3", m_ext_saveas3),
                 ("Doc.SaveAs4", m_doc_saveas4),
                 ("Doc.SaveAs3", m_doc_saveas3)]
        if self._save_mode:
            modes.sort(key=lambda kv: kv[0] != self._save_mode)

        last_err = None
        for label, fn in modes:
            try:
                fn()
                self._save_mode = label
                return True
            except pythoncom.com_error as e:
                last_err = e
        if last_err:
            raise last_err
        return False

    def convert(self, src, targets, dest_dir=None):
        results = []
        try:
            ext = os.path.splitext(src.lower())[1]
            doc = self._open(src, ext)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t, dest_dir), False, "ouverture: %s" % e))
            return results

        title = None
        try:
            title = doc.GetTitle()
        except Exception:
            pass

        for t in targets:
            out = output_path(src, t, dest_dir)
            try:
                ok = self._save(doc, out) and os.path.exists(out)
                results.append((t, out, ok, "" if ok else "SaveAs a échoué"))
            except Exception as e:
                results.append((t, out, False, str(e)))

        try:
            if title:
                self.app.CloseDoc(title)
        except Exception:
            pass
        return results

    def stop(self):
        # On laisse SolidWorks ouvert (plus rapide, et on ne tue pas une
        # session que l'utilisateur avait déjà lancée).
        pass


# --------------------------------- CATIA -----------------------------------

class CatiaHandler(Handler):
    name = "CATIA V5"

    _FORMATS = {
        "step": ["stp", "step"],
        "stl":  ["stl"],
        "dwg":  ["dwg"],
        "dxf":  ["dxf"],
        "pdf":  ["pdf", "PDF"],
    }

    def start(self):
        import win32com.client
        try:
            self.app = win32com.client.GetActiveObject("CATIA.Application")
        except Exception:
            self.app = win32com.client.Dispatch("CATIA.Application")
        try:
            self.app.Visible = True
        except Exception:
            pass
        try:
            self.app.DisplayFileAlerts = False
        except Exception:
            pass

    def convert(self, src, targets, dest_dir=None):
        results = []
        try:
            doc = self.app.Documents.Open(src)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t, dest_dir), False, "ouverture: %s" % e))
            return results

        for t in targets:
            out = output_path(src, t, dest_dir)
            ok, msg = False, ""
            for fmt in self._FORMATS.get(t, [t]):
                try:
                    doc.ExportData(out, fmt)
                    if os.path.exists(out):
                        ok = True
                        break
                    msg = "ExportData('%s') sans fichier produit" % fmt
                except Exception as e:
                    msg = "format '%s': %s" % (fmt, e)
            results.append((t, out, ok, "" if ok else msg))

        try:
            doc.Close()
        except Exception:
            pass
        return results


# -------------------------------- Inventor ---------------------------------

class InventorHandler(Handler):
    name = "Inventor"

    _KFILEBROWSE = 13059  # kFileBrowseIOMechanism

    # target -> (mot-clé DisplayName, GUID du traducteur)
    _TRANSLATORS = {
        "step": ("STEP", "{90AF7F40-0C01-11D5-B79E-0010B3AF2C1B}"),
        "stl":  ("STL",  "{81D58C09-F638-42B4-BD3E-F5DAC29F94D1}"),
        "dwg":  ("DWG",  "{C24E3AC2-122E-11D5-8E91-0010B541CD80}"),
        "dxf":  ("DXF",  "{C24E3AC4-122E-11D5-8E91-0010B541CD80}"),
        "pdf":  ("PDF",  "{0AC6FD95-2F4D-42CE-8BE0-8AEA580399E4}"),
    }

    def start(self):
        import win32com.client
        try:
            self.app = win32com.client.GetActiveObject("Inventor.Application")
        except Exception:
            self.app = win32com.client.Dispatch("Inventor.Application")
        try:
            self.app.Visible = True
        except Exception:
            pass
        self._addin_cache = {}

    def _translator(self, target):
        if target in self._addin_cache:
            return self._addin_cache[target]
        keyword, guid = self._TRANSLATORS[target]
        addins = self.app.ApplicationAddIns
        found = None
        for a in addins:
            try:
                if (a.ClassIdString or "").upper() == guid.upper():
                    found = a
                    break
            except Exception:
                continue
        if found is None:
            for a in addins:
                try:
                    if keyword.lower() in (a.DisplayName or "").lower():
                        found = a
                        break
                except Exception:
                    continue
        self._addin_cache[target] = found
        return found

    def convert(self, src, targets, dest_dir=None):
        results = []
        try:
            doc = self.app.Documents.Open(src, False)
        except Exception:
            try:
                doc = self.app.Documents.Open(src)
            except Exception as e:
                for t in targets:
                    results.append((t, output_path(src, t, dest_dir), False, "ouverture: %s" % e))
                return results

        to = self.app.TransientObjects
        for t in targets:
            out = output_path(src, t, dest_dir)
            addin = self._translator(t)
            if addin is None:
                results.append((t, out, False, "traducteur %s introuvable" % t.upper()))
                continue
            try:
                try:
                    addin.Activate()
                except Exception:
                    pass
                ctx = to.CreateTranslationContext()
                ctx.Type = self._KFILEBROWSE
                opts = to.CreateNameValueMap()
                med = to.CreateDataMedium()
                med.FileName = out
                try:
                    if addin.HasSaveCopyAsOptions(doc, ctx, opts) and t == "step":
                        opts.Add("ApplicationProtocolType", 3)  # AP214
                except Exception:
                    pass
                addin.SaveCopyAs(doc, ctx, opts, med)
                ok = os.path.exists(out)
                results.append((t, out, ok, "" if ok else "SaveCopyAs sans fichier produit"))
            except Exception as e:
                results.append((t, out, False, str(e)))

        try:
            doc.Close(True)
        except Exception:
            pass
        return results


# ---------------------------------- Creo -----------------------------------

class CreoHandler(Handler):
    """Pilotage de PTC Creo Parametric via la VB API (Object TOOLKIT), gratuite.

    Prérequis : Creo doit être DÉJÀ LANCÉ (connexion asynchrone à la session)
    et la VB API installée/enregistrée. ProgID variables selon la version :
    plusieurs sont tentés. À VÉRIFIER sur la machine.
    """
    name = "PTC Creo"

    _CONNECT_PROGIDS = [
        "pfcls.CCpfcAsyncConnection",
        "pfcAsyncConnection.pfcAsyncConnection",
        "CCpfcAsyncConnection",
    ]

    def start(self):
        import win32com.client
        self._win32 = win32com.client
        self.session = None
        last = None
        for progid in self._CONNECT_PROGIDS:
            try:
                connector = win32com.client.Dispatch(progid)
                conn = connector.Connect(None, None, None, None)
                self.session = conn.Session
                self.log("  Creo : connexion via %s" % progid)
                break
            except Exception as e:
                last = e
        if self.session is None:
            raise RuntimeError(
                "Impossible de se connecter à Creo (VB API). Creo est-il lancé "
                "et la VB API installée ? Dernière erreur : %s" % last)

    def _instructions(self, target):
        w = self._win32
        cls = {
            "step": "pfcls.CCpfcSTEP3DExportInstructions",
            "stl":  "pfcls.CCpfcSTLBinaryExportInstructions",
            "dwg":  "pfcls.CCpfcDWGExportInstructions",
            "dxf":  "pfcls.CCpfcDXFExportInstructions",
            "pdf":  "pfcls.CCpfcPDFExportInstructions",
        }[target]
        return w.Dispatch(cls).Create()

    def convert(self, src, targets, dest_dir=None):
        results = []
        try:
            from_dir = os.path.dirname(os.path.abspath(src))
            try:
                self.session.ChangeDirectory(from_dir)
            except Exception:
                pass
            descr = self._win32.Dispatch(
                "pfcls.CCpfcModelDescriptor").CreateFromFileName(os.path.basename(src))
            model = self.session.RetrieveModel(descr)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t, dest_dir), False, "ouverture: %s" % e))
            return results

        for t in targets:
            out = output_path(src, t, dest_dir)
            try:
                model.Export(out, self._instructions(t))
                ok = os.path.exists(out)
                results.append((t, out, ok, "" if ok else "Export sans fichier produit"))
            except Exception as e:
                results.append((t, out, False, str(e)))
        return results


HANDLERS = {
    "solidworks": SolidWorksHandler,
    "catia": CatiaHandler,
    "inventor": InventorHandler,
    "creo": CreoHandler,
}


# ---------------------------------------------------------------------------
# Cœur réutilisable (appelé par la CLI comme par l'interface graphique)
# ---------------------------------------------------------------------------

def collect_files(paths):
    """Développe les arguments en liste de fichiers existants."""
    files = []
    for p in paths:
        p = p.strip('"')
        if os.path.isdir(p):
            for root, dirs, names in os.walk(p):
                dirs[:] = [d for d in dirs
                           if d.lower() != "export" and not d.startswith(".")]
                for n in names:
                    files.append(os.path.join(root, n))
        elif os.path.isfile(p):
            files.append(p)
    return files


def group_by_software(files):
    """Renvoie (groups, skipped).

    groups : { logiciel : [ (chemin, nature), ... ] }
    skipped : [ chemins non gérés ]
    """
    groups, skipped = {}, []
    for f in files:
        route = route_for(f)
        if route is None:
            skipped.append(f)
            continue
        soft, kind = route
        groups.setdefault(soft, []).append((f, kind))
    return groups, skipped


def convert_groups(groups, sel3d, sel2d, dest_dir=None, log=default_log,
                   stop_flag=None):
    """Convertit tous les fichiers groupés par logiciel.

    sel3d / sel2d : listes de formats retenus pour 3D / 2D.
    stop_flag : callable() -> bool, permet d'interrompre proprement.
    Renvoie (total_ok, total).
    """
    total_ok = total = 0
    for soft, items in groups.items():
        if stop_flag and stop_flag():
            break
        log("=" * 56)
        log("%s : %d fichier(s)" % (SOFTWARE_LABELS.get(soft, soft), len(items)))

        handler = HANDLERS[soft](log=log)
        try:
            handler.start()
        except Exception as e:
            log("  Impossible de démarrer %s : %s" % (handler.name, e))
            for f, kind in items:
                total += len(sel3d if kind == "3d" else sel2d)
            continue

        for f, kind in items:
            if stop_flag and stop_flag():
                break
            targets = list(sel3d if kind == "3d" else sel2d)
            if not targets:
                continue
            log("  > %s" % os.path.basename(f))
            try:
                res = handler.convert(f, targets, dest_dir)
            except Exception as e:
                log("    ERREUR inattendue : %s" % e)
                res = [(t, output_path(f, t, dest_dir), False, str(e)) for t in targets]
            for t, out, ok, msg in res:
                total += 1
                if ok:
                    total_ok += 1
                    log("    OK  %s -> %s" % (t.upper(), out))
                else:
                    log("    KO  %s : %s" % (t.upper(), msg))

        try:
            handler.stop()
        except Exception:
            pass

    log("=" * 56)
    log("Terminé : %d/%d export(s) réussi(s)." % (total_ok, total))
    return total_ok, total


# ---------------------------------------------------------------------------
# Interface en ligne de commande
# ---------------------------------------------------------------------------

def run_cli(paths, dry_run=False):
    files = collect_files(paths)
    if not files:
        default_log("Aucun fichier à traiter.")
        return 1
    groups, skipped = group_by_software(files)
    for f in skipped:
        default_log("Format non géré, ignoré : %s" % os.path.basename(f))
    if not groups:
        default_log("Aucun fichier CAO/DAO reconnu.")
        return 1

    if dry_run:
        n = 0
        for soft, items in groups.items():
            default_log("=" * 56)
            default_log("%s : %d fichier(s)" % (SOFTWARE_LABELS.get(soft, soft), len(items)))
            for f, kind in items:
                for t in TARGETS[kind]:
                    default_log("  [dry-run] %s -> %s" % (os.path.basename(f), output_path(f, t)))
                    n += 1
        default_log("Simulation terminée : %d export(s) prévu(s)." % n)
        return 0

    convert_groups(groups, TARGETS["3d"], TARGETS["2d"], dest_dir=None, log=default_log)
    return 0


def main(argv):
    args = list(argv[1:])
    if "--help" in args or "-h" in args:
        print(__doc__)
        return 0
    if not args:
        # Aucun argument : on ouvre l'interface graphique.
        try:
            import convertisseur_gui
            return convertisseur_gui.main()
        except Exception as e:
            print("Interface graphique indisponible (%s)." % e)
            print(__doc__)
            return 1
    dry = "--dry-run" in args
    args = [a for a in args if a != "--dry-run"]
    return run_cli(args, dry_run=dry)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
