#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convertisseur_cao.py — Moteur de conversion CAO/DAO par lot (gratuit, local).

On ne convertit rien "à la main" : on PILOTE les logiciels de CAO déjà
installés via leur automation COM (pywin32). Chaque logiciel n'est lancé
qu'une seule fois, quel que soit le nombre de fichiers à traiter.

Trois natures de fichiers, chacune avec ses formats de sortie :
    - Pièces       -> STEP (.step) et/ou STL (.stl)
    - Assemblages  -> STEP (.step) et/ou STL (.stl)
    - Mises en plan-> DWG (.dwg), DXF (.dxf) et/ou PDF (.pdf)

Ce fichier contient le MOTEUR. L'interface graphique est dans
convertisseur_gui.py. Sans argument, on ouvre l'interface graphique.

Usage ligne de commande :
    python convertisseur_cao.py fichier1.sldprt dossier\\ ...
    python convertisseur_cao.py --dry-run *.sldprt      (n'ouvre aucun logiciel)
    python convertisseur_cao.py --help

Ne fonctionne réellement que sous Windows avec pywin32 et les logiciels
concernés. Le mode --dry-run tourne partout (sert à vérifier le routage).
"""

import os
import re
import shutil
import sys
import time
import traceback
import unicodedata

# ---------------------------------------------------------------------------
# Configuration / routage
# ---------------------------------------------------------------------------

# ext (en minuscules) -> (logiciel, nature)
# nature : "part" (pièce) | "assembly" (assemblage) | "drawing" (mise en plan)
ROUTES = {
    ".sldprt":     ("solidworks", "part"),
    ".sldasm":     ("solidworks", "assembly"),
    ".slddrw":     ("solidworks", "drawing"),

    ".catpart":    ("catia", "part"),
    ".catproduct": ("catia", "assembly"),
    ".catdrawing": ("catia", "drawing"),

    ".ipt":        ("inventor", "part"),
    ".iam":        ("inventor", "assembly"),
    ".idw":        ("inventor", "drawing"),

    ".prt":        ("creo", "part"),
    ".asm":        ("creo", "assembly"),
    ".drw":        ("creo", "drawing"),
}

# Formats proposés par nature (valeurs par défaut pour la CLI et l'interface).
# "part" comme cible = enregistrer un ASSEMBLAGE sous forme de fichier pièce
# (extension native du logiciel, ex. .sldprt pour SolidWorks).
TARGETS = {
    "part":     ["step", "stl"],
    "assembly": ["step", "part"],
    "drawing":  ["dwg", "dxf", "pdf"],
}

# Ordre d'affichage des natures dans l'interface
KIND_ORDER = ["part", "assembly", "drawing"]

KIND_LABELS = {
    "part":     "Pièces",
    "assembly": "Assemblages",
    "drawing":  "Mises en plan",
}

TARGET_LABELS = {
    "step": "STEP (.step)",
    "stl":  "STL (.stl)",
    "part": "Pièce (fichier natif)",
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

# Caractères interdits dans un nom de fichier Windows (pour préfixe/suffixe)
_BAD_NAME_CHARS = re.compile(r'[<>:"/\\|?*]')


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


def sanitize_affix(text):
    """Nettoie un préfixe/suffixe (retire les caractères interdits)."""
    return _BAD_NAME_CHARS.sub("", text or "")


def export_dir_for(path):
    """Renvoie le dossier « Export » à côté du fichier (créé au besoin)."""
    d = os.path.join(os.path.dirname(os.path.abspath(path)), "Export")
    os.makedirs(d, exist_ok=True)
    return d


class ExportOptions(object):
    """Options de sortie communes à toute une conversion.

    dest_dir   : dossier de destination unique (None = « Export » à côté
                 de chaque fichier source).
    subfolders : True -> ranger chaque format dans son sous-dossier
                 (STEP/, STL/, DWG/, DXF/, PDF/).
    prefix     : texte ajouté DEVANT le nom de fichier.
    suffix     : texte ajouté DERRIÈRE le nom (avant l'extension).
    """
    def __init__(self, dest_dir=None, subfolders=False, prefix="", suffix="",
                 number_mode="none", number_start=1, number_digits=3,
                 number_position="suffix", number_sep="_", category_text=None,
                 category_position="prefix", base_name=""):
        self.dest_dir = dest_dir or None
        self.subfolders = bool(subfolders)
        self.prefix = sanitize_affix(prefix)
        self.suffix = sanitize_affix(suffix)
        # nom commun : s'il est renseigné, il REMPLACE le nom d'origine
        self.base_name = sanitize_affix(base_name)
        # numérotation : "none" | "global" | "per_type"
        self.number_mode = number_mode if number_mode in ("none", "global", "per_type") else "none"
        try:
            self.number_start = int(number_start)
        except (TypeError, ValueError):
            self.number_start = 1
        try:
            self.number_digits = max(1, int(number_digits))
        except (TypeError, ValueError):
            self.number_digits = 3
        self.number_position = "prefix" if number_position == "prefix" else "suffix"
        self.number_sep = number_sep if number_sep is not None else "_"
        # texte propre à chaque catégorie : { "part"/"assembly"/"drawing": txt }
        self.category_text = {k: sanitize_affix(v)
                              for k, v in (category_text or {}).items()}
        # position du texte catégorie : "prefix" (début) ou "suffix" (fin)
        self.category_position = "suffix" if category_position == "suffix" else "prefix"
        # rempli par build_export_names : { chemin_abs : nom_de_sortie }
        self.name_for = None


def build_output_stem(stem, kind, idx, opts):
    """Construit le nom de sortie (sans extension) d'un fichier.

    idx : rang pour la numérotation (0,1,2…) ou None si pas de numéro.
    Assemblage : préfixe + texte_catégorie + [num] + nom + [num] + suffixe.
    """
    if opts.base_name:
        stem = opts.base_name
    num = ""
    if idx is not None and opts.number_mode in ("global", "per_type"):
        num = str(opts.number_start + idx).zfill(opts.number_digits)
    cat = (opts.category_text or {}).get(kind, "")
    inner = stem
    if num and opts.number_position == "prefix":
        inner = "%s%s%s" % (num, opts.number_sep, inner)
    elif num:
        inner = "%s%s%s" % (inner, opts.number_sep, num)
    # texte de catégorie au début ou à la fin (à l'intérieur du préfixe/suffixe)
    if cat and opts.category_position == "suffix":
        core = "%s%s" % (inner, cat)
    elif cat:
        core = "%s%s" % (cat, inner)
    else:
        core = inner
    full = "%s%s%s" % (opts.prefix, core, opts.suffix)
    return _BAD_NAME_CHARS.sub("", full)


def build_export_names(items, opts):
    """Pré-calcule le nom de sortie de chaque fichier (numérotation comprise).

    items : liste de (chemin, nature) dans l'ordre de numérotation voulu.
    Remplit et renvoie opts.name_for = { chemin_abs : nom_sans_extension }.
    """
    counters = {}
    global_ctr = 0
    name_for = {}
    for src, kind in items:
        if opts.number_mode == "global":
            idx = global_ctr
            global_ctr += 1
        elif opts.number_mode == "per_type":
            idx = counters.get(kind, 0)
            counters[kind] = idx + 1
        else:
            idx = None
        name_for[os.path.abspath(src)] = build_output_stem(
            clean_stem(src), kind, idx, opts)
    opts.name_for = name_for
    return name_for


def output_path(src, target, opts=None, ext=None):
    """Chemin de sortie pour un fichier source et un format cible.

    ext : extension réelle du fichier de sortie si elle diffère du nom de
    la cible (cas de "part" -> .sldprt / .CATPart / .ipt / .prt). Le
    sous-dossier, lui, garde le nom de la cible (« PART »).
    """
    if opts is None:
        opts = ExportOptions()
    base = opts.dest_dir if opts.dest_dir else export_dir_for(src)
    folder = os.path.join(base, target.upper()) if opts.subfolders else base
    os.makedirs(folder, exist_ok=True)
    real_ext = ext if ext else target
    key = os.path.abspath(src)
    if opts.name_for and key in opts.name_for:
        stem = opts.name_for[key]
    else:
        stem = "%s%s%s" % (opts.prefix, clean_stem(src), opts.suffix)
    return os.path.join(folder, "%s.%s" % (stem, real_ext))


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

    def convert(self, src, targets, opts):
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
    PART_EXT = "sldprt"   # cible "part" : assemblage -> fichier pièce

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

    def convert(self, src, targets, opts):
        results = []
        try:
            ext = os.path.splitext(src.lower())[1]
            doc = self._open(src, ext)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t, opts), False, "ouverture: %s" % e))
            return results

        title = None
        try:
            title = doc.GetTitle()
        except Exception:
            pass

        for t in targets:
            ext_out = self.PART_EXT if t == "part" else t
            out = output_path(src, t, opts, ext=ext_out)
            try:
                ok = self._save(doc, out) and os.path.exists(out)
                results.append((t, out, ok, "" if ok else "SaveAs a échoué"))
            except Exception as e:
                results.append((t, out, False, str(e)))

        # Fermer le document converti (l'utilisateur l'a demandé).
        self._close_doc(doc, title, src)
        return results

    def _close_doc(self, doc, title, src):
        """Ferme le document dans SolidWorks.

        CloseDoc exige le TITRE EXACT de la fenêtre, qui peut inclure ou
        non l'extension selon la configuration. On essaie donc plusieurs
        formes, puis on vérifie que le document n'est plus ouvert.
        """
        path = ""
        try:
            path = doc.GetPathName() or ""
        except Exception:
            pass

        candidates = []

        def add(name):
            if name and name not in candidates:
                candidates.append(name)

        add(title)
        try:
            add(doc.GetTitle())
        except Exception:
            pass
        if path:
            base = os.path.basename(path)
            add(base)
            add(os.path.splitext(base)[0])
        add(os.path.basename(src))
        add(clean_stem(src))

        for name in candidates:
            try:
                self.app.CloseDoc(name)
            except Exception:
                continue

        # Vérification : le document est-il encore ouvert ?
        if path:
            for getter in ("GetOpenDocumentByName2", "GetOpenDocumentByName"):
                try:
                    still = getattr(self.app, getter)(path)
                    if still is None:
                        return True
                    self.log("    (le document est resté ouvert : %s)"
                             % os.path.basename(path))
                    return False
                except Exception:
                    continue
        return True

    def stop(self):
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

    def convert(self, src, targets, opts):
        results = []
        try:
            doc = self.app.Documents.Open(src)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t, opts), False, "ouverture: %s" % e))
            return results

        for t in targets:
            if t == "part":
                out = output_path(src, t, opts, ext="CATPart")
                results.append((t, out, False,
                                "export 'Pièce' non pris en charge ici (SolidWorks uniquement)"))
                continue
            out = output_path(src, t, opts)
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

    def convert(self, src, targets, opts):
        results = []
        try:
            doc = self.app.Documents.Open(src, False)
        except Exception:
            try:
                doc = self.app.Documents.Open(src)
            except Exception as e:
                for t in targets:
                    results.append((t, output_path(src, t, opts), False, "ouverture: %s" % e))
                return results

        to = self.app.TransientObjects
        for t in targets:
            if t == "part":
                out = output_path(src, t, opts, ext="ipt")
                results.append((t, out, False,
                                "export 'Pièce' non pris en charge ici (SolidWorks uniquement)"))
                continue
            out = output_path(src, t, opts)
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
                options = to.CreateNameValueMap()
                med = to.CreateDataMedium()
                med.FileName = out
                try:
                    if addin.HasSaveCopyAsOptions(doc, ctx, options) and t == "step":
                        options.Add("ApplicationProtocolType", 3)  # AP214
                except Exception:
                    pass
                addin.SaveCopyAs(doc, ctx, options, med)
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

    def convert(self, src, targets, opts):
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
                results.append((t, output_path(src, t, opts), False, "ouverture: %s" % e))
            return results

        for t in targets:
            if t == "part":
                out = output_path(src, t, opts, ext="prt")
                results.append((t, out, False,
                                "export 'Pièce' non pris en charge ici (SolidWorks uniquement)"))
                continue
            out = output_path(src, t, opts)
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


def count_exports(groups, selection):
    """Nombre total d'exports prévus pour un jeu de sélections."""
    total = 0
    for items in groups.values():
        for _f, kind in items:
            total += len(selection.get(kind, []))
    return total


def convert_groups(groups, selection, opts=None, log=default_log,
                   progress=None, stop_flag=None):
    """Convertit tous les fichiers groupés par logiciel.

    selection : dict nature -> liste de formats retenus
                (ex. {"part": ["step","stl"], "assembly": ["step"], ...}).
    opts      : ExportOptions.
    progress  : callable(done, total) appelé après chaque export.
    stop_flag : callable() -> bool pour interrompre proprement.

    Renvoie (total_ok, total, failures) où failures est une liste de
    tuples (chemin_source, format, message).
    """
    if opts is None:
        opts = ExportOptions()
    total = count_exports(groups, selection)
    done = total_ok = 0
    failures = []

    for soft, items in groups.items():
        if stop_flag and stop_flag():
            break
        # ne traiter ce logiciel que s'il a au moins un format retenu
        if not any(selection.get(kind) for _f, kind in items):
            continue
        log("=" * 56)
        log("%s : %d fichier(s)" % (SOFTWARE_LABELS.get(soft, soft), len(items)))

        handler = HANDLERS[soft](log=log)
        try:
            handler.start()
        except Exception as e:
            log("  Impossible de démarrer %s : %s" % (handler.name, e))
            for f, kind in items:
                for t in selection.get(kind, []):
                    failures.append((f, t, "démarrage %s: %s" % (handler.name, e)))
                    done += 1
                    if progress:
                        progress(done, total)
            continue

        for f, kind in items:
            if stop_flag and stop_flag():
                break
            targets = list(selection.get(kind, []))
            if not targets:
                continue
            log("  > %s" % os.path.basename(f))
            try:
                res = handler.convert(f, targets, opts)
            except Exception as e:
                log("    ERREUR inattendue : %s" % e)
                res = [(t, output_path(f, t, opts), False, str(e)) for t in targets]
            for t, out, ok, msg in res:
                done += 1
                if ok:
                    total_ok += 1
                    log("    OK  %s -> %s" % (t.upper(), out))
                else:
                    failures.append((f, t, msg))
                    log("    KO  %s : %s" % (t.upper(), msg))
                if progress:
                    progress(done, total)

        try:
            handler.stop()
        except Exception:
            pass

    log("=" * 56)
    log("Terminé : %d/%d export(s) réussi(s)." % (total_ok, total))
    return total_ok, total, failures


# ---------------------------------------------------------------------------
# Renommage par lot (indépendant de la conversion)
# ---------------------------------------------------------------------------

CASE_MODES = ["none", "upper", "lower", "capitalize"]
CASE_LABELS = {
    "none": "(inchangée)",
    "upper": "MAJUSCULES",
    "lower": "minuscules",
    "capitalize": "Première lettre en majuscule",
}


def strip_accents(text):
    """Retire les accents/diacritiques (é -> e, ç -> c)."""
    nf = unicodedata.normalize("NFKD", text)
    return "".join(c for c in nf if not unicodedata.combining(c))


# séparateurs pouvant entourer une numérotation
_SEQ_SEP = " _-."


def strip_sequence(name, position="suffix"):
    """Retire UNIQUEMENT une suite de chiffres (et son séparateur adjacent).

    position="suffix" -> suite en fin de nom ; "prefix" -> suite en début.
    Si aucun chiffre n'est présent à cet endroit, le nom est renvoyé tel quel.
    """
    if position == "prefix":
        m = re.match(r"^\d+", name)
        if m:
            name = name[m.end():].lstrip(_SEQ_SEP)
    else:  # suffix / fin
        m = re.search(r"\d+$", name)
        if m:
            name = name[:m.start()].rstrip(_SEQ_SEP)
    return name


class RenameRules(object):
    """Règles de renommage appliquées au NOM (sans l'extension).

    replacements : liste de couples (rechercher, remplacer_par). Les règles
    sont appliquées successivement, dans l'ordre. Pour « supprimer », il
    suffit d'un remplacement par une chaîne vide.
    """
    def __init__(self, replacements=None, case_sensitive=False,
                 prefix="", suffix="", case_mode="none",
                 spaces_to_underscore=False, remove_accents=False,
                 number_mode="none", number_start=1, number_digits=3,
                 number_position="suffix", number_sep="_", base_name=""):
        self.replacements = list(replacements or [])
        # nom commun : s'il est renseigné, il REMPLACE le nom d'origine
        self.base_name = base_name or ""
        self.case_sensitive = bool(case_sensitive)
        self.prefix = prefix or ""
        self.suffix = suffix or ""
        self.case_mode = case_mode if case_mode in CASE_MODES else "none"
        self.spaces_to_underscore = bool(spaces_to_underscore)
        self.remove_accents = bool(remove_accents)
        # "none" | "add" (ajouter) | "remove" (supprimer la suite)
        self.number_mode = number_mode if number_mode in ("none", "add", "remove") else "none"
        try:
            self.number_start = int(number_start)
        except (TypeError, ValueError):
            self.number_start = 1
        try:
            self.number_digits = max(1, int(number_digits))
        except (TypeError, ValueError):
            self.number_digits = 3
        self.number_position = "prefix" if number_position == "prefix" else "suffix"
        self.number_sep = number_sep if number_sep is not None else "_"


def apply_rules(stem, index, rules):
    """Transforme un nom (sans extension) selon les règles. index = 0,1,2…"""
    # nom commun : remplace entièrement le nom d'origine ; on ignore alors
    # rechercher/remplacer, suppression de numéro et casse (sans objet).
    if rules.base_name:
        s = rules.base_name
    else:
        s = stem

        # 1) rechercher / remplacer (plusieurs règles, dans l'ordre)
        for find, replace in rules.replacements:
            if not find:
                continue
            if rules.case_sensitive:
                s = s.replace(find, replace)
            else:
                s = re.sub(re.escape(find), lambda _m, r=replace: r,
                           s, flags=re.IGNORECASE)

        # 2) suppression d'une numérotation existante (uniquement la suite)
        if rules.number_mode == "remove":
            s = strip_sequence(s, rules.number_position)

        # 3) casse (sur le nom d'origine, pas sur le préfixe/suffixe)
        if rules.case_mode == "upper":
            s = s.upper()
        elif rules.case_mode == "lower":
            s = s.lower()
        elif rules.case_mode == "capitalize":
            s = s.capitalize()

    # 4) assemblage : une numérotation AJOUTÉE se place À L'INTÉRIEUR, juste
    #    après le préfixe ou juste avant le suffixe (jamais au-delà).
    num = ""
    if rules.number_mode == "add":
        num = str(rules.number_start + index).zfill(rules.number_digits)
    if num and rules.number_position == "prefix":
        core = "%s%s%s%s%s" % (rules.prefix, num, rules.number_sep, s, rules.suffix)
    elif num:  # suffix
        core = "%s%s%s%s%s" % (rules.prefix, s, rules.number_sep, num, rules.suffix)
    else:
        core = "%s%s%s" % (rules.prefix, s, rules.suffix)

    # 4) nettoyage sur l'ENSEMBLE (donc aussi le préfixe/suffixe)
    if rules.spaces_to_underscore:
        core = core.replace(" ", "_")
    if rules.remove_accents:
        core = strip_accents(core)

    # sécurité : retirer les caractères interdits dans un nom de fichier
    core = _BAD_NAME_CHARS.sub("", core)
    return core


def build_rename_plan(paths, rules):
    """Construit le plan de renommage.

    Renvoie une liste de dicts :
      { path, folder, old, new, status }
    status : "ok" | "unchanged" | "empty" | "duplicate" | "exists"
    """
    plan = []
    for i, p in enumerate(paths):
        p = os.path.abspath(p)
        folder = os.path.dirname(p)
        base = os.path.basename(p)
        stem, ext = os.path.splitext(base)
        new_stem = apply_rules(stem, i, rules)
        new_base = (new_stem + ext) if new_stem else ""
        plan.append({"path": p, "folder": folder, "old": base,
                     "new": new_base, "status": "ok"})

    # cibles (Windows : comparaison insensible à la casse)
    targets = {}
    sources = set()
    for it in plan:
        sources.add((it["folder"].lower(), it["old"].lower()))
    for it in plan:
        key = (it["folder"].lower(), it["new"].lower())
        targets.setdefault(key, []).append(it)

    for it in plan:
        if not it["new"]:
            it["status"] = "empty"
            continue
        if it["new"] == it["old"]:
            it["status"] = "unchanged"
            continue
        key = (it["folder"].lower(), it["new"].lower())
        if len(targets[key]) > 1:
            it["status"] = "duplicate"
            continue
        target_path = os.path.join(it["folder"], it["new"])
        # conflit seulement si la cible existe ET n'est pas un fichier de
        # notre lot (qui sera renommé ailleurs, géré par les deux passes) ;
        # un simple changement de casse du même fichier est autorisé.
        if os.path.exists(target_path) and key not in sources:
            it["status"] = "exists"
            continue
        it["status"] = "ok"
    return plan


def rename_stats(plan):
    """Compte par statut."""
    stats = {}
    for it in plan:
        stats[it["status"]] = stats.get(it["status"], 0) + 1
    return stats


def execute_rename(plan, log=default_log, progress=None):
    """Applique le renommage SUR PLACE, en deux passes (évite les collisions).

    Ne traite que les entrées « ok ». Renvoie (nb_ok, nb_total, failures)
    où failures est une liste de (old, new, message).
    """
    todo = [it for it in plan if it["status"] == "ok"]
    total = len(todo)
    done = ok = 0
    failures = []
    temps = []

    # Passe 1 : vers des noms temporaires uniques
    for it in todo:
        src = os.path.join(it["folder"], it["old"])
        tmp = os.path.join(it["folder"], it["old"] + ".rntmp_%d" % done)
        n = 0
        while os.path.exists(tmp):
            n += 1
            tmp = os.path.join(it["folder"], it["old"] + ".rntmp_%d_%d" % (done, n))
        try:
            os.rename(src, tmp)
            temps.append((it, tmp))
        except Exception as e:
            failures.append((it["old"], it["new"], "passe 1: %s" % e))
        done += 1
        if progress:
            progress(done, total * 2)

    # Passe 2 : du temporaire vers le nom final
    for it, tmp in temps:
        dst = os.path.join(it["folder"], it["new"])
        try:
            os.rename(tmp, dst)
            ok += 1
            log("  OK  %s -> %s" % (it["old"], it["new"]))
        except Exception as e:
            failures.append((it["old"], it["new"], "passe 2: %s" % e))
            log("  KO  %s -> %s : %s" % (it["old"], it["new"], e))
            # tentative de restauration du nom d'origine
            try:
                os.rename(tmp, os.path.join(it["folder"], it["old"]))
            except Exception:
                pass
        done += 1
        if progress:
            progress(done, total * 2)

    log("Renommage terminé : %d/%d fichier(s)." % (ok, total))
    return ok, total, failures


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

    # En CLI : tous les formats par défaut. Sous-dossiers si plus d'un format
    # distinct est produit.
    selection = {k: list(v) for k, v in TARGETS.items()}
    distinct = set()
    for items in groups.values():
        for _f, kind in items:
            distinct.update(selection.get(kind, []))
    opts = ExportOptions(subfolders=len(distinct) > 1)

    if dry_run:
        n = 0
        for soft, items in groups.items():
            default_log("=" * 56)
            default_log("%s : %d fichier(s)" % (SOFTWARE_LABELS.get(soft, soft), len(items)))
            for f, kind in items:
                for t in selection.get(kind, []):
                    default_log("  [dry-run] %s -> %s" % (os.path.basename(f), output_path(f, t, opts)))
                    n += 1
        default_log("Simulation terminée : %d export(s) prévu(s)." % n)
        return 0

    _ok, _tot, failures = convert_groups(groups, selection, opts, log=default_log)
    return 0 if not failures else 2


def _launch_gui(files=None):
    try:
        import convertisseur_gui
        return convertisseur_gui.main(files)
    except Exception as e:
        print("Interface graphique indisponible (%s)." % e)
        print(__doc__)
        return 1


def main(argv):
    args = list(argv[1:])
    if "--help" in args or "-h" in args:
        print(__doc__)
        return 0
    # conversion sans fenêtre (glisser-déposer « headless », scripts)
    if "--convert" in args:
        return run_cli([a for a in args if a != "--convert"], dry_run=False)
    if "--dry-run" in args:
        return run_cli([a for a in args if a != "--dry-run"], dry_run=True)
    if not args:
        # aucun argument : fenêtre d'accueil
        return _launch_gui(None)
    # des fichiers (ou dossiers) sont fournis : on ouvre la fenêtre de
    # conversion préchargée avec ces fichiers (glisser-déposer, clic droit).
    return _launch_gui(collect_files(args))


if __name__ == "__main__":
    sys.exit(main(sys.argv))


if __name__ == "__main__":
    sys.exit(main(sys.argv))
