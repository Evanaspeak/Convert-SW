#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
convertisseur_cao.py — Convertisseur CAO/DAO par lot (gratuit, local, Windows).

Principe : on ne convertit rien "à la main", on PILOTE les logiciels de CAO
déjà installés via leur automation COM (pywin32). Chaque logiciel n'est
lancé qu'une seule fois, quel que soit le nombre de fichiers à traiter.

Formats de sortie :
    - Pièces / assemblages  -> STEP (.step) + STL (.stl)
    - Mises en plan         -> DWG (.dwg) + DXF (.dxf)

Les fichiers exportés sont rangés dans un sous-dossier « Export » créé à côté
de chaque fichier source.

Logiciels pilotés :
    - SolidWorks  (.sldprt / .sldasm / .slddrw)
    - CATIA V5    (.CATPart / .CATProduct / .CATDrawing)
    - Inventor    (.ipt / .iam / .idw)
    - PTC Creo    (.prt / .asm / .drw, y compris versionnés « .prt.3 »)

Usage :
    python convertisseur_cao.py fichier1.sldprt fichier2.CATPart ...
    python convertisseur_cao.py --dry-run *.sldprt      (n'ouvre aucun logiciel)
    python convertisseur_cao.py --help

Remarque : ce script ne fonctionne que sous Windows avec pywin32 installé et
les logiciels correspondants présents. Le mode --dry-run, lui, tourne partout
et sert à vérifier le routage des fichiers.
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

# nature -> liste des formats de sortie (extension sans point, sert aussi de nom d'export)
TARGETS = {
    "3d": ["step", "stl"],
    "2d": ["dwg", "dxf"],
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
        # on repart du vrai nom (casse d'origine) tronqué à la même longueur
        name = name[:len(m.group("stem"))]
    stem, _ = os.path.splitext(name)
    return stem


def export_dir_for(path):
    """Crée (au besoin) et renvoie le dossier « Export » à côté du fichier."""
    d = os.path.join(os.path.dirname(os.path.abspath(path)), "Export")
    os.makedirs(d, exist_ok=True)
    return d


def output_path(src, target):
    """Chemin de sortie pour un fichier source et un format cible."""
    return os.path.join(export_dir_for(src), clean_stem(src) + "." + target)


# ---------------------------------------------------------------------------
# Journalisation
# ---------------------------------------------------------------------------

class Log:
    def __init__(self):
        self.lines = []

    def __call__(self, msg):
        stamp = time.strftime("%H:%M:%S")
        line = "[%s] %s" % (stamp, msg)
        print(line, flush=True)
        self.lines.append(line)

    def dump(self, folder):
        try:
            p = os.path.join(folder, "conversion.log")
            with open(p, "a", encoding="utf-8") as f:
                f.write("\n".join(self.lines) + "\n")
            return p
        except Exception:
            return None


LOG = Log()


# ---------------------------------------------------------------------------
# Petits utilitaires COM (importés paresseusement : rien n'est chargé tant
# qu'aucun handler n'est démarré, ce qui permet --dry-run hors Windows).
# ---------------------------------------------------------------------------

def _byref_long(value=0):
    """VARIANT entier passé par référence (VT_BYREF | VT_I4).

    C'est LE correctif du DISP_E_TYPEMISMATCH de SolidWorks : les arguments
    [out] « Errors » / « Warnings » de OpenDoc6 / SaveAs doivent être passés
    comme des VARIANT byref, pas comme des entiers Python.
    """
    import pythoncom
    from win32com.client import VARIANT
    return VARIANT(pythoncom.VT_BYREF | pythoncom.VT_I4, value)


# ---------------------------------------------------------------------------
# Handlers
# ---------------------------------------------------------------------------

class Handler(object):
    name = "?"

    def __init__(self):
        self.app = None

    def start(self):
        raise NotImplementedError

    def convert(self, src, targets):
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
        # Dispatch lance SolidWorks s'il n'est pas déjà ouvert, sinon s'y attache.
        self.app = win32com.client.Dispatch("SldWorks.Application")
        self.app.Visible = True
        self._open_mode = None   # on mémorise la 1re convention qui marche
        self._save_mode = None

    # -- ouverture ---------------------------------------------------------
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
        # si une convention a déjà marché, on la tente d'abord
        if self._open_mode:
            modes.sort(key=lambda kv: kv[0] != self._open_mode)

        last_err = None
        for label, fn in modes:
            try:
                doc = fn()
                if doc is not None:
                    self._open_mode = label
                    return doc
                # OpenDoc6 renvoie parfois None mais le doc est actif
                active = self.app.ActiveDoc
                if active is not None:
                    self._open_mode = label
                    return active
            except pythoncom.com_error as e:
                last_err = e
        if last_err:
            raise last_err
        raise RuntimeError("SolidWorks n'a pas pu ouvrir le fichier (doc None).")

    # -- enregistrement ----------------------------------------------------
    def _save(self, doc, outpath):
        import pythoncom
        ver, opt = self._VER_CURRENT, self._SAVE_SILENT_COPY
        ext = doc.Extension

        def m_ext_saveas():
            errs, warns = _byref_long(), _byref_long()
            return ext.SaveAs(outpath, ver, opt, None, errs, warns)

        def m_ext_saveas3():
            errs, warns = _byref_long(), _byref_long()
            # SaveAs3(Name, Version, Options, ExportData, AdvancedOptions, Errors, Warnings)
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
                ret = fn()
                # ret peut être bool, ou un tuple (bool, errs, warns) selon la
                # convention. On considère l'écriture du fichier comme preuve.
                self._save_mode = label
                return True
            except pythoncom.com_error as e:
                last_err = e
        if last_err:
            raise last_err
        return False

    def convert(self, src, targets):
        results = []
        try:
            ext = os.path.splitext(src.lower())[1]
            doc = self._open(src, ext)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t), False, "ouverture: %s" % e))
            return results

        title = None
        try:
            title = doc.GetTitle()
        except Exception:
            pass

        for t in targets:
            out = output_path(src, t)
            try:
                ok = self._save(doc, out)
                ok = ok and os.path.exists(out)
                results.append((t, out, ok, "" if ok else "SaveAs a échoué"))
            except Exception as e:
                results.append((t, out, False, str(e)))

        # fermeture du document (on garde SolidWorks ouvert pour les suivants)
        try:
            if title:
                self.app.CloseDoc(title)
        except Exception:
            pass
        return results

    def stop(self):
        # On ne ferme pas SolidWorks : plus rapide si relancé, et évite de
        # tuer une session que l'utilisateur avait déjà ouverte.
        pass


# --------------------------------- CATIA -----------------------------------

class CatiaHandler(Handler):
    name = "CATIA V5"

    # target -> liste de chaînes de format à essayer (l'orthographe varie
    # selon la version / les licences installées).
    _FORMATS = {
        "step": ["stp", "step"],
        "stl":  ["stl"],
        "dwg":  ["dwg"],
        "dxf":  ["dxf"],
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

    def convert(self, src, targets):
        results = []
        doc = None
        try:
            doc = self.app.Documents.Open(src)
        except Exception as e:
            for t in targets:
                results.append((t, output_path(src, t), False, "ouverture: %s" % e))
            return results

        for t in targets:
            out = output_path(src, t)
            ok, msg = False, ""
            for fmt in self._FORMATS[t]:
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
    # On cherche d'abord par GUID (stable), puis par DisplayName (repli).
    _TRANSLATORS = {
        "step": ("STEP", "{90AF7F40-0C01-11D5-B79E-0010B3AF2C1B}"),
        "stl":  ("STL",  "{81D58C09-F638-42B4-BD3E-F5DAC29F94D1}"),
        "dwg":  ("DWG",  "{C24E3AC2-122E-11D5-8E91-0010B541CD80}"),
        "dxf":  ("DXF",  "{C24E3AC4-122E-11D5-8E91-0010B541CD80}"),
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
        # 1) par GUID (ClassIdString)
        for a in addins:
            try:
                cid = (a.ClassIdString or "").upper()
                if cid == guid.upper():
                    found = a
                    break
            except Exception:
                continue
        # 2) par DisplayName (contient le mot-clé, insensible à la casse)
        if found is None:
            for a in addins:
                try:
                    dn = (a.DisplayName or "")
                    if keyword.lower() in dn.lower():
                        found = a
                        break
                except Exception:
                    continue
        self._addin_cache[target] = found
        return found

    def convert(self, src, targets):
        results = []
        doc = None
        try:
            doc = self.app.Documents.Open(src, False)  # False = ne pas afficher
        except Exception:
            try:
                doc = self.app.Documents.Open(src)
            except Exception as e:
                for t in targets:
                    results.append((t, output_path(src, t), False, "ouverture: %s" % e))
                return results

        to = self.app.TransientObjects
        for t in targets:
            out = output_path(src, t)
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
                # Options éventuelles (protocole STEP AP214, etc.)
                try:
                    if addin.HasSaveCopyAsOptions(doc, ctx, opts) and t == "step":
                        # 3 = AP214 (auto/AP203 selon version) — on reste souple.
                        # NameValueMap : on renseigne via Add (map vide au départ).
                        opts.Add("ApplicationProtocolType", 3)
                except Exception:
                    pass
                addin.SaveCopyAs(doc, ctx, opts, med)
                ok = os.path.exists(out)
                results.append((t, out, ok, "" if ok else "SaveCopyAs sans fichier produit"))
            except Exception as e:
                results.append((t, out, False, str(e)))

        try:
            doc.Close(True)  # True = fermer sans enregistrer
        except Exception:
            pass
        return results


# ---------------------------------- Creo -----------------------------------

class CreoHandler(Handler):
    """Pilotage de PTC Creo Parametric via la VB API (Object TOOLKIT), gratuite.

    Prérequis côté machine :
      - Creo doit être DÉJÀ LANCÉ (la connexion asynchrone s'attache à la
        session en cours ; on ne démarre pas Creo ici).
      - La VB API doit être installée et enregistrée (pfcls).

    Vu la variabilité des ProgID / coclasses selon la version de Creo, la
    connexion est tentée sur plusieurs identifiants. À VÉRIFIER sur la machine.
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
                LOG("  Creo : connexion via %s" % progid)
                break
            except Exception as e:
                last = e
        if self.session is None:
            raise RuntimeError(
                "Impossible de se connecter à Creo (VB API). "
                "Creo est-il lancé et la VB API installée ? Dernière erreur : %s" % last)

    def _instructions(self, target):
        """Crée les instructions d'export pour le format demandé."""
        w = self._win32
        if target == "step":
            return w.Dispatch("pfcls.CCpfcSTEP3DExportInstructions").Create()
        if target == "stl":
            # STL binaire par défaut
            return w.Dispatch("pfcls.CCpfcSTLBinaryExportInstructions").Create()
        if target == "dwg":
            return w.Dispatch("pfcls.CCpfcDWGExportInstructions").Create()
        if target == "dxf":
            return w.Dispatch("pfcls.CCpfcDXFExportInstructions").Create()
        raise ValueError(target)

    def convert(self, src, targets):
        results = []
        model = None
        try:
            # RetrieveModelWithOpts / RetrieveModel selon version ; on récupère
            # le modèle par son nom de fichier.
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
                results.append((t, output_path(src, t), False, "ouverture: %s" % e))
            return results

        for t in targets:
            out = output_path(src, t)
            try:
                instr = self._instructions(t)
                model.Export(out, instr)
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
# Programme principal
# ---------------------------------------------------------------------------

def collect_files(paths):
    """Développe les arguments en liste de fichiers existants et gérés."""
    files = []
    for p in paths:
        p = p.strip('"')
        if os.path.isdir(p):
            for root, dirs, names in os.walk(p):
                # on n'entre pas dans les dossiers Export ni les dossiers cachés
                dirs[:] = [d for d in dirs
                           if d.lower() != "export" and not d.startswith(".")]
                for n in names:
                    files.append(os.path.join(root, n))
        elif os.path.isfile(p):
            files.append(p)
        else:
            LOG("Introuvable, ignoré : %s" % p)
    return files


def group_by_software(files):
    groups = {}
    skipped = []
    for f in files:
        route = route_for(f)
        if route is None:
            skipped.append(f)
            continue
        soft, kind = route
        groups.setdefault(soft, []).append((f, kind))
    return groups, skipped


def run(paths, dry_run=False):
    files = collect_files(paths)
    if not files:
        LOG("Aucun fichier à traiter.")
        return 1

    groups, skipped = group_by_software(files)
    for f in skipped:
        LOG("Format non géré, ignoré : %s" % os.path.basename(f))

    if not groups:
        LOG("Aucun fichier CAO/DAO reconnu.")
        return 1

    total_ok = total = 0
    for soft, items in groups.items():
        LOG("=" * 60)
        LOG("%s : %d fichier(s)" % (HANDLERS[soft].name, len(items)))
        if dry_run:
            for f, kind in items:
                for t in TARGETS[kind]:
                    LOG("  [dry-run] %s -> %s" % (os.path.basename(f), output_path(f, t)))
                    total += 1
            continue

        handler = HANDLERS[soft]()
        try:
            handler.start()
        except Exception as e:
            LOG("  Impossible de démarrer %s : %s" % (handler.name, e))
            for f, kind in items:
                total += len(TARGETS[kind])
            continue

        for f, kind in items:
            LOG("  > %s" % os.path.basename(f))
            try:
                res = handler.convert(f, TARGETS[kind])
            except Exception as e:
                LOG("    ERREUR inattendue : %s" % e)
                LOG(traceback.format_exc())
                res = [(t, output_path(f, t), False, str(e)) for t in TARGETS[kind]]
            for t, out, ok, msg in res:
                total += 1
                if ok:
                    total_ok += 1
                    LOG("    OK  %s -> %s" % (t.upper(), out))
                else:
                    LOG("    KO  %s : %s" % (t.upper(), msg))

        try:
            handler.stop()
        except Exception:
            pass

    LOG("=" * 60)
    if dry_run:
        LOG("Simulation terminée : %d export(s) prévu(s)." % total)
        return 0
    LOG("Terminé : %d/%d export(s) réussi(s)." % (total_ok, total))
    # journal à côté du premier fichier traité
    try:
        LOG.dump(export_dir_for(files[0]))
    except Exception:
        pass
    return 0 if total_ok == total else 2


HELP = __doc__


def main(argv):
    args = [a for a in argv[1:]]
    if not args or "--help" in args or "-h" in args:
        print(HELP)
        return 0
    dry = False
    if "--dry-run" in args:
        dry = True
        args = [a for a in args if a != "--dry-run"]
    return run(args, dry_run=dry)


if __name__ == "__main__":
    sys.exit(main(sys.argv))
