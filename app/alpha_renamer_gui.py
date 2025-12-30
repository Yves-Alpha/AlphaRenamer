#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os, sys, csv, time, re, traceback
import subprocess
import tempfile
import shutil
from pathlib import Path
from dataclasses import dataclass
from typing import Optional, List, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD  # type: ignore
    HAS_DND = True
except Exception:  # pragma: no cover - dépendance optionnelle
    DND_FILES = None
    TkinterDnD = None
    HAS_DND = False

from pdfminer.high_level import extract_text
from pypdf import PdfReader, PdfWriter
from rapidfuzz import fuzz
from rename_from_lexique import LexiqueError, rename_with_lexique

# ----------------- Config par défaut -----------------
APP_DIR = Path(__file__).resolve().parent
DEFAULT_OUT = Path.home() / "AlphaRenamer-OUT"
LOG_NAME = lambda: f"log_renommage_{time.strftime('%Y%m%d-%H%M%S')}.csv"

RE_OP   = re.compile(r"OP[ÉE]RATION\s*N[°º]?\s*([A-Z0-9]+)", re.IGNORECASE)
RE_AN   = re.compile(r"\b(20\d{2})\b")
RE_CODE = re.compile(r"Codification.*?:\s*/\s*(\d{5})", re.IGNORECASE)
RE_CODE5= re.compile(r"\b(\d{5})\b")

# -- helpers pour adresse/noms --
RE_POSTAL = re.compile(r"\b\d{5}\b")
RE_STREET_HINT = re.compile(
    r"\b(?:RUE|AVENUE|AV\.?|AVE\.?|BD|BOULEVARD|PLACE|CHEMIN|IMPASSE|ALL(?:EE|ÉE)S?|"
    r"QUAI|ROUTE|PL|SQUARE|PASSAGE|PROMENADE|VOIE|ROND[ -]?POINT|C\.C\.?|CC)\b",
    re.IGNORECASE,
)

TH_NOM  = 80  # tolérance nom si heuristique

# Police par défaut (évite le texte invisible sur certaines configs Tk/macOS)
# -----------------------------------------------------

@dataclass
class PageInfo:
    op: Optional[str]
    code: Optional[str]
    nom: Optional[str]
    an: Optional[str]
    note: str


def sanitize(name: str) -> str:
    name = name.strip()
    # Conserver uniquement lettres/chiffres/underscore/espace/tiret, supprimer le reste
    # (D.S.P.L. -> DSPL, etc.).
    name = re.sub(r"[^\w\s-]", "", name, flags=re.UNICODE)
    name = re.sub(r"\s+", " ", name)
    return name.strip(" ._-")

def unique_path(path: Path) -> Path:
    if not path.exists():
        return path
    stem, suf = path.stem, path.suffix
    i = 1
    while True:
        cand = path.with_name(f"{stem} ({i}){suf}")
        if not cand.exists():
            return cand
        i += 1


def get_bold_lines_for_page(pdf_path: Path, page_index: int) -> List[str]:
    """Retourne les lignes de texte en gras sur une page, si pdfminer le permet."""
    try:
        from pdfminer.high_level import extract_pages  # type: ignore
        from pdfminer.layout import LTTextContainer, LTChar  # type: ignore
    except Exception:
        return []

    bold_lines: List[str] = []
    try:
        for page_layout in extract_pages(str(pdf_path), page_numbers=[page_index]):
            for element in page_layout:
                if isinstance(element, LTTextContainer):
                    for text_line in element:
                        line_text = text_line.get_text().replace("\n", "").strip()
                        if not line_text:
                            continue
                        is_bold = False
                        for obj in text_line:
                            if isinstance(obj, LTChar):
                                fontname = getattr(obj, "fontname", "") or ""
                                fname = fontname.lower()
                                if "bold" in fname or "bd" in fname or "black" in fname:
                                    is_bold = True
                                    break
                        if is_bold:
                            bold_lines.append(line_text)
    except Exception:
        return []
    return bold_lines


def extract_fields_from_text(txt: str, bold_lines: Optional[List[str]] = None) -> PageInfo:
    # OP
    m_op = RE_OP.search(txt)
    op = m_op.group(1) if m_op else None

    # année (on prend la plus grande 20xx trouvée)
    years = [int(y) for y in RE_AN.findall(txt)]
    an = str(max(years)) if years else None

    # code client
    code = None
    m_c = RE_CODE.search(txt)
    if m_c:
        code = m_c.group(1)
    else:
        m_c2 = RE_CODE5.search(txt)
        if m_c2:
            code = m_c2.group(1)

    # nom client : priorité au bloc adresse
    # Schéma attendu :
    #   i:   enseigne (souvent "G20") OU directement NOM MAGASIN
    #   i+1: NOM MAGASIN si G20 en i, sinon voie
    #   i+2: voie (souvent commence par un numéro ou contient un mot de voie)
    #   i+3: code postal + ville (contient un code postal à 5 chiffres)
    nom = None
    lines = [l.strip() for l in txt.splitlines() if l.strip()]

    bold_lines = bold_lines or []

    def is_bold(line: str) -> bool:
        line = line.strip()
        if not line:
            return False
        for bl in bold_lines:
            if line == bl:
                return True
            if len(line) >= 4 and len(bl) >= 4:
                try:
                    if fuzz.ratio(line, bl) >= 90:
                        return True
                except Exception:
                    # Par sécurité, on ignore les erreurs de rapidfuzz ici
                    pass
        return False

    def is_street_line(line: str) -> bool:
        s = line.strip()
        if not s:
            return False
        parts = s.split()
        first_token = parts[0]
        # 1) Ligne commençant par un numéro séparé ("5 RUE ...")
        if first_token.isdigit():
            return True
        # 2) Cas sans espace après le numéro ("5AVENUE...", "12BD..."), insensible à la casse
        if re.match(r"^\d{1,4}[A-Z]", s, re.IGNORECASE):
            return True
        # 3) Mot-clé de voie (AV., AVE., CC, C.C., ROND-POINT, etc.)
        return RE_STREET_HINT.search(s) is not None

    # Recherche du bloc adresse robuste
    for i in range(0, max(0, len(lines) - 3)):
        l0, l1, l2, l3 = lines[i], lines[i+1], lines[i+2], lines[i+3]
        # L0 : première ligne du bloc (enseigne ou nom magasin)
        has_g20 = l0.upper() in ("G20", "G 20", "G‑20", "G–20")
        # Autoriser ponctuation fréquente dans les noms (.,:/()* etc.)
        l0_ok = has_g20 or re.fullmatch(r"[A-Z0-9À-ÖØ-Ý '&()\-./:*]{2,}", l0) is not None

        # Cas standard sur 4 lignes :
        # L2 doit ressembler à une voie, L3 à une ligne de fin d'adresse (CP ou "Paris le ...")
        l2_ok = is_street_line(l2)
        l3_ok = RE_POSTAL.search(l3) is not None or l3.lower().startswith("paris le")
        if l0_ok and l2_ok and l3_ok:
            # Bloc adresse trouvé :
            # - si G20 sur la 1re ligne, on garde la 2e (nom magasin),
            # - sinon, on prend la 1re ligne du bloc comme NomClient.
            if has_g20:
                nom = l1
            else:
                nom = l0
            break

        # Cas assoupli sur 5 lignes :
        # G20 / NOM / voie1 / complément / CP+ville
        if i + 4 < len(lines) and l0_ok:
            l4 = lines[i+4]
            street_block_ok = is_street_line(l2) or is_street_line(l3)
            l4_ok = RE_POSTAL.search(l4) is not None or l4.lower().startswith("paris le")
            if street_block_ok and l4_ok:
                if has_g20:
                    nom = l1
                else:
                    nom = l0
                break
        
    # Repli : ancienne heuristique (ligne majuscules plausibles en haut de page)
    if not nom:
        best = None
        best_idx: Optional[int] = None
        for idx, ln in enumerate(lines[:120]):
            if len(ln) < 3:
                continue
            if re.fullmatch(r"[A-Z0-9À-ÖØ-Ý '()\-.,/*:/]+", ln) and not ln.isdigit():
                if best is None or len(ln) > len(best):
                    best = ln
                    best_idx = idx
        if best is not None:
            cand = best
            stripped = cand.lstrip()
            # Un NomClient ne peut pas commencer par un chiffre : si c'est le cas, on prend la ligne du dessus
            if stripped and stripped[0].isdigit() and best_idx is not None and best_idx > 0:
                prev = lines[best_idx - 1].strip()
                if prev:
                    cand = prev
            nom = cand

    note = "PDF text v20251118-2"
    return PageInfo(op=op, code=code, nom=nom, an=an, note=note)

def extract_page_text(pdf_path: Path, page_index: int) -> str:
    # pdfminer n’extrait pas “par page” directement via high_level; on coupe via PyPDF
    # 1) on extrait la page avec PyPDF, 2) PyPDF sait extraire du texte page par page.
    try:
        reader = PdfReader(str(pdf_path))
        page = reader.pages[page_index]
        # PyPDF text extraction (suffisant pour Word/texte)
        return page.extract_text() or ""
    except Exception:
        # repli : extraire tout et heuristiquement fractionner (moins fiable)
        try:
            return extract_text(str(pdf_path)) or ""
        except Exception:
            return ""

def split_and_process(pdf_path: Path, out_dir: Path, errors_dir: Path, log_writer, dry_run: bool):
    try:
        reader = PdfReader(str(pdf_path))
    except Exception as e:
        log_writer.writerow([pdf_path.name, "", "", "", "", "", "ERROR", f"Lecture PDF: {e}"])
        return

    n = len(reader.pages)
    if n == 0:
        log_writer.writerow([pdf_path.name, "", "", "", "", "", "ERROR", "PDF vide"])
        return

    for i in range(n):
        page_txt = extract_page_text(pdf_path, i)
        bold_lines = get_bold_lines_for_page(pdf_path, i)
        info = extract_fields_from_text(page_txt, bold_lines=bold_lines)

        # Sécurité supplémentaire : si le NomClient ressemble à une ligne technique
        # (ex: "DÉSIGNATION MONTANT HT"), on l’invalide pour forcer la page en erreur.
        if info.nom:
            up_nom = info.nom.strip().upper()
            if up_nom.startswith("DÉSIGNATION") or up_nom.startswith("DESIGNATION"):
                info.nom = None

        if not (info.op and info.code and info.nom and info.an):
            # en erreur -> page extraite vers _ERREURS avec suffixe
            out_err = unique_path(errors_dir / f"{pdf_path.stem}_page{i+1}.pdf")
            if not dry_run:
                writer = PdfWriter()
                writer.add_page(reader.pages[i])
                with open(out_err, "wb") as f:
                    writer.write(f)
            log_writer.writerow([pdf_path.name, info.op or "", info.code or "", info.nom or "", info.an or "", out_err.name, "ERROR", "Champs manquants"])
            continue

        base = f"ALPHA_OP{sanitize(info.op)}_{sanitize(info.code)}_{sanitize(info.nom)}_{sanitize(info.an)}"
        out_ok = unique_path(out_dir / f"{base}.pdf")
        if not dry_run:
            writer = PdfWriter()
            writer.add_page(reader.pages[i])
            with open(out_ok, "wb") as f:
                writer.write(f)
        log_writer.writerow([pdf_path.name, info.op, info.code, info.nom, info.an, out_ok.name, "RENAMED" if not dry_run else "DRY-RUN", info.note])

# ----------------- GUI -----------------
if HAS_DND and TkinterDnD is not None:
    TK_BASE = TkinterDnD.Tk  # type: ignore[attr-defined]
else:
    TK_BASE = tk.Tk


class App(TK_BASE):
    def __init__(self):
        super().__init__()
        self.title("AlphaRenamer")

        self.geometry("560x420")
        self.minsize(520, 380)

        self.output_dir: Optional[Path] = None
        self.files: List[Path] = []
        self.dry_run_var = tk.BooleanVar(value=True)
        self.auto_lexique_var = tk.BooleanVar(value=False)
        self.lexique_path: Optional[Path] = None

        # UI
        frm = tk.Frame(self, padx=12, pady=12)
        frm.pack(fill="both", expand=True)

        # Dossier de sortie
        out_row = tk.Frame(frm)
        out_row.pack(fill="x")
        tk.Label(out_row, text="Dossier de sortie :").pack(side="left")
        self.out_label = tk.Label(out_row, text="(non défini)")
        self.out_label.pack(side="left", padx=8)
        tk.Button(out_row, text="Choisir…", command=self.choose_output).pack(side="right")

        # Liste des fichiers
        list_row = tk.Frame(frm)
        list_row.pack(fill="both", expand=True, pady=(12, 8))

        tk.Label(list_row, text="Fichiers à traiter (PDF ou Word) :").pack(anchor="w")

        lb_wrap = tk.Frame(list_row)
        lb_wrap.pack(fill="both", expand=True)

        scroll = tk.Scrollbar(lb_wrap, orient="vertical")
        scroll.pack(side="right", fill="y")

        self.listbox = tk.Listbox(
            lb_wrap,
            height=14,
            selectmode="extended",
            exportselection=False,
            activestyle="dotbox",
        )
        self.listbox.pack(side="left", fill="both", expand=True)
        self.listbox.config(yscrollcommand=scroll.set)
        scroll.config(command=self.listbox.yview)

        # Boutons de gestion de la liste
        btns = tk.Frame(frm)
        btns.pack(fill="x", pady=(6, 6))
        tk.Button(btns, text="Ajouter des fichiers…", command=self.add_files).pack(side="left")
        tk.Button(btns, text="Retirer la sélection", command=self.remove_selected).pack(side="left", padx=8)
        tk.Button(btns, text="Vider la liste", command=self.clear_list).pack(side="left", padx=8)

        # Options
        opt_row = tk.Frame(frm)
        opt_row.pack(fill="x", pady=(6,6))
        tk.Checkbutton(opt_row, text="Simulation (dry-run)", variable=self.dry_run_var).pack(side="left")

        # Lexique
        lex_row = tk.Frame(frm)
        lex_row.pack(fill="x", pady=(4, 2))
        tk.Label(lex_row, text="Lexique (Excel) :").pack(side="left")
        self.lex_label = tk.Label(lex_row, text="(aucun)")
        self.lex_label.pack(side="left", padx=8)
        tk.Button(lex_row, text="Choisir…", command=self.choose_lexique).pack(side="right")

        lex_actions = tk.Frame(frm)
        lex_actions.pack(fill="x", pady=(0, 8))
        tk.Checkbutton(lex_actions, text="Renommer via lexique après traitement", variable=self.auto_lexique_var).pack(side="left")
        tk.Button(lex_actions, text="Renommer avec le lexique maintenant", command=self.run_lexique).pack(side="right")

        # Lancer
        action_row = tk.Frame(frm)
        action_row.pack(fill="x")
        self.run_btn = tk.Button(action_row, text="Lancer le traitement", command=self.run)
        self.run_btn.pack(side="right")

        self.ok_dir = None
        self.err_dir = None

        # DnD natif (fenêtre + liste), si disponible
        self._init_dnd()


    def choose_output(self):
        path = filedialog.askdirectory(title="Choisir le dossier de sortie")
        if not path:
            return
        self.output_dir = Path(path)
        self.ok_dir = self.output_dir
        self.err_dir = self.output_dir / "_ERREURS"
        self.ok_dir.mkdir(parents=True, exist_ok=True)
        self.err_dir.mkdir(parents=True, exist_ok=True)
        self.out_label.config(text=str(self.output_dir))

    def choose_lexique(self):
        path = filedialog.askopenfilename(
            title="Choisir le fichier lexique (LEXIQUE.xlsx)",
            filetypes=[("Excel", "*.xlsx *.xls"), ("Tous", "*.*")],
        )
        if not path:
            return
        self.lexique_path = Path(path)
        self.lex_label.config(text=self.lexique_path.name)

    def run_lexique(self):
        res = self.apply_lexique(dry_run=self.dry_run_var.get())
        if res:
            messagebox.showinfo("AlphaRenamer", f"Renommage lexique terminé.\n{res}")

    def _add_paths(self, paths: List[str]):
        for p in paths:
            pth = Path(p)
            ext = pth.suffix.lower()
            if ext in (".pdf", ".docx", ".doc") and pth not in self.files:
                self.files.append(pth)
                self.listbox.insert("end", str(pth))

    def add_files(self):
        filetypes = [
            ("Documents", "*.pdf *.PDF *.docx *.DOCX *.doc *.DOC"),
            ("PDF", "*.pdf *.PDF"),
            ("Word", "*.docx *.DOCX *.doc *.DOC"),
            ("Tous", "*.*"),
        ]
        paths = filedialog.askopenfilenames(title="Sélectionne des fichiers (PDF ou Word)", filetypes=filetypes)
        if not paths:
            return
        self._add_paths(list(paths))

    def remove_selected(self):
        sel_idx = list(self.listbox.curselection())
        sel_idx.reverse()
        for i in sel_idx:
            try:
                p = Path(self.listbox.get(i))
                if p in self.files:
                    self.files.remove(p)
                self.listbox.delete(i)
            except Exception:
                pass

    def clear_list(self):
        self.files.clear()
        self.listbox.delete(0, "end")

    def _split_dnd_data(self, data: str) -> List[str]:
        """Découpe la chaîne DND_FILES (avec ou sans { } autour des chemins)."""
        res: List[str] = []
        buf = ""
        in_brace = False
        for ch in data:
            if ch == "{":
                in_brace = True
                buf = ""
            elif ch == "}":
                in_brace = False
                if buf:
                    res.append(buf)
                    buf = ""
            elif ch == " " and not in_brace:
                if buf:
                    res.append(buf)
                    buf = ""
            else:
                buf += ch
        if buf:
            res.append(buf)
        return res

    def _init_dnd(self):
        """Active le drag-and-drop de fichiers si tkinterdnd2 est disponible."""
        if not HAS_DND or DND_FILES is None:
            return
        for w in (self, self.listbox):
            try:
                w.drop_target_register(DND_FILES)
                w.dnd_bind("<<Drop>>", self.on_drop_files)
            except Exception:
                pass

    def on_drop_files(self, event):
        data = getattr(event, "data", None)
        if not data:
            return
        paths = self._split_dnd_data(str(data))
        if not paths:
            return
        self._add_paths(paths)

    def apply_lexique(self, dry_run: bool) -> Optional[str]:
        if not self.output_dir or not self.ok_dir:
            messagebox.showwarning("AlphaRenamer", "Définis d’abord un dossier de sortie.")
            return None
        if not self.lexique_path:
            messagebox.showwarning("AlphaRenamer", "Choisis d’abord un fichier de lexique (LEXIQUE.xlsx).")
            return None
        try:
            renamed, skipped, errors = rename_with_lexique(
                self.ok_dir,
                self.lexique_path,
                dry_run=dry_run,
            )
        except LexiqueError as e:
            messagebox.showerror("AlphaRenamer", f"Renommage via lexique impossible :\n{e}")
            return None
        except Exception as e:
            messagebox.showerror("AlphaRenamer", f"Erreur inattendue pendant le renommage :\n{e}")
            return None

        label = "fichiers renommés (simulation)" if dry_run else "fichiers renommés"
        return f"{renamed} {label}, {skipped} inchangés, {errors} erreurs"

    def run(self):
        if not self.output_dir:
            messagebox.showwarning("AlphaRenamer", "Définis d’abord un dossier de sortie.")
            return
        if not self.files:
            messagebox.showwarning("AlphaRenamer", "Ajoute au moins un fichier (PDF ou Word).")
            return

        dry_run = self.dry_run_var.get()
        log_path = self.ok_dir / LOG_NAME()
        processed = 0
        errors = 0
        convert_errors: List[Tuple[str, str]] = []

        def convert_word_to_pdf(src: Path, dest_dir: Path) -> Optional[Path]:
            """Convertit un fichier Word en PDF et retourne le chemin cible, ou None si échec."""
            try:
                dest_dir.mkdir(parents=True, exist_ok=True)
            except Exception:
                pass
            out_pdf = dest_dir / (src.stem + ".pdf")

            # 1) docx2pdf si dispo
            try:
                from docx2pdf import convert as docx2pdf_convert  # type: ignore
                # docx2pdf accepte src et dossier de sortie
                docx2pdf_convert(str(src), str(dest_dir))
                if out_pdf.exists():
                    return out_pdf
            except Exception:
                pass

            # 2) LibreOffice/soffice si dispo
            soffice = shutil.which("soffice") or shutil.which("libreoffice")
            # Cherche aussi les emplacements macOS courants si non trouvés dans PATH
            if not soffice and sys.platform == "darwin":
                mac_candidates = [
                    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
                    "/Applications/LibreOffice.app/Contents/MacOS/loffice",
                ]
                for cand in mac_candidates:
                    if os.path.exists(cand):
                        soffice = cand
                        break
            if soffice:
                try:
                    # --headless conversion
                    subprocess.run(
                        [soffice, "--headless", "--convert-to", "pdf", "--outdir", str(dest_dir), str(src)],
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                    )
                    if out_pdf.exists():
                        return out_pdf
                except Exception:
                    pass

            # 3) macOS + Microsoft Word via osascript
            if sys.platform == "darwin":
                try:
                    apple_script = (
                        "on run argv\n"
                        "    set inFile to POSIX file (item 1 of argv) as alias\n"
                        "    set outPath to item 2 of argv\n"
                        "    tell application \"Microsoft Word\"\n"
                        "        activate\n"
                        "        set theDoc to open inFile\n"
                        "        save as theDoc file format format PDF file name outPath\n"
                        "        close theDoc saving no\n"
                        "    end tell\n"
                        "end run\n"
                    )
                    script_path = dest_dir / "convert_word_to_pdf.scpt"
                    script_path.write_text(apple_script)
                    subprocess.run(
                        ["osascript", str(script_path), str(src), str(out_pdf)],
                        check=True,
                        stdout=subprocess.PIPE,
                        stderr=subprocess.PIPE,
                    )
                    if out_pdf.exists():
                        return out_pdf
                except Exception:
                    pass

            return None

        try:
            with open(log_path, "w", encoding="utf-8", newline="") as lf:
                lw = csv.writer(lf, delimiter=";")
                lw.writerow(["pdf_source", "OP", "CodeClient", "NomClient", "Annee", "pdf_sortie", "action", "note"])
                with tempfile.TemporaryDirectory(prefix="alpharenamer_") as tmpd:
                    tmp_dir = Path(tmpd)
                    for f in self.files:
                        try:
                            src = Path(f)
                            ext = src.suffix.lower()
                            pdf_to_process = src
                            if ext in (".docx", ".doc"):
                                conv = convert_word_to_pdf(src, tmp_dir)
                                if not conv:
                                    convert_errors.append((src.name, "Conversion Word→PDF impossible (docx2pdf/LibreOffice/Word non disponible)"))
                                    errors += 1
                                    lw.writerow([src.name, "", "", "", "", "", "ERROR", "Conversion Word→PDF échouée"])
                                    continue
                                pdf_to_process = conv

                            split_and_process(pdf_to_process, self.ok_dir, self.err_dir, lw, dry_run=dry_run)
                            processed += 1
                        
                        except Exception as e:
                            errors += 1
                            lw.writerow([Path(f).name, "", "", "", "", "", "ERROR", f"{e}"])
            msg = f"Traitement terminé.\nFichiers: {processed}\nJournal: {log_path}"
            if errors:
                msg += f"\nErreurs: {errors} (voir dossier _ERREURS et le log)"
            if convert_errors:
                msg += "\n\nConversions Word échouées :\n" + "\n".join([f"- {name} : {why}" for name, why in convert_errors])
            lex_msg = ""
            if self.auto_lexique_var.get():
                lex_res = self.apply_lexique(dry_run=dry_run)
                if lex_res:
                    lex_msg = "\n\nRenommage via lexique :\n" + lex_res

            messagebox.showinfo("AlphaRenamer", msg + lex_msg)
        except Exception as e:
            messagebox.showerror("AlphaRenamer", f"Erreur pendant le traitement:\n{e}\n\n{traceback.format_exc()}")

if __name__ == "__main__":
    app = App()
    # Fichiers passés sur la ligne de commande (ex : DnD sur l’icône .app)
    if len(sys.argv) > 1:
        app._add_paths(sys.argv[1:])
    app.mainloop()
