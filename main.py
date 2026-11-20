#!/usr/bin/env python3
"""
TALdoc Desktop - App locale Windows per interrogare documenti con AI locale
Completamente offline. Ideale per ragionieri e professionisti.
Uso base immediato: carica documenti, fai domande, verifica conti, confronta.
"""

import os
import sys
import json
import re
from pathlib import Path
from datetime import datetime
from typing import List, Optional

import threading

import customtkinter as ctk
from tkinter import filedialog, messagebox, Menu
import requests

try:
    from PIL import Image
except ImportError:
    Image = None

# Tentativo di importare psutil per monitoraggio RAM
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

# Librerie opzionali per estrazione testo
try:
    from pypdf import PdfReader
except ImportError:
    PdfReader = None
try:
    from docx import Document as DocxDocument
except ImportError:
    DocxDocument = None
try:
    import openpyxl
except ImportError:
    openpyxl = None
try:
    from pptx import Presentation
except ImportError:
    Presentation = None
try:
    import pandas as pd
except ImportError:
    pd = None

# OCR opzionale
try:
    from paddleocr import PaddleOCR
    import cv2
    PADDLE_AVAILABLE = True
except ImportError:
    PADDLE_AVAILABLE = False

APP_NAME = "TALdoc Desktop"
OLLAMA_URL = "http://localhost:11434"
DEFAULT_MODEL = "qwen2.5:3b"  # leggero, buon rapporto

BORDER_COLORS = {
    "word": "#1F6BFF",
    "powerpoint": "#FF6A00",
    "pdf": "#FF1F3D",
    "excel": "#00B341",
    "txt": "#FFC400",
    "other": "#8B5CF6"
}

MANUALE = """
TALdoc Desktop - Manuale d'uso

1. COS'È TALDOC
TALdoc è un'app per Windows che legge i tuoi documenti in locale (senza internet) e ti permette di interrogarli con un'intelligenza artificiale locale (tramite Ollama).
Puoi fare domande libere, chiedere calcoli, verificare totali e IVA, confrontare due documenti e trovare differenze.

Tutto resta sul tuo PC. Nessun dato viene inviato online.

2. PRIMO AVVIO
- Assicurati di avere Ollama installato (scaricalo da ollama.com e avvialo).
- Al primo avvio scegli l'aspetto: chiaro (testo nero) oppure scuro (testo chiaro).
- L'app analizza processore, RAM ed eventuale scheda video NVIDIA e ti mostra i modelli AI
  con la qualità e i secondi di attesa stimati per ogni risposta. Quello con il miglior
  rapporto qualità/velocità è segnato come CONSIGLIATO, ma la scelta è tua.
- Se il modello scelto non è ancora sul PC, l'app si offre di scaricarlo da sola
  (oppure puoi usare il comando: ollama pull nome-modello).
- Puoi rifare la scelta quando vuoi dal pulsante "Motore AI e aspetto" nella barra laterale.

3. CARICARE DOCUMENTI
Clicca "Carica documenti" nella barra laterale e seleziona uno o più file.
Formati supportati: PDF, Word, Excel, CSV, PowerPoint, TXT, immagini (JPG, PNG...).

I documenti appaiono in fila orizzontale, ognuno con il bordo del suo colore,
l'estensione al centro e il titolo sotto:
- Blu = Word
- Arancione = PowerPoint
- Rosso = PDF
- Verde = Excel/CSV
- Giallo = TXT
- Viola = immagini e altri formati
Per togliere un documento clicca sulla ✕ in alto a destra della card.

4. FARE DOMANDE
Scrivi una domanda nel box e premi "Chiedi" o Invio.

4b. CITARE UN DOCUMENTO CON @
Scrivi @ (o premi il pulsante @ accanto a "Chiedi"): si apre un menu con i documenti caricati,
ognuno con la sua estensione e il titolo. Sceglilo con il mouse o con le frecce + Invio.
La domanda verrà fatta solo sui documenti citati, così puoi interrogarne uno in particolare
o confrontarne due (es. "Che differenze ci sono tra @offerta_A.pdf e @offerta_B.pdf?").
Puoi anche cliccare su una card per citarla.

Nel menu trovi anche le AZIONI RAPIDE:
- Confronta due Excel
- Confronta due documenti
- Verifica i conti di un documento
- Riassumi ed estrai i dati chiave
Scegli l'azione e poi i documenti richiesti: la domanda completa si scrive da sola.
Esempi:
- Qual è il totale della fattura?
- Controlla se quantità × prezzo torna su tutte le righe
- Qual è l'indirizzo del fornitore?
- Riassumi i punti chiave
- Confronta le date dei due documenti

5. VERIFICA CONTI
Premi "Verifica conti". L'app controlla automaticamente:
- Quantità × prezzo = importo riga
- Imponibile + IVA = totale
- Somma delle righe = totale documento
Ti segnala esattamente dove non torna e la differenza.

5b. MODIFICA DOCUMENTI WORD/EXCEL
Premi "Modifica e salva copia".
- Si apre il file con Word o Excel.
- Modificalo e salvalo.
- L'app ti permette di creare una copia con il nome "nome_file_COPIA".

6. CONFRONTA DUE DOCUMENTI
Premi "Confronta 2 docs", seleziona due documenti caricati e vedi le differenze (testo aggiunto, rimosso o modificato).

7. OCR SU SCANSIONI
Se carichi un'immagine o un PDF scansionato e hai installato PaddleOCR, l'app estrae il testo automaticamente (ottimizzato per tabelle e numeri).
Per attivarlo: pip install paddleocr opencv-python-headless

8. COSA MANCA PER RENDERLA ANCORA PIÙ USABILE
- Anteprima reale del contenuto del documento (non solo icona)
- Salvataggio automatico delle sessioni e dei documenti caricati
- Ricerca full-text tra tutti i documenti
- Esportazione dei risultati in PDF o Excel
- Supporto completo a PowerPoint e ODT
- Integrazione diretta con cartelle di Windows (drag & drop)
- Storico completo delle domande e risposte
- Modelli AI più grandi se hai molta RAM
- Correzione manuale del testo OCR direttamente nell'interfaccia

9. INDICATORE RAM
Nella barra laterale vedi l'utilizzo attuale della RAM del PC (richiede psutil).
Il colore diventa arancione sopra il 70% e rosso sopra l'85%.

10. ASPETTO
In basso a sinistra puoi passare in ogni momento tra vista Chiara (sfondo chiaro, testo nero)
e Scura (sfondo scuro, testo chiaro). La scelta viene ricordata.

Buon lavoro con TALdoc!
"""

class LocalAI:
    def __init__(self, model: str = DEFAULT_MODEL):
        self.model = model
        self.available = self.check()

    def check(self) -> bool:
        try:
            r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=2)
            return r.status_code == 200
        except Exception:
            return False

    def ask(self, prompt: str, context: str = "") -> str:
        if not self.available:
            return "Ollama non disponibile. Avvia Ollama e assicurati che il modello sia scaricato."
        full = f"Contesto dai documenti:\n{context}\n\nDomanda: {prompt}\nRispondi in italiano, in modo chiaro e preciso. Se fai calcoli mostra i passaggi."
        try:
            r = requests.post(
                f"{OLLAMA_URL}/api/generate",
                json={"model": self.model, "prompt": full, "stream": False},
                timeout=300
            )
            if r.status_code == 200:
                return r.json().get("response", "Nessuna risposta")
            return f"Errore Ollama: {r.status_code}"
        except Exception as e:
            return f"Errore connessione: {e}"

class Document:
    def __init__(self, path: str):
        self.path = path
        self.name = os.path.basename(path)
        self.ext = Path(path).suffix.lower()
        self.type = self._detect_type()
        self.border_color = BORDER_COLORS.get(self.type, BORDER_COLORS["other"])
        self.text = self._extract_text()

    def _detect_type(self) -> str:
        if self.ext in [".doc", ".docx"]: return "word"
        if self.ext in [".ppt", ".pptx"]: return "powerpoint"
        if self.ext == ".pdf": return "pdf"
        if self.ext in [".xls", ".xlsx", ".csv"]: return "excel"
        if self.ext in [".txt", ".rtf"]: return "txt"
        return "other"

    def _extract_text(self) -> str:
        try:
            if self.ext in [".png", ".jpg", ".jpeg", ".bmp", ".tiff"] and PADDLE_AVAILABLE:
                ocr = PaddleOCR(use_angle_cls=True, lang='it', show_log=False)
                result = ocr.ocr(self.path, cls=True)
                if result and result[0]:
                    return "\n".join([line[1][0] for line in result[0]])
                return "[Nessun testo estratto dall'immagine]"

            if self.ext == ".pdf" and PdfReader:
                reader = PdfReader(self.path)
                text = "\n".join(page.extract_text() or "" for page in reader.pages)
                if len(text.strip()) < 50 and PADDLE_AVAILABLE:
                    return "[PDF scansionato - OCR non completo in questa versione base]"
                return text or "[PDF senza testo]"

            if self.ext == ".docx" and DocxDocument:
                doc = DocxDocument(self.path)
                return "\n".join(p.text for p in doc.paragraphs)

            if self.ext in [".xlsx", ".xls"] and openpyxl:
                wb = openpyxl.load_workbook(self.path, data_only=True)
                lines = []
                for sheet in wb.worksheets:
                    lines.append(f"--- {sheet.title} ---")
                    for row in sheet.iter_rows(values_only=True):
                        lines.append(" | ".join(str(c) if c is not None else "" for c in row))
                return "\n".join(lines)

            if self.ext == ".csv" and pd:
                return pd.read_csv(self.path).to_string()

            if self.ext == ".pptx" and Presentation:
                prs = Presentation(self.path)
                texts = []
                for slide in prs.slides:
                    for shape in slide.shapes:
                        if hasattr(shape, "text"):
                            texts.append(shape.text)
                return "\n".join(texts)

            if self.ext in [".txt", ".rtf"]:
                with open(self.path, "r", encoding="utf-8", errors="ignore") as f:
                    return f.read()

            return f"[Formato {self.ext} supportato parzialmente]"
        except Exception as e:
            return f"[Errore lettura: {e}]"

# ---------------------------------------------------------------------------
#  Impostazioni salvate (tema, modello) – così la configurazione si fa una volta
# ---------------------------------------------------------------------------
def _settings_path() -> Path:
    base = os.environ.get("APPDATA") or str(Path.home())
    folder = Path(base) / "TALdoc_Desktop"
    try:
        folder.mkdir(parents=True, exist_ok=True)
    except Exception:
        folder = Path.home()
    return folder / "settings.json"


def load_settings() -> dict:
    try:
        with open(_settings_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data: dict):
    try:
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
    except Exception:
        pass


# ---------------------------------------------------------------------------
#  Analisi della potenza del PC e stima dei tempi di risposta
# ---------------------------------------------------------------------------
# Modelli Ollama (quantizzazione Q4 predefinita).
# size_gb = spazio su disco / memoria occupata; min_ram = RAM totale minima consigliata
MODEL_CATALOG = [
    {"name": "qwen2.5:1.5b", "params": 1.5, "size_gb": 1.0, "min_ram": 4,  "quality": 2,
     "desc": "Velocissimo, per PC datati. Risposte semplici."},
    {"name": "qwen2.5:3b",   "params": 3.0, "size_gb": 1.9, "min_ram": 6,  "quality": 3,
     "desc": "Leggero e veloce. Buono per domande dirette."},
    {"name": "qwen2.5:7b",   "params": 7.6, "size_gb": 4.7, "min_ram": 12, "quality": 4,
     "desc": "Ottimo equilibrio. Calcoli e confronti affidabili."},
    {"name": "qwen2.5:14b",  "params": 14.8, "size_gb": 9.0, "min_ram": 20, "quality": 5,
     "desc": "Molto preciso. Richiede un PC potente."},
    {"name": "qwen2.5:32b",  "params": 32.8, "size_gb": 20.0, "min_ram": 40, "quality": 5,
     "desc": "Massima qualità. Solo per workstation."},
]

# Domanda "tipo" usata per la stima: documento di 2-3 pagine e risposta di ~200 parole
EST_PROMPT_TOKENS = 1500
EST_ANSWER_TOKENS = 280
MAX_GOOD_SECONDS = 45  # oltre questa attesa l'uso quotidiano diventa scomodo


def _no_window_flags():
    return 0x08000000 if os.name == "nt" else 0  # CREATE_NO_WINDOW su Windows


def detect_hardware() -> dict:
    """Rileva processore, RAM e GPU NVIDIA (se presente)."""
    import platform
    import subprocess

    hw = {"cpu_name": platform.processor() or "Processore sconosciuto",
          "cores": os.cpu_count() or 4, "threads": os.cpu_count() or 4,
          "ghz": 2.5, "ram_gb": 8.0, "gpu_name": None, "vram_gb": 0.0}

    # Nome e frequenza CPU dal registro di Windows
    if os.name == "nt":
        try:
            import winreg
            key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE,
                                 r"HARDWARE\DESCRIPTION\System\CentralProcessor\0")
            hw["cpu_name"] = winreg.QueryValueEx(key, "ProcessorNameString")[0].strip()
            hw["ghz"] = winreg.QueryValueEx(key, "~MHz")[0] / 1000
            winreg.CloseKey(key)
        except Exception:
            pass
    else:
        try:
            with open("/proc/cpuinfo") as f:
                for line in f:
                    if line.startswith("model name"):
                        hw["cpu_name"] = line.split(":", 1)[1].strip()
                        break
        except Exception:
            pass

    if PSUTIL_AVAILABLE:
        try:
            hw["ram_gb"] = psutil.virtual_memory().total / (1024 ** 3)
            hw["cores"] = psutil.cpu_count(logical=False) or hw["cores"]
            hw["threads"] = psutil.cpu_count(logical=True) or hw["threads"]
            freq = psutil.cpu_freq()
            if freq and freq.max and freq.max > 500:
                hw["ghz"] = max(hw["ghz"], freq.max / 1000)
        except Exception:
            pass

    # GPU NVIDIA (Ollama la usa automaticamente)
    try:
        out = subprocess.run(
            ["nvidia-smi", "--query-gpu=name,memory.total", "--format=csv,noheader,nounits"],
            capture_output=True, text=True, timeout=4, creationflags=_no_window_flags())
        if out.returncode == 0 and out.stdout.strip():
            name, mem = out.stdout.strip().splitlines()[0].rsplit(",", 1)
            hw["gpu_name"] = name.strip()
            hw["vram_gb"] = float(mem) / 1024
    except Exception:
        pass
    return hw


def estimate_models(hw: dict) -> List[dict]:
    """Per ogni modello stima velocità e secondi di attesa per una risposta tipo."""
    # Punteggio CPU: core fisici × GHz (la velocità su CPU è limitata anche dalla memoria)
    cpu_score = min(hw["cores"] * max(hw["ghz"], 1.0), 45)
    if hw["threads"] > hw["cores"]:
        cpu_score *= 1.1
    results = []
    for m in MODEL_CATALOG:
        on_gpu = hw["vram_gb"] >= m["size_gb"] + 1.0
        if on_gpu:
            gen_tps = max(min(260 / m["params"], 120), 8)   # stima prudente GPU media
            prefill_tps = gen_tps * 20
        else:
            gen_tps = max(cpu_score * 2.6 / m["params"], 0.3)
            prefill_tps = gen_tps * 10
        seconds = EST_PROMPT_TOKENS / prefill_tps + EST_ANSWER_TOKENS / gen_tps + 1.5
        fits = hw["ram_gb"] >= m["min_ram"] or on_gpu
        results.append({**m, "seconds": seconds, "fits": fits, "on_gpu": on_gpu,
                        "tps": gen_tps, "recommended": False})

    usable = [r for r in results if r["fits"]]
    good = [r for r in usable if r["seconds"] <= MAX_GOOD_SECONDS]
    if good:
        # la qualità più alta entro un'attesa accettabile; a parità, il più veloce
        best = max(good, key=lambda r: (r["quality"], -r["seconds"]))
    elif usable:
        best = min(usable, key=lambda r: r["seconds"])
    else:
        best = results[0]
    best["recommended"] = True
    return results


def format_seconds(s: float) -> str:
    if s < 10:
        return f"~{max(round(s), 2)} s"
    if s < 60:
        return f"~{int(round(s / 5.0) * 5)} s"
    m = s / 60
    return f"~{m:.0f} min" if m >= 1.5 else "~1 min"


# ---------------------------------------------------------------------------
#  Grafica disegnata al volo con Pillow: logo glossy e card dei documenti
# ---------------------------------------------------------------------------
try:
    from PIL import ImageDraw, ImageFont, ImageFilter
except ImportError:
    ImageDraw = ImageFont = ImageFilter = None


def _hex_rgb(h: str):
    h = h.lstrip("#")
    return tuple(int(h[i:i + 2], 16) for i in (0, 2, 4))


def _mix(c1, c2, t: float):
    return tuple(int(a + (b - a) * t) for a, b in zip(c1, c2))


_FONT_CACHE = {}


def pil_font(size: int, bold: bool = False):
    key = (size, bold)
    if key in _FONT_CACHE:
        return _FONT_CACHE[key]
    names = (["segoeuib.ttf", "arialbd.ttf", "DejaVuSans-Bold.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"] if bold else
             ["segoeui.ttf", "arial.ttf", "DejaVuSans.ttf",
              "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf"])
    font = None
    for n in names:
        for path in (n, os.path.join(os.environ.get("WINDIR", r"C:\Windows"), "Fonts", n)):
            try:
                font = ImageFont.truetype(path, size)
                break
            except Exception:
                continue
        if font:
            break
    if font is None:
        try:
            font = ImageFont.load_default(size)
        except Exception:
            font = ImageFont.load_default()
    _FONT_CACHE[key] = font
    return font


def make_logo(size: int = 256):
    """Logo TALdoc: tessera arrotondata con gradiente viola-fucsia, documento e riflesso glossy."""
    if Image is None or ImageDraw is None:
        return None
    S = size * 4  # disegno in alta risoluzione e poi riduco (bordi morbidi)
    r = int(S * 0.24)

    # gradiente diagonale
    c1, c2, c3 = _hex_rgb("#4F46E5"), _hex_rgb("#9333EA"), _hex_rgb("#EC4899")
    grad = Image.new("RGB", (S, S))
    gp = grad.load()
    step = 4
    for y in range(0, S, step):
        for x in range(0, S, step):
            t = (x + y) / (2 * S)
            col = _mix(c1, c2, t * 2) if t < 0.5 else _mix(c2, c3, (t - 0.5) * 2)
            for dy in range(step):
                for dx in range(step):
                    gp[x + dx, y + dy] = col
    mask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, S - 1, S - 1), r, fill=255)
    logo = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    logo.paste(grad, (0, 0), mask)

    d = ImageDraw.Draw(logo)
    # ombra del documento
    shadow = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    sd = ImageDraw.Draw(shadow)
    dx0, dy0, dx1, dy1 = int(S * .27), int(S * .19), int(S * .73), int(S * .83)
    sd.rounded_rectangle((dx0 + S * .02, dy0 + S * .035, dx1 + S * .02, dy1 + S * .035),
                         int(S * .06), fill=(30, 10, 60, 110))
    shadow = shadow.filter(ImageFilter.GaussianBlur(S * .025))
    logo.alpha_composite(shadow)
    d = ImageDraw.Draw(logo)

    # documento bianco con angolo piegato
    fold = int(S * .14)
    d.rounded_rectangle((dx0, dy0, dx1, dy1), int(S * .06), fill=(255, 255, 255, 255))
    d.polygon([(dx1 - fold, dy0), (dx1 + 2, dy0), (dx1 + 2, dy0 + fold)], fill=(0, 0, 0, 0))
    logo_px = logo.load()
    # ripristina il gradiente nell'angolo tagliato
    for y in range(dy0, dy0 + fold + 3):
        for x in range(dx1 - fold, dx1 + 3):
            if x - (dx1 - fold) > (y - dy0) and logo_px[x, y][3] == 0 and mask.getpixel((x, y)):
                logo_px[x, y] = grad.getpixel((x, y)) + (255,)
    d.polygon([(dx1 - fold, dy0), (dx1 - fold, dy0 + fold), (dx1, dy0 + fold)],
              fill=(214, 204, 255, 255))

    # righe di testo + "T" in gradiente
    tx = dx0 + int(S * .07)
    d.rounded_rectangle((tx, int(S * .30), tx + int(S * .17), int(S * .345)), int(S * .02),
                        fill=_hex_rgb("#7C3AED") + (255,))
    d.rounded_rectangle((tx + int(S * .06), int(S * .30), tx + int(S * .11), int(S * .55)),
                        int(S * .02), fill=_hex_rgb("#7C3AED") + (255,))
    for i, w in enumerate((.25, .29, .21)):
        y = int(S * (.61 + i * .065))
        d.rounded_rectangle((tx, y, tx + int(S * w), y + int(S * .03)), int(S * .015),
                            fill=(196, 181, 253, 255))
    # scintilla AI
    cx, cy, k = int(S * .62), int(S * .42), S * .085
    star = [(cx, cy - k), (cx + k * .28, cy - k * .28), (cx + k, cy), (cx + k * .28, cy + k * .28),
            (cx, cy + k), (cx - k * .28, cy + k * .28), (cx - k, cy), (cx - k * .28, cy - k * .28)]
    d.polygon(star, fill=_hex_rgb("#EC4899") + (255,))

    # riflesso glossy sulla metà superiore
    gloss = Image.new("RGBA", (S, S), (0, 0, 0, 0))
    gl = gloss.load()
    gmask = Image.new("L", (S, S), 0)
    ImageDraw.Draw(gmask).ellipse((-S * .35, -S * .75, S * 1.35, S * .48), fill=255)
    for y in range(0, int(S * .5)):
        a = int(150 * (1 - y / (S * .5)) ** 1.2)
        for x in range(0, S, 2):
            if gmask.getpixel((x, y)) and mask.getpixel((x, y)):
                gl[x, y] = (255, 255, 255, a)
                if x + 1 < S:
                    gl[x + 1, y] = (255, 255, 255, a)
    logo.alpha_composite(gloss)
    # bordo luminoso sottile
    ImageDraw.Draw(logo).rounded_rectangle((2, 2, S - 3, S - 3), r, outline=(255, 255, 255, 70),
                                           width=max(2, S // 90))
    return logo.resize((size, size), Image.LANCZOS)


CARD_W, CARD_H = 170, 134


def doc_title(doc) -> str:
    """Titolo del file mostrato in card e nel menu @ (nome senza estensione)."""
    return Path(doc.name).stem or doc.name


def wrap_title(d, text, font, max_w, max_lines):
    """Va a capo preferibilmente su spazi, '_' o '-'; aggiunge … se il testo è troppo lungo."""
    tokens = [t for t in re.split(r"(?<=[ _\-.])", text) if t]
    pieces = []
    for t in tokens:  # spezza i token troppo lunghi
        while d.textlength(t, font=font) > max_w:
            n = len(t)
            while n > 1 and d.textlength(t[:n], font=font) > max_w:
                n -= 1
            pieces.append(t[:n])
            t = t[n:]
        pieces.append(t)
    lines, cur = [], ""
    for i, t in enumerate(pieces):
        if d.textlength(cur + t, font=font) <= max_w:
            cur += t
        else:
            lines.append(cur)
            cur = t
    if cur:
        lines.append(cur)
    lines = [l.strip() for l in lines]
    if len(lines) > max_lines:
        last = lines[max_lines - 1]
        while last and d.textlength(last + "…", font=font) > max_w:
            last = last[:-1]
        lines = lines[:max_lines - 1] + [last.rstrip(" _-.") + "…"]
    return lines


def make_card(doc, hover: bool = False):
    """Card documento: bordo acceso, interno con gradiente molto chiaro dello stesso colore,
    estensione al centro e titolo sotto."""
    if Image is None or ImageDraw is None:
        return None
    k = 3
    W, H = CARD_W * k, CARD_H * k
    base = _hex_rgb(doc.border_color)
    white = (255, 255, 255)
    top = _mix(base, white, 0.80 if not hover else 0.72)
    bottom = _mix(base, white, 0.95 if not hover else 0.88)

    grad = Image.new("RGB", (W, H))
    gd = ImageDraw.Draw(grad)
    for y in range(H):
        gd.line([(0, y), (W, y)], fill=_mix(top, bottom, y / H))
    mask = Image.new("L", (W, H), 0)
    rad = 16 * k
    ImageDraw.Draw(mask).rounded_rectangle((0, 0, W - 1, H - 1), rad, fill=255)
    card = Image.new("RGBA", (W, H), (0, 0, 0, 0))
    card.paste(grad, (0, 0), mask)
    d = ImageDraw.Draw(card)
    bw = 3 * k
    d.rounded_rectangle((bw // 2, bw // 2, W - 1 - bw // 2, H - 1 - bw // 2), rad, outline=base, width=bw)

    # estensione al centro
    ext = (doc.ext.lstrip(".") or "FILE").upper()
    f_ext = pil_font(30 * k if len(ext) <= 4 else 24 * k, bold=True)
    tw = d.textlength(ext, font=f_ext)
    d.text(((W - tw) / 2, 30 * k), ext, font=f_ext, fill=base)

    # titolo sotto (max 2 righe)
    f_name = pil_font(11 * k, bold=True)
    max_w = W - 24 * k
    lines = wrap_title(d, doc_title(doc), f_name, max_w, 2)
    text_col = _mix(base, (15, 23, 42), 0.75)
    y = 82 * k
    for line in lines[:2]:
        lw = d.textlength(line, font=f_name)
        d.text(((W - lw) / 2, y), line, font=f_name, fill=text_col)
        y += 15 * k

    # pulsante chiudi in alto a destra
    cx, cy, cr = W - 18 * k, 18 * k, 9 * k
    if hover:
        d.ellipse((cx - cr, cy - cr, cx + cr, cy + cr), fill=_mix(base, white, 0.6))
    s = 3.5 * k
    d.line([(cx - s, cy - s), (cx + s, cy + s)], fill=text_col, width=int(1.6 * k))
    d.line([(cx - s, cy + s), (cx + s, cy - s)], fill=text_col, width=int(1.6 * k))
    return card


# ---------------------------------------------------------------------------
#  Tema grafico moderno (ogni colore è una coppia: modalità chiara, scura)
# ---------------------------------------------------------------------------
FONT_FAMILY = "Segoe UI"

C = {
    "bg":        ("#F4F6FB", "#0B1020"),
    "sidebar":   ("#FFFFFF", "#111827"),
    "surface":   ("#FFFFFF", "#161E31"),
    "surface2":  ("#EEF1F8", "#1E2840"),
    "border":    ("#E2E6F0", "#263250"),
    "text":      ("#000000", "#E5E9F5"),
    "muted":     ("#64748B", "#8C97B2"),
    "accent":    ("#6366F1", "#7C83FF"),
    "accent_h":  ("#4F46E5", "#6366F1"),
    "accent_bg": ("#EEF0FF", "#232A55"),
    "danger":    ("#EF4444", "#F87171"),
    "danger_h":  ("#DC2626", "#EF4444"),
    "ok":        ("#16A34A", "#4ADE80"),
    "warn":      ("#EA580C", "#FB923C"),
    "user_bg":   ("#EEF0FF", "#232A55"),
}


def F(size=13, weight="normal"):
    return ctk.CTkFont(family=FONT_FAMILY, size=size, weight=weight)


class SidebarButton(ctk.CTkButton):
    def __init__(self, master, text, icon, command, **kw):
        super().__init__(
            master, text=f"  {icon}   {text}", command=command, anchor="w",
            height=36, corner_radius=10, font=F(13),
            fg_color="transparent", hover_color=C["surface2"], text_color=C["text"], **kw
        )


class TALdocApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.settings = load_settings()
        ctk.set_appearance_mode(self.settings.get("theme", "dark"))
        ctk.set_default_color_theme("blue")

        self.title(APP_NAME)
        self.geometry("1240x800")
        self.minsize(1000, 640)
        self.configure(fg_color=C["bg"])

        self.documents: List[Document] = []
        self.ai = LocalAI()
        self.busy = False

        self._build_menu()
        self._build_ui()
        self.after(300, self._start_up)

    # ------------------------------------------------------------------ UI
    def _build_menu(self):
        menubar = Menu(self)
        self.config(menu=menubar)
        help_menu = Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Aiuto", menu=help_menu)
        help_menu.add_command(label="Manuale", command=self.show_manual)
        help_menu.add_separator()
        help_menu.add_command(label="Informazioni", command=self.show_about)

    def _build_ui(self):
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self._build_sidebar()
        self._build_main()

    def _build_sidebar(self):
        sb = ctk.CTkFrame(self, width=250, corner_radius=0, fg_color=C["sidebar"],
                          border_width=0)
        sb.grid(row=0, column=0, sticky="nsw")
        sb.grid_propagate(False)
        sb.pack_propagate(False)

        # Logo
        brand = ctk.CTkFrame(sb, fg_color="transparent")
        brand.pack(fill="x", padx=20, pady=(20, 18))
        self._logo_pil = make_logo(176)
        if self._logo_pil is not None:
            self._logo_img = ctk.CTkImage(light_image=self._logo_pil, dark_image=self._logo_pil,
                                          size=(46, 46))
            ctk.CTkLabel(brand, image=self._logo_img, text="").pack(side="left")
            try:
                from PIL import ImageTk
                self._icon_tk = ImageTk.PhotoImage(self._logo_pil.resize((64, 64), Image.LANCZOS))
                self.iconphoto(True, self._icon_tk)
            except Exception:
                pass
        else:
            ctk.CTkLabel(brand, text="T", width=46, height=46, corner_radius=12,
                         fg_color=C["accent"], text_color="white",
                         font=F(20, "bold")).pack(side="left")
        names = ctk.CTkFrame(brand, fg_color="transparent")
        names.pack(side="left", padx=12)
        ctk.CTkLabel(names, text="TALdoc", font=F(19, "bold"),
                     text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(names, text="Desktop · AI locale", font=F(11),
                     text_color=C["muted"]).pack(anchor="w")

        # Azione principale
        self.load_btn = ctk.CTkButton(
            sb, text="＋  Carica documenti", command=self.load_documents, height=44,
            corner_radius=12, font=F(14, "bold"),
            fg_color=C["accent"], hover_color=C["accent_h"], text_color="white")
        self.load_btn.pack(fill="x", padx=16, pady=(0, 14))

        ctk.CTkLabel(sb, text="STRUMENTI", font=F(11, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=24, pady=(4, 6))
        self.action_buttons = [
            SidebarButton(sb, "Verifica conti", "✓", self.verify_accounts),
            SidebarButton(sb, "Confronta 2 documenti", "⇄", self.compare_docs),
            SidebarButton(sb, "Modifica e salva copia", "✎", self.edit_and_save_copy),
            SidebarButton(sb, "Pulisci tutto", "⌫", self.clear_docs),
        ]
        for b in self.action_buttons:
            b.pack(fill="x", padx=12, pady=2)

        ctk.CTkLabel(sb, text="AIUTO", font=F(11, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=24, pady=(12, 4))
        SidebarButton(sb, "Motore AI e aspetto", "⚙",
                      lambda: self.open_setup_wizard()).pack(fill="x", padx=12, pady=2)
        SidebarButton(sb, "Manuale", "?", self.show_manual).pack(fill="x", padx=12, pady=2)
        SidebarButton(sb, "Informazioni", "i", self.show_about).pack(fill="x", padx=12, pady=2)

        # Parte bassa: stato sistema
        bottom = ctk.CTkFrame(sb, fg_color="transparent")
        bottom.pack(side="bottom", fill="x", padx=16, pady=(8, 14))

        card = ctk.CTkFrame(bottom, fg_color=C["surface2"], corner_radius=12)
        card.pack(fill="x", pady=(0, 12))

        row = ctk.CTkFrame(card, fg_color="transparent")
        row.pack(fill="x", padx=14, pady=(12, 2))
        self.ollama_dot = ctk.CTkLabel(row, text="●", font=F(14),
                                       text_color=C["ok"] if self.ai.available else C["danger"])
        self.ollama_dot.pack(side="left")
        self.ollama_label = ctk.CTkLabel(
            row, text="Ollama attivo" if self.ai.available else "Ollama non attivo",
            font=F(12, "bold"), text_color=C["text"])
        self.ollama_label.pack(side="left", padx=6)

        self.model_label = ctk.CTkLabel(card, text=f"Modello: {self.ai.model}",
                                        font=F(12), text_color=C["muted"])
        self.model_label.pack(anchor="w", padx=14)

        self.ram_label = ctk.CTkLabel(card, text="RAM: --", font=F(12), text_color=C["muted"])
        self.ram_label.pack(anchor="w", padx=14, pady=(6, 2))
        self.ram_bar = ctk.CTkProgressBar(card, height=6, corner_radius=3,
                                          progress_color=C["ok"], fg_color=C["border"])
        self.ram_bar.set(0)
        self.ram_bar.pack(fill="x", padx=14, pady=(2, 14))

        ctk.CTkLabel(bottom, text="Aspetto", font=F(11, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=4, pady=(0, 4))
        self.theme_seg = ctk.CTkSegmentedButton(
            bottom, values=["☀ Chiaro", "☾ Scuro"], height=32, corner_radius=10, font=F(12),
            selected_color=C["accent"], selected_hover_color=C["accent_h"],
            unselected_color=C["surface2"], unselected_hover_color=C["border"],
            fg_color=C["surface2"], text_color=C["text"],
            command=lambda v: self._apply_theme("light" if "Chiaro" in v else "dark"))
        self.theme_seg.set("☀ Chiaro" if self.settings.get("theme", "dark") == "light" else "☾ Scuro")
        self.theme_seg.pack(fill="x")

        self.after(1000, self.update_ram_indicator)

    def _build_main(self):
        main = ctk.CTkFrame(self, fg_color="transparent")
        main.grid(row=0, column=1, sticky="nsew", padx=24, pady=(16, 10))
        main.grid_columnconfigure(0, weight=1)
        main.grid_rowconfigure(2, weight=1)

        # Intestazione
        head = ctk.CTkFrame(main, fg_color="transparent")
        head.grid(row=0, column=0, sticky="ew")
        ctk.CTkLabel(head, text="I tuoi documenti", font=F(24, "bold"),
                     text_color=C["text"]).pack(side="left")
        self.count_badge = ctk.CTkLabel(head, text="0", width=30, height=26, corner_radius=13,
                                        fg_color=C["accent_bg"], text_color=C["accent"],
                                        font=F(12, "bold"))
        self.count_badge.pack(side="left", padx=10)
        ctk.CTkLabel(head, text="🔒 100% offline – i dati restano sul tuo PC",
                     font=F(12), text_color=C["muted"]).pack(side="right")

        # Striscia documenti
        docs_wrap = ctk.CTkFrame(main, fg_color=C["surface"], corner_radius=16,
                                 border_width=1, border_color=C["border"])
        docs_wrap.grid(row=1, column=0, sticky="ew", pady=(14, 16))
        self.docs_frame = ctk.CTkScrollableFrame(docs_wrap, orientation="horizontal",
                                                 height=140, fg_color="transparent",
                                                 scrollbar_button_color=C["border"])
        self.docs_frame.pack(fill="x", padx=8, pady=8)
        self._show_empty_docs()

        # Area risposta (stile chat)
        chat = ctk.CTkFrame(main, fg_color=C["surface"], corner_radius=16,
                            border_width=1, border_color=C["border"])
        chat.grid(row=2, column=0, sticky="nsew")
        chat.grid_columnconfigure(0, weight=1)
        chat.grid_rowconfigure(1, weight=1)

        chat_head = ctk.CTkFrame(chat, fg_color="transparent")
        chat_head.grid(row=0, column=0, sticky="ew", padx=20, pady=(16, 4))
        ctk.CTkLabel(chat_head, text="Assistente", font=F(16, "bold"),
                     text_color=C["text"]).pack(side="left")
        ctk.CTkButton(chat_head, text="Copia risposta", width=120, height=30,
                      corner_radius=8, font=F(12), fg_color=C["surface2"],
                      hover_color=C["border"], text_color=C["text"],
                      command=self._copy_response).pack(side="right")
        ctk.CTkButton(chat_head, text="Svuota", width=80, height=30,
                      corner_radius=8, font=F(12), fg_color="transparent",
                      hover_color=C["surface2"], text_color=C["muted"],
                      command=self._clear_chat).pack(side="right", padx=6)

        self.response_text = ctk.CTkTextbox(
            chat, wrap="word", fg_color="transparent", text_color=C["text"],
            font=F(14), border_width=0, activate_scrollbars=True,
            scrollbar_button_color=C["border"])
        self.response_text.grid(row=1, column=0, sticky="nsew", padx=14, pady=4)
        self._setup_text_tags()
        self._welcome_message()

        self.progress = ctk.CTkProgressBar(chat, mode="indeterminate", height=3,
                                           progress_color=C["accent"], fg_color=C["surface"])
        self.progress.grid(row=2, column=0, sticky="ew", padx=20)
        self.progress.grid_remove()

        # Barra domanda
        ask = ctk.CTkFrame(chat, fg_color=C["surface2"], corner_radius=14)
        self.ask_bar = ask
        ask.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 16))
        ask.grid_columnconfigure(0, weight=1)
        self.question_entry = ctk.CTkEntry(
            ask, placeholder_text="Fai una domanda… scrivi @ per citare un documento o scegliere un'azione rapida",
            fg_color="transparent", border_width=0, height=46, font=F(14),
            text_color=C["text"], placeholder_text_color=C["muted"])
        self.question_entry.grid(row=0, column=0, sticky="ew", padx=(14, 6), pady=4)
        ctk.CTkButton(ask, text="@", width=38, height=38, corner_radius=10, font=F(16, "bold"),
                      fg_color=C["accent_bg"], hover_color=C["border"], text_color=C["accent"],
                      command=self._open_at_button).grid(row=0, column=1, padx=(0, 2), pady=4)
        self.ask_btn = ctk.CTkButton(ask, text="Chiedi  ➤", command=self.ask_question,
                                     width=110, height=38, corner_radius=10, font=F(13, "bold"),
                                     fg_color=C["accent"], hover_color=C["accent_h"],
                                     text_color="white")
        self.ask_btn.grid(row=0, column=2, padx=6, pady=4)

        # Suggerimenti rapidi
        chips = ctk.CTkFrame(chat, fg_color="transparent")
        chips.grid(row=4, column=0, sticky="w", padx=16, pady=(0, 14))
        self.suggest_row = chips
        self.tag_row = ctk.CTkFrame(chat, fg_color="transparent")
        self.tag_row.grid(row=4, column=0, sticky="w", padx=16, pady=(0, 14))
        self.tag_row.grid_remove()
        for s in ["Qual è il totale?", "Riassumi i punti chiave", "Chi è il fornitore?",
                  "Controlla quantità × prezzo"]:
            ctk.CTkButton(chips, text=s, height=28, corner_radius=14, font=F(12),
                          fg_color="transparent", border_width=1, border_color=C["border"],
                          hover_color=C["surface2"], text_color=C["muted"],
                          command=lambda t=s: self._quick_ask(t)).pack(side="left", padx=(0, 8))

        # Barra di stato
        self.status = ctk.CTkLabel(main, text="Pronto. Carica i documenti e inizia a fare domande.",
                                   anchor="w", font=F(12), text_color=C["muted"])
        self.status.grid(row=3, column=0, sticky="ew", pady=(6, 0))

        self._init_mention()

    # ----------------------------------------------------------- helpers UI
    def _mode_index(self):
        return 1 if ctk.get_appearance_mode() == "Dark" else 0

    def _setup_text_tags(self):
        i = self._mode_index()
        tb = self.response_text
        tb.tag_config("user", foreground=C["accent"][i], spacing1=14, spacing3=6)
        tb.tag_config("label", foreground=C["muted"][i], spacing1=10)
        tb.tag_config("answer", foreground=C["text"][i], spacing3=10, lmargin1=4, lmargin2=4)
        tb.tag_config("info", foreground=C["muted"][i])
        tb.tag_config("error", foreground=C["danger"][i])

    def _apply_theme(self, mode: str, save: bool = True):
        """mode = 'light' (sfondo chiaro, testo nero) oppure 'dark' (sfondo scuro, testo chiaro)."""
        ctk.set_appearance_mode(mode)
        self._setup_text_tags()
        if hasattr(self, "theme_seg"):
            self.theme_seg.set("☀ Chiaro" if mode == "light" else "☾ Scuro")
        if save:
            self.settings["theme"] = mode
            save_settings(self.settings)

    def _welcome_message(self):
        self._write("Benvenuto in TALdoc Desktop 👋\n\n", "answer")
        self._write("1. Carica uno o più documenti (PDF, Word, Excel, CSV, PowerPoint, TXT, immagini)\n"
                    "2. Scrivi una domanda nel box qui sotto oppure usa gli strumenti a sinistra\n"
                    "3. Le risposte compaiono qui, tutto elaborato in locale con Ollama\n", "info")

    def _write(self, text, tag="answer"):
        self.response_text.configure(state="normal")
        self.response_text.insert("end", text, tag)
        self.response_text.see("end")
        self.response_text.configure(state="disabled")

    def _clear_chat(self):
        self.response_text.configure(state="normal")
        self.response_text.delete("1.0", "end")
        self.response_text.configure(state="disabled")

    def _copy_response(self):
        text = self.response_text.get("1.0", "end").strip()
        if text:
            self.clipboard_clear()
            self.clipboard_append(text)
            self.status.configure(text="Testo copiato negli appunti.")

    def _quick_ask(self, text):
        self.question_entry.delete(0, "end")
        self.question_entry.insert(0, text)
        self.ask_question()

    def _set_busy(self, busy: bool, msg: str = ""):
        self.busy = busy
        state = "disabled" if busy else "normal"
        self.ask_btn.configure(state=state)
        self.load_btn.configure(state=state)
        for b in self.action_buttons:
            b.configure(state=state)
        if busy:
            self.progress.grid()
            self.progress.start()
        else:
            self.progress.stop()
            self.progress.grid_remove()
        if msg:
            self.status.configure(text=msg)

    def _run_ai(self, title: str, prompt: str, context: str = "", done_msg="Risposta pronta."):
        """Esegue la richiesta all'AI in background, senza bloccare la finestra."""
        if self.busy:
            return
        self._write(f"\n{title}\n", "user")
        self._write("TALdoc sta pensando…\n", "info")
        self._set_busy(True, "Elaborazione in corso…")

        def worker():
            answer = self.ai.ask(prompt, context)
            self.after(0, lambda: self._show_answer(answer, done_msg))

        threading.Thread(target=worker, daemon=True).start()

    def _show_answer(self, answer: str, done_msg: str):
        # rimuove la riga "sta pensando"
        self.response_text.configure(state="normal")
        idx = self.response_text.search("TALdoc sta pensando…", "1.0", backwards=False)
        last = None
        while idx:
            last = idx
            idx = self.response_text.search("TALdoc sta pensando…", f"{idx}+1c", stopindex="end")
        if last:
            self.response_text.delete(last, f"{last} lineend +1c")
        self.response_text.configure(state="disabled")
        tag = "error" if answer.startswith(("Errore", "Ollama non disponibile")) else "answer"
        self._write("TALdoc\n", "label")
        self._write(answer.strip() + "\n", tag)
        self._set_busy(False, done_msg)

    def _refresh_count(self):
        n = len(self.documents)
        self.count_badge.configure(text=str(n))

    def _show_empty_docs(self):
        self.empty_docs = ctk.CTkFrame(self.docs_frame, fg_color="transparent", height=130)
        self.empty_docs.pack(side="left", padx=10, pady=10)
        ctk.CTkLabel(self.empty_docs, text="📂", font=F(34)).pack(side="left", padx=(6, 14))
        txt = ctk.CTkFrame(self.empty_docs, fg_color="transparent")
        txt.pack(side="left")
        ctk.CTkLabel(txt, text="Nessun documento caricato", font=F(15, "bold"),
                     text_color=C["text"]).pack(anchor="w")
        ctk.CTkLabel(txt, text="Usa “Carica documenti” per iniziare. PDF, Word, Excel, CSV, PowerPoint, TXT e immagini.",
                     font=F(12), text_color=C["muted"]).pack(anchor="w")

    def update_ram_indicator(self):
        """Aggiorna l'indicatore di utilizzo RAM."""
        try:
            if PSUTIL_AVAILABLE:
                mem = psutil.virtual_memory()
                used_gb = mem.used / (1024 ** 3)
                total_gb = mem.total / (1024 ** 3)
                percent = mem.percent
                color = C["ok"] if percent < 70 else C["warn"] if percent < 85 else C["danger"]
                self.ram_label.configure(
                    text=f"RAM: {used_gb:.1f} / {total_gb:.1f} GB  ({percent:.0f}%)")
                self.ram_bar.configure(progress_color=color)
                self.ram_bar.set(percent / 100)
            else:
                self.ram_label.configure(text="RAM: psutil non installato")
        except Exception:
            self.ram_label.configure(text="RAM: n/d")
        # aggiorna anche lo stato di Ollama
        self.ai.available = self.ai.check()
        self.ollama_dot.configure(text_color=C["ok"] if self.ai.available else C["danger"])
        self.ollama_label.configure(text="Ollama attivo" if self.ai.available else "Ollama non attivo")
        self.after(5000, self.update_ram_indicator)

    # ------------------------------------------------------- dialog moderni
    def _modal(self, title, w=480, h=360):
        win = ctk.CTkToplevel(self)
        win.title(title)
        win.geometry(f"{w}x{h}")
        win.configure(fg_color=C["bg"])
        win.transient(self)
        win.resizable(False, False)
        win.after(100, win.grab_set)
        win.after(100, win.lift)
        return win

    def _start_up(self):
        if self.settings.get("configured") and self.settings.get("model"):
            self._set_model(self.settings["model"])
        else:
            self.open_setup_wizard(first_run=True)

    def _set_model(self, model: str):
        self.ai.model = model
        self.model_label.configure(text=f"Modello: {model}")

    def open_setup_wizard(self, first_run: bool = False):
        """Configurazione: aspetto + analisi del PC + scelta del motore AI."""
        win = self._modal("Configurazione TALdoc", 700, 760)
        win.resizable(True, True)
        win.minsize(640, 560)

        footer = ctk.CTkFrame(win, fg_color="transparent")
        footer.pack(side="bottom", fill="x", padx=28, pady=(6, 20))
        body = ctk.CTkScrollableFrame(win, fg_color="transparent",
                                      scrollbar_button_color=C["border"])
        body.pack(fill="both", expand=True, padx=10, pady=(10, 0))

        title = "Benvenuto in TALdoc 👋" if first_run else "Impostazioni"
        ctk.CTkLabel(body, text=title, font=F(22, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=18, pady=(12, 2))
        ctk.CTkLabel(body, text="Scegli l'aspetto dell'app e il motore AI più adatto al tuo PC.",
                     font=F(13), text_color=C["muted"]).pack(anchor="w", padx=18)

        # ---- 1. Aspetto
        ctk.CTkLabel(body, text="1 · ASPETTO", font=F(11, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=18, pady=(20, 6))
        theme_var = ctk.StringVar(value=self.settings.get("theme", "dark"))
        tiles = ctk.CTkFrame(body, fg_color="transparent")
        tiles.pack(fill="x", padx=18)
        tiles.grid_columnconfigure((0, 1), weight=1)

        def theme_tile(col, value, label, sub, bg, fg, fg2):
            t = ctk.CTkFrame(tiles, fg_color=C["surface"], corner_radius=12,
                             border_width=1, border_color=C["border"])
            t.grid(row=0, column=col, sticky="ew", padx=(0 if col == 0 else 6, 6 if col == 0 else 0))
            prev = ctk.CTkFrame(t, fg_color=bg, corner_radius=8, height=72,
                                border_width=1, border_color="#CBD5E1")
            prev.pack(fill="x", padx=12, pady=(12, 8))
            prev.pack_propagate(False)
            ctk.CTkLabel(prev, text="Aa  Totale fattura", font=F(14, "bold"),
                         text_color=fg).pack(anchor="w", padx=12, pady=(10, 0))
            ctk.CTkLabel(prev, text="1.220,00 EUR · IVA 22%", font=F(11),
                         text_color=fg2).pack(anchor="w", padx=12, pady=(0, 8))
            ctk.CTkRadioButton(t, text=label, variable=theme_var, value=value, font=F(13, "bold"),
                               text_color=C["text"], fg_color=C["accent"], hover_color=C["accent_h"],
                               command=lambda: self._apply_theme(theme_var.get(), save=False)
                               ).pack(anchor="w", padx=12)
            ctk.CTkLabel(t, text=sub, font=F(11), text_color=C["muted"]).pack(anchor="w", padx=40, pady=(0, 12))

        theme_tile(0, "light", "Chiara", "Sfondo chiaro, testo nero", "#FFFFFF", "#000000", "#475569")
        theme_tile(1, "dark", "Scura", "Sfondo scuro, testo chiaro", "#0B1020", "#F1F5F9", "#94A3B8")

        # ---- 2. Il tuo PC
        ctk.CTkLabel(body, text="2 · IL TUO PC", font=F(11, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=18, pady=(20, 6))
        pc_card = ctk.CTkFrame(body, fg_color=C["surface"], corner_radius=12,
                               border_width=1, border_color=C["border"])
        pc_card.pack(fill="x", padx=18)
        pc_label = ctk.CTkLabel(pc_card, text="Analisi del processore e della RAM in corso…",
                                font=F(13), text_color=C["text"], justify="left", anchor="w")
        pc_label.pack(fill="x", padx=14, pady=12)

        # ---- 3. Motore AI
        ctk.CTkLabel(body, text="3 · MOTORE AI", font=F(11, "bold"),
                     text_color=C["muted"]).pack(anchor="w", padx=18, pady=(20, 6))
        models_box = ctk.CTkFrame(body, fg_color="transparent")
        models_box.pack(fill="x", padx=18)
        ctk.CTkLabel(body, text=(f"Tempi stimati per una risposta di circa 200 parole su un documento di 2–3 pagine. "
                                 "Sono valori approssimativi: documenti lunghi richiedono più tempo e la prima "
                                 "domanda dopo l'avvio può impiegare qualche secondo in più per caricare il modello."),
                     font=F(11), text_color=C["muted"], wraplength=600, justify="left"
                     ).pack(anchor="w", padx=18, pady=(8, 10))

        model_var = ctk.StringVar(value="")
        confirm_btn = ctk.CTkButton(footer, text="Conferma e inizia" if first_run else "Salva",
                                    height=44, corner_radius=12, font=F(14, "bold"),
                                    fg_color=C["accent"], hover_color=C["accent_h"],
                                    text_color="white", state="disabled")
        confirm_btn.pack(fill="x")

        def fill_results(hw, results):
            gpu = (f"{hw['gpu_name']} ({hw['vram_gb']:.0f} GB)" if hw["gpu_name"]
                   else "nessuna GPU NVIDIA (l'AI userà il processore)")
            pc_label.configure(text=(f"Processore: {hw['cpu_name']}\n"
                                     f"{hw['cores']} core / {hw['threads']} thread · {hw['ghz']:.1f} GHz\n"
                                     f"RAM: {hw['ram_gb']:.0f} GB   ·   GPU: {gpu}"))
            current = self.settings.get("model")
            for r in results:
                row = ctk.CTkFrame(models_box, corner_radius=12, border_width=2 if r["recommended"] else 1,
                                   fg_color=C["accent_bg"] if r["recommended"] else C["surface"],
                                   border_color=C["accent"] if r["recommended"] else C["border"])
                row.pack(fill="x", pady=4)
                row.grid_columnconfigure(1, weight=1)
                rb = ctk.CTkRadioButton(row, text="", variable=model_var, value=r["name"], width=24,
                                        fg_color=C["accent"], hover_color=C["accent_h"],
                                        state="normal" if r["fits"] else "disabled")
                rb.grid(row=0, column=0, rowspan=2, padx=(14, 4), pady=12)

                head = ctk.CTkFrame(row, fg_color="transparent")
                head.grid(row=0, column=1, sticky="w", pady=(10, 0))
                ctk.CTkLabel(head, text=r["name"], font=F(14, "bold"),
                             text_color=C["text"] if r["fits"] else C["muted"]).pack(side="left")
                ctk.CTkLabel(head, text="  " + "★" * r["quality"] + "☆" * (5 - r["quality"]),
                             font=F(12), text_color=C["warn"] if r["fits"] else C["muted"]).pack(side="left")
                if r["recommended"]:
                    ctk.CTkLabel(head, text=" CONSIGLIATO ", font=F(10, "bold"), corner_radius=8,
                                 fg_color=C["accent"], text_color="white").pack(side="left", padx=8)
                sub = f"{r['desc']}  ·  {r['size_gb']:.1f} GB"
                if not r["fits"]:
                    sub = f"RAM insufficiente (servono almeno {r['min_ram']} GB)  ·  {r['size_gb']:.1f} GB"
                elif r["on_gpu"]:
                    sub += "  ·  accelerato da GPU"
                ctk.CTkLabel(row, text=sub, font=F(11), text_color=C["muted"], anchor="w"
                             ).grid(row=1, column=1, sticky="w", pady=(0, 10))

                if r["fits"]:
                    secs = r["seconds"]
                    col = C["ok"] if secs <= 20 else C["warn"] if secs <= 60 else C["danger"]
                    tbox = ctk.CTkFrame(row, fg_color="transparent")
                    tbox.grid(row=0, column=2, rowspan=2, padx=16)
                    ctk.CTkLabel(tbox, text=format_seconds(secs), font=F(16, "bold"),
                                 text_color=col).pack(anchor="e")
                    ctk.CTkLabel(tbox, text="attesa stimata", font=F(10),
                                 text_color=C["muted"]).pack(anchor="e")
                for w in (row, head):
                    if r["fits"]:
                        w.bind("<Button-1>", lambda e, n=r["name"]: model_var.set(n))

            rec = next(r for r in results if r["recommended"])
            fit_names = [r["name"] for r in results if r["fits"]]
            model_var.set(current if (current in fit_names and not first_run) else rec["name"])
            confirm_btn.configure(state="normal")

        def analyse():
            hw = detect_hardware()
            results = estimate_models(hw)
            self.after(0, lambda: fill_results(hw, results) if win.winfo_exists() else None)

        threading.Thread(target=analyse, daemon=True).start()

        def confirm():
            model = model_var.get()
            if not model:
                return
            theme = theme_var.get()
            win.destroy()
            self.settings.update({"configured": True, "model": model, "theme": theme})
            save_settings(self.settings)
            self._apply_theme(theme)
            self._set_model(model)
            self.status.configure(text=f"Motore AI impostato: {model}")
            self._check_model_installed(model)

        def on_close():
            if first_run:
                # se l'utente chiude, si usa la scelta corrente
                if model_var.get():
                    confirm()
                else:
                    self._apply_theme(theme_var.get())
                    win.destroy()
            else:
                self._apply_theme(self.settings.get("theme", "dark"), save=False)
                win.destroy()

        confirm_btn.configure(command=confirm)
        win.protocol("WM_DELETE_WINDOW", on_close)

    # --------------------------------------------- download modello Ollama
    def _installed_models(self) -> Optional[List[str]]:
        try:
            r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=3)
            if r.status_code == 200:
                return [m.get("name", "") for m in r.json().get("models", [])]
        except Exception:
            pass
        return None

    def _check_model_installed(self, model: str):
        installed = self._installed_models()
        if installed is None:
            messagebox.showinfo("Ollama non attivo",
                f"Hai scelto {model}.\n\nOllama non risulta avviato: avvialo e poi scarica il modello "
                f"aprendo un terminale e lanciando:\n\nollama pull {model}")
            return
        if any(n == model or n.startswith(model + "-") for n in installed):
            return
        size = next((m["size_gb"] for m in MODEL_CATALOG if m["name"] == model), None)
        size_txt = f" (circa {size:.1f} GB)" if size else ""
        if messagebox.askyesno("Scarica modello",
                               f"Il modello {model} non è ancora presente sul PC{size_txt}.\n\n"
                               "Vuoi scaricarlo adesso? Serve una connessione internet solo per il download."):
            self._download_model(model)

    def _download_model(self, model: str):
        win = self._modal("Download modello", 480, 210)
        ctk.CTkLabel(win, text=f"Scarico {model}", font=F(18, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=24, pady=(22, 4))
        info = ctk.CTkLabel(win, text="Connessione a Ollama…", font=F(12), text_color=C["muted"])
        info.pack(anchor="w", padx=24)
        bar = ctk.CTkProgressBar(win, height=10, corner_radius=5,
                                 progress_color=C["accent"], fg_color=C["surface2"])
        bar.set(0)
        bar.pack(fill="x", padx=24, pady=16)
        state = {"stop": False}
        ctk.CTkButton(win, text="Annulla", height=34, corner_radius=10, font=F(12),
                      fg_color=C["surface2"], hover_color=C["border"], text_color=C["text"],
                      command=lambda: (state.update(stop=True), win.destroy())
                      ).pack(anchor="e", padx=24)

        def ui(text, frac=None):
            if win.winfo_exists():
                info.configure(text=text)
                if frac is not None:
                    bar.set(frac)

        def worker():
            try:
                with requests.post(f"{OLLAMA_URL}/api/pull", json={"name": model, "stream": True},
                                   stream=True, timeout=30) as r:
                    for line in r.iter_lines():
                        if state["stop"]:
                            return
                        if not line:
                            continue
                        data = json.loads(line)
                        if data.get("error"):
                            self.after(0, lambda e=data["error"]: ui(f"Errore: {e}"))
                            return
                        total, done = data.get("total"), data.get("completed")
                        if total and done:
                            txt = f"{done / 1e9:.2f} / {total / 1e9:.2f} GB"
                            self.after(0, lambda t=txt, f=done / total: ui(t, f))
                        elif data.get("status"):
                            self.after(0, lambda s=data["status"]: ui(s))
                        if data.get("status") == "success":
                            break
                self.after(0, lambda: (win.destroy() if win.winfo_exists() else None,
                                       self.status.configure(text=f"Modello {model} pronto all'uso.")))
            except Exception as e:
                self.after(0, lambda: ui(f"Download non riuscito: {e}"))

        threading.Thread(target=worker, daemon=True).start()

    # ------------------------------------------------------ menu "@" (tag)
    QUICK_ACTIONS = [
        {"key": "cmp_excel", "icon": "⇄", "label": "Confronta due Excel",
         "desc": "Righe aggiunte, rimosse, valori e totali diversi", "types": ["excel"], "n": 2,
         "template": "Confronta {0} e {1}: elenca le righe aggiunte, rimosse o con valori diversi "
                     "e le differenze nei totali, indicando foglio e riga."},
        {"key": "cmp_any", "icon": "≠", "label": "Confronta due documenti",
         "desc": "Differenze di testo, importi, date e condizioni", "types": None, "n": 2,
         "template": "Confronta {0} con {1} e segnala tutte le differenze importanti di testo, "
                     "importi, date e condizioni."},
        {"key": "check", "icon": "✓", "label": "Verifica i conti di un documento",
         "desc": "Quantità × prezzo, imponibile + IVA, totali", "types": None, "n": 1,
         "template": "Verifica i conti di {0}: quantità × prezzo, imponibile + IVA, somma righe = "
                     "totale. Indica cosa non torna e la differenza esatta."},
        {"key": "summary", "icon": "≡", "label": "Riassumi ed estrai i dati chiave",
         "desc": "5 punti + date, importi, scadenze, controparti", "types": None, "n": 1,
         "template": "Riassumi {0} in 5 punti ed estrai date, importi, scadenze e controparti."},
    ]

    def _init_mention(self):
        self._menu = None
        self._menu_items = []      # (kind, payload, row)
        self._menu_index = 0
        self._menu_at = None       # posizione della @ nel testo
        self._act = None           # azione rapida in corso
        e = self.question_entry
        e.bind("<KeyRelease>", self._on_key_release)
        e.bind("<Return>", self._on_return)
        e.bind("<Up>", lambda ev: self._menu_move(-1))
        e.bind("<Down>", lambda ev: self._menu_move(1))
        e.bind("<Tab>", self._on_tab)
        e.bind("<Escape>", lambda ev: self._close_menu(cancel=True))
        e.bind("<FocusOut>", lambda ev: self.after(250, self._close_if_unfocused))
        self.bind("<Configure>", lambda ev: self._close_menu() if ev.widget is self else None, add="+")

    def _docs_of(self, types):
        return [d for d in self.documents if types is None or d.type in types]

    def _close_if_unfocused(self):
        try:
            f = self.focus_get()
        except Exception:
            f = None
        if f is None or not str(f).startswith(str(self.question_entry)):
            self._close_menu()

    def _mention_token(self):
        """Restituisce (posizione @, testo cercato) se il cursore è dentro un tag, altrimenti None."""
        e = self.question_entry
        text = e.get()
        cur = e.index("insert")
        before = text[:cur]
        at = before.rfind("@")
        if at < 0:
            return None
        if at > 0 and before[at - 1] not in " (:,;\n":
            return None
        token = before[at + 1:]
        if len(token) > 60:
            return None
        # tag già completato (nome documento esatto seguito da altro testo)
        for d in self.documents:
            if token.startswith(d.name) and len(token) > len(d.name):
                return None
        return at, token

    def _on_key_release(self, ev):
        if ev.keysym in ("Up", "Down", "Return", "Escape", "Tab", "Shift_L", "Shift_R"):
            return
        self._update_tag_chips()
        tok = self._mention_token()
        if tok is None:
            self._close_menu()
            return
        self._menu_at, query = tok
        self._open_menu(query)

    def _on_return(self, ev):
        if self._menu is not None and self._menu_items:
            self._menu_select(self._menu_index)
            return "break"
        self.ask_question()
        return "break"

    def _on_tab(self, ev):
        if self._menu is not None and self._menu_items:
            self._menu_select(self._menu_index)
            return "break"

    def _open_at_button(self):
        e = self.question_entry
        e.focus_set()
        text = e.get()
        if text and not text.endswith(" "):
            e.insert("end", " ")
        e.insert("end", "@")
        e.icursor("end")
        self._menu_at = len(e.get()) - 1
        self._open_menu("")

    def _open_menu(self, query: str):
        q = query.lower().strip()
        act = self._act
        types = act["action"]["types"] if act else None
        docs = [d for d in self._docs_of(types)
                if (not q or q in d.name.lower()) and (not act or d not in act["picked"])]
        actions = [] if act else [a for a in self.QUICK_ACTIONS if not q or q in a["label"].lower()]
        if q and not docs and not actions:
            self._close_menu()
            return

        if self._menu is None:
            self._menu = ctk.CTkFrame(self, corner_radius=14, fg_color=C["surface"],
                                      border_width=1, border_color=C["accent"])
        m = self._menu
        for w in m.winfo_children():
            w.destroy()
        self._menu_items = []
        self.update_idletasks()
        bar_w = max(int(self.ask_bar.winfo_width() / ctk.ScalingTracker.get_widget_scaling(self)), 480)
        ctk.CTkFrame(m, fg_color="transparent", width=bar_w - 16, height=2).pack(padx=8, pady=(4, 0))

        def header(text):
            ctk.CTkLabel(m, text=text, font=F(11, "bold"), text_color=C["muted"],
                         anchor="w").pack(fill="x", padx=16, pady=(10, 2))

        if act:
            a = act["action"]
            nth = len(act["picked"]) + 1
            kind = "file Excel" if a["types"] == ["excel"] else "documento"
            header(f"{a['label'].upper()} · SCEGLI IL {nth}° {kind.upper()}  (Esc per annullare)")
        else:
            header("DOCUMENTI CARICATI")

        if not self.documents:
            ctk.CTkLabel(m, text="Nessun documento caricato: usa “Carica documenti” nella barra laterale.",
                         font=F(12), text_color=C["muted"], anchor="w").pack(fill="x", padx=16, pady=6)
        elif not docs:
            ctk.CTkLabel(m, text="Nessun documento corrispondente.", font=F(12),
                         text_color=C["muted"], anchor="w").pack(fill="x", padx=16, pady=6)
        for d in docs[:7]:
            row = ctk.CTkFrame(m, fg_color="transparent", corner_radius=10, height=46)
            row.pack(fill="x", padx=8, pady=1)
            ext = (d.ext.lstrip(".") or "FILE").upper()
            ctk.CTkLabel(row, text=ext, width=52, height=30, corner_radius=8,
                         fg_color=d.border_color, text_color="white",
                         font=F(11, "bold")).pack(side="left", padx=(8, 10), pady=7)
            ctk.CTkLabel(row, text=doc_title(d), font=F(13, "bold"),
                         text_color=C["text"], anchor="w").pack(side="left")
            ctk.CTkLabel(row, text=d.name, font=F(11), text_color=C["muted"],
                         anchor="e").pack(side="right", padx=12)
            self._menu_items.append(("doc", d, row))

        if actions:
            header("AZIONI RAPIDE")
            for a in actions:
                available = len(self._docs_of(a["types"])) >= a["n"]
                row = ctk.CTkFrame(m, fg_color="transparent", corner_radius=10, height=46)
                row.pack(fill="x", padx=8, pady=1)
                ctk.CTkLabel(row, text=a["icon"], width=52, height=30, corner_radius=8,
                             fg_color=C["accent_bg"], text_color=C["accent"],
                             font=F(15, "bold")).pack(side="left", padx=(8, 10), pady=7)
                ctk.CTkLabel(row, text=a["label"], font=F(13, "bold"),
                             text_color=C["text"] if available else C["muted"],
                             anchor="w").pack(side="left")
                if available:
                    note = a["desc"]
                else:
                    need = "file Excel" if a["types"] == ["excel"] else "documenti"
                    note = f"Servono {a['n']} {need}" if a["n"] > 1 else "Carica un documento"
                ctk.CTkLabel(row, text=note, font=F(11), text_color=C["muted"],
                             anchor="e").pack(side="right", padx=12)
                if available:
                    self._menu_items.append(("action", a, row))
        ctk.CTkFrame(m, fg_color="transparent", width=bar_w - 16, height=4).pack(padx=8, pady=(0, 4))

        for i, (_, _, row) in enumerate(self._menu_items):
            for w in [row] + row.winfo_children():
                w.bind("<Button-1>", lambda ev, i=i: self._menu_select(i))
                w.bind("<Enter>", lambda ev, i=i: self._menu_highlight(i))
                try:
                    w.configure(cursor="hand2")
                except Exception:
                    pass
        self._menu_index = 0
        self._menu_highlight(0)

        bar = self.ask_bar
        self.update_idletasks()
        x = bar.winfo_rootx() - self.winfo_rootx()
        y = bar.winfo_rooty() - self.winfo_rooty() - 8
        sc = ctk.ScalingTracker.get_widget_scaling(self)
        m.place(x=x / sc, y=y / sc, anchor="sw")
        m.lift()

    def _menu_highlight(self, i):
        if not self._menu_items:
            return
        self._menu_index = max(0, min(i, len(self._menu_items) - 1))
        for j, (_, _, row) in enumerate(self._menu_items):
            row.configure(fg_color=C["accent_bg"] if j == self._menu_index else "transparent")

    def _menu_move(self, step):
        if self._menu is None or not self._menu_items:
            return
        self._menu_highlight((self._menu_index + step) % len(self._menu_items))
        return "break"

    def _close_menu(self, cancel=False):
        if self._menu is not None:
            self._menu.place_forget()
        self._menu_items = []
        if cancel and self._act:
            # annulla l'azione rapida e ripristina il testo
            e = self.question_entry
            e.delete(self._act["start"], "end")
            self._act = None
            self._update_tag_chips()
        return "break" if cancel else None

    def _replace_token(self, new_text):
        """Sostituisce '@testo-cercato' (dalla @ al cursore) con new_text."""
        e = self.question_entry
        cur = e.index("insert")
        start = self._menu_at if self._menu_at is not None else cur
        e.delete(start, cur)
        e.insert(start, new_text)
        e.icursor(start + len(new_text))
        return start

    def _menu_select(self, i):
        if not self._menu_items:
            return
        kind, payload, _ = self._menu_items[i]
        e = self.question_entry
        e.focus_set()
        if kind == "action":
            start = self._replace_token("")
            self._act = {"action": payload, "start": start, "picked": []}
            lead = f"{payload['label']}: @"
            e.insert(start, lead)
            e.icursor(start + len(lead))
            self._menu_at = start + len(lead) - 1
            self._open_menu("")
            return

        doc = payload
        if self._act:
            act = self._act
            act["picked"].append(doc)
            a = act["action"]
            tags = [f"@{d.name}" for d in act["picked"]]
            e.delete(act["start"], "end")
            if len(act["picked"]) < a["n"]:
                txt = f"{a['label']}: " + " e ".join(tags) + " e @"
                e.insert(act["start"], txt)
                e.icursor("end")
                self._menu_at = len(e.get()) - 1
                self._open_menu("")
                return
            e.insert(act["start"], a["template"].format(*tags))
            e.icursor("end")
            self._act = None
        else:
            self._replace_token(f"@{doc.name} ")
        self._close_menu()
        self._update_tag_chips()

    def _tagged_docs(self, text: str) -> List[Document]:
        found = []
        for d in sorted(self.documents, key=lambda d: len(d.name), reverse=True):
            if f"@{d.name}" in text and d not in found:
                found.append(d)
        found.sort(key=lambda d: text.find(f"@{d.name}"))
        return found

    def _update_tag_chips(self):
        tagged = self._tagged_docs(self.question_entry.get())
        if tagged:
            self.suggest_row.grid_remove()
            for w in self.tag_row.winfo_children():
                w.destroy()
            ctk.CTkLabel(self.tag_row, text="Documenti citati:", font=F(12),
                         text_color=C["muted"]).pack(side="left", padx=(2, 8))
            for d in tagged:
                ctk.CTkLabel(self.tag_row, text=f" {(d.ext.lstrip('.') or 'FILE').upper()} · {doc_title(d)} ",
                             height=26, corner_radius=13, fg_color=d.border_color,
                             text_color="white", font=F(12, "bold")).pack(side="left", padx=(0, 6))
            self.tag_row.grid()
        else:
            self.tag_row.grid_remove()
            self.suggest_row.grid()

    # ------------------------------------------------------------ documenti
    def load_documents(self):
        files = filedialog.askopenfilenames(
            title="Seleziona documenti",
            filetypes=[("Documenti", "*.pdf *.docx *.xlsx *.xls *.csv *.pptx *.txt *.png *.jpg *.jpeg"), ("Tutti", "*.*")]
        )
        if not files:
            return
        self.status.configure(text="Lettura documenti in corso…")
        self.update_idletasks()
        for f in files:
            try:
                doc = Document(f)
                self.documents.append(doc)
                self._add_card(doc)
            except Exception as e:
                messagebox.showerror("Errore", str(e))
        self._refresh_count()
        self.status.configure(text=f"{len(self.documents)} documento/i caricati.")

    def _add_card(self, doc: Document):
        if getattr(self, "empty_docs", None) is not None:
            self.empty_docs.destroy()
            self.empty_docs = None

        img_n, img_h = make_card(doc), make_card(doc, hover=True)
        if img_n is None:  # Pillow non disponibile: card semplice
            card = ctk.CTkLabel(self.docs_frame, width=CARD_W, height=CARD_H, corner_radius=14,
                                text=f"{doc.ext.upper()}\n{doc_title(doc)}", fg_color=C["surface2"])
            card.pack(side="left", padx=6, pady=4)
            card.bind("<Button-3>", lambda e: self._remove_doc(doc, card))
            return
        ci_n = ctk.CTkImage(light_image=img_n, dark_image=img_n, size=(CARD_W, CARD_H))
        ci_h = ctk.CTkImage(light_image=img_h, dark_image=img_h, size=(CARD_W, CARD_H))
        card = ctk.CTkLabel(self.docs_frame, image=ci_n, text="", cursor="hand2")
        card._imgs = (ci_n, ci_h)
        card.pack(side="left", padx=6, pady=4)

        def on_click(e):
            # coordinate logiche (la card può essere ridimensionata dallo scaling di Windows)
            sx = CARD_W / max(card.winfo_width(), 1)
            sy = CARD_H / max(card.winfo_height(), 1)
            if e.x * sx > CARD_W - 32 and e.y * sy < 32:
                self._remove_doc(doc, card)
            else:  # clic sulla card: cita il documento nella domanda
                self._open_at_button()
                self._replace_token(f"@{doc.name} ")
                self._close_menu()
                self._update_tag_chips()

        card.bind("<Enter>", lambda e: card.configure(image=ci_h))
        card.bind("<Leave>", lambda e: card.configure(image=ci_n))
        card.bind("<Button-1>", on_click)

    def _remove_doc(self, doc, card):
        if doc in self.documents:
            self.documents.remove(doc)
        card.destroy()
        self._refresh_count()
        if not self.documents:
            self._show_empty_docs()
        self.status.configure(text=f"{len(self.documents)} documento/i caricati.")

    def clear_docs(self):
        self.documents.clear()
        for w in self.docs_frame.winfo_children():
            w.destroy()
        self.empty_docs = None
        self._show_empty_docs()
        self._refresh_count()
        self.status.configure(text="Documenti rimossi.")

    def _get_context(self) -> str:
        if not self.documents:
            return "Nessun documento caricato."
        parts = []
        for d in self.documents:
            parts.append(f"=== {d.name} ===\n{d.text[:8000]}")
        return "\n\n".join(parts)

    # --------------------------------------------------------------- azioni
    def ask_question(self):
        q = self.question_entry.get().strip()
        if not q or self.busy:
            return
        self._close_menu()
        self.question_entry.delete(0, "end")
        self._update_tag_chips()
        tagged = self._tagged_docs(q)
        if tagged:
            names = ", ".join(d.name for d in tagged)
            context = "\n\n".join(f"=== {d.name} ===\n{d.text[:12000]}" for d in tagged)
            prompt = (f"{q}\n\n(I documenti citati con @ sono: {names}. "
                      "Usa solo questi documenti per rispondere.)")
            self._run_ai(f"Tu: {q}", prompt, context)
        else:
            self._run_ai(f"Tu: {q}", q, self._get_context())

    def verify_accounts(self):
        if not self.documents:
            messagebox.showwarning("Attenzione", "Carica almeno un documento.")
            return
        prompt = ("Analizza i documenti e verifica tutti i calcoli contabili possibili: "
                  "quantità × prezzo, imponibile + IVA, somma righe = totale, aliquote IVA. "
                  "Segnala chiaramente cosa torna e cosa non torna, con le differenze esatte.")
        self._run_ai("✓ Verifica conti", prompt, self._get_context(), "Verifica completata.")

    def compare_docs(self):
        if len(self.documents) < 2:
            messagebox.showwarning("Attenzione", "Carica almeno due documenti.")
            return
        names = [f"{i+1}. {d.name}" for i, d in enumerate(self.documents)]

        win = self._modal("Confronta documenti", 520, 300)
        ctk.CTkLabel(win, text="Confronta due documenti", font=F(20, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=28, pady=(24, 14))
        v1 = ctk.StringVar(value=names[0])
        v2 = ctk.StringVar(value=names[1])
        for lbl, var in (("Documento 1", v1), ("Documento 2", v2)):
            ctk.CTkLabel(win, text=lbl, font=F(12), text_color=C["muted"]).pack(anchor="w", padx=28)
            ctk.CTkOptionMenu(win, values=names, variable=var, height=36, corner_radius=10,
                              font=F(13), fg_color=C["surface2"], button_color=C["accent"],
                              button_hover_color=C["accent_h"], text_color=C["text"]
                              ).pack(fill="x", padx=28, pady=(2, 10))

        def go():
            i1 = names.index(v1.get())
            i2 = names.index(v2.get())
            if i1 == i2:
                messagebox.showwarning("Attenzione", "Scegli due documenti diversi.", parent=win)
                return
            win.destroy()
            d1, d2 = self.documents[i1], self.documents[i2]
            prompt = ("Confronta questi due documenti e segnala tutte le differenze importanti "
                      f"(testo, importi, date, totali):\n\nDOCUMENTO 1:\n{d1.text[:5000]}"
                      f"\n\nDOCUMENTO 2:\n{d2.text[:5000]}")
            self._run_ai(f"⇄ Confronto: {d1.name} ↔ {d2.name}", prompt, "", "Confronto completato.")

        ctk.CTkButton(win, text="Confronta", height=42, corner_radius=12, font=F(14, "bold"),
                      fg_color=C["accent"], hover_color=C["accent_h"], text_color="white",
                      command=go).pack(fill="x", padx=28, pady=(8, 20))

    def edit_and_save_copy(self):
        """Chiede all'AI di modificare un documento Word e permette di salvare o sovrascrivere."""
        if not self.documents:
            messagebox.showwarning("Attenzione", "Carica almeno un documento.")
            return
        word_docs = [d for d in self.documents if d.ext == ".docx"]
        if not word_docs:
            messagebox.showinfo("Info", "Carica un file Word (.docx) per poterlo modificare con l'AI.")
            return

        names = [f"{i+1}. {d.name}" for i, d in enumerate(word_docs)]
        win = self._modal("Modifica con AI", 560, 380)
        ctk.CTkLabel(win, text="Modifica con AI", font=F(20, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=28, pady=(24, 12))
        ctk.CTkLabel(win, text="Documento Word", font=F(12), text_color=C["muted"]).pack(anchor="w", padx=28)
        sel = ctk.StringVar(value=names[0])
        ctk.CTkOptionMenu(win, values=names, variable=sel, height=36, corner_radius=10, font=F(13),
                          fg_color=C["surface2"], button_color=C["accent"],
                          button_hover_color=C["accent_h"], text_color=C["text"]
                          ).pack(fill="x", padx=28, pady=(2, 12))
        ctk.CTkLabel(win, text="Cosa vuoi modificare?", font=F(12),
                     text_color=C["muted"]).pack(anchor="w", padx=28)
        box = ctk.CTkTextbox(win, height=90, corner_radius=10, font=F(13),
                             fg_color=C["surface"], border_width=1, border_color=C["border"],
                             text_color=C["text"])
        box.pack(fill="x", padx=28, pady=(2, 4))
        ctk.CTkLabel(win, text="Es.: “Aggiungi alla fine una nota di ringraziamento” oppure “Cancella il secondo paragrafo”",
                     font=F(11), text_color=C["muted"], wraplength=500, justify="left").pack(anchor="w", padx=28)

        def go():
            modifica = box.get("1.0", "end").strip()
            if not modifica:
                return
            doc = word_docs[names.index(sel.get())]
            win.destroy()
            self._do_edit(doc, modifica)

        ctk.CTkButton(win, text="Applica modifica", height=42, corner_radius=12, font=F(14, "bold"),
                      fg_color=C["accent"], hover_color=C["accent_h"], text_color="white",
                      command=go).pack(fill="x", padx=28, pady=(12, 20))

    def _do_edit(self, doc: Document, modifica: str):
        prompt = (
            f"Documento originale:\n{doc.text}\n\n"
            f"Modifica richiesta: {modifica}\n\n"
            "Riscrivi l'INTERO documento applicando la modifica richiesta. "
            "Mantieni la formattazione il più possibile semplice (paragrafi). "
            "Restituisci SOLO il testo finale del documento, senza commenti."
        )
        self._write(f"\n✎ Modifica di {doc.name}: {modifica}\n", "user")
        self._write("Sto modificando il documento…\n", "info")
        self._set_busy(True, "Modifica in corso…")

        def worker():
            nuovo = self.ai.ask(prompt)
            self.after(0, lambda: self._edit_done(doc, nuovo))

        threading.Thread(target=worker, daemon=True).start()

    def _edit_done(self, doc: Document, nuovo_testo: str):
        self._set_busy(False, "Anteprima pronta.")
        if not nuovo_testo or "Errore" in nuovo_testo or nuovo_testo.startswith("Ollama non disponibile"):
            self._write("Impossibile ottenere la modifica dall'AI.\n", "error")
            messagebox.showerror("Errore", "Impossibile ottenere la modifica dall'AI.")
            return
        self._write("Anteprima aperta in una nuova finestra.\n", "info")

        preview = ctk.CTkToplevel(self)
        preview.title("Anteprima modifica")
        preview.geometry("780x580")
        preview.configure(fg_color=C["bg"])
        preview.transient(self)
        preview.after(100, preview.lift)
        ctk.CTkLabel(preview, text="Anteprima modifica", font=F(20, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=20, pady=(18, 2))
        ctk.CTkLabel(preview, text=doc.name, font=F(12), text_color=C["muted"]).pack(anchor="w", padx=20)
        txt = ctk.CTkTextbox(preview, wrap="word", corner_radius=12, font=F(13),
                             fg_color=C["surface"], border_width=1, border_color=C["border"],
                             text_color=C["text"])
        txt.pack(fill="both", expand=True, padx=20, pady=12)
        txt.insert("end", nuovo_testo)
        txt.configure(state="disabled")

        def salva(sovrascrivi=False):
            try:
                from docx import Document as DocxDocument
                new_doc = DocxDocument()
                for paragrafo in nuovo_testo.split("\n"):
                    if paragrafo.strip():
                        new_doc.add_paragraph(paragrafo.strip())

                if sovrascrivi:
                    save_path = doc.path
                else:
                    base, ext = os.path.splitext(doc.path)
                    save_path = f"{base}_copia{ext}"
                    counter = 1
                    while os.path.exists(save_path):
                        save_path = f"{base}_copia_{counter}{ext}"
                        counter += 1

                new_doc.save(save_path)
                messagebox.showinfo("Salvato", f"Documento salvato:\n{os.path.basename(save_path)}")
                preview.destroy()

                if not sovrascrivi:
                    new_loaded = Document(save_path)
                    self.documents.append(new_loaded)
                    self._add_card(new_loaded)
                    self._refresh_count()
                else:
                    doc.text = nuovo_testo
                self.status.configure(text="Documento modificato e salvato.")
            except Exception as e:
                messagebox.showerror("Errore", f"Errore nel salvataggio:\n{e}")

        bar = ctk.CTkFrame(preview, fg_color="transparent")
        bar.pack(fill="x", padx=20, pady=(0, 18))
        ctk.CTkButton(bar, text="Salva come copia (_copia)", height=40, corner_radius=10,
                      font=F(13, "bold"), fg_color=C["accent"], hover_color=C["accent_h"],
                      text_color="white", command=lambda: salva(False)).pack(side="left")
        ctk.CTkButton(bar, text="Sovrascrivi originale", height=40, corner_radius=10,
                      font=F(13), fg_color=C["danger"], hover_color=C["danger_h"],
                      text_color="white", command=lambda: salva(True)).pack(side="left", padx=10)
        ctk.CTkButton(bar, text="Annulla", height=40, corner_radius=10, font=F(13),
                      fg_color=C["surface2"], hover_color=C["border"], text_color=C["text"],
                      command=preview.destroy).pack(side="right")

    # ---------------------------------------------------------------- aiuto
    def show_manual(self):
        win = ctk.CTkToplevel(self)
        win.title("Manuale TALdoc Desktop")
        win.geometry("760x640")
        win.configure(fg_color=C["bg"])
        win.transient(self)
        win.after(100, win.lift)
        ctk.CTkLabel(win, text="Manuale d'uso", font=F(22, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=22, pady=(18, 6))
        txt = ctk.CTkTextbox(win, wrap="word", corner_radius=12, font=F(13),
                             fg_color=C["surface"], border_width=1, border_color=C["border"],
                             text_color=C["text"])
        txt.pack(fill="both", expand=True, padx=20, pady=(4, 20))
        txt.insert("end", MANUALE.strip())
        txt.configure(state="disabled")

    def show_about(self):
        messagebox.showinfo("Informazioni",
            f"{APP_NAME}\n\nApp locale per interrogare documenti con AI.\n"
            "Tutto avviene sul tuo PC.\n\nVersione 1.3.")


if __name__ == "__main__":
    app = TALdocApp()
    app.mainloop()
