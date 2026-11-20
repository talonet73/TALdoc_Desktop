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
    "word": "#3B82F6",
    "powerpoint": "#F97316",
    "pdf": "#EF4444",
    "excel": "#22C55E",
    "txt": "#EAB308",
    "other": "#94A3B8"
}

TYPE_ICONS = {
    "word": "W",
    "powerpoint": "P",
    "pdf": "PDF",
    "excel": "X",
    "txt": "TXT",
    "other": "IMG"
}

MANUALE = """
TALdoc Desktop - Manuale d'uso

1. COS'È TALDOC
TALdoc è un'app per Windows che legge i tuoi documenti in locale (senza internet) e ti permette di interrogarli con un'intelligenza artificiale locale (tramite Ollama).
Puoi fare domande libere, chiedere calcoli, verificare totali e IVA, confrontare due documenti e trovare differenze.

Tutto resta sul tuo PC. Nessun dato viene inviato online.

2. PRIMO AVVIO
- Assicurati di avere Ollama installato (scaricalo da ollama.com e avvialo).
- All'avvio l'app ti chiede quanta RAM ha il PC e ti suggerisce un modello leggero.
- Scarica il modello consigliato con il comando che ti indica (es. ollama pull qwen2.5:3b).
- Poi puoi usare subito l'app.

3. CARICARE DOCUMENTI
Clicca "Carica documenti" nella barra laterale e seleziona uno o più file.
Formati supportati: PDF, Word, Excel, CSV, PowerPoint, TXT, immagini (JPG, PNG...).

I documenti appaiono in fila orizzontale con un bordo colorato:
- Blu = Word
- Arancione = PowerPoint
- Rosso = PDF
- Verde = Excel/CSV
- Giallo = TXT

4. FARE DOMANDE
Scrivi una domanda nel box e premi "Chiedi" o Invio.
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
- Modalità scura/chiara più personalizzabile
- Integrazione diretta con cartelle di Windows (drag & drop)
- Storico completo delle domande e risposte
- Modelli AI più grandi se hai molta RAM
- Correzione manuale del testo OCR direttamente nell'interfaccia

9. INDICATORE RAM
Nella barra laterale vedi l'utilizzo attuale della RAM del PC (richiede psutil).
Il colore diventa arancione sopra il 70% e rosso sopra l'85%.

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
                timeout=120
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
#  Tema grafico moderno (ogni colore è una coppia: modalità chiara, scura)
# ---------------------------------------------------------------------------
FONT_FAMILY = "Segoe UI"

C = {
    "bg":        ("#F4F6FB", "#0B1020"),
    "sidebar":   ("#FFFFFF", "#111827"),
    "surface":   ("#FFFFFF", "#161E31"),
    "surface2":  ("#EEF1F8", "#1E2840"),
    "border":    ("#E2E6F0", "#263250"),
    "text":      ("#0F172A", "#E5E9F5"),
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
        ctk.set_appearance_mode("dark")
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
        self.after(300, self._ask_ram_on_start)

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
        logo_set = False
        if Image is not None and os.path.exists("logo.jpg"):
            try:
                logo_img = ctk.CTkImage(light_image=Image.open("logo.jpg"),
                                        dark_image=Image.open("logo.jpg"), size=(40, 40))
                ctk.CTkLabel(brand, image=logo_img, text="").pack(side="left")
                logo_set = True
            except Exception:
                pass
        if not logo_set:
            ctk.CTkLabel(brand, text="T", width=40, height=40, corner_radius=12,
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

        theme_row = ctk.CTkFrame(bottom, fg_color="transparent")
        theme_row.pack(fill="x")
        ctk.CTkLabel(theme_row, text="Tema scuro", font=F(12),
                     text_color=C["muted"]).pack(side="left", padx=4)
        self.theme_switch = ctk.CTkSwitch(theme_row, text="", command=self._toggle_theme,
                                          progress_color=C["accent"], width=44)
        self.theme_switch.select()
        self.theme_switch.pack(side="right")

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
        ask.grid(row=3, column=0, sticky="ew", padx=16, pady=(8, 16))
        ask.grid_columnconfigure(0, weight=1)
        self.question_entry = ctk.CTkEntry(
            ask, placeholder_text="Fai una domanda sui documenti… es. Qual è il totale? Controlla i calcoli IVA",
            fg_color="transparent", border_width=0, height=46, font=F(14),
            text_color=C["text"], placeholder_text_color=C["muted"])
        self.question_entry.grid(row=0, column=0, sticky="ew", padx=(14, 6), pady=4)
        self.question_entry.bind("<Return>", lambda e: self.ask_question())
        self.ask_btn = ctk.CTkButton(ask, text="Chiedi  ➤", command=self.ask_question,
                                     width=110, height=38, corner_radius=10, font=F(13, "bold"),
                                     fg_color=C["accent"], hover_color=C["accent_h"],
                                     text_color="white")
        self.ask_btn.grid(row=0, column=1, padx=6, pady=4)

        # Suggerimenti rapidi
        chips = ctk.CTkFrame(chat, fg_color="transparent")
        chips.grid(row=4, column=0, sticky="w", padx=16, pady=(0, 14))
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

    def _toggle_theme(self):
        ctk.set_appearance_mode("dark" if self.theme_switch.get() else "light")
        self._setup_text_tags()

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

    def _ask_ram_on_start(self):
        win = self._modal("Configurazione iniziale", 520, 400)
        ctk.CTkLabel(win, text="Benvenuto 👋", font=F(22, "bold"),
                     text_color=C["text"]).pack(anchor="w", padx=28, pady=(26, 4))
        ctk.CTkLabel(win, text="Quanta RAM ha questo PC? Ti suggerisco il modello AI più adatto.",
                     font=F(13), text_color=C["muted"], wraplength=460,
                     justify="left").pack(anchor="w", padx=28)

        options = [
            ("8 GB o meno", 8, "qwen2.5:3b", "leggero e veloce"),
            ("16 GB", 16, "qwen2.5:7b", "buon equilibrio"),
            ("32 GB o più", 32, "qwen2.5:14b", "più preciso"),
        ]
        choice = ctk.IntVar(value=8)
        box = ctk.CTkFrame(win, fg_color="transparent")
        box.pack(fill="x", padx=28, pady=18)
        for label, val, model, desc in options:
            card = ctk.CTkFrame(box, fg_color=C["surface"], corner_radius=12,
                                border_width=1, border_color=C["border"])
            card.pack(fill="x", pady=4)
            ctk.CTkRadioButton(card, text=f"{label}   ·   {model}  ({desc})", variable=choice,
                               value=val, font=F(13), text_color=C["text"],
                               fg_color=C["accent"], hover_color=C["accent_h"]
                               ).pack(anchor="w", padx=14, pady=12)

        def confirm():
            ram = choice.get()
            win.destroy()
            self._apply_ram(ram)

        ctk.CTkButton(win, text="Continua", height=42, corner_radius=12, font=F(14, "bold"),
                      fg_color=C["accent"], hover_color=C["accent_h"], text_color="white",
                      command=confirm).pack(fill="x", padx=28, pady=(0, 20))
        win.protocol("WM_DELETE_WINDOW", confirm)

    def _apply_ram(self, ram: int):
        if ram <= 8:
            model, desc = "qwen2.5:3b", "leggero e veloce"
        elif ram <= 16:
            model, desc = "qwen2.5:7b", "buon equilibrio"
        else:
            model, desc = "qwen2.5:14b", "più preciso"
        self.ai.model = model
        self.model_label.configure(text=f"Modello: {model}")
        messagebox.showinfo("Modello consigliato",
            f"RAM {ram} GB → ti consiglio {model} ({desc}).\n\n"
            f"Se non lo hai ancora, apri un terminale e lancia:\nollama pull {model}\n\n"
            "Poi puoi usare subito l'app.")

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

        card = ctk.CTkFrame(self.docs_frame, width=170, height=134, corner_radius=14,
                            fg_color=C["surface2"], border_width=2, border_color=doc.border_color)
        card.pack(side="left", padx=6, pady=4)
        card.pack_propagate(False)

        top = ctk.CTkFrame(card, fg_color="transparent")
        top.pack(fill="x", padx=12, pady=(12, 6))
        ctk.CTkLabel(top, text=TYPE_ICONS.get(doc.type, "DOC"), width=44, height=44,
                     corner_radius=10, fg_color=doc.border_color, text_color="white",
                     font=F(13 if len(TYPE_ICONS.get(doc.type, "")) <= 1 else 11, "bold")
                     ).pack(side="left")
        ctk.CTkButton(top, text="✕", width=26, height=26, corner_radius=13, font=F(11),
                      fg_color="transparent", hover_color=C["border"], text_color=C["muted"],
                      command=lambda: self._remove_doc(doc, card)).pack(side="right", anchor="n")

        name = doc.name if len(doc.name) <= 20 else doc.name[:18] + "…"
        ctk.CTkLabel(card, text=name, font=F(12, "bold"), text_color=C["text"],
                     anchor="w").pack(fill="x", padx=12)
        words = len(doc.text.split()) if doc.text else 0
        ctk.CTkLabel(card, text=f"{doc.type.upper()} · {words:,} parole".replace(",", "."),
                     font=F(11), text_color=C["muted"], anchor="w").pack(fill="x", padx=12)

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
        self.question_entry.delete(0, "end")
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
            "Tutto avviene sul tuo PC.\n\nVersione 1.1 – nuova interfaccia grafica.")


if __name__ == "__main__":
    app = TALdocApp()
    app.mainloop()
