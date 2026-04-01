#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Interface gráfica para sanitizar RTF (remoção de lixo DDE após {\*\bkmkstart __DdeLink__).
"""

from __future__ import annotations

import ctypes
import os
import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import quote_plus

from db_sanitize import (
    export_batch_report_csv,
    get_batch_report,
    list_postgres_columns,
    list_postgres_tables,
    rollback_batch,
    sanitize_documento_mesclado,
    test_postgres_connection,
)
from rtf_sanitize import (
    AGGRESSIVE_LEVEL,
    INTERMEDIATE_LEVEL,
    SAFE_LEVEL,
    DEFAULT_MARKERS,
    MARKER_DDE_BOOKMARK,
    analisar_limpeza,
    carregar_marcadores_de_json,
    limpar_arquivo_rtf,
    parece_rtf,
)


def _ler_texto_preservando_bytes(caminho: Path) -> tuple[str, str]:
    """
    Lê ficheiro em binário e decodifica como latin-1 (preserva todos os bytes).
    Retorna (texto, etiqueta do modo) para mostrar ao utilizador.
    """
    data = caminho.read_bytes()
    return data.decode("latin-1"), "latin-1 (preservação byte-a-byte)"


def _guardar_texto_preservando_bytes(caminho: Path, texto: str) -> None:
    caminho.write_bytes(texto.encode("latin-1"))


class App(tk.Tk):
    def __init__(self) -> None:
        super().__init__()
        self._apply_dark_theme()
        self.title("Sanitizador RTF — DDE / bookmark")
        self.minsize(920, 620)
        self.geometry("1040x720")
        self.update_idletasks()
        self._schedule_windows_titlebar_dark()
        self.bind("<Map>", lambda _e: self._apply_windows_titlebar_dark(), add="+")

        self._queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._last_batch_id = ""
        self._stop_requested = threading.Event()
        self._progress_animating = False
        self._progress_base_text = "Progresso: parado"
        self._progress_anim_step = 0
        self._build()
        self.after(200, self._poll_queue)

    def _apply_windows_titlebar_dark(self) -> None:
        """
        Barra de título nativa (minimizar/maximizar/fechar) no tema escuro.

        No Tk, winfo_id() é o HWND do cliente interno; o DWM precisa do HWND
        da janela top-level (GetParent), senão a API não altera a barra.
        Opcionalmente aplica DWMWA_CAPTION_COLOR no Windows 11 22H2+ para
        harmonizar com o fundo da app (#1f2128).
        """
        if os.name != "nt":
            return
        try:
            wid = int(self.winfo_id())
            hwnd = int(ctypes.windll.user32.GetParent(wid))
            if hwnd == 0:
                hwnd = wid

            dark = ctypes.c_int(1)
            dwm = ctypes.windll.dwmapi
            DWMWA_USE_IMMERSIVE_DARK_MODE = 20
            DWMWA_USE_IMMERSIVE_DARK_MODE_LEGACY = 19
            # Win11 / Win10 20H1+
            if dwm.DwmSetWindowAttribute(
                hwnd,
                DWMWA_USE_IMMERSIVE_DARK_MODE,
                ctypes.byref(dark),
                ctypes.sizeof(dark),
            ) != 0:
                dwm.DwmSetWindowAttribute(
                    hwnd,
                    DWMWA_USE_IMMERSIVE_DARK_MODE_LEGACY,
                    ctypes.byref(dark),
                    ctypes.sizeof(dark),
                )

        except Exception:
            return
        # Windows 11 22H2+: cor da legenda; falha em builds antigas sem efeito na linha seguinte
        try:
            DWMWA_CAPTION_COLOR = 35
            caption_bgr = ctypes.c_int(0x0028211F)  # #1f2128 (0x00BBGGRR)
            dwm = ctypes.windll.dwmapi
            wid = int(self.winfo_id())
            hwnd = int(ctypes.windll.user32.GetParent(wid)) or wid
            dwm.DwmSetWindowAttribute(
                hwnd,
                DWMWA_CAPTION_COLOR,
                ctypes.byref(caption_bgr),
                ctypes.sizeof(caption_bgr),
            )
        except Exception:
            pass

    def _schedule_windows_titlebar_dark(self) -> None:
        """Reaplica DWM após o conteúdo existir — o HWND às vezes só estabiliza depois do map."""
        if os.name != "nt":
            return
        self._apply_windows_titlebar_dark()
        self.after(150, self._apply_windows_titlebar_dark)
        self.after(450, self._apply_windows_titlebar_dark)

    def _apply_dark_theme(self) -> None:
        bg = "#1f2128"
        panel = "#262a33"
        panel2 = "#2d3340"
        fg = "#e6eaf2"
        muted = "#a8b0c0"
        accent = "#5b8cff"
        field = "#2a2f3a"
        border = "#353c4a"
        border_soft = "#2f3542"

        style = ttk.Style(self)
        try:
            style.theme_use("clam")
        except tk.TclError:
            pass

        self.configure(bg=bg)
        self.option_add("*Foreground", fg)
        self.option_add("*Background", bg)
        self.option_add("*selectBackground", accent)
        self.option_add("*selectForeground", "#ffffff")
        self.option_add("*Entry.Background", field)
        self.option_add("*Entry.Foreground", fg)
        self.option_add("*Text.Background", field)
        self.option_add("*Text.Foreground", fg)
        self.option_add("*insertBackground", "#ffffff")
        self.option_add("*highlightBackground", bg)
        self.option_add("*highlightColor", accent)

        style.configure(".", background=bg, foreground=fg)
        style.configure(".", focuscolor=border_soft)
        style.configure("TFrame", background=bg)
        style.configure(
            "TLabelframe",
            background=bg,
            foreground=fg,
            bordercolor=border_soft,
            lightcolor=border_soft,
            darkcolor=border_soft,
            borderwidth=1,
            relief="solid",
        )
        style.configure("TLabelframe.Label", background=bg, foreground=fg)
        style.configure("TLabel", background=bg, foreground=fg)
        style.configure("TCheckbutton", background=bg, foreground=fg)
        style.configure("TRadiobutton", background=bg, foreground=fg)
        style.map(
            "TCheckbutton",
            background=[("active", bg)],
            foreground=[("disabled", muted), ("active", fg)],
        )
        style.map(
            "TRadiobutton",
            background=[("active", bg)],
            foreground=[("disabled", muted), ("active", fg)],
        )

        style.configure(
            "TButton",
            background=panel2,
            foreground=fg,
            bordercolor=border_soft,
            lightcolor=panel2,
            darkcolor=panel2,
            padding=(10, 5),
        )
        style.map(
            "TButton",
            background=[("active", "#384053"), ("pressed", "#333a4a")],
            foreground=[("disabled", muted), ("active", "#ffffff")],
        )

        style.configure(
            "TEntry",
            fieldbackground=field,
            foreground=fg,
            insertcolor="#ffffff",
            bordercolor=border_soft,
            lightcolor=border_soft,
            darkcolor=border_soft,
        )
        style.map(
            "TEntry",
            fieldbackground=[("readonly", "#242934")],
            foreground=[("readonly", "#d5d8df")],
            bordercolor=[("focus", "#4f79d9"), ("!focus", border_soft)],
            lightcolor=[("focus", "#4f79d9"), ("!focus", border_soft)],
            darkcolor=[("focus", "#4f79d9"), ("!focus", border_soft)],
        )

        style.configure(
            "TCombobox",
            fieldbackground=field,
            background=panel2,
            foreground=fg,
            bordercolor=border_soft,
            lightcolor=border_soft,
            darkcolor=border_soft,
            arrowsize=14,
        )
        style.map(
            "TCombobox",
            fieldbackground=[("readonly", "#242934")],
            foreground=[("readonly", "#d5d8df")],
            background=[("active", "#384053")],
            bordercolor=[("focus", "#4f79d9"), ("!focus", border_soft)],
            lightcolor=[("focus", "#4f79d9"), ("!focus", border_soft)],
            darkcolor=[("focus", "#4f79d9"), ("!focus", border_soft)],
        )

        style.configure(
            "Dark.TNotebook",
            background=bg,
            borderwidth=1,
            bordercolor=border_soft,
            lightcolor=border_soft,
            darkcolor=border_soft,
            tabmargins=(6, 6, 6, 0),
        )
        style.configure(
            "Dark.TNotebook.Tab",
            background="#2a3040",
            foreground="#c9d3e7",
            padding=(14, 8),
            bordercolor=border_soft,
            lightcolor=border_soft,
            darkcolor=border_soft,
            borderwidth=1,
        )
        style.map(
            "Dark.TNotebook.Tab",
            background=[
                ("selected", "#3b4b69"),
                ("active", "#33425e"),
                ("!selected", "#2a3040"),
            ],
            foreground=[
                ("selected", "#ffffff"),
                ("active", "#f3f6ff"),
                ("!selected", "#c9d3e7"),
            ],
            bordercolor=[("selected", "#4f79d9"), ("!selected", border_soft)],
            lightcolor=[("selected", "#4f79d9"), ("!selected", border_soft)],
            darkcolor=[("selected", "#4f79d9"), ("!selected", border_soft)],
        )

        style.configure("Horizontal.TScrollbar", background=panel2, troughcolor=panel)
        style.configure("Vertical.TScrollbar", background=panel2, troughcolor=panel)

    def _build(self) -> None:
        pad = {"padx": 10, "pady": 8}
        frm = ttk.Frame(self, padding=12)
        frm.pack(fill=tk.BOTH, expand=True)

        header = ttk.Frame(frm)
        header.pack(fill=tk.X, **pad)
        ttk.Label(
            header,
            text=(
                "Remove o lixo após a primeira ocorrência de "
                r"{\*\bkmkstart __DdeLink__ e fecha grupos RTF pendentes."
            ),
            wraplength=760,
        ).pack(side=tk.LEFT, anchor=tk.W)
        ttk.Button(
            header,
            text="?",
            width=3,
            command=self._mostrar_manual,
        ).pack(side=tk.RIGHT, padx=(8, 0))
        ttk.Button(
            header,
            text="Manual completo",
            command=self._mostrar_manual_completo,
        ).pack(side=tk.RIGHT, padx=(8, 0))

        regras = ttk.LabelFrame(frm, text="Regras avançadas de limpeza", padding=8)
        regras.pack(fill=tk.X, **pad)
        rr = ttk.Frame(regras)
        rr.pack(fill=tk.X)
        ttk.Label(rr, text="Nível:").pack(side=tk.LEFT)
        self._cleaning_level = tk.StringVar(value=SAFE_LEVEL)
        ttk.Combobox(
            rr,
            textvariable=self._cleaning_level,
            values=[SAFE_LEVEL, INTERMEDIATE_LEVEL, AGGRESSIVE_LEVEL],
            width=14,
            state="readonly",
        ).pack(side=tk.LEFT, padx=(6, 12))
        ttk.Label(rr, text="Marcadores extras (;):").pack(side=tk.LEFT)
        self._markers_text = tk.StringVar(value="")
        ttk.Entry(rr, textvariable=self._markers_text).pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8)
        )
        ttk.Button(rr, text="Carregar JSON...", command=self._carregar_markers_json).pack(side=tk.LEFT)

        notebook = ttk.Notebook(frm, style="Dark.TNotebook")
        notebook.pack(fill=tk.BOTH, expand=True, **pad)

        aba_arquivo = ttk.Frame(notebook, padding=10)
        aba_pasta = ttk.Frame(notebook, padding=10)
        aba_banco = ttk.Frame(notebook, padding=10)
        notebook.add(aba_arquivo, text="Arquivo")
        notebook.add(aba_pasta, text="Pasta (lote)")
        notebook.add(aba_banco, text="Banco de dados")

        # Área rolável da aba Banco de dados para caber em ecrãs menores.
        banco_canvas = tk.Canvas(
            aba_banco,
            highlightthickness=0,
            borderwidth=0,
            bg="#1f2128",
        )
        banco_scroll = ttk.Scrollbar(aba_banco, orient=tk.VERTICAL, command=banco_canvas.yview)
        banco_canvas.configure(yscrollcommand=banco_scroll.set)
        banco_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        banco_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        banco_content = ttk.Frame(banco_canvas)
        banco_window_id = banco_canvas.create_window((0, 0), window=banco_content, anchor="nw")

        def _sync_banco_scroll_region(_event: object) -> None:
            banco_canvas.configure(scrollregion=banco_canvas.bbox("all"))

        def _sync_banco_content_width(event: object) -> None:
            banco_canvas.itemconfigure(banco_window_id, width=event.width)

        banco_content.bind("<Configure>", _sync_banco_scroll_region)
        banco_canvas.bind("<Configure>", _sync_banco_content_width)

        def _on_banco_mousewheel(event: object) -> None:
            delta = getattr(event, "delta", 0)
            if delta:
                banco_canvas.yview_scroll(int(-delta / 120), "units")
            else:
                num = getattr(event, "num", 0)
                if num == 4:
                    banco_canvas.yview_scroll(-1, "units")
                elif num == 5:
                    banco_canvas.yview_scroll(1, "units")

        banco_canvas.bind_all("<MouseWheel>", _on_banco_mousewheel)
        banco_canvas.bind_all("<Button-4>", _on_banco_mousewheel)
        banco_canvas.bind_all("<Button-5>", _on_banco_mousewheel)

        # —— Aba Arquivo ——
        f1 = ttk.LabelFrame(aba_arquivo, text="Um ficheiro", padding=8)
        f1.pack(fill=tk.X, **pad)
        row1 = ttk.Frame(f1)
        row1.pack(fill=tk.X)
        self._path_um = tk.StringVar()
        ttk.Entry(row1, textvariable=self._path_um, state="readonly").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8)
        )
        ttk.Button(row1, text="Escolher…", command=self._escolher_um).pack(side=tk.LEFT)
        row1b = ttk.Frame(f1)
        row1b.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(
            row1b,
            text="Limpar e guardar como…",
            command=self._processar_um,
        ).pack(side=tk.LEFT)

        self._var_backup = tk.BooleanVar(value=True)
        ttk.Checkbutton(
            f1,
            text="Se guardar por cima do original, criar cópia .bak",
            variable=self._var_backup,
        ).pack(anchor=tk.W, pady=(6, 0))

        # —— Aba Pasta (lote) ——
        f2 = ttk.LabelFrame(aba_pasta, text="Pasta (vários ficheiros)", padding=8)
        f2.pack(fill=tk.X, **pad)
        row2 = ttk.Frame(f2)
        row2.pack(fill=tk.X)
        self._path_pasta = tk.StringVar()
        ttk.Entry(row2, textvariable=self._path_pasta, state="readonly").pack(
            side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8)
        )
        ttk.Button(row2, text="Escolher pasta…", command=self._escolher_pasta).pack(
            side=tk.LEFT
        )
        row2b = ttk.Frame(f2)
        row2b.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(row2b, text="Extensões:").pack(side=tk.LEFT)
        self._ext_vars = {}
        for ext in (".rtf", ".txt"):
            v = tk.BooleanVar(value=ext == ".rtf")
            self._ext_vars[ext] = v
            ttk.Checkbutton(row2b, text=ext, variable=v).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(
            row2b,
            text="Processar pasta",
            command=self._processar_pasta,
        ).pack(side=tk.RIGHT)

        row2c = ttk.Frame(f2)
        row2c.pack(fill=tk.X, pady=(8, 0))
        self._batch_mode = tk.StringVar(value="overwrite")
        ttk.Radiobutton(
            row2c,
            text="Sobrescrever originais",
            value="overwrite",
            variable=self._batch_mode,
            command=self._atualizar_estado_destino_lote,
        ).pack(side=tk.LEFT)
        ttk.Radiobutton(
            row2c,
            text="Guardar em nova pasta",
            value="new_folder",
            variable=self._batch_mode,
            command=self._atualizar_estado_destino_lote,
        ).pack(side=tk.LEFT, padx=(12, 0))

        row2d = ttk.Frame(f2)
        row2d.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(row2d, text="Destino:").pack(side=tk.LEFT)
        self._path_destino_lote = tk.StringVar()
        self._entry_destino_lote = ttk.Entry(
            row2d, textvariable=self._path_destino_lote, state="readonly"
        )
        self._entry_destino_lote.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 8))
        self._btn_destino_lote = ttk.Button(
            row2d, text="Escolher destino…", command=self._escolher_destino_lote
        )
        self._btn_destino_lote.pack(side=tk.LEFT)
        self._atualizar_estado_destino_lote()

        # —— Aba Banco de dados (PostgreSQL) ——
        f3 = ttk.Frame(banco_content, padding=2)
        f3.pack(fill=tk.X, **pad)

        conn_box = ttk.LabelFrame(f3, text="1) Conexão", padding=8)
        conn_box.pack(fill=tk.X)
        db1 = ttk.Frame(conn_box)
        db1.pack(fill=tk.X)
        ttk.Label(db1, text="Host").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._db_host = tk.StringVar(value="127.0.0.1")
        ttk.Entry(db1, textvariable=self._db_host, width=22).grid(row=0, column=1, sticky="w", padx=(0, 14))
        ttk.Label(db1, text="Porta").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self._db_port = tk.StringVar(value="5432")
        ttk.Entry(db1, textvariable=self._db_port, width=8).grid(row=0, column=3, sticky="w", padx=(0, 14))
        ttk.Label(db1, text="Banco").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self._db_name = tk.StringVar()
        ttk.Entry(db1, textvariable=self._db_name).grid(row=0, column=5, sticky="ew")
        db1.columnconfigure(5, weight=1)

        db1b = ttk.Frame(conn_box)
        db1b.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db1b, text="Usuário").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._db_user = tk.StringVar()
        ttk.Entry(db1b, textvariable=self._db_user, width=22).grid(row=0, column=1, sticky="w", padx=(0, 14))
        ttk.Label(db1b, text="Senha").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self._db_pass = tk.StringVar()
        ttk.Entry(db1b, textvariable=self._db_pass, show="*").grid(row=0, column=3, sticky="ew")
        db1b.columnconfigure(3, weight=1)

        scope_box = ttk.LabelFrame(f3, text="2) Escopo e filtros", padding=8)
        scope_box.pack(fill=tk.X, pady=(8, 0))
        db2 = ttk.Frame(scope_box)
        db2.pack(fill=tk.X)
        ttk.Label(db2, text="Tabela").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._db_table = tk.StringVar(value="documento_mesclado")
        self._db_table_combo = ttk.Combobox(db2, textvariable=self._db_table, width=40, state="normal")
        self._db_table_combo.grid(row=0, column=1, sticky="ew", padx=(0, 12))
        self._db_table_combo.bind("<<ComboboxSelected>>", self._on_tabela_change)
        ttk.Label(db2, text="Coluna conteúdo").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self._db_content_col = tk.StringVar(value="conteudo")
        self._db_content_combo = ttk.Combobox(db2, textvariable=self._db_content_col, width=22, state="normal")
        self._db_content_combo.grid(row=0, column=3, sticky="w")
        db2.columnconfigure(1, weight=1)

        db2b = ttk.Frame(scope_box)
        db2b.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db2b, text="Min chars").grid(row=0, column=0, sticky="w", padx=(0, 6))
        self._db_min_length = tk.StringVar(value="1000000")
        ttk.Entry(db2b, textvariable=self._db_min_length, width=12).grid(row=0, column=1, sticky="w", padx=(0, 12))
        ttk.Label(db2b, text="Min MB").grid(row=0, column=2, sticky="w", padx=(0, 6))
        self._db_min_mb = tk.StringVar(value="")
        ttk.Entry(db2b, textvariable=self._db_min_mb, width=10).grid(row=0, column=3, sticky="w", padx=(0, 12))
        ttk.Label(db2b, text="Limite").grid(row=0, column=4, sticky="w", padx=(0, 6))
        self._db_limit = tk.StringVar(value="")
        ttk.Entry(db2b, textvariable=self._db_limit, width=10).grid(row=0, column=5, sticky="w", padx=(0, 12))
        ttk.Label(db2b, text="Commit por lote").grid(row=0, column=6, sticky="w", padx=(0, 6))
        self._db_batch_size = tk.StringVar(value="200")
        ttk.Entry(db2b, textvariable=self._db_batch_size, width=8).grid(row=0, column=7, sticky="w")

        db2c = ttk.Frame(scope_box)
        db2c.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db2c, text="Campos relatório").pack(side=tk.LEFT)
        self._db_report_cols = tk.StringVar(value="")
        ttk.Entry(db2c, textvariable=self._db_report_cols).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        ttk.Label(
            scope_box,
            text="Dica: use colunas locais (ex.: id_documentomesclado) ou relacionadas em formato tabela.coluna (ex.: protocolo_documentomesclado.id_protocolo).",
            foreground="#aeb4c0",
            wraplength=900,
        ).pack(anchor=tk.W, pady=(6, 0))
        self._db_sql_size_only = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            scope_box,
            text="WHERE só por tamanho (sem marcadores no SQL) — ative com Min MB/chars: evita position() lento em colunas gigantes",
            variable=self._db_sql_size_only,
        ).pack(anchor=tk.W, pady=(8, 0))

        run_box = ttk.LabelFrame(f3, text="3) Execução", padding=8)
        run_box.pack(fill=tk.X, pady=(8, 0))
        run_checks = ttk.Frame(run_box)
        run_checks.pack(fill=tk.X)
        run_checks_2 = ttk.Frame(run_box)
        run_checks_2.pack(fill=tk.X, pady=(6, 0))
        self._db_only_rtf = tk.BooleanVar(value=True)
        ttk.Checkbutton(run_checks, text="Apenas registros que parecem RTF", variable=self._db_only_rtf).pack(side=tk.LEFT)
        self._db_full_scan = tk.BooleanVar(value=False)
        ttk.Checkbutton(run_checks, text="Varredura geral", variable=self._db_full_scan).pack(side=tk.LEFT, padx=(18, 0))
        self._db_execute = tk.BooleanVar(value=False)
        ttk.Checkbutton(run_checks_2, text="Aplicar UPDATE", variable=self._db_execute).pack(side=tk.LEFT)
        self._db_strict_validation = tk.BooleanVar(value=True)
        ttk.Checkbutton(run_checks_2, text="Validação RTF estrita", variable=self._db_strict_validation).pack(side=tk.LEFT, padx=(18, 0))

        run_progress = ttk.Frame(run_box)
        run_progress.pack(fill=tk.X, pady=(8, 0))
        self._db_progress_text = tk.StringVar(value="Progresso: parado")
        ttk.Label(run_progress, textvariable=self._db_progress_text, foreground="#aeb4c0").pack(side=tk.LEFT)

        run_actions = ttk.Frame(run_box)
        run_actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(run_actions, text="Carregar tabelas/colunas", command=self._carregar_metadata_banco).pack(side=tk.LEFT)
        ttk.Button(run_actions, text="Testar conexão", command=self._testar_conexao_banco).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(run_actions, text="Parar", command=self._parar_processamento_banco).pack(side=tk.RIGHT)
        ttk.Button(run_actions, text="Higienizar banco", command=self._processar_banco).pack(side=tk.RIGHT, padx=(0, 8))

        audit_box = ttk.LabelFrame(f3, text="4) Batch e auditoria", padding=8)
        audit_box.pack(fill=tk.X, pady=(8, 0))
        db5 = ttk.Frame(audit_box)
        db5.pack(fill=tk.X)
        ttk.Label(db5, text="Batch ID").pack(side=tk.LEFT)
        self._db_batch_id = tk.StringVar()
        ttk.Entry(db5, textvariable=self._db_batch_id).pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)
        ttk.Label(
            audit_box,
            text="Use o Batch ID para consultar, exportar ou desfazer um lote.",
            foreground="#aeb4c0",
        ).pack(anchor=tk.W, pady=(6, 0))
        db5_actions = ttk.Frame(audit_box)
        db5_actions.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(db5_actions, text="Ver relatório", command=self._ver_relatorio_batch).pack(side=tk.LEFT)
        ttk.Button(db5_actions, text="Exportar CSV", command=self._exportar_relatorio_csv).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(db5_actions, text="Rollback", command=self._rollback_batch).pack(side=tk.RIGHT)

        f3b = ttk.LabelFrame(banco_content, text="URL gerada (opcional)", padding=8)
        f3b.pack(fill=tk.X, **pad)
        self._db_url_preview = tk.StringVar()
        ttk.Entry(f3b, textvariable=self._db_url_preview, state="readonly").pack(
            fill=tk.X
        )

        # —— Log ——
        ttk.Label(frm, text="Registo").pack(anchor=tk.W, **{**pad, "pady": (12, 2)})
        log_frame = ttk.Frame(frm)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=(0, 6))
        self._log = tk.Text(log_frame, height=12, wrap=tk.WORD, state=tk.DISABLED)
        self._log.configure(
            bg="#262c37",
            fg="#e8e8e8",
            insertbackground="#ffffff",
            selectbackground="#5b8cff",
            selectforeground="#ffffff",
            relief=tk.FLAT,
            borderwidth=0,
        )
        sb = ttk.Scrollbar(log_frame, command=self._log.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._log.configure(yscrollcommand=sb.set)

        ttk.Label(
            frm,
            text="Marcador: " + MARKER_DDE_BOOKMARK[:40] + "…",
            font=("TkDefaultFont", 8),
            foreground="#aeb4c0",
        ).pack(anchor=tk.W, padx=10)

    def _log_line(self, msg: str) -> None:
        self._log.configure(state=tk.NORMAL)
        self._log.insert(tk.END, msg + "\n")
        self._log.see(tk.END)
        self._log.configure(state=tk.DISABLED)

    def _som_conclusao_limpeza(self) -> None:
        """Som breve ao terminar limpeza/atualização (Windows: som padrão; resto: bell Tk)."""
        try:
            if os.name == "nt":
                import winsound

                winsound.MessageBeep(winsound.MB_OK)
            else:
                self.bell()
        except Exception:
            try:
                self.bell()
            except Exception:
                pass

    def _msg_limpeza_merece_som(self, title: str, text: str) -> bool:
        """True para conclusões de higienização em ficheiro, pasta ou banco (não erros nem cancelamentos)."""
        low = text.lower()
        if title in ("Concluído", "Lote"):
            return True
        if title == "Banco":
            if "cancelado" in low or "nenhuma mudança necessária" in low:
                return False
            return "atualizados:" in low or "simulados:" in low
        return False

    def _animate_progress_text(self) -> None:
        if not self._progress_animating:
            return
        dots = ["", ".", "..", "..."]
        suffix = dots[self._progress_anim_step % len(dots)]
        self._db_progress_text.set(f"{self._progress_base_text}{suffix}")
        self._progress_anim_step += 1
        self.after(380, self._animate_progress_text)

    def _set_progress_status(self, base_text: str, *, animate: bool) -> None:
        self._progress_base_text = base_text
        self._progress_anim_step = 0
        if animate:
            if not self._progress_animating:
                self._progress_animating = True
                self._animate_progress_text()
            else:
                # Atualiza imediatamente para novo contexto sem esperar próximo ciclo.
                self._db_progress_text.set(base_text)
        else:
            self._progress_animating = False
            self._db_progress_text.set(base_text)

    def _mostrar_manual(self) -> None:
        texto = (
            "Manual rápido — Sanitizador RTF\n\n"
            "TIPOS DE HIGIENIZAÇÃO (Nível)\n"
            "- seguro: só grupos DDE/__DdeLink__ em marcadores.\n"
            "- intermediario: seguro + remove grupos auxiliares (generator, rsidtbl, etc.).\n"
            "- agressivo: intermediario + em ficheiros > ~5 MB remove blocos de imagem/OLE\n"
            "  (\\*\\shppict, nonshppict, \\pict{, \\*\\pict, \\*\\objdata) e corta\n"
            "  massas enormes só com hex (500k+ dígitos). APAGA fotos embutidas no RTF.\n"
            "  No PostgreSQL, o conteúdo anterior fica na auditoria (rollback por Batch ID).\n\n"
            "Marcadores padrão (deteção no texto e no banco): DDE, shppict, nonshppict,\n"
            "\\pict{, objdata, \\*\\pict. Pode acrescentar mais com Marcadores extras ou JSON.\n"
            "Com Min MB/chars, use 'WHERE só por tamanho' para o PostgreSQL não fazer\n"
            "position(marcador) em cada linha (muito lento em conteúdos de 100 MB).\n\n"
            "1) Aba Arquivo\n"
            "- Escolha o Nível antes de limpar.\n"
            "- Limpar e guardar como; .bak opcional ao sobrescrever.\n"
            "- Ao terminar arquivo/pasta ou higienização no banco, toca um som breve.\n\n"
            "2) Aba Pasta (lote)\n"
            "- Mesmo Nível e marcadores que em Arquivo.\n"
            "- Sobrescrever originais ou nova pasta (conserva subpastas).\n\n"
            "3) Aba Banco de dados\n"
            "- O Nível aplica-se a cada registo higienizado.\n"
            "- Simulação antes de UPDATE; guarde o Batch ID para relatório e rollback.\n"
            "- Commit por lote, parar e export CSV.\n\n"
            "Dica:\n"
            "- RTF muito grande com imagens: use 'agressivo' se precisar de texto sem binário.\n"
            "- Documento corrompido por outra ferramenta: restaure .bak/original e tente de novo."
        )
        messagebox.showinfo("Ajuda / Manual", texto)

    def _mostrar_manual_completo(self) -> None:
        janela = tk.Toplevel(self)
        janela.title("Manual completo — Sanitizador RTF")
        janela.geometry("860x620")
        janela.minsize(760, 520)
        janela.transient(self)

        container = ttk.Frame(janela, padding=12)
        container.pack(fill=tk.BOTH, expand=True)

        txt = tk.Text(container, wrap=tk.WORD)
        txt.configure(
            bg="#262c37",
            fg="#e8e8e8",
            insertbackground="#ffffff",
            selectbackground="#5b8cff",
            selectforeground="#ffffff",
            relief=tk.FLAT,
            borderwidth=0,
        )
        sb = ttk.Scrollbar(container, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        conteudo = (
            "SANITIZADOR RTF — MANUAL COMPLETO\n\n"
            "Objetivo\n"
            "Reduzir RTF/TXT corrompidos ou inchados por DDE, metadados repetidos, imagens\n"
            "em hex (Word/LibreOffice: \\\\*\\\\shppict, \\\\nonshppict, \\\\pict{, pngblip)\n"
            "e objetos OLE. O nível 'agressivo' remove blocos binários grandes; o texto\n"
            "visível da ata ou documento costuma preservar-se.\n\n"
            "========================================\n"
            "NÍVEIS DE HIGIENIZAÇÃO (campo Nível na interface)\n"
            "========================================\n"
            "seguro\n"
            "- Remove apenas grupos de bookmark DDE associados a __DdeLink__ (regex interna).\n"
            "- Não remove imagens nem metadados de stylesheet.\n\n"
            "intermediario\n"
            "- Inclui tudo do nível seguro.\n"
            "- Remove grupos RTF frequentemente supérfluos: \\\\*\\\\generator, \\\\*\\\\userprops,\n"
            "  \\\\*\\\\xmlnstbl, \\\\*\\\\rsidtbl, \\\\*\\\\themedata, \\\\*\\\\colorschememapping.\n\n"
            "agressivo\n"
            "- Inclui tudo do nível intermediário.\n"
            "- Se o conteúdo (após as etapas acima) tiver mais de ~5 MB:\n"
            "  remove grupos completos que começam em \\\\*\\\\shppict, \\\\nonshppict,\n"
            "  \\\\*\\\\objdata, \\\\*\\\\pict e \\\\pict{ (várias passagens até não haver mais).\n"
            "  Isto ELIMINA imagens e OLE embutidos; use só quando aceitar perder anexos visuais.\n"
            "- Em qualquer tamanho: se existir uma sequência de 500000+ dígitos hex (0-9 A-F)\n"
            "  com espaços ou mudanças de linha entre dígitos, trunca ANTES dessa massa e\n"
            "  fecha chaves RTF em falta (proteção contra lixo hex/corrupção).\n"
            "- Se um marcador da lista padrão aparecer nos últimos 10% do ficheiro, pode\n"
            "  aplicar-se corte tardio com equilíbrio de grupos.\n\n"
            "Marcadores padrão (filtros no PostgreSQL, precisa_limpeza, corte tardio)\n"
            "- __DdeLink__ / bkmkstart; \\\\*\\\\shppict; \\\\nonshppict; \\\\pict{;\n"
            "  \\\\*\\\\objdata; \\\\*\\\\pict. Extras: campo 'Marcadores extras' ou JSON.\n\n"
            "PostgreSQL e risco\n"
            "- Cada UPDATE grava o texto anterior em rtf_sanitize_audit (old_content).\n"
            "- Rollback por Batch ID repõe esses valores. Teste sempre em simulação antes.\n\n"
            "Alerta sonoro\n"
            "- Após concluir: um ficheiro, um lote de pasta, ou higienização no banco\n"
            "  (simulação ou UPDATE), a aplicação emite um som breve (Windows: beep sistema).\n"
            "- Não toca em avisos como UPDATE cancelado ou 'nenhuma mudança necessária'.\n\n"
            "Barra de título (Windows)\n"
            "- A janela tenta usar o tema escuro nativo da barra (minimizar/maximizar/fechar),\n"
            "  alinhada ao resto da interface.\n\n"
            "========================================\n"
            "1) ABA ARQUIVO\n"
            "========================================\n"
            "Quando usar:\n"
            "- Para higienizar um único documento.\n\n"
            "Passo a passo:\n"
            "1. Clique em 'Escolher...'.\n"
            "2. Selecione o arquivo .rtf ou .txt.\n"
            "3. Clique em 'Limpar e guardar como...'.\n"
            "4. Escolha o destino do arquivo limpo.\n\n"
            "Backup:\n"
            "- Se salvar por cima do original e a opção estiver marcada, é criado .bak.\n\n"
            "========================================\n"
            "2) ABA PASTA (LOTE)\n"
            "========================================\n"
            "Quando usar:\n"
            "- Para processar vários arquivos de uma vez.\n\n"
            "Passo a passo:\n"
            "1. Clique em 'Escolher pasta...'.\n"
            "2. Marque extensões (.rtf / .txt).\n"
            "3. Escolha modo:\n"
            "   - Sobrescrever originais\n"
            "   - Guardar em nova pasta\n"
            "4. Clique em 'Processar pasta'.\n\n"
            "Observação:\n"
            "- No modo nova pasta, a estrutura de subpastas é preservada.\n\n"
            "========================================\n"
            "3) ABA BANCO DE DADOS\n"
            "========================================\n"
            "Quando usar:\n"
            "- Para higienizar registros diretamente no PostgreSQL.\n\n"
            "Passo a passo recomendado:\n"
            "1. Preencha Host, Porta, Banco, Usuário e Senha.\n"
            "2. Clique em 'Testar conexão'.\n"
            "3. Clique em 'Carregar tabelas/colunas'.\n"
            "4. Selecione tabela e coluna de conteúdo.\n"
            "5. Ajuste filtros (Min chars, Min MB, limite).\n"
            "6. Rode primeiro em simulação (UPDATE desmarcado).\n"
            "7. Revise a prévia e só então marque UPDATE.\n"
            "8. Guarde o Batch ID para relatório/rollback.\n\n"
            "Campos importantes:\n"
            "- Apenas registros que parecem RTF: restringe para conteúdos com cara de RTF.\n"
            "- Varredura geral: ignora filtros de tamanho e analisa toda a tabela.\n"
            "- WHERE só por tamanho: SQL sem marcadores; use com Min MB (recomendado p/ ficheiros enormes).\n"
            "- Validar RTF após limpeza (estrito): evita gravar saída inválida.\n"
            "- Commit por lote: define frequência de commit durante UPDATE.\n"
            "- Parar processamento: envia solicitação de interrupção segura.\n\n"
            "========================================\n"
            "4) RELATÓRIO E ROLLBACK\n"
            "========================================\n"
            "- Ver relatório: mostra registros alterados por um Batch ID.\n"
            "- Exportar CSV: gera arquivo para auditoria externa.\n"
            "- Rollback batch: restaura conteúdo anterior daquele lote (old_content).\n\n"
            "========================================\n"
            "5) BOAS PRÁTICAS\n"
            "========================================\n"
            "- Sempre começar em simulação.\n"
            "- Fazer backup antes de operações grandes.\n"
            "- Executar lotes grandes fora do horário de pico.\n"
            "- Guardar os Batch IDs das execuções.\n\n"
            "========================================\n"
            "6) RESOLUÇÃO DE PROBLEMAS\n"
            "========================================\n"
            "Arquivo ainda não abre após higienização:\n"
            "- Tente novamente a partir do original ou .bak.\n"
            "- Confirme se o documento de origem realmente contém RTF válido.\n\n"
            "Banco sem alterações:\n"
            "- Reduza filtros de tamanho.\n"
            "- Desmarque temporariamente 'Apenas registros que parecem RTF'.\n"
            "- Verifique se a coluna escolhida é a coluna correta de conteúdo.\n"
        )

        txt.insert("1.0", conteudo)
        txt.configure(state=tk.DISABLED)

        footer = ttk.Frame(janela, padding=(12, 0, 12, 12))
        footer.pack(fill=tk.X)
        ttk.Button(footer, text="Fechar", command=janela.destroy).pack(side=tk.RIGHT)

    def _poll_queue(self) -> None:
        try:
            while True:
                kind, data = self._queue.get_nowait()
                if kind == "log":
                    self._log_line(data)
                elif kind == "msg":
                    title, text, is_error = data
                    if is_error:
                        messagebox.showerror(title, text)
                    else:
                        messagebox.showinfo(title, text)
                        if self._msg_limpeza_merece_som(title, text):
                            self._som_conclusao_limpeza()
                elif kind == "confirm":
                    title, text, result_box, evt = data
                    result_box["ok"] = messagebox.askyesno(title, text)
                    evt.set()
        except queue.Empty:
            pass
        self.after(200, self._poll_queue)

    def _escolher_um(self) -> None:
        p = filedialog.askopenfilename(
            title="Ficheiro RTF ou texto",
            filetypes=[
                ("RTF e texto", "*.rtf *.txt"),
                ("RTF", "*.rtf"),
                ("Texto", "*.txt"),
                ("Todos", "*.*"),
            ],
        )
        if p:
            self._path_um.set(p)

    def _escolher_pasta(self) -> None:
        p = filedialog.askdirectory(title="Pasta com ficheiros")
        if p:
            self._path_pasta.set(p)

    def _escolher_destino_lote(self) -> None:
        p = filedialog.askdirectory(title="Pasta de destino dos ficheiros limpos")
        if p:
            self._path_destino_lote.set(p)

    def _markers_ativos(self) -> list[str]:
        extras = [x.strip() for x in self._markers_text.get().split(";") if x.strip()]
        out: list[str] = []
        for m in [*DEFAULT_MARKERS, *extras]:
            if m not in out:
                out.append(m)
        return out

    def _carregar_markers_json(self) -> None:
        p = filedialog.askopenfilename(
            title="Selecionar JSON de marcadores",
            filetypes=[("JSON", "*.json"), ("Todos", "*.*")],
        )
        if not p:
            return
        try:
            markers = carregar_marcadores_de_json(p)
            extras = [m for m in markers if m not in DEFAULT_MARKERS]
            self._markers_text.set(";".join(extras))
            messagebox.showinfo(
                "Marcadores",
                f"Marcadores carregados: {len(markers)}\n"
                f"Marcador padrão sempre ativo: {DEFAULT_MARKERS[0]}",
            )
        except Exception as e:  # noqa: BLE001
            messagebox.showerror("Erro ao carregar marcadores", str(e))

    def _atualizar_estado_destino_lote(self) -> None:
        enabled = self._batch_mode.get() == "new_folder"
        state = tk.NORMAL if enabled else tk.DISABLED
        self._entry_destino_lote.configure(state="readonly" if enabled else "disabled")
        self._btn_destino_lote.configure(state=state)

    def _processar_um(self) -> None:
        orig = self._path_um.get().strip()
        if not orig:
            messagebox.showwarning("Falta ficheiro", "Escolha um ficheiro primeiro.")
            return

        dest = filedialog.asksaveasfilename(
            title="Guardar ficheiro limpo",
            defaultextension=".rtf",
            filetypes=[
                ("RTF", "*.rtf"),
                ("Texto", "*.txt"),
                ("Todos", "*.*"),
            ],
            initialfile=Path(orig).stem + "_limpo" + Path(orig).suffix,
        )
        if not dest:
            return

        def job() -> None:
            try:
                po, pd = Path(orig), Path(dest)
                if po.resolve() == pd.resolve() and self._var_backup.get():
                    bak = po.with_suffix(po.suffix + ".bak")
                    bak.write_bytes(po.read_bytes())
                    self._queue.put(("log", f"Backup: {bak}"))

                texto, modo = _ler_texto_preservando_bytes(po)
                n_antes = len(texto)
                analise = analisar_limpeza(
                    texto,
                    markers=self._markers_ativos(),
                    cleaning_level=self._cleaning_level.get(),
                )
                limpo = limpar_arquivo_rtf(
                    texto,
                    markers=self._markers_ativos(),
                    cleaning_level=self._cleaning_level.get(),
                )
                n_depois = len(limpo)
                _guardar_texto_preservando_bytes(pd, limpo)

                rem_ini = (analise["removed_preview_start"] or "").replace("\n", " ")[:120]
                rem_fim = (analise["removed_preview_end"] or "").replace("\n", " ")[:120]
                self._queue.put(
                    (
                        "log",
                        f"{po.name}: {n_antes:,} → {n_depois:,} chars | "
                        f"RTF provável: {parece_rtf(texto)} | marcador: {analise['marker_found']} | leitura: {modo}",
                    )
                )
                if analise["marker_found"]:
                    self._queue.put(
                        (
                            "log",
                            f"Prévia removida (início): {rem_ini if rem_ini else '(vazio)'}",
                        )
                    )
                    self._queue.put(
                        (
                            "log",
                            f"Prévia removida (fim): {rem_fim if rem_fim else '(vazio)'}",
                        )
                    )
                self._queue.put(("msg", ("Concluído", f"Guardado:\n{pd}", False)))
            except OSError as e:
                self._queue.put(("msg", ("Erro", str(e), True)))
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro", str(e), True)))

        threading.Thread(target=job, daemon=True).start()

    def _extensoes_escolhidas(self) -> list[str]:
        return [e for e, v in self._ext_vars.items() if v.get()]

    def _build_database_url(self) -> str:
        host = self._db_host.get().strip()
        port = self._db_port.get().strip()
        dbname = self._db_name.get().strip()
        user = self._db_user.get().strip()
        password = self._db_pass.get()

        if not host or not port or not dbname or not user:
            raise ValueError("Preencha Host, Porta, Banco e Usuário.")
        if not port.isdigit():
            raise ValueError("Porta deve ser numérica.")

        user_q = quote_plus(user)
        pass_q = quote_plus(password)
        url = f"postgresql://{user_q}:{pass_q}@{host}:{port}/{dbname}"
        self._db_url_preview.set(url)
        return url

    def _testar_conexao_banco(self) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError as e:
            messagebox.showwarning("Conexão", str(e))
            return

        def job() -> None:
            try:
                resumo = test_postgres_connection(db_url)
                self._queue.put(("msg", ("Conexão", resumo, False)))
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro de conexão", str(e), True)))

        threading.Thread(target=job, daemon=True).start()

    def _carregar_metadata_banco(self) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError as e:
            messagebox.showwarning("Conexão", str(e))
            return

        tabela_atual = self._db_table.get().strip()

        def job() -> None:
            try:
                tabelas = list_postgres_tables(db_url)
                self.after(0, lambda: self._db_table_combo.configure(values=tabelas))

                target_table = tabela_atual or (tabelas[0] if tabelas else "")
                if target_table:
                    self.after(0, lambda: self._db_table.set(target_table))
                    self._carregar_colunas_da_tabela(db_url, target_table)

                self._queue.put(
                    (
                        "msg",
                        (
                            "Metadados carregados",
                            f"Tabelas encontradas: {len(tabelas)}",
                            False,
                        ),
                    )
                )
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro ao carregar metadados", str(e), True)))

        threading.Thread(target=job, daemon=True).start()

    def _carregar_colunas_da_tabela(self, db_url: str, table_name: str) -> None:
        def job_cols() -> None:
            try:
                cols = list_postgres_columns(db_url, table_name)
                self.after(0, lambda: self._db_content_combo.configure(values=cols))

                current_content = self._db_content_col.get().strip()

                if not current_content or current_content not in cols:
                    selected_content = None
                    for c in ("conteudo", "content", "texto", "text"):
                        if c in cols:
                            selected_content = c
                            break
                    if selected_content is None and cols:
                        selected_content = cols[0]
                    if selected_content:
                        self.after(0, lambda col=selected_content: self._db_content_col.set(col))
            except Exception:
                # Não interrompe fluxo principal por falha de sugestão de colunas.
                pass

        threading.Thread(target=job_cols, daemon=True).start()

    def _on_tabela_change(self, _event: object) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError:
            return
        table_name = self._db_table.get().strip()
        if table_name:
            self._carregar_colunas_da_tabela(db_url, table_name)

    def _processar_banco(self) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError as e:
            messagebox.showwarning("Conexão", str(e))
            return
        try:
            min_length = int(self._db_min_length.get().strip())
            if min_length < 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Min chars", "Informe um número inteiro >= 0.")
            return
        min_mb_raw = self._db_min_mb.get().strip()
        min_mb = None
        if min_mb_raw:
            try:
                min_mb = float(min_mb_raw.replace(",", "."))
                if min_mb < 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Min MB", "Informe vazio ou um número >= 0.")
                return

        limit_raw = self._db_limit.get().strip()
        limit = None
        if limit_raw:
            try:
                limit = int(limit_raw)
                if limit <= 0:
                    raise ValueError
            except ValueError:
                messagebox.showwarning("Limite", "Informe vazio ou um número inteiro > 0.")
                return

        table_name = self._db_table.get().strip() or "documento_mesclado"
        content_col = self._db_content_col.get().strip() or "conteudo"
        report_cols = [c.strip() for c in self._db_report_cols.get().split(",") if c.strip()]
        markers = self._markers_ativos()
        cleaning_level = self._cleaning_level.get()
        execute = self._db_execute.get()
        only_rtf = self._db_only_rtf.get()
        full_scan = self._db_full_scan.get()
        strict_rtf_validation = self._db_strict_validation.get()
        try:
            batch_size = int(self._db_batch_size.get().strip())
            if batch_size <= 0:
                raise ValueError
        except ValueError:
            messagebox.showwarning("Commit por lote", "Informe um número inteiro > 0.")
            return
        sql_size_only = self._db_sql_size_only.get() and not full_scan
        if sql_size_only and min_length <= 0 and min_mb is None:
            messagebox.showwarning(
                "Filtro SQL",
                "Marque 'WHERE só por tamanho' apenas com Min chars > 0 ou Min MB preenchido,\n"
                "ou desmarque essa opção para usar marcadores no SQL.",
            )
            return
        self._stop_requested.clear()
        if execute:
            self._set_progress_status("Progresso: prévia — analisando candidatos", animate=True)
        else:
            self._set_progress_status("Progresso: simulando higienização", animate=True)

        def job() -> None:
            try:
                if execute:
                    self._queue.put(
                        (
                            "log",
                            "Prévia (1.ª passagem): a percorrer candidatos no PostgreSQL. "
                            "Com Varredura geral isto lê a tabela inteira; em cada linha simula a limpeza "
                            "(pode demorar muitos minutos se houver milhares de registos ou conteúdos enormes). "
                            "O contador na barra de progresso atualiza ao processar cada registo.",
                        )
                    )
                else:
                    self._queue.put(("log", "Simulação: a analisar registos candidatos no banco..."))
                if execute:
                    preview_lines: list[str] = []
                    total_preview, skipped_preview, _ = sanitize_documento_mesclado(
                        db_url,
                        execute=False,
                        min_length=min_length,
                        min_megabytes=min_mb,
                        full_scan=full_scan,
                        only_rtf=only_rtf,
                        limit=limit,
                        table_name=table_name,
                        content_column=content_col,
                        id_column=None,
                        report_columns=report_cols,
                        markers=markers,
                        batch_size=batch_size,
                        strict_rtf_validation=strict_rtf_validation,
                        cleaning_level=cleaning_level,
                        sql_where_size_only=sql_size_only,
                        should_stop=lambda: self._stop_requested.is_set(),
                        progress=lambda s, u, k: self.after(
                            0,
                            lambda: self._set_progress_status(
                                f"Progresso: prévia lidos={s} a alterar={u} ignorados={k}",
                                animate=True,
                            ),
                        ),
                        log=lambda m: preview_lines.append(m) if len(preview_lines) < 20 else None,
                    )
                    if total_preview == 0:
                        self._queue.put(
                            (
                                "msg",
                                (
                                    "Banco",
                                    f"Nenhuma mudança necessária.\nIgnorados sem mudança: {skipped_preview}",
                                    False,
                                ),
                            )
                        )
                        return

                    preview_txt = "\n".join(preview_lines) if preview_lines else "(sem detalhes)"
                    if total_preview > len(preview_lines):
                        preview_txt += f"\n... e mais {total_preview - len(preview_lines)} item(ns)."

                    result_box: dict[str, bool] = {"ok": False}
                    evt = threading.Event()
                    self._queue.put(
                        (
                            "confirm",
                            (
                                "Confirmar UPDATE",
                                f"Prévia concluída.\nRegistros que serão atualizados: {total_preview}\n"
                                f"Ignorados sem mudança: {skipped_preview}\n\n"
                                f"Exemplos:\n{preview_txt}\n\n"
                                "Deseja executar o UPDATE agora?",
                                result_box,
                                evt,
                            ),
                        )
                    )
                    evt.wait()
                    if not result_box.get("ok", False):
                        self._queue.put(("msg", ("Banco", "UPDATE cancelado pelo usuário.", False)))
                        return

                    self._queue.put(("log", "Prévia confirmada. Iniciando higienização (UPDATE)..."))
                    self.after(0, lambda: self._set_progress_status("Progresso: higienizando registros", animate=True))
                    updated, skipped, batch_id = sanitize_documento_mesclado(
                        db_url,
                        execute=True,
                        min_length=min_length,
                        min_megabytes=min_mb,
                        full_scan=full_scan,
                        only_rtf=only_rtf,
                        limit=limit,
                        table_name=table_name,
                        content_column=content_col,
                        id_column=None,
                        report_columns=report_cols,
                        markers=markers,
                        batch_size=batch_size,
                        strict_rtf_validation=strict_rtf_validation,
                        cleaning_level=cleaning_level,
                        sql_where_size_only=sql_size_only,
                        should_stop=lambda: self._stop_requested.is_set(),
                        progress=lambda s, u, k: self.after(
                            0, lambda: self._set_progress_status(
                                f"Progresso: lidos={s} atualizados={u} ignorados={k}",
                                animate=True,
                            )
                        ),
                        log=lambda m: self._queue.put(("log", m)),
                    )
                    self._last_batch_id = batch_id or ""
                    if batch_id:
                        self._queue.put(("log", f"batch_id={batch_id}"))
                        self.after(0, lambda bid=batch_id: self._db_batch_id.set(bid))
                    self._queue.put(
                        (
                            "msg",
                            (
                                "Banco",
                                f"Atualizados: {updated}\nIgnorados sem mudança: {skipped}\nBatch ID: {batch_id}",
                                False,
                            ),
                        )
                    )
                else:
                    self._queue.put(("log", "Iniciando simulação de higienização..."))
                    self.after(0, lambda: self._set_progress_status("Progresso: simulando higienização", animate=True))
                    updated, skipped, _ = sanitize_documento_mesclado(
                        db_url,
                        execute=False,
                        min_length=min_length,
                        min_megabytes=min_mb,
                        full_scan=full_scan,
                        only_rtf=only_rtf,
                        limit=limit,
                        table_name=table_name,
                        content_column=content_col,
                        id_column=None,
                        report_columns=report_cols,
                        markers=markers,
                        batch_size=batch_size,
                        strict_rtf_validation=strict_rtf_validation,
                        cleaning_level=cleaning_level,
                        sql_where_size_only=sql_size_only,
                        should_stop=lambda: self._stop_requested.is_set(),
                        progress=lambda s, u, k: self.after(
                            0, lambda: self._set_progress_status(
                                f"Progresso: lidos={s} simulados={u} ignorados={k}",
                                animate=True,
                            )
                        ),
                        log=lambda m: self._queue.put(("log", m)),
                    )
                    self._queue.put(
                        (
                            "msg",
                            (
                                "Banco",
                                f"Simulados: {updated}\nIgnorados sem mudança: {skipped}",
                                False,
                            ),
                        )
                    )
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro", str(e), True)))
            finally:
                self.after(0, lambda: self._set_progress_status("Progresso: parado", animate=False))
        threading.Thread(target=job, daemon=True).start()

    def _ver_relatorio_batch(self) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError as e:
            messagebox.showwarning("Conexão", str(e))
            return
        batch_id = self._db_batch_id.get().strip() or self._last_batch_id
        if not batch_id:
            messagebox.showwarning("Batch ID", "Informe um Batch ID para ver o relatório.")
            return

        def job() -> None:
            try:
                rows = get_batch_report(db_url, batch_id, limit=500)
                if not rows:
                    self._queue.put(("msg", ("Relatório", "Nenhum registro encontrado para esse batch.", False)))
                    return
                self._queue.put(("log", f"--- Relatório batch {batch_id} ({len(rows)} itens) ---"))
                for r in rows[:100]:
                    extras = ", ".join(f"{k}={v}" for k, v in (r.get("report_data") or {}).items())
                    self._queue.put(
                        (
                            "log",
                            f"{r['table_name']} {r['key_column']}={r['key_value']} "
                            f"{r['old_len']}->{r['new_len']} {extras}".strip(),
                        )
                    )
                if len(rows) > 100:
                    self._queue.put(("log", f"... e mais {len(rows) - 100} linhas"))
                self._queue.put(("msg", ("Relatório", f"Relatório carregado para batch {batch_id}.", False)))
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro relatório", str(e), True)))

        threading.Thread(target=job, daemon=True).start()

    def _exportar_relatorio_csv(self) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError as e:
            messagebox.showwarning("Conexão", str(e))
            return
        batch_id = self._db_batch_id.get().strip() or self._last_batch_id
        if not batch_id:
            messagebox.showwarning("Batch ID", "Informe um Batch ID para exportar.")
            return
        pasta = filedialog.askdirectory(title="Pasta para exportar CSV")
        if not pasta:
            return

        def job() -> None:
            try:
                out_path = export_batch_report_csv(db_url, batch_id, output_dir=pasta)
                self._queue.put(("msg", ("Exportação", f"CSV exportado:\n{out_path}", False)))
                self._queue.put(("log", f"Relatório CSV exportado: {out_path}"))
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro exportação", str(e), True)))

        threading.Thread(target=job, daemon=True).start()

    def _parar_processamento_banco(self) -> None:
        self._stop_requested.set()
        self._set_progress_status("Progresso: solicitação de parada enviada", animate=True)
        self._queue.put(("log", "Solicitação de parada enviada. Aguarde o término da etapa atual."))

    def _rollback_batch(self) -> None:
        try:
            db_url = self._build_database_url()
        except ValueError as e:
            messagebox.showwarning("Conexão", str(e))
            return
        batch_id = self._db_batch_id.get().strip() or self._last_batch_id
        if not batch_id:
            messagebox.showwarning("Batch ID", "Informe um Batch ID para rollback.")
            return
        if not messagebox.askyesno(
            "Confirmar rollback",
            f"Tem certeza que deseja desfazer o batch {batch_id}?",
        ):
            return

        def job() -> None:
            try:
                total = rollback_batch(db_url, batch_id)
                self._queue.put(("msg", ("Rollback", f"Rollback concluído. Linhas restauradas: {total}", False)))
                self._queue.put(("log", f"Rollback batch {batch_id}: {total} linhas restauradas"))
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro rollback", str(e), True)))

        threading.Thread(target=job, daemon=True).start()

    def _processar_pasta(self) -> None:
        raiz = self._path_pasta.get().strip()
        if not raiz:
            messagebox.showwarning("Falta pasta", "Escolha uma pasta primeiro.")
            return
        modo_lote = self._batch_mode.get()
        destino_lote = self._path_destino_lote.get().strip()
        if modo_lote == "new_folder" and not destino_lote:
            messagebox.showwarning(
                "Destino em falta",
                "Escolha a pasta de destino para guardar os ficheiros limpos.",
            )
            return
        exts = self._extensoes_escolhidas()
        if not exts:
            messagebox.showwarning("Extensões", "Selecione pelo menos uma extensão.")
            return

        def job() -> None:
            try:
                base = Path(raiz)
                destino_base = Path(destino_lote) if modo_lote == "new_folder" else None
                ficheiros = sorted(
                    p
                    for p in base.rglob("*")
                    if p.is_file() and p.suffix.lower() in exts
                )
                if not ficheiros:
                    self._queue.put(("msg", ("Pasta", "Nenhum ficheiro encontrado.", False)))
                    return

                alterados = 0
                markers = self._markers_ativos()
                cleaning_level = self._cleaning_level.get()
                for po in ficheiros:
                    texto, modo = _ler_texto_preservando_bytes(po)
                    limpo = limpar_arquivo_rtf(texto, markers=markers, cleaning_level=cleaning_level)
                    if modo_lote == "new_folder":
                        rel = po.relative_to(base)
                        pd = destino_base / rel
                        pd.parent.mkdir(parents=True, exist_ok=True)
                        _guardar_texto_preservando_bytes(pd, limpo)
                    else:
                        pd = po
                        if limpo == texto:
                            self._queue.put(("log", f"{po.name}: sem alterações"))
                            continue
                        if self._var_backup.get():
                            bak = po.with_suffix(po.suffix + ".bak")
                            if not bak.exists():
                                bak.write_bytes(po.read_bytes())
                        _guardar_texto_preservando_bytes(po, limpo)
                    alterados += 1
                    self._queue.put(
                        (
                            "log",
                            f"{po.name}: {len(texto):,} → {len(limpo):,} chars ({modo})"
                            f"{' -> ' + str(pd) if modo_lote == 'new_folder' else ''}",
                        )
                    )

                self._queue.put(
                    (
                        "msg",
                        (
                            "Lote",
                            f"Processados {len(ficheiros)} ficheiros.\n"
                            f"{'Gerados' if modo_lote == 'new_folder' else 'Alterados'}: {alterados}.",
                            False,
                        ),
                    )
                )
            except OSError as e:
                self._queue.put(("msg", ("Erro", str(e), True)))
            except Exception as e:  # noqa: BLE001
                self._queue.put(("msg", ("Erro", str(e), True)))

        threading.Thread(target=job, daemon=True).start()


def main() -> None:
    app = App()
    app.mainloop()


if __name__ == "__main__":
    main()
