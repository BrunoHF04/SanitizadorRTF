#!/usr/bin/env python3
# -*- coding: utf-8 -*-
r"""
Interface gráfica para sanitizar RTF (remoção de lixo DDE após {\*\bkmkstart __DdeLink__).
"""

from __future__ import annotations

import queue
import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from urllib.parse import quote_plus

from db_sanitize import (
    get_batch_report,
    list_postgres_columns,
    list_postgres_tables,
    rollback_batch,
    sanitize_documento_mesclado,
    test_postgres_connection,
)
from rtf_sanitize import MARKER_DDE_BOOKMARK, limpar_arquivo_rtf, parece_rtf


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
        self.title("Sanitizador RTF — DDE / bookmark")
        self.minsize(860, 560)
        self.geometry("980x650")

        self._queue: queue.Queue[tuple[str, str]] = queue.Queue()
        self._last_batch_id = ""
        self._build()
        self.after(200, self._poll_queue)

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
                r"{\*\bkmkstart __DdeLink__ e garante que o RTF termina com }."
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

        notebook = ttk.Notebook(frm)
        notebook.pack(fill=tk.BOTH, expand=True, **pad)

        aba_arquivo = ttk.Frame(notebook, padding=10)
        aba_pasta = ttk.Frame(notebook, padding=10)
        aba_banco = ttk.Frame(notebook, padding=10)
        notebook.add(aba_arquivo, text="Arquivo")
        notebook.add(aba_pasta, text="Pasta (lote)")
        notebook.add(aba_banco, text="Banco de dados")

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
        f3 = ttk.LabelFrame(aba_banco, text="Conexão PostgreSQL", padding=8)
        f3.pack(fill=tk.X, **pad)
        db1 = ttk.Frame(f3)
        db1.pack(fill=tk.X)
        ttk.Label(db1, text="Host:").pack(side=tk.LEFT)
        self._db_host = tk.StringVar(value="127.0.0.1")
        ttk.Entry(db1, textvariable=self._db_host, width=20).pack(side=tk.LEFT, padx=(8, 12))
        ttk.Label(db1, text="Porta:").pack(side=tk.LEFT)
        self._db_port = tk.StringVar(value="5432")
        ttk.Entry(db1, textvariable=self._db_port, width=8).pack(side=tk.LEFT, padx=(8, 12))
        ttk.Label(db1, text="Banco:").pack(side=tk.LEFT)
        self._db_name = tk.StringVar()
        ttk.Entry(db1, textvariable=self._db_name, width=20).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(8, 0))

        db1b = ttk.Frame(f3)
        db1b.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db1b, text="Usuário:").pack(side=tk.LEFT)
        self._db_user = tk.StringVar()
        ttk.Entry(db1b, textvariable=self._db_user, width=20).pack(side=tk.LEFT, padx=(8, 12))
        ttk.Label(db1b, text="Senha:").pack(side=tk.LEFT)
        self._db_pass = tk.StringVar()
        ttk.Entry(db1b, textvariable=self._db_pass, show="*", width=20).pack(
            side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True
        )

        db2 = ttk.Frame(f3)
        db2.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db2, text="Min chars:").pack(side=tk.LEFT)
        self._db_min_length = tk.StringVar(value="1000000")
        ttk.Entry(db2, textvariable=self._db_min_length, width=12).pack(side=tk.LEFT, padx=(8, 16))
        ttk.Label(db2, text="Min MB:").pack(side=tk.LEFT)
        self._db_min_mb = tk.StringVar(value="")
        ttk.Entry(db2, textvariable=self._db_min_mb, width=10).pack(side=tk.LEFT, padx=(8, 16))
        ttk.Label(db2, text="Limite:").pack(side=tk.LEFT)
        self._db_limit = tk.StringVar(value="")
        ttk.Entry(db2, textvariable=self._db_limit, width=10).pack(side=tk.LEFT, padx=(8, 16))
        ttk.Label(db2, text="Campos relatório:").pack(side=tk.LEFT)
        self._db_report_cols = tk.StringVar(value="")
        ttk.Entry(db2, textvariable=self._db_report_cols, width=36).pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)
        ttk.Label(
            f3,
            text="Dica: use colunas locais (ex.: id_documentomesclado) ou relacionadas em formato tabela.coluna (ex.: protocolo_documentomesclado.id_protocolo).",
            foreground="#555",
            wraplength=900,
        ).pack(anchor=tk.W, pady=(6, 0))

        db2b = ttk.Frame(f3)
        db2b.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db2b, text="Tabela:").pack(side=tk.LEFT)
        self._db_table = tk.StringVar(value="documento_mesclado")
        self._db_table_combo = ttk.Combobox(
            db2b, textvariable=self._db_table, width=42, state="normal"
        )
        self._db_table_combo.pack(side=tk.LEFT, padx=(8, 0), fill=tk.X, expand=True)
        self._db_table_combo.bind("<<ComboboxSelected>>", self._on_tabela_change)

        db2c = ttk.Frame(f3)
        db2c.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db2c, text="Coluna conteúdo:").pack(side=tk.LEFT)
        self._db_content_col = tk.StringVar(value="conteudo")
        self._db_content_combo = ttk.Combobox(
            db2c, textvariable=self._db_content_col, width=24, state="normal"
        )
        self._db_content_combo.pack(side=tk.LEFT, padx=(8, 16))

        db3 = ttk.Frame(f3)
        db3.pack(fill=tk.X, pady=(8, 0))
        self._db_only_rtf = tk.BooleanVar(value=True)
        ttk.Checkbutton(db3, text="Apenas registos que parecem RTF", variable=self._db_only_rtf).pack(
            side=tk.LEFT
        )
        self._db_full_scan = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            db3,
            text="Varredura geral (toda a tabela)",
            variable=self._db_full_scan,
        ).pack(side=tk.LEFT, padx=(12, 0))
        self._db_execute = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            db3,
            text="Aplicar UPDATE (desmarcado = simulação)",
            variable=self._db_execute,
        ).pack(side=tk.LEFT, padx=(12, 0))
        db4 = ttk.Frame(f3)
        db4.pack(fill=tk.X, pady=(8, 0))
        ttk.Button(
            db4,
            text="Higienizar banco",
            command=self._processar_banco,
        ).pack(side=tk.RIGHT)
        ttk.Button(
            db4,
            text="Testar conexão",
            command=self._testar_conexao_banco,
        ).pack(side=tk.RIGHT, padx=(0, 8))
        ttk.Button(
            db4,
            text="Carregar tabelas/colunas",
            command=self._carregar_metadata_banco,
        ).pack(side=tk.RIGHT, padx=(0, 8))

        db5 = ttk.Frame(f3)
        db5.pack(fill=tk.X, pady=(8, 0))
        ttk.Label(db5, text="Batch ID:").pack(side=tk.LEFT)
        self._db_batch_id = tk.StringVar()
        ttk.Entry(db5, textvariable=self._db_batch_id).pack(side=tk.LEFT, padx=(8, 16), fill=tk.X, expand=True)
        ttk.Button(db5, text="Ver relatório", command=self._ver_relatorio_batch).pack(side=tk.RIGHT)
        ttk.Button(db5, text="Rollback batch", command=self._rollback_batch).pack(side=tk.RIGHT, padx=(0, 8))

        f3b = ttk.LabelFrame(aba_banco, text="URL gerada (opcional)", padding=8)
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
        sb = ttk.Scrollbar(log_frame, command=self._log.yview)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        self._log.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self._log.configure(yscrollcommand=sb.set)

        ttk.Label(
            frm,
            text="Marcador: " + MARKER_DDE_BOOKMARK[:40] + "…",
            font=("TkDefaultFont", 8),
            foreground="#555",
        ).pack(anchor=tk.W, padx=10)

    def _log_line(self, msg: str) -> None:
        self._log.configure(state=tk.NORMAL)
        self._log.insert(tk.END, msg + "\n")
        self._log.see(tk.END)
        self._log.configure(state=tk.DISABLED)

    def _mostrar_manual(self) -> None:
        texto = (
            "Manual rápido — Sanitizador RTF\n\n"
            "1) Aba Arquivo\n"
            "- Clique em Escolher para selecionar um .rtf/.txt.\n"
            "- Clique em Limpar e guardar como.\n"
            "- Se salvar por cima do original, a opção .bak cria backup.\n\n"
            "2) Aba Pasta (lote)\n"
            "- Selecione a pasta e as extensões.\n"
            "- Escolha sobrescrever ou gerar em nova pasta.\n"
            "- Em nova pasta, a estrutura de subpastas é mantida.\n\n"
            "3) Aba Banco de dados\n"
            "- Preencha conexão e teste.\n"
            "- Carregue tabelas/colunas, ajuste filtros e rode primeiro em simulação.\n"
            "- Marque UPDATE somente após revisar a prévia.\n"
            "- Guarde o Batch ID para relatório e rollback.\n\n"
            "Dica:\n"
            "- Se um documento já foi salvo corrompido por outra ferramenta,\n"
            "  restaure do .bak/original e execute novamente a higienização."
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
        sb = ttk.Scrollbar(container, command=txt.yview)
        txt.configure(yscrollcommand=sb.set)
        sb.pack(side=tk.RIGHT, fill=tk.Y)
        txt.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        conteudo = (
            "SANITIZADOR RTF — MANUAL COMPLETO\n\n"
            "Objetivo\n"
            "A aplicação remove blocos de lixo que começam em:\n"
            r"{\*\bkmkstart __DdeLink__" "\n"
            "Após encontrar a primeira ocorrência, o sistema mantém apenas a parte válida\n"
            "anterior e fecha os grupos RTF pendentes.\n\n"
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
            "- Varredura geral: ignora filtros de tamanho e analisa toda a tabela.\n\n"
            "========================================\n"
            "4) RELATÓRIO E ROLLBACK\n"
            "========================================\n"
            "- Ver relatório: mostra registros alterados por um Batch ID.\n"
            "- Rollback batch: restaura conteúdo anterior daquele lote.\n\n"
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
                tem_marcador = MARKER_DDE_BOOKMARK in texto
                limpo = limpar_arquivo_rtf(texto)
                n_depois = len(limpo)
                _guardar_texto_preservando_bytes(pd, limpo)

                self._queue.put(
                    (
                        "log",
                        f"{po.name}: {n_antes:,} → {n_depois:,} chars | "
                        f"RTF provável: {parece_rtf(texto)} | marcador: {tem_marcador} | leitura: {modo}",
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
        execute = self._db_execute.get()
        only_rtf = self._db_only_rtf.get()
        full_scan = self._db_full_scan.get()

        def job() -> None:
            try:
                self._queue.put(("log", "Iniciando higienização no banco..."))
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
                for po in ficheiros:
                    texto, modo = _ler_texto_preservando_bytes(po)
                    limpo = limpar_arquivo_rtf(texto)
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
