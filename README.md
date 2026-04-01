# Sanitizador RTF (DDE/Bookmark) para TXT/RTF e PostgreSQL

Ferramenta para higienizar conteúdos RTF/TXT com corrupção por metadados de DDE/bookmark (ex.: `{\*\bkmkstart __DdeLink__...}`), reduzindo tamanho de registros e evitando erros de memória em editores e consultas.

![Python](https://img.shields.io/badge/Python-3.13%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Plataforma-Windows-0078D6?logo=windows&logoColor=white)
![Status](https://img.shields.io/badge/Status-Em%20uso-2EA043)
![Database](https://img.shields.io/badge/PostgreSQL-16%2B-4169E1?logo=postgresql&logoColor=white)
![License](https://img.shields.io/badge/Licen%C3%A7a-definir-lightgrey)

## Problema que a ferramenta resolve

Em alguns cenários, o conteúdo RTF passa a conter blocos repetitivos e invisíveis no final do texto, geralmente iniciando em:

`{\*\bkmkstart __DdeLink__`

Isso causa:

- arquivos com dezenas de MB sem aumento real do texto visível;
- travamentos em editores;
- erros de memória em ferramentas de banco.

A estratégia adotada é cirúrgica:

1. localizar a **primeira ocorrência** do marcador;
2. manter somente o conteúdo anterior;
3. fechar os grupos RTF pendentes de forma estrutural.

---

## Tecnologias usadas

- **Python 3.13+**
- **Tkinter** (GUI desktop)
- **PostgreSQL** (via `psycopg2-binary`)
- **PyInstaller** (geração de `.exe`)

---

## Estrutura do projeto

- `rtf_sanitize.py`  
  Núcleo de higienização de string RTF/TXT.

- `db_sanitize.py`  
  Operações de banco: varredura, update, auditoria por lote, relatório e rollback.

- `rtf_sanitize_gui.py`  
  Interface gráfica com abas: arquivo, pasta e banco de dados.

- `batch_sanitize_rtf.py`  
  Script de manutenção em lote (modo CLI).

- `build_exe.ps1`  
  Script PowerShell para compilar o executável.

- `requirements.txt`  
  Dependências Python.

---

## Funcionalidades principais

### 1) Higienização de arquivo único

- abre `.rtf`/`.txt`;
- limpa o conteúdo;
- salva em novo arquivo;
- opção de backup `.bak` ao sobrescrever;
- mostra prévia do trecho removido no log (início/fim).

### 2) Higienização de pasta (lote)

- processa subpastas recursivamente;
- filtra por extensão (`.rtf`, `.txt`);
- dois modos:
  - sobrescrever originais;
  - gerar em nova pasta.

### 3) Higienização no PostgreSQL

- conexão por campos (`host`, `porta`, `banco`, `usuário`, `senha`);
- teste de conexão;
- carregamento automático de tabelas/colunas (`information_schema`);
- filtros de seleção:
  - `Min chars`
  - `Min MB`
  - `Varredura geral (toda a tabela)` (ignora filtros e analisa todos os registros);
- opção **Apenas registros que parecem RTF**;
- opção **Validar RTF após limpeza (estrito)** para evitar gravação inválida;
- opção **Commit por lote** (transações parciais);
- botão **Parar processamento** (interrupção segura);
- execução com **simulação (dry-run)** ou **UPDATE real**;
- confirmação obrigatória com prévia antes do UPDATE.

### 4) Auditoria por lote, relatório e rollback

Cada UPDATE gera um `batch_id` único e grava auditoria em `rtf_sanitize_audit`.

Com isso é possível:

- consultar relatório de um lote;
- exportar relatório em **CSV**;
- registrar campos extras no relatório (`Campos relatório`);
- fazer rollback completo do lote pelo `batch_id`.

### 5) Regras avançadas de marcadores

- campo **Marcadores extras (; )** na GUI para novos padrões;
- carregamento por JSON (`{"markers": ["..."]}` ou `["..."]`);
- marcador padrão DDE permanece sempre ativo.

### 6) Interface (UX)

- tema **dark** com contraste suave para uso prolongado;
- aba de banco reorganizada em blocos:
  - `1) Conexão`
  - `2) Escopo e filtros`
  - `3) Execução`
  - `4) Batch e auditoria`
- botão `?` (manual rápido) e **Manual completo** com rolagem.

---

## Instalação

## 1. Clonar e entrar no projeto

```bash
git clone <URL_DO_REPOSITORIO>
cd d-Conversor-txt-docx
```

## 2. Criar ambiente virtual (opcional, recomendado)

```bash
python -m venv .venv
```

Windows (PowerShell):

```powershell
.venv\Scripts\Activate.ps1
```

## 3. Instalar dependências

```bash
pip install -r requirements.txt
```

---

## Como usar (GUI)

## Rodar pela fonte

```bash
python rtf_sanitize_gui.py
```

## Compilar `.exe`

Opção 1 (manual):

```bash
pyinstaller --noconfirm --clean --onefile --windowed --name SanitizadorRTF_novo rtf_sanitize_gui.py
```

Opção 2 (PowerShell):

```powershell
powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
```

Executável gerado em:

`dist/SanitizadorRTF.exe`

---

## Uso no banco (passo a passo)

1. Abra a aba **Banco de dados**.  
2. Preencha conexão e clique em **Testar conexão**.  
3. Clique em **Carregar tabelas/colunas**.  
4. Selecione:
   - tabela;
   - coluna de conteúdo;
   - filtros (`Min chars`, `Min MB`) ou marque **Varredura geral**.
5. (Opcional) Informe **Campos relatório**.
6. Rode primeiro em simulação (desmarcando UPDATE).
7. Marque UPDATE e execute.
8. Guarde o **Batch ID** para relatório/rollback.

---

## Campos relatório (exemplos)

Aceita:

- colunas da própria tabela:  
  `id_documentomesclado`

- colunas relacionadas no formato `tabela.coluna`:  
  `protocolo_documentomesclado.id_protocolo`

Exemplo com múltiplos campos:

`id_documentomesclado,protocolo_documentomesclado.id_protocolo`

---

## Rollback

- Informe o `Batch ID` (ou use o último preenchido automaticamente).
- Clique em **Rollback batch**.
- O sistema restaura o conteúdo anterior das linhas daquele lote.

Observação: rollback atua sobre o lote informado; use com atenção em ambientes concorrentes.

---

## Tabela de auditoria

Criada automaticamente quando necessário:

`rtf_sanitize_audit`

Campos principais:

- `batch_id`
- `table_name`
- `key_column`
- `key_value`
- `content_column`
- `report_data` (JSONB)
- `old_content`
- `new_content`
- `old_len`
- `new_len`
- `changed_at`

---

## Script CLI (opcional)

Para manutenção em linha de comando:

```bash
python batch_sanitize_rtf.py
python batch_sanitize_rtf.py --execute --min-length 1000000
```

---

## Boas práticas de operação

- Sempre rodar simulação antes do UPDATE.
- Manter backup recente do banco.
- Executar em horário de menor carga.
- Monitorar tamanho de transações/lotes.
- Guardar o `batch_id` de cada execução.

---

## Limitações e notas

- A limpeza é baseada no marcador DDE/bookmark conhecido.
- Registros grandes sem esse padrão não serão alterados.
- O processamento em banco depende de permissões de leitura/escrita na tabela alvo.

---

## Licença

Defina aqui a licença do projeto (ex.: MIT, Apache-2.0, Proprietária).

