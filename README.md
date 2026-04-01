# Sanitizador RTF (DDE/Bookmark) para TXT/RTF e PostgreSQL

Ferramenta para higienizar conteúdos RTF/TXT com corrupção por metadados de DDE/bookmark (ex.: `{\*\bkmkstart __DdeLink__...}`), reduzindo tamanho de registros e evitando erros de memória em editores e consultas.

![Python](https://img.shields.io/badge/Python-3.13%2B-3776AB?logo=python&logoColor=white)
![Platform](https://img.shields.io/badge/Plataforma-Windows-0078D6?logo=windows&logoColor=white)
![Status](https://img.shields.io/badge/Status-Em%20uso-2EA043)
![Database](https://img.shields.io/badge/PostgreSQL-16%2B-4169E1?logo=postgresql&logoColor=white)
![License](https://img.shields.io/badge/Licen%C3%A7a-definir-lightgrey)

## Problema que a ferramenta resolve

Em alguns cenários, o conteúdo RTF/TXT fica com:

- **DDE / bookmarks** repetitivos (`{\*\bkmkstart __DdeLink__` …);
- **Metadados** RTF redundantes (generator, rsidtbl, etc.);
- **Imagens e OLE** embutidos em hex (muito comuns em RTF exportado por **LibreOffice/Collabora**: `{\*\shppict` + `{\nonshppict{\pict{` + `pngblip`/binário), inchando o ficheiro para **dezenas ou centenas de MB** sem aumentar o texto legível.

Consequências típicas: travamentos em editores, timeouts e erros de memória em consultas ao PostgreSQL.

A limpeza combina **remoção de grupos RTF** (conforme o **nível** escolhido), **deteção de massas hexadecimais** suspeitas e **corte tardio** quando um marcador conhecido aparece nos últimos 10% do documento, com **fecho de chaves** para manter estrutura básica.

---

## Tecnologias usadas

- **Python 3.13+**
- **Tkinter** (GUI desktop)
- **PostgreSQL** (via `psycopg2-binary`)
- **PyInstaller** (geração de `.exe`)

---

## Estrutura do projeto

- `rtf_sanitize.py`  
  Núcleo de higienização: níveis `seguro` / `intermediario` / `agressivo`, `DEFAULT_MARKERS`, limiar **5 MB** para remoção pesada de pict/OLE, deteção de hex órfão (500k+ dígitos).

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
  - **WHERE só por tamanho** (opcional): o `WHERE` do PostgreSQL usa apenas os limiares de tamanho, **sem** `position(marcador)` — muito mais rápido quando a coluna tem dezenas de MB por linha (caso contrário o motor pode fazer várias procuras de texto em cada registo).
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

- campo **Marcadores extras (;)** e JSON (`{"markers": ["..."]}` ou `["..."]`);
- os **marcadores padrão** (DDE, shppict, nonshppict, `{\pict{`, objdata, `\*\pict`) incorporam-se sempre à lista ativa na GUI e na lógica de deteção.

### 6) Interface (UX)

- tema **dark** com contraste suave para uso prolongado;
- no **Windows**, barra de título nativa (min/max/fechar) alinhada ao tema escuro (**DWM**), também nas janelas de ajuda;
- **alerta sonoro** breve ao concluir: limpeza de um ficheiro, lote de pasta, ou higienização no banco (simulação ou UPDATE);
- aba **Banco de dados**:
  - **Área rolável** (canvas) para caber em ecrãs mais baixos, com barra vertical coerente com o resto da app;
  - quatro **cartões colapsáveis** (clique no título **▶** / **▼** para expandir ou recolher):
    - `1) Conexão` — credenciais, **Testar conexão** e **Carregar tabelas/colunas**;
    - `2) Escopo e filtros` — tabela, coluna, limiares, **WHERE só por tamanho**, campos de relatório;
    - `3) Execução` — opções RTF/varredura/UPDATE/validação, progresso, **Higienizar banco** e **Parar**;
    - `4) Batch e auditoria` — Batch ID, relatório, CSV, rollback;
  - por **defeito os quatro cartões vêm recolhidos** (só os títulos visíveis); a secção **URL gerada (opcional)** mantém-se sempre visível por baixo;
- **Barras de rolagem** verticais com estilo escuro unificado (trilho alinhado ao fundo do registo `#262c37`, polegar mais largo e legível), na aba Banco, no **Registo** e nas janelas de ajuda;
- botão **`?`** (manual rápido) e **Manual completo**: janelas em **tema escuro** (`Toplevel` com texto rolável, não `messagebox`); no manual rápido, **Escape** fecha e a janela é modal.

---

## Níveis de higienização

| Nível | O que faz |
|-------|-----------|
| **seguro** | Remove apenas grupos de bookmark **DDE** associados a `__DdeLink__` (regex interna). Não remove imagens. |
| **intermediario** | Tudo do seguro + remove grupos auxiliares (`\*\generator`, `\*\userprops`, `\*\xmlnstbl`, `\*\rsidtbl`, `\*\themedata`, `\*\colorschememapping`). |
| **agressivo** | Tudo do intermediário + se o texto ainda tiver **> ~5 MB**: remove grupos completos que começam em `\*\shppict`, `\nonshppict`, `\*\objdata`, `\*\pict`, `{\pict{` (**elimina imagens/OLE** embutidos). Em qualquer tamanho: se existir uma massa **≥ 500 000** dígitos hex (com espaços/newlines entre dígitos), **trunca antes** dessa massa. Pode aplicar-se corte tardio se um marcador padrão aparecer nos **últimos 10%** do ficheiro. |

**Regra dos ~5 MB:** a remoção de blocos pict/obj/shppict (que apaga figuras) só corre no nível agressivo quando o conteúdo já ultrapassa esse limiar — ficheiros pequenos não são “esmiuçados” da mesma forma.

**Marcadores padrão** (para filtros SQL, `precisa_limpeza` e fallback de corte): bookmark DDE, `\*\shppict`, `\nonshppict`, `{\pict{`, `\*\objdata`, `\*\pict`. É possível acrescentar marcadores na GUI (**extras** ou JSON).

**PostgreSQL:** o conteúdo anterior a cada UPDATE fica em `rtf_sanitize_audit.old_content`; use **rollback** por `batch_id` se um lote agressivo não for o desejado.

---

## Instalação

## 1. Clonar e entrar no projeto

```bash
git clone https://github.com/BrunoHF04/SanitizadorRTF.git
cd SanitizadorRTF
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
2. Expanda **1) Conexão** (clique na linha do título para mudar **▶** → **▼**), preencha host/porta/banco/utilizador/senha e clique em **Testar conexão**.  
3. No mesmo cartão, clique em **Carregar tabelas/colunas**.  
4. Expanda **2) Escopo e filtros** e selecione:
   - tabela;
   - coluna de conteúdo;
   - filtros (`Min chars`, `Min MB`) ou marque **Varredura geral**;
   - (opcional) **WHERE só por tamanho** quando usar limiares em colunas muito grandes.
5. (Opcional) Informe **Campos relatório**.  
6. Expanda **3) Execução**, rode primeiro em **simulação** (desmarcando **Aplicar UPDATE**).  
7. Marque **Aplicar UPDATE** e **Higienizar banco** quando estiver pronto. Use **Parar** se precisar interromper.  
8. O **Batch ID** pode ser copiado do registo; em **4) Batch e auditoria** use relatório, CSV ou **Rollback**.

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

Para manutenção em linha de comando (`documento_mesclado` no PostgreSQL). O `WHERE` considera `LENGTH(conteudo)` **ou** qualquer marcador em `DEFAULT_MARKERS`.

```bash
python batch_sanitize_rtf.py
python batch_sanitize_rtf.py --execute --min-length 1000000
python batch_sanitize_rtf.py --execute --cleaning-level agressivo --min-length 500000
```

`--cleaning-level`: `seguro` (padrão) | `intermediario` | `agressivo`.

---

## Boas práticas de operação

- Sempre rodar simulação antes do UPDATE.
- Manter backup recente do banco.
- Executar em horário de menor carga.
- Monitorar tamanho de transações/lotes.
- Guardar o `batch_id` de cada execução.

---

## Limitações e notas

- **Nível seguro** só trata DDE/bookmarks; ficheiros enormes só por imagens precisam de **agressivo** (ou outra ferramenta).
- **Agressivo** remove **figuras e objetos** embutidos quando o conteúdo ultrapassa ~5 MB; o texto da ata costuma manter-se, mas faça **simulação** e use **rollback** no banco se necessário.
- O critério de candidatos no banco usa **marcadores padrão + tamanho** (`--min-length`); sem correspondência, o registo não entra no lote.
- Permissões de leitura/escrita na tabela alvo e extensão `psycopg2` são necessárias.

---

## Licença

Defina aqui a licença do projeto (ex.: MIT, Apache-2.0, Proprietária).

