# Editor em Lote de Informativos de Atividade Programada (Outlook Clássico)

Aplicação desktop em Python com **interface web no navegador** (distribuída também como executável `.exe`) para leitura em lote de e-mails de notificações/informativos de atividades programadas no **Microsoft Outlook Clássico**, detecção automática de campos em comum, edição em lote ou individual e abertura controlada de rascunhos no Outlook (`.Display()`, sem envio automático).

---

## Principais Funcionalidades

- **Seleção de Origem dos E-mails (Pasta do Outlook ou Pasta Local com `.msg`):**
  - **Pasta do Outlook:** Navegação pela árvore de todas as contas e caixas compartilhadas (como `sir.fibra`) diretamente na interface web ou por meio da janela nativa do Outlook (`Namespace.PickFolder()`).
  - **Pasta Local (`.msg`):** Opção para buscar ou digitar uma pasta do computador contendo arquivos `.msg` salvos localmente (com botão **📂 Procurar pasta…** que abre o seletor de pastas do Windows). Lê todos os arquivos `.msg` preservando HTML, destinatários, anexos e logotipo embutido, sem modificar os arquivos originais no disco.
  - Filtro por assunto e limite de itens carregados, além de memorização da última pasta utilizada.
- **Detecção Inteligente de Campos em Comum:**
  - **Assunto:** identifica prefixos de cenário (`REPROGRAMAÇÃO:`, `CANCELAMENTO:`) e assuntos base padronizados (`INFORMATIVO DE ATIVIDADE PROGRAMADA - AMERICAN TOWER` e `MANUTENÇÃO PROGRAMADA`), permitindo edição global para o lote inteiro ou individual por e-mail.
  - **Tópicos do Corpo (`Rótulo: valor`):** extrai automaticamente os tópicos em negrito do corpo do e-mail:
    - `Registro da atividade`
    - `Localidade(s) de execução`
    - `Data/Hora de início` (com seletor de data e hora mantendo o padrão `dd/mm/aaaa – HHhMM`)
    - `Data/Hora de término` (com seletor de data e hora mantendo o padrão `dd/mm/aaaa – HHhMM`)
    - `Descrição/Justificativa`
    - `Listagem dos equipamentos, clientes e/ou nodes afetados do cliente` — protegido automaticamente com cadeado (`🔒`) por variar entre clientes, mas editável individualmente quando necessário.
- **Atalhos de Cenário (Reprogramação / Cancelamento):**
  - Botões rápidos que aplicam o prefixo correspondente no assunto, destacam os campos de data/hora e oferecem a opção de inserir um aviso em destaque logo após `"Prezado cliente,"`.
- **Preservação Total de Formatação e Imagens Embutidas (CID):**
  - Edição cirúrgica do HTML via `BeautifulSoup`, mantendo fontes, cores, tabelas, anexos e o logotipo embutido no cabeçalho do informativo.
  - Pré-visualização em tempo real (Original vs. Alterado) com destaque amarelo nas mudanças.
- **Abertura Segura no Outlook (`.Display()`):**
  - Cria cada rascunho a partir do item original preservando destinatários (`Para`, `Cc`, `Cco`), anexos e imagens inline.
  - Preenche automaticamente o remetente com a caixa compartilhada (`sir.fibra` via `SentOnBehalfOfName` / `SendUsingAccount`) ou deixa em branco para escolha manual.
  - Suporta abertura **um por vez** (fila passo a passo para não travar a tela com dezenas de janelas) ou **todos de uma vez**.
  - **Nunca envia automaticamente** — o clique em *Enviar* é sempre feito manualmente pelo operador na janela do Outlook.
- **Modo de Demonstração Integrado:**
  - Permite testar todo o fluxo com 6 informativos fictícios mesmo sem estar conectado ao Outlook.
- **Registro de Logs (CSV):**
  - Gera logs detalhados de todos os rascunhos abertos em `logs/log_informativos_YYYY-MM-DD.csv` e permite download direto pela interface.

---

## Estrutura do Projeto

```text
email_batch_editor/
├── app.py                  # Servidor local Flask (127.0.0.1), rotas da API e inicialização
├── config.json             # Configurações padrão (caixa compartilhada, rótulos, assuntos base)
├── requirements.txt        # Dependências Python
├── Iniciar_Editor.bat      # Atalho para iniciar o .exe ou rodar via Python
├── core/
│   ├── models.py           # Modelos de dados (EmailItem, ExtractedField, CommonFieldInfo)
│   ├── detector.py         # Extração de campos do HTML e análise de lote
│   ├── html_editor.py      # Substituição de valores no HTML preservando formatação
│   └── demo_data.py        # Lote de demonstração com 6 informativos de exemplo
├── outlook/
│   ├── worker.py           # Thread dedicada STA com inicialização COM (pythoncom)
│   ├── reader.py           # Leitura de caixas, pastas, destinatários e imagens CID
│   └── writer.py           # Criação de rascunhos no Outlook e abertura com .Display()
├── static/
│   ├── index.html          # Interface web principal
│   ├── style.css           # Estilos da aplicação
│   └── app.js              # Lógica de estado, pré-visualização e fila de rascunhos
├── prototipo/
│   └── index.html          # Protótipo estático inicial de validação de interface
└── test_core.py            # Testes automatizados de detecção, edição HTML e API
```

---

## Pré-requisitos

- **Sistema Operacional:** Windows 10 ou 11
- **Microsoft Outlook Clássico (Desktop)** instalado e com sessão ativa (para uso com e-mails reais)
- **Python 3.10+** (apenas para execução via código-fonte ou geração do `.exe`)

---

## Como Executar

### 1. Usando o Executável (`.exe`)
Se o executável já estiver gerado em `dist\EditorInformativos.exe`:
1. Dê dois cliques em `Iniciar_Editor.bat` ou diretamente em `dist\EditorInformativos.exe`.
2. O navegador padrão abrirá automaticamente em `http://127.0.0.1:54123`.

### 2. Executando via Código-Fonte (Python)

```powershell
# Criar ambiente virtual
python -m venv .venv

# Instalar dependências
.\.venv\Scripts\pip install -r requirements.txt

# Iniciar a aplicação
.\.venv\Scripts\python app.py
```

### 3. Gerando um Novo Executável (`.exe`) com PyInstaller

```powershell
.\.venv\Scripts\pyinstaller --noconfirm --clean --onefile --name EditorInformativos `
  --add-data "static;static" `
  --add-data "config.json;." `
  --hidden-import win32timezone `
  app.py
```
O arquivo final será salvo em `dist\EditorInformativos.exe`.

---

## Executando os Testes

```powershell
.\.venv\Scripts\python test_core.py
```
