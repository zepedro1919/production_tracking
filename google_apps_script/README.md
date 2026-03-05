# 🏭 Production Tracking — Google Apps Script + Print Agent

Sistema de rastreamento de produção com picagem por QR Code e impressão automática de etiquetas.

---

## 📋 Como funciona

```
Trabalhador lê QR Code
        ↓
Telemóvel abre URL do Apps Script
        ↓
Apps Script:
  ├─ Valida se há picagens por fazer
  ├─ Incrementa quantidadePicada
  ├─ Gera PDF com etiqueta posicionada na folha A4
  └─ Envia PDF via HTTP POST ao Print Agent
        ↓
Print Agent (PC da impressora):
  ├─ Recebe o PDF
  └─ Imprime automaticamente (SumatraPDF silent print)
        ↓
Trabalhador pega na etiqueta e cola no produto
```

**Zero cliques. O trabalhador só lê o QR code e a etiqueta sai da impressora.**

---

## 🏷️ Layout das Etiquetas

Folha A4 com **2 colunas × 7 linhas = 14 etiquetas**:

```
┌─────────────────────────────┐
│  margem 1.5cm topo          │
│  ┌──────────┐ ┌──────────┐  │
│  │  1       │ │  2       │  │
│  ├──────────┤ ├──────────┤  │
│  │  3       │ │  4       │  │
│  ├──────────┤ ├──────────┤  │
│  │  5       │ │  6       │  │
│  ├──────────┤ ├──────────┤  │
│  │  7       │ │  8       │  │
│  ├──────────┤ ├──────────┤  │
│  │  9       │ │ 10       │  │
│  ├──────────┤ ├──────────┤  │
│  │ 11       │ │ 12       │  │
│  ├──────────┤ ├──────────┤  │
│  │ 13       │ │ 14       │  │
│  └──────────┘ └──────────┘  │
│  margem 1.5cm fundo         │
└─────────────────────────────┘
```

- Largura de cada etiqueta: ~9.85 cm
- Altura de cada etiqueta: ~3.81 cm
- Espaço entre colunas: 3 mm

---

## 🚀 Setup passo a passo

### 1. Criar o projeto Apps Script

1. Abra o Google Sheets com os dados
2. Vá a **Extensões → Apps Script**
3. Apague o código existente no `Code.gs`

### 2. Criar os ficheiros

No editor do Apps Script, crie os seguintes ficheiros (use **+** → **Script**):

| Ficheiro | Tipo | Descrição |
|---|---|---|
| `Config.gs` | Script | Configurações, dimensões, URL do print agent |
| `Code.gs` | Script | Lógica principal do web app (backend puro) |
| `LabelPrinter.gs` | Script | Geração do PDF da etiqueta |
| `QRGenerator.gs` | Script | Geração dos QR codes |
| `Setup.gs` | Script | Setup inicial e utilitários |

Copie o conteúdo de cada ficheiro da pasta `google_apps_script/` para o editor.

### 3. Verificar nomes das folhas

Certifique-se que as folhas no Google Sheets têm estes nomes exactos:
- `encomendas` — folha com os dados das encomendas
- `current_position` — folha com a posição actual da etiqueta

> ⚠️ Se os nomes forem diferentes, altere em `Config.gs`

### 4. Executar o setup

1. No editor Apps Script, selecione a função `setup`
2. Clique **Executar** (▶)
3. Autorize as permissões quando pedido
4. Verifique a mensagem de sucesso

### 4.5. Instalar o Print Agent no PC da impressora

1. No PC ligado à impressora, instale Python 3.8+
2. Instale o [SumatraPDF](https://www.sumatrapdfreader.org/) (grátis)
3. Copie a pasta `print_agent/` para o PC
4. Abra um terminal nessa pasta e execute:
   ```
   pip install -r requirements.txt
   python print_agent.py
   ```
5. Anote o IP do PC (ex: `192.168.1.50`)
6. Verifique que aparece "Waiting for print jobs..."

### 4.6. Configurar o URL do Print Agent

1. Em `Config.gs`, altere `PRINT_AGENT_URL` para o IP do PC:
   ```javascript
   PRINT_AGENT_URL: "http://192.168.1.50:5555",
   ```
2. Verifique que `PRINT_AGENT_TOKEN` é igual nos dois lados (`"producao2026"`)

### 5. Deploy como Web App

1. Clique **Deploy** → **New deployment**
2. Tipo: **Web app**
3. Configurações:
   - **Description**: Produção Tracking v1
   - **Execute as**: Me (a sua conta)
   - **Who has access**: Anyone
4. Clique **Deploy**
5. **Copie o URL** do web app

### 6. Configurar o URL

1. No editor, execute a função `setWebAppUrl`
2. Cole o URL copiado no passo anterior
3. Ou no menu: **Produção → ...** (o menu aparece ao reabrir a folha)

### 7. Gerar QR Codes

1. Reabra o Google Sheets (para carregar o menu)
2. Vá ao menu **🏭 Produção → Gerar QR Codes**
3. Os QR codes aparecem na coluna G

### 8. Imprimir QR Codes

1. Selecione a coluna dos QR codes
2. Imprima e recorte cada QR code
3. Distribua aos trabalhadores (um QR por linha/produto)

---

## 🔧 Testar

1. **`testPrintAgentConnection()`** — verifica se o Apps Script consegue falar com o print agent
2. **`testLabelGeneration()`** — gera uma etiqueta de teste e envia para a impressora
3. **`testScanFlow()`** — simula uma picagem completa na linha 2

---

## 📱 Fluxo do trabalhador

1. Pega no QR code do produto
2. Lê com a câmara do telemóvel
3. **Etiqueta sai automaticamente da impressora** (sem tocar em mais nada)
4. Cola a etiqueta no produto

---

## 🔄 Menu de gestão (na folha)

| Opção | O que faz |
|---|---|
| Gerar QR Codes | Cria QR codes para todas as linhas |
| Limpar QR Codes | Remove todos os QR codes |
| Reset Posição Etiqueta | Volta a posição para 1 |
| Reset Todas Picagens | Coloca quantidadePicada = 0 em tudo |

---

## ⚙️ Personalização

### Alterar dimensões das etiquetas
Edite os valores em `Config.gs`:
```javascript
MARGIN_TOP_CM: 1.5,
MARGIN_BOTTOM_CM: 1.5,
MARGIN_LEFT_CM: 0.5,
MARGIN_RIGHT_CM: 0.5,
GAP_HORIZONTAL_CM: 0.3,
LABEL_ROWS: 7,
LABEL_COLS: 2,
```

### Alterar o conteúdo da etiqueta
Edite a função `buildLabelHtml()` em `LabelPrinter.gs`.

### Impressão automática (sem Ctrl+P)
Já está implementada! O `print_agent.py` usa SumatraPDF para impressão silenciosa.
Certifique-se que:
- O SumatraPDF está instalado no PC da impressora
- O caminho em `print_agent.py` (`SUMATRA_PATH`) está correcto
- A impressora predefinida do Windows é a correcta (ou configure `PRINTER_NAME`)
