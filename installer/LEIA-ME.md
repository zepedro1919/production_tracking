# Print Agent Nautilus - Guia de Instalação

## O que é isto?

Este é o sistema de impressão automática de etiquetas da Nautilus.
Quando alguém lê um QR Code na fábrica, a etiqueta é impressa automaticamente.

---

## Requisitos

- **Windows 7 ou superior**
- **Impressora** ligada ao PC (ex: EPSON ET-3700 Series)
- **Ligação à internet**

---

## Instalação (só precisa fazer 1 vez)

### Passo 1: Preparar a pasta

Copie toda a pasta `installer` para o PC (por exemplo, para `C:\PrintAgent\`).

### Passo 2: Descarregar ferramentas necessárias

Antes de correr o instalador, descarregue estes programas:

| Programa | Download | Onde colocar |
|----------|----------|-------------|
| **SumatraPDF** | [sumatrapdfreader.org/download](https://www.sumatrapdfreader.org/download-free-pdf-viewer) - versão **Portable** | `tools\SumatraPDF.exe` |
| **Python 3.8** | [python.org/downloads](https://www.python.org/downloads/release/python-3820/) | Instalar normalmente (marcar **"Add to PATH"**) |

> **NOTA:** O túnel (LocalTunnel) é 100% Python — **não precisa de descarregar
> nenhum binário extra**. Compatível com Windows 7, 8, 10 e 11.

> **NOTA Windows 7:** Use Python **3.8** (última versão compatível).

### Passo 3: Correr o instalador

1. Clique duas vezes em **`setup.bat`**
2. Siga as instruções no ecrã
3. Quando pedir a impressora, copie o nome exacto que aparece na lista

---

## Utilização Diária

### Iniciar o sistema

1. Clique duas vezes em **`iniciar.bat`** (ou no atalho do Desktop)
2. Vai abrir **1 janela** com tudo:
   - **Print Agent** - o servidor de impressão
   - **LocalTunnel** - o túnel de internet (automático, 100% Python)
3. O URL do túnel é mostrado e copiado automaticamente

> 💡 **LocalTunnel é gratuito e sem conta** — pode correr infinitamente!
> Mas o URL muda se reiniciar o programa.

### Actualizar o URL no Google Apps Script

⚠️ **O URL muda cada vez que reinicia!** (versão gratuita)

1. Copie o URL que aparece quando inicia (ex: `https://xxxx-xxxx-xxxx.loca.lt`)
2. Abra o **Google Apps Script** do projecto
3. No ficheiro **Config.gs**, altere a linha:
   ```javascript
   PRINT_AGENT_URL: "https://NOVO-URL-AQUI.loca.lt"
   ```
4. **Guarde** o ficheiro (Ctrl+S)
5. Vá a **Implementar → Gerir implementações**
6. Clique no **lápis (editar)** na implementação existente
7. Mude a versão para **"Nova versão"**
8. Clique **Implementar**

> ⚠️ **IMPORTANTE:** Nunca crie uma "Nova implementação"! Isso muda o URL dos QR Codes e todos deixam de funcionar!

### Parar o sistema

Feche a janela do terminal (ou pressione Ctrl+C).

---

## Resolução de Problemas

### "SumatraPDF não encontrado"
- Verifique que `SumatraPDF.exe` está na pasta `tools\`
- Ou corra `setup.bat` novamente e indique o caminho

### "Falha ao criar túnel"
- Verifique a ligação à internet
- O serviço localtunnel.me pode estar temporariamente em baixo
- O launcher tenta 3 vezes automaticamente
- Verifique o ficheiro `launcher.log` para erros detalhados

### "Impressão não funciona"
- Verifique que a impressora está ligada e é a predefinida do Windows
- Abra `http://localhost:5555/health` no browser para verificar o estado
- Verifique o ficheiro `launcher.log` para erros detalhados

### "QR Code dá erro 404"
- O túnel não está a correr. Inicie com `iniciar.bat`
- Se mudou o PC, o URL mudou. Actualize o Config.gs

### "Erro de autorização"
- O token no Config.gs tem de ser: `producao2026`

---

## Estrutura de Ficheiros

```
installer/
├── setup.bat           ← Instalador (correr 1 vez)
├── iniciar.bat         ← Iniciar sistema (uso diário)
├── print_agent.py      ← Servidor de impressão
├── launcher.py         ← Arranca tudo (print agent + tunnel)
├── config.json         ← Configuração (criado automaticamente)
├── config.ini          ← Config do instalador
├── launcher.log        ← Log de actividade
├── tunnel_url.txt      ← Último URL do túnel
├── LEIA-ME.md          ← Este ficheiro
└── tools/
    └── SumatraPDF.exe  ← Descarregar e colocar aqui
```

---

## Contacto

Em caso de dúvida, contactar a equipa Kaizen.
