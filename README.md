# NF-e Destinadas · Consulta SEFAZ

> Consulta NF-e **Modelo 55** emitidas **contra o seu CNPJ** via Web Service
> `NFeDistribuicaoDFe` do Ambiente Nacional (SEFAZ Federal).

---

## ✅ O que o sistema faz

- Carrega seu **certificado digital A1** (`.pfx` ou `.p12`) e extrai o CNPJ automaticamente
- Conecta ao WebService **NFeDistribuicaoDFe** da SEFAZ com autenticação mútua TLS
- Consulta **NF-e Modelo 55 destinadas** ao seu CNPJ a partir do último NSU registrado
- Salva localmente o **NSU** e as notas para não perder o controle sequencial
- Exibe as notas em uma tabela com: chave de acesso, emitente, valor, data, situação
- Respeita o **cooldown de 60 minutos** do SEFAZ com timer regressivo
- Permite **exportar as notas em JSON**

---

## 📋 Pré-requisitos

- Windows 10 ou 11
- Certificado digital **A1** (arquivo `.pfx` ou `.p12`) com senha
- Acesso à internet para o WebService da SEFAZ

---

## 🚀 Como instalar e rodar (Windows)

### Passo 1 — Instalar o Python

1. Acesse **https://www.python.org/downloads/**
2. Clique em **"Download Python 3.x.x"** (versão mais recente)
3. Execute o instalador baixado
4. ⚠️ **IMPORTANTE:** marque a opção **"Add Python to PATH"** antes de clicar em *Install Now*
5. Conclua a instalação normalmente

> O Tkinter (interface gráfica) já vem incluído no instalador oficial — não precisa instalar separado.

---

### Passo 2 — Descompactar o projeto

1. Clique com o botão direito no arquivo `nfe_consulta.zip`
2. Selecione **"Extrair tudo..."** e escolha uma pasta (ex: `C:\nfe_consulta`)

---

### Passo 3 — Instalar as dependências

1. Abra o **Prompt de Comando** (`cmd`) ou o **PowerShell`**
   - Pressione `Win + R`, digite `cmd` e pressione Enter
2. Navegue até a pasta do projeto:
   ```cmd
   cd C:\nfe_consulta
   ```
3. Instale as dependências:
   ```cmd
   pip install -r requirements.txt
   ```
   Aguarde o download e instalação (requer internet).

---

### Passo 4 — Executar o sistema

Ainda no Prompt de Comando, dentro da pasta do projeto:

```cmd
python main.py
```

A janela do sistema abrirá automaticamente.

---

### Atalho opcional — criar um `.bat` para abrir com duplo clique

Crie um arquivo chamado `iniciar.bat` dentro da pasta do projeto com o conteúdo:

```bat
@echo off
python main.py
pause
```

Depois basta dar **duplo clique** no `iniciar.bat` para abrir o sistema.

---

## 🖥️ Como usar

1. **Clique em 📁** e selecione o arquivo `.pfx` ou `.p12` do seu certificado
2. **Digite a senha** do certificado e clique em **CARREGAR CERTIFICADO**
   - O CNPJ é extraído automaticamente do certificado
3. Verifique o **CNPJ** e selecione o ambiente (**Produção** ou **Homologação**)
4. Clique em **⟳ CONSULTAR NF-e**
5. Aguarde o resultado na tabela; dê **duplo clique** numa nota para ver os detalhes

### Regras do SEFAZ

| Situação | Comportamento |
|---|---|
| Primeira consulta | NSU = 0 (retorna últimos 90 dias) |
| Consultas subsequentes | Usa o último NSU salvo (sequência obrigatória) |
| Rejeição 656 | Consumo indevido — aguarde **60 minutos** |
| Sem documentos novos | cStat 137 — normal, nada novo desde o último NSU |

---

## 📁 Estrutura

```
nfe_consulta/
├── main.py              # Ponto de entrada
├── requirements.txt
├── src/
│   ├── app.py           # Interface gráfica (Tkinter)
│   ├── certificado.py   # Leitura do .pfx e extração de CNPJ
│   ├── sefaz_client.py  # Cliente SOAP NFeDistribuicaoDFe
│   ├── parser.py        # Parser XML (resNFe, procNFe, procEventoNFe)
│   └── storage.py       # Persistência NSU e notas (JSON local)
├── data/
│   ├── nsu_estado.json  # NSU por CNPJ (criado automaticamente)
│   └── notas/           # Notas por CNPJ (criado automaticamente)
└── logs/                # (reservado para logs futuros)
```

---

## ⚙️ Controle de NSU

O NSU (Número Sequencial Único) é **crítico**:

- O SEFAZ exige que as consultas sejam feitas em **ordem crescente e sem pular valores**
- O sistema salva automaticamente o último NSU após cada consulta em `data/nsu_estado.json`
- Se outra aplicação (ERP, portal da SEFAZ) tiver feito consultas, use **"Forçar NSU"** no painel
  lateral para sincronizar o ponto de partida

---

## 🔒 Segurança

- O arquivo `.pfx` **não é copiado** — apenas lido em memória
- Os arquivos PEM temporários usados para a conexão TLS são **deletados imediatamente** após cada consulta
- A senha do certificado **não é salva** em nenhum arquivo

---

## 🐞 Erros comuns

| Erro | Solução |
|---|---|
| `ImportError: cryptography` | `pip install cryptography` |
| `No module named tkinter` | `sudo apt install python3-tk` |
| Certificado inválido | Verifique se é A1 e se a senha está correta |
| Rejeição 656 | Aguarde 60 min — outra aplicação usou o serviço |
| HTTP 500 | Instabilidade do servidor SEFAZ — tente mais tarde |
