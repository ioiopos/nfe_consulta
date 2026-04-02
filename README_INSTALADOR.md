Com certeza! Preparei o conteúdo para você copiar e colar no Word. Organizei com uma estrutura profissional, pronta para o seu manual de suporte ou documentação interna do **Abbas PDV**.

---

# GUIA TÉCNICO: GERAÇÃO DE EXECUTÁVEL (NF-E DESTINADAS)

Este documento descreve o procedimento para transformar o código-fonte em um aplicativo funcional (`.exe`) para Windows, garantindo que todas as dependências de criptografia e comunicação com a SEFAZ sejam incluídas.

## 1. PRÉ-REQUISITOS DE AMBIENTE
Antes de iniciar, certifique-se de que o ambiente de desenvolvimento possui as ferramentas necessárias:
* **Python 3.x instalado**: Certifique-se de que a opção "Add Python to PATH" foi marcada na instalação.
* **Bibliotecas Base**: As dependências do arquivo `requirements.txt` (cryptography, requests, lxml, signxml) devem estar instaladas.
* **PyInstaller**: Ferramenta responsável por empacotar o projeto. Instale via terminal:
    > `pip install pyinstaller`

---

## 2. PROCEDIMENTO DE COMPILAÇÃO
Como o comando direto do PyInstaller pode apresentar erros de "caminho não encontrado" em algumas máquinas, utilize o comando via módulo do Python:

### Passo a Passo:
1. Abra o **PowerShell** ou **CMD** na pasta raiz do projeto (`C:\nfe_consulta`).
2. Execute o seguinte comando:
   ```powershell
   python -m PyInstaller --noconsole --onefile --name "ConsultaNFe" --add-data "src;src" main.py
   ```

### Detalhes do Comando:
* **--noconsole**: Impede que a janela preta do terminal apareça ao abrir o programa.
* **--onefile**: Agrupa tudo em um único arquivo executável facilitando a distribuição.
* **--name "ConsultaNFe"**: Define o nome do arquivo final que será gerado.
* **--add-data "src;src"**: Garante que os módulos internos (consulta, parser e persistência) sejam incluídos no pacote.

---

## 3. LOCALIZAÇÃO DO ARQUIVO FINAL
Após o término do processo, duas pastas serão criadas: `build` e `dist`.
* O executável pronto para o cliente estará em: **`C:\nfe_consulta\dist\ConsultaNFe.exe`**
* As pastas `build` e o arquivo `.spec` podem ser deletados após a geração.

---

## 4. FAQ – PERGUNTAS FREQUENTES (SUPORTE)

**O comando "pyinstaller" não é reconhecido pelo sistema?**
* Isso ocorre quando os scripts do Python não estão no PATH do Windows. Utilize o prefixo `python -m PyInstaller` para contornar o problema.

**O sistema precisa de instalação de Python no computador do cliente?**
* Não. O executável gerado com `--onefile` já contém o interpretador Python e todas as DLLs necessárias para rodar de forma independente.

**Como o sistema gerencia o NSU e as notas baixadas?**
* O executável gerencia automaticamente a criação das pastas `data/` e `notas/` para salvar o estado das consultas e os arquivos JSON.

**O que fazer em caso de "Rejeição 656" (Consumo Indevido)?**
* Esta é uma trava da SEFAZ por excesso de consultas em curto prazo. O usuário deve aguardar 60 minutos conforme o timer exibido na interface.

**O certificado digital A1 precisa estar instalado no Windows?**
* Não é obrigatório. O sistema permite carregar o arquivo `.pfx` ou `.p12` diretamente de qualquer pasta, solicitando a senha apenas em memória.