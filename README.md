# 🤖 WhatsApp Service Quote Bot

> **Automação inteligente em Python para envio de solicitações de orçamentos e cotações em massa via WhatsApp sem a necessidade de salvar contatos na agenda.**

---

## 📌 Contexto & Problema

Ao cotar serviços com múltiplos profissionais (pintores, eletricistas, instaladores de ar-condicionado, reformas, redes de proteção, etc.), geralmente organizamos uma planilha com dezenas de contatos e telefones.

No entanto, o processo manual enfrenta grandes gargalos:
1. **Limitação da Lista de Transmissão**: O WhatsApp exige que o destinatário tenha o seu número salvo na agenda dele para receber mensagens de listas de transmissão.
2. **Perda de Tempo e Erros Manuais**: Salvar contato por contato, abrir conversa individual, copiar o texto, ajustar o nome, anexar a foto de referência do projeto e disparar consome horas.
3. **Risco de Enviar Foto como Figurinha**: Muitas automações ou envios manuais rápidos convertem a foto em figurinha/sticker ou perdem a resolução da imagem técnica.

---

## 💡 A Solução

Este projeto automatiza todo o fluxo de contato comercial e solicitação de orçamentos:
- 🚀 **Sem Salvar Contato**: Acessa diretamente a API do WhatsApp Web (`web.whatsapp.com/send?phone=...`).
- ✍️ **Mensagem Personalizada**: Reconhece a coluna de nome da planilha e personaliza a mensagem com a tag `{nome}` (ex.: *"Bom dia! Tudo bem? É da Diego - Nippon Ar?"*).
- 📷 **Envio Duplo Inteligente (Texto + Foto Real)**:
  - **1ª Mensagem**: Envia o texto da cotação com todas as especificações do serviço.
  - **2ª Mensagem**: Em seguida, anexa e dispara a foto técnica/modelo de referência no formato de **Foto Oficial em Alta Resolução** (não vira figurinha).
- 🦊 **Multi-Navegador**: Suporte nativo e testado para **Mozilla Firefox**, **Google Chrome**, **Microsoft Edge** ou detecção automática do navegador padrão do Windows.
- 🔑 **Sessão Persistente**: Salva os cookies e a autenticação localmente. Você escaneia o **QR Code apenas na primeira execução**; depois, ele entra direto sem pedir login.
- 🛡️ **Proteção Anti-Bloqueio**: Intervalos randômicos e inteligentes entre envios (configuráveis) para simular comportamento humano e proteger o seu número de bloqueios por spam.
- 📊 **Auditoria e Relatório em Excel**: Gera automaticamente um relatório detalhado (`relatorio_envios.xlsx`) registrando horário, status do envio do texto, status da imagem e alertas sobre números que não possuem WhatsApp.

---

## 📂 Estrutura do Projeto

```text
whatsapp_bot/
├── bot.py                     # Script principal de automação
├── config.ini                 # Configuração de navegador, intervalos e foto padrão
├── mensagem.txt               # Modelo do texto padrão enviado aos prestadores
├── contatos.xlsx              # Planilha com colunas: Nome, Telefone, Imagem
├── executar.bat               # Atalho para executar com dois cliques no Windows
├── requirements.txt           # Dependências do projeto
├── imagens/                   # Pasta para armazenar fotos e projetos técnicos
│   └── exemplo.jpeg           # Imagem de referência a ser anexada
└── relatorio_envios.xlsx      # Relatório gerado automaticamente pós-execução
```

---

## 🚀 Como Usar

### 1. Pré-requisitos
Certifique-se de ter o Python 3.10+ instalado no seu computador. Em seguida, instale as dependências:
```bash
pip install -r requirements.txt
python -m playwright install firefox chromium
```

### 2. Configuração dos Contatos e Mensagem
1. **Planilha (`contatos.xlsx` ou `contatos.csv`)**: Preencha as colunas:
   - `Telefone`: Número com DDD (ex: `16 99182-0848`, `(11) 98765-4321`, `11987654321`). O bot limpa e formata automaticamente com o código do país (`55`).
   - `Nome` *(opcional)*: Nome do profissional ou empresa.
   - `Imagem` *(opcional)*: Nome da foto técnica salva na pasta `imagens/` (ex: `exemplo.jpeg`).
2. **Mensagem (`mensagem.txt`)**: Escreva o texto desejado. Use `{nome}` onde quiser que o nome do prestador seja inserido.
3. **Navegador & Intervalos (`config.ini`)**:
   ```ini
   [CONFIGURACOES]
   navegador = firefox      # Opções: firefox, chrome, edge ou auto
   delay_minimo = 8         # Segundos mínimos entre contatos
   delay_maximo = 15        # Segundos máximos entre contatos
   imagem_padrao = exemplo.jpeg
   ```

### 3. Execução
- No Windows, dê um duplo clique no arquivo **`executar.bat`**.
- Ou execute diretamente pelo terminal:
  ```bash
  python bot.py
  ```

---

## 🔒 Segurança e Boas Práticas
- Os dados da sessão do navegador ficam salvos em pastas locais (`perfil_firefox/`, etc.) e são ignorados pelo `.gitignore`.
- Nunca compartilhe seus arquivos de perfil ou relatórios contendo dados de terceiros.
- Mantenha intervalos de segurança entre envios para respeitar as diretrizes do WhatsApp.

---

## 🛠️ Tecnologias Utilizadas
- [Python 3](https://www.python.org/)
- [Playwright](https://playwright.dev/python/) — Automação de navegadores moderna e resiliente
- [Pandas](https://pandas.pydata.org/) — Tratamento de planilhas e dados tabulares
- [OpenPyXL](https://openpyxl.readthedocs.io/) — Leitura e escrita de arquivos `.xlsx`
