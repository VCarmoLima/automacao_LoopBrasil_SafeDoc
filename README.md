# 🤖 Automação LoopBrasil SafeDoc

> **RPA para Cálculo de Rotas, Geração de PDFs e Upload em Portal Bancário.**

Este projeto é uma solução robusta de automação desenvolvida em Python para otimizar o processo de restituição e remoção de veículos. Ele integra leitura de planilhas, cálculo de rotas via Google Maps, geração de evidências em PDF e inserção automática de dados em portal corporativo.

---

## 🚀 Funcionalidades Principais

*   **🔄 Sincronização Inteligente de Dados**
    *   Fluxo de dados: `Planilha Local` -> `Base de Rede` -> `Histórico Geral`.
    *   Garante integridade dos dados e evita reprocessamento desnecessário.
    *   Interação com usuário para resolução de conflitos de dados.

*   **🗺️ Google Maps & Cálculo de Custos**
    *   Extração automática de quilometragem (KM) via Selenium.
    *   Cálculo de valores baseado em **Ranges de KM** e **Tabelas de Custo JPR**.
    *   Geração automática de PDFs das rotas como evidência.

*   **🏦 Automação Bancária (Portal)**
    *   Login automático e navegação em menus complexos (GCA).
    *   Preenchimento de formulários e upload de arquivos PDF.

*   **📢 Notificações & Logs**
    *   **Telegram:** Envio de resumo da execução (Sucessos, Falhas e Valores Totais).
    *   **Logs Diários:** Organização automática de logs em pastas por data (`logs/YYYY-MM-DD/`).

---

## 🛠️ Pré-requisitos

*   **Python 3.8+**
*   **Google Chrome** instalado.

### Instalação das Dependências

Execute o comando abaixo para instalar as bibliotecas necessárias:

```bash
pip install pandas selenium python-dotenv openpyxl python-telegram-bot
```

---

## ⚙️ Configuração (.env)

Crie um arquivo `.env` na raiz do projeto para armazenar suas credenciais e caminhos. **Este arquivo não deve ser versionado.**

```ini
# --- Caminhos e Arquivos ---
PASTA_DOWNLOADS="C:\Caminho\Para\Downloads"
CAMINHO_BASE_EXTERNA="Z:\Rede\remocao-restituicao.xlsx"
CAMINHO_CUSTO_RESTITUICAO="C:\Dados\Custo_Restituicao.xlsx"

# --- Acesso ao Portal Bancário ---
URL_BANCO="https://seu.portal.banco.com.br"
USUARIO_BANCO="seu_usuario"
SENHA_BANCO="sua_senha"

# --- Notificações Telegram (Opcional) ---
TELEGRAM_BOT_TOKEN="seu_token_do_bot"
TELEGRAM_CHAT_ID="seu_chat_id"
```

---

## 📂 Estrutura de Arquivos Importantes

*   `automacao.py`: Script principal.
*   `Base_Restituicoes.xlsx`: Planilha de entrada (Local).
*   `historico_processamento.xlsx`: Base de dados histórica (Gerada/Atualizada automaticamente).
*   `logs/`: Diretório onde os logs de execução são salvos diariamente.

---

## ▶️ Como Executar

1.  Certifique-se de que o arquivo `.env` está configurado corretamente.
2.  Feche qualquer arquivo Excel que possa estar sendo usado pelo script.
3.  Execute o script:

```bash
python automacao.py
```

O robô iniciará o processo, exibindo o progresso no terminal e salvando logs detalhados.

---

## 📝 Licença

Este projeto está licenciado sob a licença **MIT**. Consulte o arquivo LICENSE para mais detalhes.
