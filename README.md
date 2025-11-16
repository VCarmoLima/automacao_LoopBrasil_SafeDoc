Este projeto é um script de automação (RPA) desenvolvido em Python. Sua principal função é ler uma planilha do Excel, calcular rotas no Google Maps para gerar PDFs e, em seguida, fazer o upload desses documentos e preencher um formulário em um portal web.

## Funcionalidades
* **Leitura de Dados:** Lê informações de uma planilha Excel (`.xlsm`).
* **Cálculo de Rota:** Abre o Google Maps (via URL) para calcular a quilometragem (KM) de rotas de "Remoção" e "Restituição".
* **Cálculo de Valor:** Determina o valor do serviço com base na categoria (leve, moto, pesado) e na quilometragem.
* **Geração de PDF:** Salva um PDF da página do mapa com a rota calculada.
* **Login Automatizado:** Entra no portal de acesso (`seu.acesso.io`) usando Selenium.
* **Navegação:** Navega pela estrutura de menus (GCA) até o formulário de upload.
* **Upload de Documento:** Preenche o formulário com dados (placa, contrato, valor) e anexa o PDF gerado.
* **Controle de Log:** Mantém um log em Excel (`log_processados.xlsx`) para evitar processar a mesma placa duas vezes.

## Libs Utilizadas
* **Python 3**
* **Selenium:** Para automação e controle do navegador (Google Chrome).
* **Pandas:** Para leitura e manipulação da planilha Excel `.xlsm`.
* **Python-dotenv:** Para gerenciamento seguro de credenciais e caminhos.
* **Openpyxl:** (Dependência do Pandas) para manipulação de arquivos `.xlsx`.

## Configuração do Ambiente
1.  **Clone o repositório:**
    ```bash
    git clone [https://github.com/seu-usuario/seu-repositorio.git](https://github.com/seu-usuario/seu-repositorio.git)
    cd seu-repositorio
    ```
2.  **Crie um ambiente virtual (Recomendado):**
    ```bash
    python -m venv venv
    .\venv\Scripts\activate   
    ```
3.  **Instale as bibliotecas necessárias:**
    ```bash
    pip install selenium pandas python-dotenv openpyxl
    ```
    
## Arquivo de Configuração
Para que o script funcione, você **deve** criar um arquivo chamado `.env` na raiz do projeto. Este arquivo **não** é enviado ao GitHub e contém todos os seus dados sensíveis.

Copie o conteúdo abaixo e cole no seu arquivo `.env`, substituindo com seus dados:

```ini
# Configuração de credenciais e caminhos
PASTA_DOWNLOADS="C:\Caminho\Para\Sua\Pasta\Downloads"
URL_BANCO="link_safe_doc"
USUARIO_BANCO="seu_usuario_de_login"
SENHA_BANCO="sua_senha_secreta_123"
```

## Como Executar
Após instalar as dependências e configurar o `.env`, basta executar o script principal:
```bash
python seu_script.py
```
*(Substitua `seu_script.py` pelo nome real do seu arquivo .py)*
