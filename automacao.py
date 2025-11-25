import pandas as pd
import os
import time
import json
import base64
import datetime
import logging
import telegram
import asyncio
import glob
import math
import unicodedata
import re
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from concurrent.futures import ThreadPoolExecutor, as_completed

load_dotenv()

# --- Configurações ---
NOME_ARQUIVO_EXCEL = "Base Restituições Vinícius.xlsm"
NOME_ABA_CALCULOS = "Calculos"
NOME_ABA_BASES = "Bases"
PASTA_DOWNLOADS = os.getenv("PASTA_DOWNLOADS")
NOME_ARQUIVO_HISTORICO = "historico_processamento.xlsx"

CAMINHO_BASE_EXTERNA = os.getenv("CAMINHO_BASE_EXTERNA")
CAMINHO_CUSTO_RESTITUICAO = os.getenv("CAMINHO_CUSTO_RESTITUICAO")

COLUNAS_EXTERNAS_MAP = {
    'Placa': 'placa_key',
    'Guincheiro': 'transp_raw',
    'nm': 'patio_raw',
    'CidadeOrigem': 'cidade_raw',
    'financiado': 'financiado_db',
    'cpf': 'cpf_db',
    'Contrato': 'contrato_externo',
    'ValorGuincheiro': 'valor_base_db',

    # --- NOVAS COLUNAS ---
    'DataSolicitacao': 'Data de Remoção',
    'Marca': 'Marca',
    'Modelo': 'Modelo',
    'Categoria': 'Categoria_Ext',
    'Chassi': 'Chassi'
}

COLUNA_PLACA = "Placa"
COLUNA_CONTRATO = "Contrato"
COLUNA_CATEGORIA = "Categoria"
COLUNA_TESTE = "Teste"
COLUNA_END1 = "Endereço transportadora"
COLUNA_END2 = "Endereço Pátio"
COLUNA_END3 = "Cidade convertida"

URL_BANCO = os.getenv("URL_BANCO")
USUARIO_BANCO = os.getenv("USUARIO_BANCO")
SENHA_BANCO = os.getenv("SENHA_BANCO")

VALOR_RANGES = {
    "leve": [(200, 241), (500, 468), (700, 620), (1000, 900), (9999, 1320)],
    "moto": [(200, 230), (500, 438), (700, 580), (1000, 795), (9999, 880)],
    "pesado": [(200, 665), (500, 1045), (700, 2020), (1000, 3235), (9999, 4175)]
}

SELECTORS = {
    "google_maps": {
        "km_xpaths_list": [
            # 1. SEU XPATH ESPECÍFICO (Prioridade Máxima)
            "/html/body/div[1]/div[2]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[5]/div[1]/div[1]/div/div[1]/div[2]/div",

            # 2. Variações comuns (caso mude para div[3] ou div[10])
            "/html/body/div[1]/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[5]/div[1]/div[1]/div/div[1]/div[2]/div",
            "//div[contains(@id, 'section-directions-trip-0')]//div[contains(text(), 'km')]",
            "//div[contains(@class, 'ivN21e')]",
            "//div[contains(text(), 'km') and contains(text(), 'min')]"
        ],
        "canvas_map": "canvas"
    },
    "login": {
        "usuario": "/html/body/div/main/div/div[2]/form/div/div[2]/div/input",
        "senha": "/html/body/div/main/div/div[3]/div/input",
        "botao": "/html/body/div/main/div/div[2]/form/div/div[4]/input[1]"
    },
    "gca_menu": {
        "link_1": "/html/body/main/section/form/div/div/div/div/div/div/div[1]/div[2]/div/nobr/div/div[1]/span[1]/img",
        "link_2": "/html/body/main/section/form/div/div/div/div/div/div/div[1]/div[2]/div/nobr/div/div[2]/div[3]/span[1]/img[2]",
        "link_3": "/html/body/main/section/form/div/div/div/div/div/div/div[1]/div[2]/div/nobr/div/div[2]/div[4]/div[2]/img"
    },
    "iframes": {
        "externo": "ifrmForm",
        "interno": "ifrmObject"
    },
    "form_upload": {
        "input_arquivo": "_ctl0_ContentPlaceHolder_IDX_FILE",
        "select_status": "/html/body/form/div/div/div/div/div[2]/div[3]/select",
        "input_data": "/html/body/form/div/div/div/div/div[2]/div[4]/div/input",
        "input_contrato": "/html/body/form/div/div/div/div/div[2]/div[5]/input",
        "input_placa": "/html/body/form/div/div/div/div/div[2]/div[6]/input",
        "select_tipo_despesa": "/html/body/form/div/div/div/div/div[2]/div[7]/select",
        "input_valor": "/html/body/form/div/div/div/div/div[2]/div[8]/input",
        "input_caixa_arquivo": "/html/body/form/div/div/div/div/div[2]/div[9]/input",
        "input_observacao": "/html/body/form/div/div/div/div/div[2]/div[10]/input",
        "botao_salvar": "/html/body/form/div/div/div/div/div[3]/input",
        "mensagem_sucesso": "/html/body/form/div/div/div/div/div[4]/div/span"
    }
}

# --- Funções de Suporte ---
def configurar_logger_dinamico():
    try: diretorio_script = os.path.dirname(os.path.abspath(__file__))
    except: diretorio_script = os.getcwd()
    pasta_logs = os.path.join(diretorio_script, "logs")
    os.makedirs(pasta_logs, exist_ok=True)
    hoje_str = datetime.date.today().strftime("%Y-%m-%d")
    padrao = os.path.join(pasta_logs, f"log_{hoje_str}_v*.txt")
    maior = 0
    for arq in glob.glob(padrao):
        try: maior = max(maior, int(os.path.splitext(os.path.basename(arq))[0].split("_v")[-1]))
        except: pass
    nome_log = os.path.join(pasta_logs, f"log_{hoje_str}_v{maior + 1}.txt")
    for h in logging.root.handlers[:]: logging.root.removeHandler(h)
    logging.basicConfig(filename=nome_log, filemode='w', level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s', datefmt='%Y-%m-%d %H:%M:%S')
    print(f"--- Log: {os.path.basename(nome_log)} ---")

def limpar_texto_estilo_excel(texto):
    if not isinstance(texto, str): return ""
    nfkd = unicodedata.normalize('NFKD', texto)
    texto_sem_acento = "".join([c for c in nfkd if not unicodedata.combining(c)])
    return " ".join(re.sub(r'[^A-Z0-9\s]', '', texto_sem_acento.upper()).split())

def carregar_bases_de_enderecos():
    try:
        df_bases = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA_BASES, header=None)
        dict_transp = dict(zip(df_bases[0].dropna().astype(str).str.strip().str.upper(), df_bases[1].dropna().astype(str).str.strip()))
        dict_patio = dict(zip(df_bases[2].dropna().astype(str).apply(limpar_texto_estilo_excel), df_bases[3].dropna().astype(str).str.strip()))
        return dict_transp, dict_patio
    except Exception as e:
        logging.error(f"Erro Base Endereços: {e}")
        return {}, {}

def carregar_base_externa_rede():
    logging.info("Lendo base externa da rede...")
    if not os.path.exists(CAMINHO_BASE_EXTERNA):
        logging.critical(f"Arquivo não encontrado: {CAMINHO_BASE_EXTERNA}")
        return pd.DataFrame()
    try:
        df_ext = pd.read_excel(CAMINHO_BASE_EXTERNA, sheet_name="remocao", usecols=list(COLUNAS_EXTERNAS_MAP.keys()), engine='openpyxl', dtype=str)
        df_ext.rename(columns=COLUNAS_EXTERNAS_MAP, inplace=True)
        if 'placa_key' in df_ext.columns:
            df_ext['placa_key'] = df_ext['placa_key'].str.strip().str.upper()
            df_ext.drop_duplicates(subset=['placa_key'], inplace=True)
        return df_ext
    except Exception as e:
        logging.critical(f"Erro leitura externa: {e}")
        return pd.DataFrame()

# --- NOVA FUNÇÃO: CARREGAR CUSTOS JPR ---
def carregar_tabela_custos_jpr():
    logging.info("Carregando tabela de custos JPR...")
    if not os.path.exists(CAMINHO_CUSTO_RESTITUICAO):
        logging.warning(f"Arquivo JPR não encontrado em {CAMINHO_CUSTO_RESTITUICAO}")
        return {}
    try:
        # Carrega a aba 'Todos'. Estrutura esperada:
        # Col B (1): Cidade convertida | Col C (2): Pátio | Col D (3): Transp
        # Col E (4): Moto | Col F (5): Leve | Col G (6): Caminhonete/Pesado
        df = pd.read_excel(CAMINHO_CUSTO_RESTITUICAO, sheet_name='Todos', header=None, skiprows=1) # Skiprows assumindo cabeçalho

        tabela_jpr = {}
        for _, row in df.iterrows():
            # Chave composta: (Cidade Limpa, Pátio Limpo, Transp Limpo)
            cid = limpar_texto_estilo_excel(str(row[1]))
            pat = limpar_texto_estilo_excel(str(row[2]))
            tra = limpar_texto_estilo_excel(str(row[3]))

            vals = {
                'Moto': row[4],
                'Leve': row[5],
                'Caminhonete': row[6]
            }
            tabela_jpr[(cid, pat, tra)] = vals

        logging.info(f"Custos JPR carregados: {len(tabela_jpr)} rotas.")
        return tabela_jpr
    except Exception as e:
        logging.error(f"Erro ao ler JPR: {e}")
        return {}

# --- NOVA FUNÇÃO: Sincronizar Colunas Dinâmicas Locais ---
def sincronizar_dados_dinamicos_local(df_historico):
    logging.info("Sincronizando colunas dinâmicas da planilha local...")
    try:
        # Colunas que devem ser sempre atualizadas (sobrescritas)
        cols_dinamicas = [
            'Status atual',
            'Tipo de restituição',
            'Data restituição',
            'Fechamento Solicitação',
            'Tipo de liberação'
        ]

        # Lê a planilha local original
        df_local = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA_CALCULOS, dtype=str)
        df_local[COLUNA_PLACA] = df_local[COLUNA_PLACA].str.strip().str.upper()

        # Filtra apenas o que interessa
        cols_existentes = [c for c in cols_dinamicas if c in df_local.columns]
        if not cols_existentes:
            return df_historico

        df_local_resumo = df_local[[COLUNA_PLACA] + cols_existentes].drop_duplicates(subset=[COLUNA_PLACA])

        # Remove as colunas antigas do histórico para evitar duplicação (_x, _y) e forçar atualização
        for col in cols_existentes:
            if col in df_historico.columns:
                df_historico.drop(columns=[col], inplace=True)

        # Faz o merge para trazer os dados frescos
        df_atualizado = pd.merge(df_historico, df_local_resumo, on=COLUNA_PLACA, how='left')

        return df_atualizado
    except Exception as e:
        logging.error(f"Erro ao sincronizar dados dinâmicos: {e}")
        return df_historico

def calcular_valor_restituicao_final(transp_nome, cidade_nome, patio_nome, categoria, valor_remocao, tabela_jpr):
    # 1. Se NÃO for JPR, valor é igual ao de remoção
    if 'JPR' not in transp_nome.upper():
        return valor_remocao

    # 2. Se FOR JPR, busca na tabela
    chave = (
        limpar_texto_estilo_excel(cidade_nome),
        limpar_texto_estilo_excel(patio_nome),
        limpar_texto_estilo_excel(transp_nome)
    )

    valores = tabela_jpr.get(chave)
    if not valores:
        return "Não encontrada"

    # Mapeia categoria para a coluna correta
    cat_lower = str(categoria).strip().lower()
    if 'moto' in cat_lower:
        return valores.get('Moto', 0)
    elif 'leve' in cat_lower:
        return valores.get('Leve', 0)
    else:
        # Assume Caminhonete/Pesado para outros casos
        return valores.get('Caminhonete', 0)

def sincronizar_dados_dinamicos_local(df_historico):
    logging.info("Sincronizando colunas dinâmicas da planilha local...")
    try:
        # Colunas que mudam sempre e devem ser atualizadas
        cols_dinamicas = [
            'Status atual', 'Tipo de restituição', 'Data restituição',
            'Fechamento Solicitação', 'Tipo de liberação'
        ]

        # Lê a planilha local original
        df_local = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA_CALCULOS, dtype=str)
        df_local[COLUNA_PLACA] = df_local[COLUNA_PLACA].str.strip().str.upper()

        # Filtra apenas o que interessa
        cols_existentes = [c for c in cols_dinamicas if c in df_local.columns]
        if not cols_existentes:
            return df_historico

        df_local_resumo = df_local[[COLUNA_PLACA] + cols_existentes].drop_duplicates(subset=[COLUNA_PLACA])

        # Remove as colunas antigas do histórico para forçar a atualização fresca
        for col in cols_existentes:
            if col in df_historico.columns:
                df_historico.drop(columns=[col], inplace=True)

        # Faz o merge para trazer os dados frescos
        df_atualizado = pd.merge(df_historico, df_local_resumo, on=COLUNA_PLACA, how='left')

        return df_atualizado
    except Exception as e:
        logging.error(f"Erro ao sincronizar dados dinâmicos: {e}")
        return df_historico

def atualizar_historico_existente(df_ext, tabela_jpr, dict_patio):
    logging.info("Verificando dados faltantes e calculando valores no Histórico...")

    if not os.path.exists(NOME_ARQUIVO_HISTORICO):
        df_hist = pd.DataFrame(columns=[COLUNA_PLACA])
    else:
        try: df_hist = pd.read_excel(NOME_ARQUIVO_HISTORICO)
        except: df_hist = pd.DataFrame(columns=[COLUNA_PLACA])

    try:
        if df_hist.empty and df_ext.empty: return

        df_hist[COLUNA_PLACA] = df_hist[COLUNA_PLACA].astype(str).str.strip().str.upper()

        # Mapeamento dos dados da Rede (df_ext) para o Histórico (Excel)
        mapa_colunas_final = {
            'cpf_db': 'CPF_Banco',
            'financiado_db': 'Financiado_Banco',
            'contrato_externo': 'Contrato_Externo',
            'valor_base_db': 'Valor_Base_Guincho',
            'transp_raw': 'Transportadora',
            'patio_raw': 'Pátio',
            'cidade_raw': 'Cidade convertida', # <--- ESSENCIAL: Salva a cidade da rede no histórico
            'Data de Remoção': 'Data de Remoção',
            'Marca': 'Marca',
            'Modelo': 'Modelo',
            'Categoria_Ext': 'Categoria',
            'Chassi': 'Chassi'
        }

        if 'Valor_Base_Guincho2' not in df_hist.columns:
            df_hist['Valor_Base_Guincho2'] = None

        df_hist['Valor_Base_Guincho2'] = df_hist['Valor_Base_Guincho2'].astype('object')

        if not df_ext.empty:
            df_ext_indexed = df_ext.set_index('placa_key').to_dict('index')

            for idx, row in df_hist.iterrows():
                placa = str(row[COLUNA_PLACA]).strip().upper()

                if placa in df_ext_indexed:
                    dados_novos = df_ext_indexed[placa]

                    # 1. Preenche colunas (incluindo a Cidade agora)
                    for col_origem, col_destino in mapa_colunas_final.items():
                        if col_destino not in df_hist.columns: df_hist[col_destino] = None

                        valor_atual = str(row.get(col_destino, '')).strip()
                        valor_novo = str(dados_novos.get(col_origem, '')).strip()

                        # Se histórico estiver vazio, preenche com o da Rede (que você garantiu estar certo)
                        if valor_atual in ['nan', 'None', '', 'NaT'] and valor_novo not in ['nan', 'None', '']:
                            df_hist.at[idx, col_destino] = valor_novo

                    # 2. Cálculo JPR
                    transp_raw = str(dados_novos.get('transp_raw', '')).strip()
                    valor_guincho_orig = dados_novos.get('valor_base_db', 0)

                    if 'JPR' not in transp_raw.upper():
                        df_hist.at[idx, 'Valor_Base_Guincho2'] = valor_guincho_orig
                    else:
                        cidade_raw = str(dados_novos.get('cidade_raw', '')).strip()
                        patio_raw = str(dados_novos.get('patio_raw', '')).strip()
                        cat_raw = str(dados_novos.get('Categoria_Ext', row.get('Categoria', 'Leve'))).strip()

                        patio_limpo = limpar_texto_estilo_excel(patio_raw)
                        patio_oficial = dict_patio.get(patio_limpo, patio_limpo)

                        val_calculado = calcular_valor_restituicao_final(
                            transp_raw, cidade_raw, patio_oficial, cat_raw, valor_guincho_orig, tabela_jpr
                        )

                        if val_calculado != "Não encontrada":
                            df_hist.at[idx, 'Valor_Base_Guincho2'] = val_calculado

        df_hist = sincronizar_dados_dinamicos_local(df_hist)

        # Garante a ordem e existência das colunas
        cols_prioridade = [
            COLUNA_PLACA, 'Status atual', 'Transportadora', 'Pátio', 'Cidade convertida',
            'Marca', 'Modelo', 'Categoria', 'Data de Remoção',
            'Valor_Base_Guincho2', 'km_remocao', 'valor_rem', 'km_restituicao', 'valor_rest',
            'CPF_Banco', 'Financiado_Banco', 'Contrato_Externo', 'Valor_Base_Guincho'
        ]

        for c in cols_prioridade:
            if c not in df_hist.columns: df_hist[c] = None

        restante = [c for c in df_hist.columns if c not in cols_prioridade]
        df_hist = df_hist[cols_prioridade + restante]

        df_hist.to_excel(NOME_ARQUIVO_HISTORICO, index=False)
        logging.info("Histórico retroativo atualizado (Cidades incluídas).")

    except Exception as e:
        logging.error(f"Erro retrofit histórico: {e}", exc_info=True)

def configurar_driver(headless=True):
    chrome_options = Options()

    if headless:
        # "new" é muito mais estável e evita bugs de detecção de elemento
        chrome_options.add_argument("--headless=new")

        # Argumentos para evitar travamento e erro de memória
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--log-level=3") # Silencia erros inúteis do console

    app_state = {
        "recentDestinations": [{"id": "Save as PDF", "origin": "local", "account": ""}],
        "selectedDestinationId": "Save as PDF",
        "version": 2
    }
    chrome_options.add_experimental_option("prefs", {
        "printing.print_preview_sticky_settings.appState": json.dumps(app_state),
        "savefile.default_directory": PASTA_DOWNLOADS
    })

    try:
        return webdriver.Chrome(options=chrome_options)
    except Exception as e:
        logging.critical(f"Falha ao iniciar o Selenium: {e}", exc_info=True)
        return None

def extrair_km_do_mapa(driver):
    # Seu XPath exato
    meu_xpath = "/html/body/div[1]/div[2]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[5]/div[1]/div[1]/div/div[1]/div[2]/div"

    # Tenta esperar o elemento específico aparecer
    try:
        # Espera até 8 segundos para o elemento exato aparecer
        element = WebDriverWait(driver, 8).until(
            ec.presence_of_element_located((By.XPATH, meu_xpath))
        )
        texto = element.text.strip()

        if any(char.isdigit() for char in texto) and "km" in texto.lower():
            match = re.search(r"([\d\.]+)(?:,(\d+))?", texto)
            if match:
                base = match.group(1).replace('.', '')
                decimal = match.group(2) if match.group(2) else "0"
                km_final = float(f"{base}.{decimal}")

                logging.info(f"KM encontrado (Seu XPath): {km_final}")
                return km_final, str(int(km_final))

    except:
        # Se falhar o seu XPath, tenta uma variação genérica segura (apenas div com classe e texto km)
        # Isso é um fallback caso o Google mude div[9] para div[10]
        try:
            fallback = driver.find_element(By.XPATH, "//div[contains(text(), 'km')]")
            if "min" not in fallback.text: # Garante que não é o tempo
                texto = fallback.text.strip()
                match = re.search(r"([\d\.]+)(?:,(\d+))?", texto)
                if match:
                    base = match.group(1).replace('.', '')
                    decimal = match.group(2) if match.group(2) else "0"
                    km_final = float(f"{base}.{decimal}")
                    logging.info(f"KM encontrado (Fallback): {km_final}")
                    return km_final, str(int(km_final))
        except: pass

    return None, "KM_NAO_ENCONTRADO"

def get_valor_por_range(categoria, km_numerico):
    if km_numerico is None: return "VALOR_PENDENTE"
    categoria_limpa = str(categoria).strip().lower()
    ranges_da_categoria = VALOR_RANGES.get(categoria_limpa, [])
    for limite_km, valor in ranges_da_categoria:
        if km_numerico <= limite_km: return valor
    logging.warning(f"Valor não encontrado para Categoria: {categoria_limpa}, KM: {km_numerico}")
    return "VALOR_NAO_ENCONTRADO"

def gerar_pdf_mapa(driver, nome_arquivo_pdf):
    try:
        WebDriverWait(driver, 20).until(
            ec.visibility_of_element_located((By.TAG_NAME, SELECTORS["google_maps"]["canvas_map"]))
        )
        time.sleep(1)
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "landscape": False,
            "displayHeaderFooter": True,
            "printBackground": True,
            "marginTop": 1, "marginBottom": 1, "marginLeft": 0.5, "marginRight": 0.5
        })
        caminho_completo = os.path.join(PASTA_DOWNLOADS, nome_arquivo_pdf)
        with open(caminho_completo, "wb") as f:
            f.write(base64.b64decode(result['data']))
        return caminho_completo
    except Exception as e:
        logging.error(f"ERRO ao gerar PDF: {e}", exc_info=True)
        return None

def fazer_login_banco(driver):
    try:
        driver.get(URL_BANCO)
        WebDriverWait(driver, 20).until(ec.presence_of_element_located((By.XPATH, SELECTORS["login"]["usuario"])))
        driver.find_element(By.XPATH, SELECTORS["login"]["usuario"]).send_keys(USUARIO_BANCO)
        driver.find_element(By.XPATH, SELECTORS["login"]["senha"]).send_keys(SENHA_BANCO)
        driver.find_element(By.XPATH, SELECTORS["login"]["botao"]).click()
        WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_1"])))
        return True
    except Exception as e:
        logging.error(f"Erro Login: {e}", exc_info=True)
        return False

def navegar_menu_gca(driver):
    try:
        el1 = WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_1"])))
        el1.click()
        el2 = WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_2"])))
        el2.click()
        el3 = WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_3"])))
        el3.click()
        return True
    except Exception as e:
        logging.error(f"Erro Navegação GCA: {e}", exc_info=True)
        return False

def preencher_formulario_com_upload(driver, dados_upload, texto_anterior_ignorar=None):
    try:
        WebDriverWait(driver, 20).until(ec.element_to_be_clickable((By.XPATH, SELECTORS["form_upload"]["select_status"])))

        upload_element = driver.find_element(By.ID, SELECTORS["form_upload"]["input_arquivo"])
        caminho_pdf = dados_upload['caminho_pdf']

        driver.execute_script(
            "arguments[0].style.display = 'block'; arguments[0].style.visibility = 'visible'; arguments[0].style.opacity = 1; arguments[0].style.height = '1px'; arguments[0].style.width = '1px';",
            upload_element
        )
        time.sleep(1)
        upload_element.send_keys(caminho_pdf)
        driver.execute_script("validate(arguments[0]);", upload_element)
        time.sleep(1)

        Select(driver.find_element(By.XPATH, SELECTORS["form_upload"]["select_status"])).select_by_visible_text("Cadastrar")

        data_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_data"])
        data_field.clear()
        data_field.send_keys(dados_upload['data'])

        contrato_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_contrato"])
        contrato_field.clear()
        contrato_field.send_keys(dados_upload['contrato'])

        placa_sem_hifen = dados_upload['placa'].replace('-', '')
        placa_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_placa"])
        placa_field.clear()
        placa_field.send_keys(placa_sem_hifen)

        Select(driver.find_element(By.XPATH, SELECTORS["form_upload"]["select_tipo_despesa"])).select_by_visible_text("018 - GUINCHO")

        valor_formatado = f"{dados_upload['valor']},00"
        valor_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_valor"])
        valor_field.clear()
        valor_field.send_keys(valor_formatado)

        caixa_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_caixa_arquivo"])
        caixa_field.clear()
        caixa_field.send_keys("0")

        obs_field = driver.find_element(By.XPATH, SELECTORS["form_upload"]["input_observacao"])
        obs_field.clear()
        obs_field.send_keys(dados_upload['tipo_str'])

        time.sleep(1)
        driver.find_element(By.XPATH, SELECTORS["form_upload"]["botao_salvar"]).click()

        def mensagem_nova_apareceu(d):
            try:
                el = d.find_element(By.XPATH, SELECTORS["form_upload"]["mensagem_sucesso"])
                txt = el.text.strip()
                if "criado" in txt.lower():
                    if texto_anterior_ignorar and txt == texto_anterior_ignorar:
                        return False
                    return txt
                return False
            except:
                return False

        texto_sucesso = WebDriverWait(driver, 30).until(mensagem_nova_apareceu)
        logging.info(f"Sucesso Banco ({dados_upload['placa']}): {texto_sucesso}")
        return True, texto_sucesso

    except Exception as e:
        logging.error(f"Erro Upload ({dados_upload['placa']}): {e}", exc_info=True)
        return False, None

def enviar_resumo_telegram(lista_sucesso, lista_falha):
    logging.info("Enviando resumo Telegram...")
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return
    try:
        mensagem = ["--- 🤖 Resumo da Automação (Híbrida + Lote) ---"]
        if lista_sucesso:
            mensagem.append("\n✅ SUCESSOS:")
            total_reembolsado = 0
            for item in lista_sucesso:
                mensagem.append(f"\nPlaca: {item[COLUNA_PLACA]}")
                for k in ['valor_rem', 'valor_rest']:
                    try:
                        val = int(item.get(k, 0))
                        if val > 0:
                            total_reembolsado += val
                            tipo = "Remoção" if k == 'valor_rem' else "Restituição"
                            mensagem.append(f"  • {tipo}: R$ {val},00")
                    except: pass

                # Adiciona info extra no Telegram se for JPR
                if 'JPR' in str(item.get('Transportadora', '')).upper():
                    mensagem.append(f"  • (JPR Tabela): R$ {item.get('Valor_Base_Guincho2')}")


            mensagem.append("\n-----------------------------------")
            mensagem.append(f"💰 Total Reembolsado: R$ {total_reembolsado},00")

        if lista_falha:
            mensagem.append("\n\n❌ FALHAS:")
            for item in lista_falha:
                mensagem.append(f"  • {item.get('placa', '?')}: {item.get('motivo', 'Erro')}")

        async def enviar_async(tk, cid, texto):
            bot = telegram.Bot(token=tk)
            await bot.send_message(chat_id=cid, text=texto)
        asyncio.run(enviar_async(token, chat_id, "\n".join(mensagem)))
    except Exception as e:
        logging.error(f"Falha Telegram: {e}")

def processar_mapa_single_instance(driver, placa, contrato, categoria, url_mapa, tipo_acao, data_hoje):
    try:
        driver.get(url_mapa)
        km_num, km_str = extrair_km_do_mapa(driver)
        valor = get_valor_por_range(categoria, km_num)
        if valor == "VALOR_NAO_ENCONTRADO":
            return False, None, None, f"KM {km_num} fora do range"

        suffix = "REMO" if tipo_acao == "Remocao" else "REST"
        nome_arquivo = f"{placa}_{contrato}_{data_hoje}_{km_str}_{valor}_{suffix}.pdf"

        caminho_pdf = gerar_pdf_mapa(driver, nome_arquivo)

        if not caminho_pdf:
            return False, None, None, "Falha PDF"

        return True, km_str, valor, caminho_pdf
    except Exception as e:
        return False, None, None, str(e)

def executar_lote_banco(lote_dados):
    driver = None
    resultados_lote = []
    try:
        driver = configurar_driver(headless=True)
        if not driver: raise Exception("Falha init driver Banco")

        if not fazer_login_banco(driver): raise Exception("Falha Login")
        if not navegar_menu_gca(driver): raise Exception("Falha Menu GCA")

        WebDriverWait(driver, 10).until(ec.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["externo"])))
        WebDriverWait(driver, 10).until(ec.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["interno"])))

        ultimo_texto_sucesso = None

        for dados in lote_dados:
            logging.info(f"Iniciando upload no lote: {dados['placa']} ({dados['tipo_str']})")
            sucesso, texto_msg = preencher_formulario_com_upload(driver, dados, ultimo_texto_sucesso)

            if sucesso:
                resultados_lote.append((True, None, dados))
                ultimo_texto_sucesso = texto_msg
            else:
                resultados_lote.append((False, "Falha no preenchimento/salvamento", dados))

        return resultados_lote

    except Exception as e:
        msg_fatal = str(e)
        for dados in lote_dados:
            if not any(r[2] == dados for r in resultados_lote):
                resultados_lote.append((False, f"Erro Fatal Sessão: {msg_fatal}", dados))
        return resultados_lote
    finally:
        if driver: driver.quit()

def reprocessar_itens_pendentes_historico():
    logging.info("--- Iniciando Reprocessamento de Itens em Branco/Pendentes ---")

    if not os.path.exists(NOME_ARQUIVO_HISTORICO):
        return

    try:
        df_hist = pd.read_excel(NOME_ARQUIVO_HISTORICO)
        dict_transp, dict_patio = carregar_bases_de_enderecos()
    except: return

    # Converte colunas para texto para evitar erros
    for c in ['km_remocao', 'valor_rem', 'km_restituicao', 'valor_rest']:
        if c in df_hist.columns: df_hist[c] = df_hist[c].astype('object')

    driver = None
    alteracoes = False
    dt_hoje = datetime.date.today().strftime("%d-%m-%Y")

    def eh_pendente(valor):
        if pd.isna(valor): return True
        s = str(valor).strip().upper()
        if s in ["", "NONE", "NAN", "NAT", "KM_NAO_ENCONTRADO", "VALOR_PENDENTE", "ERRO_MAPS", "VALOR_NAO_ENCONTRADO"]:
            return True
        try: return float(s.replace(',', '.')) <= 0.1
        except: return True

    indices = [i for i, r in df_hist.iterrows() if (eh_pendente(r.get('km_remocao')) or eh_pendente(r.get('km_restituicao'))) and r.get(COLUNA_PLACA)]

    if not indices:
        logging.info("Nenhum item pendente encontrado.")
        return

    logging.info(f"Reprocessando {len(indices)} itens pendentes...")

    try:
        driver = configurar_driver(headless=True)

        for idx in indices:
            row = df_hist.loc[idx]
            placa = row[COLUNA_PLACA]

            contrato = row.get('Contrato_Externo') or row.get(COLUNA_CONTRATO) or "S_CONTRATO"
            cat = row.get('Categoria') or "Leve"
            transp = str(row.get('Transportadora', '')).strip().upper()

            patio_raw = str(row.get('Pátio', ''))
            patio_limpo = limpar_texto_estilo_excel(patio_raw)
            patio_oficial = dict_patio.get(patio_limpo, patio_limpo)

            end1 = dict_transp.get(transp, "")
            end2 = dict_patio.get(patio_limpo, "")
            if not end2 or len(end2) < 5: end2 = dict_patio.get(patio_oficial, "")

            # Tenta pegar a cidade do histórico (que agora está preenchida)
            end3 = str(row.get('Cidade convertida', row.get('CidadeOrigem', ''))).strip()
            if len(end3) < 3 or end3 == 'nan': end3 = ""

            if len(end1) < 5:
                logging.warning(f"Pular {placa}: Sem endereço transp ({transp})")
                continue

            base_maps = "https://www.google.com/maps/dir"

            # --- 1. REPROCESSAR REMOÇÃO ---
            if eh_pendente(row.get('km_remocao')) or eh_pendente(row.get('valor_rem')):
                logging.info(f"Refazendo Remoção: {placa}")
                # Transp -> Pátio -> Cidade -> Transp
                if end3:
                    url = f"{base_maps}/{end1}/{end2}/{end3}/{end1}"
                else:
                    url = f"{base_maps}/{end1}/{end2}/{end1}"

                url = url.replace(" ", "+")
                ok, km, val, pdf = processar_mapa_single_instance(driver, placa, contrato, cat, url, "Remocao", dt_hoje)

                if ok:
                    df_hist.at[idx, 'km_remocao'] = km
                    df_hist.at[idx, 'valor_rem'] = val
                    alteracoes = True
                    logging.info(f" > Remoção OK: {km} km | R$ {val}")
                else:
                    logging.warning(f" > Falha Remoção: {pdf}")

            # --- 2. REPROCESSAR RESTITUIÇÃO ---
            if eh_pendente(row.get('km_restituicao')) or eh_pendente(row.get('valor_rest')):
                logging.info(f"Refazendo Restituição: {placa}")

                # Se tem cidade, faz a rota completa inversa: Transp -> Cidade -> Transp -> Pátio
                if len(end3) > 2:
                    url = f"{base_maps}/{end1}/{end3}/{end1}/{end2}".replace(" ", "+")
                else:
                    # Fallback: Transp -> Pátio -> Transp (Melhor que nada)
                    logging.info(" > S/ Cidade. Usando rota fallback.")
                    url = f"{base_maps}/{end1}/{end2}/{end1}".replace(" ", "+")

                ok, km, val, pdf = processar_mapa_single_instance(driver, placa, contrato, cat, url, "Restituicao", dt_hoje)

                if ok:
                    df_hist.at[idx, 'km_restituicao'] = km
                    df_hist.at[idx, 'valor_rest'] = val
                    alteracoes = True
                    logging.info(f" > Restituição OK: {km} km | R$ {val}")
                else:
                    logging.warning(f" > Falha Restituição: {pdf}")

    except Exception as e:
        logging.error(f"Erro reprocessamento: {e}")

    finally:
        if driver: driver.quit()
        if alteracoes:
            try:
                df_hist.to_excel(NOME_ARQUIVO_HISTORICO, index=False)
                logging.info("Histórico atualizado e salvo com sucesso!")
            except Exception as e:
                logging.error(f"Erro ao salvar Excel: {e}")

def iniciar_automacao_completa():
    configurar_logger_dinamico()
    logging.info("--- Iniciando Automação Completa (Python + Rede Q:) ---")

    # 1. Carrega todas as bases (Endereços, Rede e agora JPR)
    dict_transp, dict_patio = carregar_bases_de_enderecos()
    df_ext = carregar_base_externa_rede()
    tabela_jpr = carregar_tabela_custos_jpr()

    # ... (código anterior de carregar bases e atualizar histórico retroativo) ...

    atualizar_historico_existente(df_ext, tabela_jpr, dict_patio)

    # --- NOVO: Reprocessa o que ficou em branco no histórico ---
    reprocessar_itens_pendentes_historico()
    # -----------------------------------------------------------

    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA_CALCULOS)
    # ... (restante do código continua igual)
        for c in [COLUNA_PLACA, COLUNA_CATEGORIA, COLUNA_CONTRATO]: df[c] = df[c].astype(str).str.strip()
        df[COLUNA_TESTE] = pd.to_numeric(df[COLUNA_TESTE], errors='coerce').fillna(0).astype(int)

        if not df_ext.empty:
            logging.info("Cruzando dados com a rede...")
            df = pd.merge(df, df_ext, left_on=COLUNA_PLACA, right_on='placa_key', how='left')

        for i, row in df.iterrows():
            transp_nome = str(row.get('transp_raw', '')).strip().upper()
            df.at[i, COLUNA_END1] = dict_transp.get(transp_nome, "")
            patio_nome = limpar_texto_estilo_excel(str(row.get('patio_raw', '')))
            df.at[i, COLUNA_END2] = dict_patio.get(patio_nome, "")
            cidade_nome = str(row.get('cidade_raw', ''))
            df.at[i, COLUNA_END3] = limpar_texto_estilo_excel(cidade_nome) if cidade_nome != 'nan' else ""

            logging.info(f"Processando: {row[COLUNA_PLACA]}")
            logging.info(f"   > Transp: '{transp_nome}' -> '{df.at[i, COLUNA_END1]}'")

    except Exception as e:
        logging.critical(f"Erro Prep Dados: {e}")
        return

    try: hist = pd.read_excel(NOME_ARQUIVO_HISTORICO)[COLUNA_PLACA].astype(str).tolist()
    except: hist = []

    fila = [row for _, row in df.iterrows() if str(row[COLUNA_PLACA]) not in hist]
    if not fila: return logging.info("Nada a processar.")

    logging.info(f"Fase Maps: {len(fila)} placas")
    res_final, uploads = {}, []
    driver = configurar_driver(headless=True)

    if driver:
        dt = datetime.date.today().strftime("%d-%m-%Y")
        for row in fila:
            placa = row[COLUNA_PLACA]
            end1, end2, end3 = row[COLUNA_END1], row[COLUNA_END2], row[COLUNA_END3]

            # Pega o nome da transportadora para usar na lógica do JPR
            transp_atual = str(row.get('transp_raw', '')).strip()

            if placa not in res_final:
                res_final[placa] = {
                    COLUNA_PLACA: placa,
                    'valor_rem': 0, 'km_remocao': 0,
                    'valor_rest': 0, 'km_restituicao': 0,
                    'falhas': [],
                    'CPF_Banco': row.get('cpf_db', ''),
                    'Financiado_Banco': row.get('financiado_db', ''),
                    'Contrato_Externo': row.get('contrato_externo', ''),
                    'Valor_Base_Guincho': row.get('valor_base_db', ''),
                    'Transportadora': transp_atual, # Salva o nome da transportadora
                    'Valor_Base_Guincho2': 0 # Será calculado abaixo
                }

            if not end1 or len(str(end1)) < 3:
                msg_err = "FALHA: Endereço Transportadora VAZIO/NÃO ENCONTRADO na Base."
                logging.error(msg_err)
                res_final[placa]['falhas'].append(msg_err)
                continue

            url_rem = f"https://www.google.com/maps/dir/{end1.replace(' ','+')}/{end2.replace(' ','+')}/{end3.replace(' ','+')}/{end1.replace(' ','+')}"
            url_rest = f"https://www.google.com/maps/dir/{end1.replace(' ','+')}/{end3.replace(' ','+')}/{end1.replace(' ','+')}/{end2.replace(' ','+')}"

            # --- REMOÇÃO ---
            ok, km, val, pdf = processar_mapa_single_instance(driver, placa, row[COLUNA_CONTRATO], row[COLUNA_CATEGORIA], url_rem, "Remocao", dt)
            if ok:
                res_final[placa]['valor_rem'] = val
                res_final[placa]['km_remocao'] = km
                uploads.append({'placa': placa, 'contrato': row[COLUNA_CONTRATO], 'data': dt, 'valor': str(val), 'tipo_str': "Remocao", 'caminho_pdf': pdf})

                # --- CÁLCULO DA RESTITUIÇÃO (BASE GUINCHO 2) ---
                # Pega o valor original da base (ValorGuincheiro) para usar caso não seja JPR
                # Pega valor do banco caso não seja JPR
                valor_guincho_original = row.get('valor_base_db', 0)

                val_restituicao_final = calcular_valor_restituicao_final(
                    transp_atual,
                    row.get('cidade_raw', ''),
                    row.get('patio_raw', ''),
                    row[COLUNA_CATEGORIA],
                    valor_guincho_original, # <--- CORRETO (Usa valor da base)
                    tabela_jpr
                )

                res_final[placa]['Valor_Base_Guincho2'] = val_restituicao_final

            else:
                logging.error(f"FALHA MAPS REMOÇÃO ({placa}): {pdf}")
                res_final[placa]['falhas'].append(f"Maps Remo: {pdf}")

            # --- RESTITUIÇÃO (MAPS) ---
            # Mantemos a lógica de rodar o mapa da restituição para pegar o KM e gerar o PDF
            if row[COLUNA_TESTE] == 1 and ok:
                if not end3: res_final[placa]['falhas'].append("Sem Cidade Destino")
                else:
                    ok2, km2, val2, pdf2 = processar_mapa_single_instance(driver, placa, row[COLUNA_CONTRATO], row[COLUNA_CATEGORIA], url_rest, "Restituicao", dt)
                    if ok2:
                        res_final[placa]['valor_rest'] = val2
                        res_final[placa]['km_restituicao'] = km2
                        uploads.append({'placa': placa, 'contrato': row[COLUNA_CONTRATO], 'data': dt, 'valor': str(val2), 'tipo_str': "Restituicao", 'caminho_pdf': pdf2})
                    else:
                        logging.error(f"FALHA MAPS RESTITUIÇÃO ({placa}): {pdf2}")
                        res_final[placa]['falhas'].append(f"Maps Rest: {pdf2}")
        driver.quit()

    logging.info(f"Fase Banco: {len(uploads)} uploads")

    if uploads:
        chunk_size = max(1, math.ceil(len(uploads)/5))
        lotes = [uploads[i:i + chunk_size] for i in range(0, len(uploads), chunk_size)]

        with ThreadPoolExecutor(max_workers=5) as executor:
            futures = {
                executor.submit(executar_lote_banco, lote): lote
                for lote in lotes
            }
            for future in as_completed(futures):
                resultados_lote = future.result()
                for sucesso, erro_msg, dados_orig in resultados_lote:
                    placa_atual = dados_orig['placa']
                    tipo_atual = dados_orig['tipo_str']
                    if sucesso:
                        logging.info(f"Upload OK: {placa_atual} ({tipo_atual})")
                    else:
                        logging.error(f"Falha Upload {placa_atual}: {erro_msg}")
                        res_final[placa_atual]['falhas'].append(f"Banco {tipo_atual}: {erro_msg}")
    else:
        logging.info("Nenhum upload gerado (Erros de Maps ou Endereço acima).")

    sucessos = [d for d in res_final.values() if not d['falhas']]
    falhas = [{'placa': k, 'motivo': v['falhas']} for k, v in res_final.items() if v['falhas']]

    if sucessos:
        try:
            df_novo = pd.DataFrame(sucessos)
            try: df_final = pd.concat([pd.read_excel(NOME_ARQUIVO_HISTORICO), df_novo], ignore_index=True)
            except FileNotFoundError: df_final = df_novo

            # Adicionadas Transportadora e Valor_Base_Guincho2 na prioridade
            cols_prio = [COLUNA_PLACA, 'Transportadora', 'Valor_Base_Guincho2', 'km_remocao', 'valor_rem', 'km_restituicao', 'valor_rest', 'CPF_Banco', 'Financiado_Banco', 'Contrato_Externo', 'Valor_Base_Guincho']
            df_final = df_final[cols_prio + [c for c in df_final.columns if c not in cols_prio]]

            df_final.drop_duplicates(subset=[COLUNA_PLACA], keep='last').to_excel(NOME_ARQUIVO_HISTORICO, index=False)
            logging.info("Histórico Atualizado!")
        except Exception as e: logging.error(f"Erro Salvar Excel: {e}")

    enviar_resumo_telegram(sucessos, falhas)
    logging.info("FIM")

if __name__ == "__main__":
    iniciar_automacao_completa()