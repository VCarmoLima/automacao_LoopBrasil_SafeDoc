import asyncio
import base64
import datetime
import glob
import json
import logging
import os
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

import pandas as pd
import telegram
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import Select
from selenium.webdriver.support.ui import WebDriverWait

load_dotenv()

NOME_ARQUIVO_EXCEL = "Base Restituições Vinícius.xlsm"
NOME_ABA = "Calculos"
PASTA_DOWNLOADS = os.getenv("PASTA_DOWNLOADS")
NOME_ARQUIVO_HISTORICO = "historico_processamento.xlsx"

COLUNA_PLACA = "Placa"
COLUNA_END1 = "Endereço transportadora"
COLUNA_END2 = "Endereço Pátio"
COLUNA_END3 = "Cidade convertida"
COLUNA_TESTE = "Teste"
COLUNA_STATUS_SAFE_DOC = "STATUS SAFE DOC"
COLUNA_CONTRATO = "Contrato"
COLUNA_CATEGORIA = "Categoria"

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
        "km_xpath": "/html/body/div[1]/div[3]/div[9]/div[9]/div/div/div[1]/div[2]/div/div[1]/div/div/div[5]/div[1]/div[1]/div/div[1]/div[2]/div",
        "canvas_map": "canvas"
    },
    "login": {
        "usuario": "/html/body/div/main/div/div[2]/form/div/div[2]/div/input",
        "senha": "/html/body/div/main/div/div[2]/form/div/div[3]/div/input",
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
        "botao_salvar": "/html/body/form/div/div/div/div/div[3]/input"
    }
}

def configurar_logger_dinamico():
    try:
        diretorio_script = os.path.dirname(os.path.abspath(__file__))
    except NameError:
        diretorio_script = os.getcwd()

    pasta_logs = os.path.join(diretorio_script, "logs")
    os.makedirs(pasta_logs, exist_ok=True)

    hoje_str = datetime.date.today().strftime("%Y-%m-%d")
    padrao_busca = os.path.join(pasta_logs, f"log_{hoje_str}_v*.txt")
    arquivos_existentes = glob.glob(padrao_busca)

    maior_versao = 0
    for arquivo_path in arquivos_existentes:
        try:
            nome_arquivo = os.path.basename(arquivo_path)
            nome_sem_ext = os.path.splitext(nome_arquivo)[0]
            numero_str = nome_sem_ext.split("_v")[-1]
            numero = int(numero_str)
            if numero > maior_versao:
                maior_versao = numero
        except (ValueError, IndexError):
            continue

    proxima_versao = maior_versao + 1
    nome_log = os.path.join(pasta_logs, f"log_{hoje_str}_v{proxima_versao}.txt")

    for handler in logging.root.handlers[:]:
        logging.root.removeHandler(handler)

    logging.basicConfig(
        filename=nome_log,
        filemode='w',
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    print(f"--- Log atual: {os.path.basename(nome_log)} ---")
    logging.info(f"Iniciando log versão: v{proxima_versao}")

def configurar_driver(headless=True):
    chrome_options = Options()
    if headless:
        chrome_options.add_argument("--headless")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--window-size=1920,1080")

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
    try:
        xpath_km = SELECTORS["google_maps"]["km_xpath"]
        WebDriverWait(driver, 10).until(ec.presence_of_element_located((By.XPATH, xpath_km)))
        km_element = driver.find_element(By.XPATH, xpath_km)
        km_bruto_texto = km_element.text
        km_bruto_num_str = km_bruto_texto.split(' ')[0]
        km_sem_milhar = km_bruto_num_str.replace('.', '')
        km_para_float = km_sem_milhar.replace(',', '.')
        km_num = float(km_para_float)
        km_str_arquivo = str(int(km_num))
        return km_num, km_str_arquivo
    except Exception as e:
        logging.warning(f"Não foi possível extrair o KM: {e}")
        return None, "KM_NAO_ENCONTRADO"

def get_valor_por_range(categoria, km_numerico):
    if km_numerico is None:
        return "VALOR_PENDENTE"
    categoria_limpa = categoria.strip().lower()
    ranges_da_categoria = VALOR_RANGES.get(categoria_limpa, [])
    for limite_km, valor in ranges_da_categoria:
        if km_numerico <= limite_km:
            return valor
    logging.warning(f"Valor não encontrado para Categoria: {categoria_limpa}, KM: {km_numerico}")
    return "VALOR_NAO_ENCONTRADO"

def gerar_pdf_mapa(driver, nome_arquivo_pdf):
    try:
        WebDriverWait(driver, 20).until(
            ec.visibility_of_element_located((By.TAG_NAME, SELECTORS["google_maps"]["canvas_map"]))
        )
        time.sleep(1) # Necessário para renderização do Canvas
        result = driver.execute_cdp_cmd("Page.printToPDF", {
            "landscape": False, "printBackground": True, "displayHeaderFooter": True,
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

def preencher_formulario_com_upload(driver, dados_upload):
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
        time.sleep(2)

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
        time.sleep(4) # Necessário para processamento do site do banco
        return True
    except Exception as e:
        logging.error(f"Erro Upload ({dados_upload['placa']}): {e}", exc_info=True)
        return False

def enviar_resumo_telegram(lista_sucesso, lista_falha):
    logging.info("Enviando resumo Telegram...")
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    chat_id = os.getenv("TELEGRAM_CHAT_ID")
    if not token or not chat_id:
        return
    try:
        mensagem = ["--- 🤖 Resumo da Automação (Híbrida) ---"]
        if lista_sucesso:
            mensagem.append("\n✅ SUCESSOS:")
            total_reembolsado = 0
            for item in lista_sucesso:
                mensagem.append(f"\nPlaca: {item[COLUNA_PLACA]}")
                for k in ['valor_rem', 'valor_rest']:
                    try:
                        val = int(item[k])
                        if val > 0:
                            total_reembolsado += val
                            tipo = "Remoção" if k == 'valor_rem' else "Restituição"
                            mensagem.append(f"  • {tipo}: R$ {val},00")
                    except: pass

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
    """Usa o driver compartilhado para processar um mapa (Fase 1)."""
    try:
        driver.get(url_mapa)
        km_num, km_str = extrair_km_do_mapa(driver)
        valor = get_valor_por_range(categoria, km_num)

        suffix = "REMO" if tipo_acao == "Remocao" else "REST"
        nome_arquivo = f"{placa}_{contrato}_{data_hoje}_{km_str}_{valor}_{suffix}.pdf"

        caminho_pdf = gerar_pdf_mapa(driver, nome_arquivo)

        if not caminho_pdf:
            return False, None, None, "Falha PDF"

        return True, km_str, valor, caminho_pdf
    except Exception as e:
        return False, None, None, str(e)

def executar_apenas_upload_banco(dados_prontos):
    """Abre driver exclusivo, faz login e upload (Fase 2)."""
    driver = None
    try:
        driver = configurar_driver(headless=True)
        if not driver: raise Exception("Falha init driver Banco")

        if not fazer_login_banco(driver): raise Exception("Falha Login")
        if not navegar_menu_gca(driver): raise Exception("Falha Menu GCA")

        WebDriverWait(driver, 10).until(ec.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["externo"])))
        WebDriverWait(driver, 10).until(ec.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["interno"])))

        if not preencher_formulario_com_upload(driver, dados_prontos):
            raise Exception("Falha Preenchimento/Salvar")

        return True, None
    except Exception as e:
        return False, str(e)
    finally:
        if driver: driver.quit()

def iniciar_automacao_completa():
    configurar_logger_dinamico()
    logging.info("--- Iniciando Automação Híbrida (Maps Único + Banco Paralelo) ---")

    lista_placas_log = []
    try:
        df_hist = pd.read_excel(NOME_ARQUIVO_HISTORICO)
        if not df_hist.empty: lista_placas_log = df_hist[COLUNA_PLACA].astype(str).tolist()
    except: pass

    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA)
        df[COLUNA_PLACA] = df[COLUNA_PLACA].astype(str)
        df[COLUNA_CONTRATO] = df[COLUNA_CONTRATO].astype(str)
        df[COLUNA_CATEGORIA] = df[COLUNA_CATEGORIA].astype(str)
        df[COLUNA_STATUS_SAFE_DOC] = df[COLUNA_STATUS_SAFE_DOC].astype(str)
    except Exception as e:
        logging.critical(f"Erro Excel: {e}")
        return

    data_hoje = datetime.date.today().strftime("%d-%m-%Y")
    tarefas_upload = []
    resultados_finais = {}

    registros_para_processar = []
    for idx, row in df.iterrows():
        placa = str(row[COLUNA_PLACA]).strip()
        if placa in lista_placas_log: continue
        registros_para_processar.append(row)

    if not registros_para_processar:
        logging.info("Nada a processar.")
        return

    # FASE 1: GOOGLE MAPS (SERIAL)
    logging.info(f"--- FASE 1: Maps para {len(registros_para_processar)} placas (Instância Única) ---")
    driver_maps = configurar_driver(headless=True)

    if driver_maps:
        for row in registros_para_processar:
            placa = str(row[COLUNA_PLACA]).strip()
            contrato = str(row[COLUNA_CONTRATO]).strip()
            categoria = str(row[COLUNA_CATEGORIA]).strip()
            status = str(row[COLUNA_STATUS_SAFE_DOC]).strip()
            teste = row.get(COLUNA_TESTE, 0)
            end1 = str(row[COLUNA_END1]).replace(" ", "+")
            end2 = str(row[COLUNA_END2]).replace(" ", "+")
            end3 = str(row[COLUNA_END3]).replace(" ", "+")

            url_remocao = f"https://www.google.com/maps/dir/{end1}/{end2}/{end3}/{end1}"
            url_restituicao = f"https://www.google.com/maps/dir/{end1}/{end3}/{end1}/{end2}"

            run_rem = status == "Pendente remoção" or teste == 1 or (status != "Pendente restituição")
            run_rest = status == "Pendente restituição" or teste == 1

            if placa not in resultados_finais:
                resultados_finais[placa] = {
                    COLUNA_PLACA: placa, 'valor_rem': 0, 'km_remocao': 0,
                    'valor_rest': 0, 'km_restituicao': 0, 'falhas': []
                }

            sucesso_rem = True

            if run_rem:
                ok, km, val, pdf = processar_mapa_single_instance(driver_maps, placa, contrato, categoria, url_remocao, "Remocao", data_hoje)
                if ok:
                    resultados_finais[placa]['km_remocao'] = km
                    resultados_finais[placa]['valor_rem'] = val
                    tarefas_upload.append({
                        'placa': placa, 'contrato': contrato, 'data': data_hoje,
                        'valor': str(val), 'tipo_str': "Remocao", 'caminho_pdf': pdf
                    })
                else:
                    sucesso_rem = False
                    resultados_finais[placa]['falhas'].append(f"Maps Remoção: {pdf}")

            if run_rest and sucesso_rem:
                ok, km, val, pdf = processar_mapa_single_instance(driver_maps, placa, contrato, categoria, url_restituicao, "Restituicao", data_hoje)
                if ok:
                    resultados_finais[placa]['km_restituicao'] = km
                    resultados_finais[placa]['valor_rest'] = val
                    tarefas_upload.append({
                        'placa': placa, 'contrato': contrato, 'data': data_hoje,
                        'valor': str(val), 'tipo_str': "Restituicao", 'caminho_pdf': pdf
                    })
                else:
                    resultados_finais[placa]['falhas'].append(f"Maps Restituição: {pdf}")

        driver_maps.quit()
        logging.info("--- FASE 1 Concluída ---")
    else:
        logging.critical("Não foi possível abrir driver do Maps.")
        return

    # FASE 2: BANCO (PARALELO)
    logging.info(f"--- FASE 2: Uploads no Banco ({len(tarefas_upload)} itens) ---")
    QTD_WORKERS = 5

    with ThreadPoolExecutor(max_workers=QTD_WORKERS) as executor:
        futures = {
            executor.submit(executar_apenas_upload_banco, dados): dados
            for dados in tarefas_upload
        }

        for future in as_completed(futures):
            dados_orig = futures[future]
            placa_atual = dados_orig['placa']
            tipo_atual = dados_orig['tipo_str']
            try:
                sucesso, erro_msg = future.result()
                if sucesso:
                    logging.info(f"Upload OK: {placa_atual} ({tipo_atual})")
                else:
                    logging.error(f"Falha Upload {placa_atual}: {erro_msg}")
                    resultados_finais[placa_atual]['falhas'].append(f"Banco {tipo_atual}: {erro_msg}")
            except Exception as e:
                logging.error(f"Erro Thread Banco: {e}")
                resultados_finais[placa_atual]['falhas'].append(f"Crash Thread {tipo_atual}")

    lista_sucessos_final = []
    lista_falhas_final = []

    for placa, dados in resultados_finais.items():
        if dados['falhas']:
            for f in dados['falhas']:
                lista_falhas_final.append({'placa': placa, 'motivo': f})
        else:
            lista_sucessos_final.append(dados)

    if lista_sucessos_final:
        try:
            df_novas = pd.DataFrame(lista_sucessos_final)
            try:
                df_antigo = pd.read_excel(NOME_ARQUIVO_HISTORICO)
                df_final = pd.concat([df_antigo, df_novas], ignore_index=True)
            except FileNotFoundError:
                df_final = df_novas

            df_final.drop_duplicates(subset=[COLUNA_PLACA], keep='last', inplace=True)
            colunas = [COLUNA_PLACA, 'km_remocao', 'valor_rem', 'km_restituicao', 'valor_rest']
            df_final = df_final.reindex(columns=colunas)
            df_final.to_excel(NOME_ARQUIVO_HISTORICO, index=False)
            logging.info("Histórico Atualizado.")
        except Exception as e:
            logging.error(f"Erro ao salvar Excel final: {e}")

    enviar_resumo_telegram(lista_sucessos_final, lista_falhas_final)
    logging.info("--- FIM DO PROCESSO ---")

if __name__ == "__main__":
    iniciar_automacao_completa()