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
from dotenv import load_dotenv
from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select

load_dotenv()

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
        except Exception:
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
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.XPATH, xpath_km)))
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
            EC.visibility_of_element_located((By.TAG_NAME, SELECTORS["google_maps"]["canvas_map"]))
        )
        time.sleep(1)
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

def enviar_resumo_telegram(lista_sucesso, lista_falha):
    logging.info("Enviando resumo Telegram...")
    TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
    CHAT_ID = os.getenv("TELEGRAM_CHAT_ID")
    if not TOKEN or not CHAT_ID:
        return
    try:
        mensagem = ["--- 🤖 Resumo da Automação ---"]
        if lista_sucesso:
            mensagem.append("\n✅ PLACAS PROCESSADAS COM SUCESSO:")
            total_reembolsado = 0
            for item in lista_sucesso:
                mensagem.append(f"\nPlaca: {item[COLUNA_PLACA]}")
                valor_rem_num = 0
                v_rem_raw = item['valor_rem']
                if v_rem_raw not in ["N/A", "VALOR_PENDENTE", "VALOR_NAO_ENCONTRADO", 0, "0"]:
                    try: valor_rem_num = int(v_rem_raw)
                    except (ValueError, TypeError): valor_rem_num = 0

                valor_rest_num = 0
                v_rest_raw = item['valor_rest']
                if v_rest_raw not in ["N/A", "VALOR_PENDENTE", "VALOR_NAO_ENCONTRADO", 0, "0"]:
                    try: valor_rest_num = int(v_rest_raw)
                    except (ValueError, TypeError): valor_rest_num = 0

                if valor_rem_num > 0:
                    mensagem.append(f"  • Remoção: R$ {valor_rem_num},00")
                    total_reembolsado += valor_rem_num
                if valor_rest_num > 0:
                    mensagem.append(f"  • Restituição: R$ {valor_rest_num},00")
                    total_reembolsado += valor_rest_num

            mensagem.append("\n-----------------------------------")
            mensagem.append(f"💰 Total Reembolsado: R$ {total_reembolsado},00")

        if lista_falha:
            mensagem.append("\n\n❌ FALHAS:")
            for item in lista_falha:
                placa_falha = item.get(COLUNA_PLACA, item.get('placa', 'N/A'))
                mensagem.append(f"  • Placa: {placa_falha} (Motivo: {item['motivo']})")

        total_s = len(lista_sucesso)
        total_f = len(lista_falha)
        if not lista_sucesso and not lista_falha:
            mensagem.append("\nNenhuma placa nova.")
        else:
            mensagem.append("\n-----------------------------------")
            mensagem.append(f"Resumo: {total_s} sucesso(s) | {total_f} falha(s).")

        async def enviar_async(token, chat_id, texto):
            bot = telegram.Bot(token=token)
            await bot.send_message(chat_id=chat_id, text=texto)
        asyncio.run(enviar_async(TOKEN, CHAT_ID, "\n".join(mensagem)))
        logging.info("Resumo Telegram enviado.")
    except Exception as e:
        logging.error(f"Falha Telegram: {e}", exc_info=True)

def fazer_login_banco(driver):
    try:
        driver.get(URL_BANCO)
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.XPATH, SELECTORS["login"]["usuario"])))
        driver.find_element(By.XPATH, SELECTORS["login"]["usuario"]).send_keys(USUARIO_BANCO)
        driver.find_element(By.XPATH, SELECTORS["login"]["senha"]).send_keys(SENHA_BANCO)
        driver.find_element(By.XPATH, SELECTORS["login"]["botao"]).click()
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_1"])))
        return True
    except Exception as e:
        logging.error(f"Erro Login: {e}", exc_info=True)
        return False

def navegar_menu_gca(driver):
    try:
        el1 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_1"])))
        el1.click()
        el2 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_2"])))
        el2.click()
        el3 = WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["gca_menu"]["link_3"])))
        el3.click()
        return True
    except Exception as e:
        logging.error(f"Erro Navegação GCA: {e}", exc_info=True)
        return False

def preencher_formulario_com_upload(driver, dados_upload):
    try:
        WebDriverWait(driver, 20).until(EC.element_to_be_clickable((By.XPATH, SELECTORS["form_upload"]["select_status"])))
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
        time.sleep(4)
        return True

    except Exception as e:
        logging.error(f"Erro Upload ({dados_upload['placa']}): {e}", exc_info=True)
        return False

def iniciar_automacao_completa():
    configurar_logger_dinamico()
    logging.info("--- Iniciando Automação Completa ---")

    lista_placas_log = []
    df_historico_antigo = pd.DataFrame()

    try:
        df_historico_antigo = pd.read_excel(NOME_ARQUIVO_HISTORICO)
        if not df_historico_antigo.empty:
            lista_placas_log = df_historico_antigo[COLUNA_PLACA].astype(str).tolist()
    except FileNotFoundError:
        logging.warning("Histórico não encontrado. Será criado.")
    except Exception as e:
        logging.warning(f"Erro leitura histórico: {e}")

    try:
        df = pd.read_excel(NOME_ARQUIVO_EXCEL, sheet_name=NOME_ABA)
    except Exception as e:
        logging.critical(f"Erro planilha principal: {e}", exc_info=True)
        enviar_resumo_telegram([], [{'placa': 'N/A', 'motivo': 'Erro leitura Excel'}])
        return

    df[COLUNA_STATUS_SAFE_DOC] = df[COLUNA_STATUS_SAFE_DOC].astype(str)
    df[COLUNA_CATEGORIA] = df[COLUNA_CATEGORIA].astype(str)

    placas_sucesso_info = []
    placas_falha_info = []
    data_hoje_formatada = datetime.date.today().strftime("%d-%m-%Y")

    for index, linha in df.iterrows():
        placa = "N/A"
        try:
            placa = str(linha[COLUNA_PLACA]).strip()
            if placa in lista_placas_log:
                continue

            contrato = str(linha[COLUNA_CONTRATO]).strip()
            categoria = str(linha[COLUNA_CATEGORIA]).strip()
            end1 = str(linha[COLUNA_END1])
            end2 = str(linha[COLUNA_END2])
            end3 = str(linha[COLUNA_END3])
            status_safe_doc = str(linha[COLUNA_STATUS_SAFE_DOC]).strip()
            teste_val = linha[COLUNA_TESTE]

        except KeyError as e:
            logging.error(f"KeyError Linha {index + 2}: {e}")
            placas_falha_info.append({'placa': f'Linha {index + 2}', 'motivo': f'Coluna faltante: {e}'})
            continue

        end1_url = end1.replace(" ", "+")
        end2_url = end2.replace(" ", "+")
        end3_url = end3.replace(" ", "+")

        logging.info(f"Processando: {placa}")

        url_remocao = f"https://www.google.com/maps/dir/{end1_url}/{end2_url}/{end3_url}/{end1_url}"
        url_restituicao = f"https://www.google.com/maps/dir/{end1_url}/{end3_url}/{end1_url}/{end2_url}"

        run_rem = status_safe_doc == "Pendente remoção" or teste_val == 1 or (status_safe_doc != "Pendente restituição")
        run_rest = status_safe_doc == "Pendente restituição" or teste_val == 1

        sucesso_final_placa = True
        km_str_rem_log = 0
        valor_rem_log = 0
        km_str_rest_log = 0
        valor_rest_log = 0

        if run_rem:
            logging.info(f"Iniciando Remoção: {placa}")
            driver_rem = None
            try:
                driver_rem = configurar_driver(headless=True)
                if not driver_rem: raise Exception("Driver Remoção falhou")

                driver_rem.get(url_remocao)
                km_num_rem, km_str_rem = extrair_km_do_mapa(driver_rem)
                valor_rem = get_valor_por_range(categoria, km_num_rem)

                km_str_rem_log = km_str_rem
                valor_rem_log = valor_rem

                nome_arquivo_rem = f"{placa}_{contrato}_{data_hoje_formatada}_{km_str_rem}_{valor_rem}_REMO.pdf"
                caminho_pdf_rem = gerar_pdf_mapa(driver_rem, nome_arquivo_rem)
                if not caminho_pdf_rem: raise Exception("Falha PDF Remoção")

                if not fazer_login_banco(driver_rem): raise Exception("Login falhou")
                if not navegar_menu_gca(driver_rem): raise Exception("Navegação falhou")

                dados_rem = {
                    "placa": placa, "contrato": contrato, "data": data_hoje_formatada,
                    "valor": str(valor_rem), "tipo_str": "Remocao", "caminho_pdf": caminho_pdf_rem
                }

                WebDriverWait(driver_rem, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["externo"])))
                WebDriverWait(driver_rem, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["interno"])))

                if not preencher_formulario_com_upload(driver_rem, dados_rem): raise Exception("Upload Remoção falhou")
                logging.info(f"Sucesso Remoção: {placa}")

            except Exception as e:
                logging.error(f"Erro Remoção {placa}: {e}", exc_info=True)
                placas_falha_info.append({'placa': placa, 'motivo': f'Falha Remoção: {e}'})
                sucesso_final_placa = False
            finally:
                if driver_rem: driver_rem.quit()

        if run_rest:
            logging.info(f"Iniciando Restituição: {placa}")
            driver_rest = None
            try:
                driver_rest = configurar_driver(headless=True)
                if not driver_rest: raise Exception("Driver Restituição falhou")

                driver_rest.get(url_restituicao)
                km_num_rest, km_str_rest = extrair_km_do_mapa(driver_rest)
                valor_rest = get_valor_por_range(categoria, km_num_rest)

                km_str_rest_log = km_str_rest
                valor_rest_log = valor_rest

                nome_arquivo_rest = f"{placa}_{contrato}_{data_hoje_formatada}_{km_str_rest}_{valor_rest}_REST.pdf"
                caminho_pdf_rest = gerar_pdf_mapa(driver_rest, nome_arquivo_rest)
                if not caminho_pdf_rest: raise Exception("Falha PDF Restituição")

                if not fazer_login_banco(driver_rest): raise Exception("Login falhou")
                if not navegar_menu_gca(driver_rest): raise Exception("Navegação falhou")

                dados_rest = {
                    "placa": placa, "contrato": contrato, "data": data_hoje_formatada,
                    "valor": str(valor_rest), "tipo_str": "Restituicao", "caminho_pdf": caminho_pdf_rest
                }

                WebDriverWait(driver_rest, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["externo"])))
                WebDriverWait(driver_rest, 10).until(EC.frame_to_be_available_and_switch_to_it((By.ID, SELECTORS["iframes"]["interno"])))

                if not preencher_formulario_com_upload(driver_rest, dados_rest): raise Exception("Upload Restituição falhou")
                logging.info(f"Sucesso Restituição: {placa}")

            except Exception as e:
                logging.error(f"Erro Restituição {placa}: {e}", exc_info=True)
                placas_falha_info.append({'placa': placa, 'motivo': f'Falha Restituição: {e}'})
                sucesso_final_placa = False
            finally:
                if driver_rest: driver_rest.quit()

        if sucesso_final_placa and (run_rem or run_rest):
            placas_sucesso_info.append({
                COLUNA_PLACA: placa, 'km_remocao': km_str_rem_log, 'valor_rem': valor_rem_log,
                'km_restituicao': km_str_rest_log, 'valor_rest': valor_rest_log
            })
            logging.info(f"Placa {placa} OK.")
        elif not (run_rem or run_rest):
            logging.info(f"Nenhuma ação para {placa}.")
        else:
            logging.warning(f"Placa {placa} falhou.")

    logging.info("--- Fim Processamento ---")

    if placas_sucesso_info:
        try:
            df_novas_placas = pd.DataFrame(placas_sucesso_info)
            df_historico_completo = pd.concat([df_historico_antigo, df_novas_placas], ignore_index=True)
            df_historico_completo.drop_duplicates(subset=[COLUNA_PLACA], keep='last', inplace=True)
            df_historico_final = df_historico_completo.reindex(columns=[COLUNA_PLACA, 'km_remocao', 'valor_rem', 'km_restituicao', 'valor_rest'])
            df_historico_final.to_excel(NOME_ARQUIVO_HISTORICO, index=False)
            logging.info("Histórico salvo.")
        except Exception as e:
            logging.critical(f"Erro salvar histórico: {e}", exc_info=True)

    enviar_resumo_telegram(placas_sucesso_info, placas_falha_info)

if __name__ == "__main__":
    iniciar_automacao_completa()