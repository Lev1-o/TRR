import os
import time
import shutil
import logging
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl import load_workbook
import win32com.client
import pythoncom

from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import tkinter as tk
from tkinter import messagebox


# ======================================================
# CONFIGURAÇÕES
# ======================================================
ARQUIVO_ORIGEM = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\REVISÕES.xlsx"
PASTA_DESTINO = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\KM ATUAL"
ARQUIVO_COPIA = os.path.join(PASTA_DESTINO, "Km_atual.xlsx")
URL = "https://dealer.scania.com"

# Arquivo de log ao lado do executável/script
BASE_DIR = Path(__file__).resolve().parent
LOG_FILE = BASE_DIR / "km_atual.log"

logging.basicConfig(
    filename=str(LOG_FILE),
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s"
)


# ======================================================
# UTILITÁRIOS
# ======================================================
def limpar_km(km):
    if km is None:
        return 0

    km_limpo = str(km).strip()
    km_limpo = km_limpo.replace(".", "").replace(",", "")

    if km_limpo in ["-", "–", "—"]:
        return 0

    return int(km_limpo) if km_limpo.isdigit() else 0


def mostrar_aviso_login():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo(
        "Login necessário",
        "O navegador será aberto.\n\n"
        "Faça login no site da Scania e depois clique em OK para continuar."
    )
    root.destroy()


def atualizar_arquivo_excel(arquivo_origem):
    """
    Abre o arquivo no Excel, atualiza as conexões e salva.
    Requer Excel instalado na máquina.
    """
    xlapp = None
    try:
        pythoncom.CoInitialize()
        xlapp = win32com.client.DispatchEx("Excel.Application")
        xlapp.DisplayAlerts = False
        xlapp.Visible = False

        arquivo = xlapp.Workbooks.Open(arquivo_origem)
        arquivo.RefreshAll()
        xlapp.CalculateUntilAsyncQueriesDone()
        arquivo.Save()
        arquivo.Close(SaveChanges=True)

        logging.info("Arquivo Excel atualizado com sucesso.")
    finally:
        if xlapp is not None:
            try:
                xlapp.Quit()
            except Exception:
                pass
        try:
            pythoncom.CoUninitialize()
        except Exception:
            pass


def copiar_arquivo(arquivo_origem, arquivo_copia):
    os.makedirs(os.path.dirname(arquivo_copia), exist_ok=True)
    shutil.copy2(arquivo_origem, arquivo_copia)
    logging.info(f"Cópia criada: {arquivo_copia}")


def carregar_dados(arquivo_copia):
    df = pd.read_excel(arquivo_copia, sheet_name="PROSPECÇÃO")
    df.columns = df.columns.str.strip().str.lower()

    if "chassi" not in df.columns:
        raise ValueError(
            f"Coluna 'Chassi' não encontrada. Colunas disponíveis: {df.columns.tolist()}"
        )

    df = df[["chassi"]].copy()
    df.columns = ["Veiculo"]

    df["Veiculo"] = df["Veiculo"].astype(str).str.strip()
    df["Veiculo"] = df["Veiculo"].str.replace(".0", "", regex=False)

    df = df[(df["Veiculo"] != "") & (df["Veiculo"].str.lower() != "nan")]
    df = df.drop_duplicates(subset="Veiculo").reset_index(drop=True)
    df["KM Atual"] = 0

    return df


def abrir_navegador():
    options = webdriver.ChromeOptions()
    options.add_argument("--start-maximized")
    options.add_argument("--disable-notifications")

    # Se quiser rodar sem abrir a janela, descomente:
    # options.add_argument("--headless=new")

    driver = webdriver.Chrome(service=Service(), options=options)
    wait = WebDriverWait(driver, 20)
    return driver, wait


def buscar_km_no_site(driver, wait, busca):
    campo_busca = wait.until(
        EC.presence_of_element_located((By.XPATH, "//input"))
    )
    campo_busca.clear()
    campo_busca.send_keys(busca)

    driver.find_element(By.XPATH, "//button").click()

    wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//div[@data-test-id='search-result-row']")
        )
    )

    lista = driver.find_elements(
        By.XPATH, "//div[@data-test-id='search-result-row']"
    )

    encontrado = False
    for item in lista:
        if busca.lower() in item.text.lower():
            driver.execute_script("arguments[0].click();", item)
            encontrado = True
            break

    if not encontrado and lista:
        driver.execute_script("arguments[0].click();", lista[0])

    time.sleep(3)

    detalhe = wait.until(
        EC.presence_of_element_located(
            (By.XPATH, "//p[@data-test-id='lead-preview-text-block-mileage']")
        )
    )

    return limpar_km(detalhe.text)


def salvar_resultados(arquivo_copia, df, historico_novo):
    # Lê histórico anterior antes de sobrescrever
    try:
        df_historico_antigo = pd.read_excel(arquivo_copia, sheet_name="HISTORICO_KM")
    except Exception:
        df_historico_antigo = pd.DataFrame(columns=["Veiculo", "KM", "Data"])

    df_historico_final = pd.concat(
        [df_historico_antigo, historico_novo],
        ignore_index=True
    )

    # Regrava as duas planilhas, preservando as outras abas do arquivo
    with pd.ExcelWriter(
        arquivo_copia,
        engine="openpyxl",
        mode="a",
        if_sheet_exists="replace"
    ) as writer:
        df.to_excel(writer, sheet_name="PROSPECÇÃO", index=False)
        df_historico_final.to_excel(writer, sheet_name="HISTORICO_KM", index=False)


# ======================================================
# PROCESSO PRINCIPAL
# ======================================================
def KmAtual():
    try:
        logging.info("Início do processo KmAtual.")

        # 1) Atualiza o arquivo original
        atualizar_arquivo_excel(ARQUIVO_ORIGEM)

        # 2) Copia o arquivo para a pasta de trabalho
        copiar_arquivo(ARQUIVO_ORIGEM, ARQUIVO_COPIA)

        # 3) Carrega dados
        df = carregar_dados(ARQUIVO_COPIA)
        print(f"🔎 Total de chassis únicos: {len(df)}")
        logging.info(f"Total de chassis únicos: {len(df)}")

        # 4) Abre navegador e pede login
        driver, wait = abrir_navegador()
        driver.get(URL)

        mostrar_aviso_login()

        historico = []

        # 5) Loop principal
        for idx, row in df.iterrows():
            busca = row["Veiculo"]

            try:
                resultado = buscar_km_no_site(driver, wait, busca)
            except Exception as e:
                resultado = 0
                msg = f"Erro em {busca}: {e}"
                print(f"❌ {msg}")
                logging.exception(msg)

            df.at[idx, "KM Atual"] = resultado

            historico.append({
                "Veiculo": busca,
                "KM": resultado,
                "Data": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            })

            print(f"✅ {busca} → {resultado}")
            logging.info(f"{busca} → {resultado}")

        # 6) Salva tudo
        df_historico_novo = pd.DataFrame(historico)
        salvar_resultados(ARQUIVO_COPIA, df, df_historico_novo)

        print("Finalizado")
        logging.info("Processo finalizado com sucesso.")

    except Exception as e:
        logging.exception(f"Falha geral no processo: {e}")
        print(f"❌ Erro geral: {e}")

    finally:
        try:
            driver.quit()
        except Exception:
            pass


if __name__ == "__main__":
    KmAtual()
