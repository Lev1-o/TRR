import pandas as pd
import os
import glob
import shutil
import tempfile
import shutil

def historico():
    escolha = input("Deseja atualizar o contador geral? (S/N): ").strip().upper()

    # =====================================================
    # FUNÇÃO PARA LER ARQUIVO SLK
    # =====================================================

    def ler_slk(caminho):
        dados = {}

        with open(caminho, "r", encoding="latin1") as f:
            for linha in f:
                if not linha.startswith("C;"):
                    continue

                partes = linha.strip().split(";")

                col = row = valor = None

                for p in partes:
                    if p.startswith("X"):
                        col = int(p[1:])
                    elif p.startswith("Y"):
                        row = int(p[1:])
                    elif p.startswith("K"):
                        valor = p[1:].strip('"')

                if row and col:
                    dados[(row, col)] = valor

        max_row = max(r for r, c in dados.keys())
        max_col = max(c for r, c in dados.keys())

        tabela = [
            [dados.get((r, c), None) for c in range(1, max_col + 1)]
            for r in range(1, max_row + 1)
        ]

        return pd.DataFrame(tabela)


    # =====================================================
    # FUNÇÃO PARA CONVERTER SLK
    # =====================================================

    def converter_slk(pasta_origem, caminho_destino):

        arquivos_slk = glob.glob(os.path.join(pasta_origem, "*.slk"))

        if not arquivos_slk:
            print("Nenhum SLK encontrado.")
            return False

        arquivo_slk = arquivos_slk[0]
        print(f"Convertendo SLK: {arquivo_slk}")

        df_origem = ler_slk(arquivo_slk)

        # cria arquivo temporário
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
            caminho_temp = tmp.name

        df_origem.to_excel(caminho_temp, index=False)

        # substituição atômica
        shutil.move(caminho_temp, caminho_destino)

        os.remove(arquivo_slk)

        print("SLK convertido com sucesso.")
        return True

    # =====================================================
    # CAMINHOS
    # =====================================================

    pasta_origem = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\BASE MESCLADA 2.0\RELATORIO HISTORICO"
    arquivo_excel = os.path.join(pasta_origem, "RelatorioRelacaoOrdensServico.xlsx")

    # =====================================================
    # 1. GARANTIR QUE EXISTE UM ARQUIVO PARA TRABALHAR
    # =====================================================

    slk_convertido = converter_slk(pasta_origem, arquivo_excel)

    if not os.path.exists(arquivo_excel):
        raise FileNotFoundError(
            f"Nenhum arquivo disponível para processamento:\n{arquivo_excel}"
        )

    # =====================================================
    # 2. TRATAMENTO DOS DADOS
    # =====================================================

    df = pd.read_excel(arquivo_excel, header=None)

    df = df.dropna(how="all")
    df = df.iloc[7:].reset_index(drop=True)

    colunas_remover = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 15, 16, 17]
    df = df.drop(columns=colunas_remover, errors="ignore")

    df = df.iloc[:, :3]
    df.columns = ["Código", "Descrição", "Placa/Chassi"]

    df[["Placa", "Chassi"]] = df["Placa/Chassi"].astype(str).str.split("/", n=1, expand=True)

    df["Placa"] = df["Placa"].ffill()
    df["Chassi"] = df["Chassi"].ffill()

    df["Chassi"] = df["Chassi"].str.replace(" ", "").str.strip()

    df = df[df["Descrição"].astype(str).str.contains("KIT DE FILTROS", na=False)]

    df["Código"] = df["Código"].astype(str).str.split("-").str[-1].str.strip()

    df = df.drop_duplicates(subset=["Chassi"])
    df = df.drop(columns=["Placa/Chassi"])

    # =====================================================
    # 3. SALVAR BASE TRATADA
    # =====================================================

    temp_file = arquivo_excel.replace(".xlsx", "_temp.xlsx")

    df.to_excel(temp_file, index=False)

    os.remove(arquivo_excel)
    os.rename(temp_file, arquivo_excel)

    print("Base tratada com sucesso!")

    # =====================================================
    # 4. HISTÓRICO DE SELEÇÃO
    # =====================================================

    arquivo_origem = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\REVISÕES.xlsx"
    pasta_destino = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\HISTÓRICO DE VEÍCULOS"
    arquivo_historico = os.path.join(pasta_destino, "HISTORICO_SELECAO.xlsx")

    # ✅ NÃO sobrescrever sempre (ESSENCIAL)
    if not os.path.exists(arquivo_historico):
        shutil.copy(arquivo_origem, arquivo_historico)
        print("Histórico criado.")

    # 🔹 Ler histórico (aba correta)
    df_hist = pd.read_excel(arquivo_historico, sheet_name="PROSPECÇÃO")

    # 🔹 Garantir que coluna Seleção existe
    if "Seleção" not in df_hist.columns:
        df_hist["Seleção"] = 0

    # 🔹 Manter apenas colunas necessárias
    df_hist = df_hist[["Chassi", "Seleção"]]

    df_hist["Seleção"] = df_hist["Seleção"].fillna(0)

    # Base tratada
    df_base = pd.read_excel(arquivo_excel)

    df_hist["Chassi"] = df_hist["Chassi"].astype(str)
    df_base["Chassi"] = df_base["Chassi"].astype(str)

    chassis_base = set(df_base["Chassi"])

    if escolha.upper() == "S":
        # Atualizar contador
        df_hist["Seleção"] = df_hist.apply(
            lambda row: row["Seleção"] + 1 if row["Chassi"] in chassis_base else row["Seleção"],
            axis=1
        )
        print("Contador atualizado")
    else:
        print("Contador não atualizado")

    df_hist.to_excel(arquivo_historico, sheet_name="PROSPECÇÃO", index=False)

    print("Histórico atualizado com sucesso!")
