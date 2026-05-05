import pandas as pd
import win32com.client as win32
import re
import os
import shutil
from datetime import datetime


def Prospeccao():
    # ======================================================
    # FUNÇÕES AUXILIARES
    # ======================================================

    def limpar_emails(email_str):
        if pd.isna(email_str) or not str(email_str).strip():
            return ""

        texto = str(email_str).replace(";", " ").replace("/", " ")

        emails = re.findall(
            r"\b[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}\b",
            texto
        )

        return ";".join(dict.fromkeys(email.lower() for email in emails))

    # ======================================================
    # CONFIGURAÇÕES
    # ======================================================

    arquivo_origem = r"Z:\PV - Pos Vendas\EQUIPE\2026\Gestão\REVISÕES DE VEÍCULOS\REVISÕES.xlsx"
    pasta_destino = r"C:\Users\Levi.oliveira\OneDrive - WLM Industria e Comercio\APOIO\CÓDIGOS E AFINS\CÓDIGOS\PROSPECÇÃO"

    os.makedirs(pasta_destino, exist_ok=True)

    arquivo_copia = os.path.join(pasta_destino, "prospecção_copia.xlsx")

    # ======================================================
    # CRIAR CÓPIA
    # ======================================================

    try:
        shutil.copy2(arquivo_origem, arquivo_copia)
        print(f"✅ Cópia criada com sucesso:\n{arquivo_copia}")
    except Exception as e:
        print(f"❌ Erro ao criar cópia: {e}")
        exit()

    # ======================================================
    # CARREGA ABA "prospecção"
    # ======================================================

    try:
        df = pd.read_excel(arquivo_copia, sheet_name="PROSPEC FINAL")
        print("✅ Aba 'prospecção' carregada com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao carregar planilha: {e}")
        exit()

    # ======================================================
    # ENVIO DE EMAILS
    # ======================================================

    modo_envio = input("Deseja enviar os e-mails agora? (S = Enviar / N = Rascunho): ").strip().upper()

    if modo_envio not in ["S", "N"]:
        print("❌ Opção inválida.")
        exit()

    outlook = win32.Dispatch("Outlook.Application")

    for idx, row in df.iterrows():

        try:
            # ================================
            # REGRA DE STATUS
            # ================================
            status = int(row["Status"])

            if status != 1:
                print(f"⛔ Cliente {row['Razão Social Cliente']} ignorado (Status = {status})")
                continue

            # ================================
            # DADOS
            # ================================
            chassi = str(row["Chassi"]).strip()
            cliente = str(row["Razão Social Cliente"]).strip()
            revisao = str(row["Revisão"]).strip()
            placa = str(row["Placas"]).strip()

            # valores numéricos para cálculo
            km_revisao_int = int(row["Km Revisão"])
            km_atual_int = int(row["Km Atual.1"])

            # cálculo da diferença
            diferenca_int = km_revisao_int - km_atual_int

            # formatação brasileira (milhar com ponto)
            km_revisao = f"{km_revisao_int:,}".replace(",", ".")
            km_atual = f"{km_atual_int:,}".replace(",", ".")
            diferenca = f"{diferenca_int:,}".replace(",", ".")

            # ================================
            # EMAILS
            # ================================
            emails = []

            for col in ["E-mails"]:
                if col in df.columns:
                    e_limpo = limpar_emails(row[col])
                    if e_limpo:
                        emails.extend(e_limpo.split(";"))

            emails = list(dict.fromkeys(emails))

            if not emails:
                print(f"⚠️ Cliente {cliente} sem e-mail. Pulando...")
                continue

            # ======================================================
            # HTML
            # ======================================================

            html_body = f"""
            <!DOCTYPE html>
            <html lang="pt-BR">
            <head>
            <meta charset="UTF-8">
            <title>SOLICITAÇÃO DE AGENDAMENTO</title>
            <style>
            body {{
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f8f9fb;
                padding: 20px;
            }}
            .container {{
                max-width: 800px;
                margin: auto;
                background-color: #fff;
                border-radius: 10px;
                box-shadow: 0 2px 6px rgba(0, 0, 0, 0.1);
            }}
            .header {{
                background-color: #041E42;
                color: #fff;
                padding: 15px;
                border-radius: 10px 10px 0 0;
                text-align: center;
            }}
            .content {{
                padding: 20px;
                font-size: 14px;
            }}
            table {{
                border-collapse: collapse;
                width: 100%;
                margin-top: 10px;
            }}
            td {{
                padding: 6px 8px;
                font-size: 1.3em;
            }}

            .detalhamento > table.info th,
            .detalhamento > table.info td  {{
                text-align: center;
                margin: 15px 0;
            }}

            .detalhamento > table.info td{{
                border: 1px solid #E9ECEF;
            }}

            .detalhamento > table.info th {{
                color: #041E42;
                font-size: 1.3em;
                background-color: #E9ECEF;
                padding: 3px;
                border: 1px solid #ced1d4;
            }}

            #Chassi{{
                font-weight: bolder;
            }}

            #Status{{
                font-weight: bolder;
                font-size: 1.3em;
            }}

            #Cliente{{
                font-weight: bolder;
            }}

            .Cliente_nome, #Cliente{{
                font-size: 1.2em;
            }}

            .image img {{
                max-width: 100%;
                border-radius: 8px;
                border: 1px solid #ddd;
            }}
            .alert {{
                background-color: #fff3cd;
                border-left: 4px solid #ffc107;
                padding: 10px 15px;
                margin: 15px 0;
            }}
            .alert p {{
                margin: 0;
                color: #856404;
            }}
            b {{
                color: #00DB81;
            }}
            .informacao > b{{
                color: #041E42;
            }}

            .informacao {{
                background-color: #00DB81;
                padding: 10px 15px;
                margin: 15px 0;
                border-radius: 7px 7px 7px 7px;
                color: #fff;
                font-weight: bolder;
            }}
            </style>
            </head>
            <body>
            <div class="container">
                <div class="header">
                <h2>SOLICITAÇÃO DE AGENDAMENTO SCANIA</h2>
                </div>
                <div class="content">
                <div class="detalhamento">
                    <table class="info">
                    <tr>
                        <th>VEÍCULO</th>
                        <th>KM ATUAL</th>
                        <th>REVISÃO PREVISTA</th>
                        <th>STATUS</th>
                    </tr>
                    <tr>
                        <td id="Chassi">{chassi}</td>
                        <td>{km_atual} Km</td>
                        <td>{km_revisao} Km</td>
                        <td id="Status" rowspan="2">APTO A EXECUTAR</td>
                    </tr>
                    <tr>
                        <td>Placa: {placa}</td>
                        <td>Diferença: {diferenca} Km</td>
                        <td>Revisão {revisao}</td>
                    </tr>
                    <tr>
                        <td class="Cliente_nome">Cliente:</td>
                        <td id="Cliente" colspan="3">{cliente}</td>
                    </tr>
                    </table>
                </div>

                <div class="alert">
                    <p><b>Atenção:</b> A telemetria do veículo pode não estar comunicando o KM atual.<br>
                    Caso se confirme, pedimos, por gentileza, que nos informe o KM atual do veículo.</p>
                </div>

                <p class="informacao"> <b>Novidade Equipo:</b> agora contamos com maior flexibilidade para as revisões preventivas, com possibilidade de estender o horário de atendimento até as 19h, conforme a sua necessidade.</p>

                <div style="text-align:center; margin:15px 0;">
                    <img src="cid:minha_imagem" style="display:block; margin:auto; max-width:500px; width:100%; border-radius:8px;">
                </div>

                <p>Agradecemos a atenção e colaboração.</p>
                <p>Atenciosamente,<br><b>Equipo</b></p>
                </div>
            </div>
            </body>
            </html>
            """

            # ======================================================
            # CRIA EMAIL
            # ======================================================

            mail = outlook.CreateItem(0)
            mail.Subject = f"SOLICITAÇÃO DE AGENDAMENTO SCANIA | {chassi} – {cliente}"
            mail.HTMLBody = html_body

            BASE_DIR = os.path.dirname(os.path.abspath(__file__))
            img_path = os.path.join(BASE_DIR, "Imagem", "minha_imagem.png")

            if os.path.exists(img_path):
                attachment = mail.Attachments.Add(img_path)
                attachment.PropertyAccessor.SetProperty(
                    "http://schemas.microsoft.com/mapi/proptag/0x3712001F",
                    "minha_imagem"
                )
            else:
                print(f"Imagem não encontrada")

            for e in emails:
                mail.Recipients.Add(e).Type = 1

            cc_fixos = [
                "bibiane.silva@wlmequipo.com.br",
                "celio.freitas@wlmequipo.com.br",
                "rodrigo.cabral@wlmequipo.com.br"
            ]

            for email_cc in cc_fixos:
                mail.Recipients.Add(email_cc).Type = 2

            mail.Recipients.ResolveAll()

            if modo_envio == "S":
                mail.Send()
                print(f"📧 Email enviado para {cliente}")

                # 🔥 Atualiza status para 0 após envio
                df.at[idx, "Status"] = 0

            else:
                mail.Save()
                print(f"📝 Email aberto como rascunho para {cliente}")

        except Exception as e:
            print(f"❌ Erro na linha {idx}: {e}")

    # ======================================================
    # SALVA ALTERAÇÕES NA CÓPIA
    # ======================================================

    try:
        df.to_excel(arquivo_copia, sheet_name="prospecção", index=False)
        print("💾 Planilha atualizada com sucesso!")
    except Exception as e:
        print(f"❌ Erro ao salvar planilha: {e}")

    print("✅ Processo finalizado!")
