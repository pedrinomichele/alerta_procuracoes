# -*- coding: utf-8 -*-
import pandas as pd
from datetime import datetime
import win32com.client as win32

# ===== LER PLANILHA =====
df = pd.read_excel("procuracoes.xlsx")
df.columns = df.columns.str.strip()

df["data_vencimento"] = pd.to_datetime(df["data_vencimento"], errors="coerce")
df = df.dropna(subset=["data_vencimento"])

df["criticidade"] = df["criticidade"].str.upper().str.strip()

# ===== CALCULAR DIAS =====
hoje = datetime.today()
df["dias_restantes"] = (df["data_vencimento"] - hoje).dt.days

# ===== FILTRAR ALERTAS =====
alertas = []

for _, row in df.iterrows():
    dias = pd.to_numeric(row["dias_restantes"], errors="coerce")

    if pd.isna(dias):
        continue

    criticidade = row["criticidade"]

    enviar = False

    if criticidade == "ALTA":
        if dias <= 45:
            enviar = True

    elif criticidade in ["MEDIA", "BAIXA"]:
        if dias <= 30:
            enviar = True

    if enviar:
        alertas.append(row.to_dict())

# transformar em dataframe
alertas_df = pd.DataFrame(alertas)

# ===== ENVIO EMAIL =====
if alertas_df.empty:
    print("Nenhum alerta hoje.")
else:
    html = """
    <html>
    <head>
    <style>
    body { font-family: Arial; }

    table { border-collapse: collapse; width: 100%; }

    th, td {
        border: 1px solid #ddd;
        padding: 8px;
    }

    th { background-color: #f2f2f2; }

    .alta { background-color: #f8d7da; }
    .media { background-color: #fff3cd; }
    .baixa { background-color: #d4edda; }

    .urgente {
        background-color: #ff4d4d;
        color: white;
        font-weight: bold;
    }

    .atencao {
        background-color: #ffa500;
        font-weight: bold;
    }
    </style>
    </head>
    <body>

    <p>A/C Dra. Gisele</p>

    <h3>🚨 Procurações com alerta</h3>

    <table>
    <tr>
        <th>Outorgante</th>
        <th>Tipo</th>
        <th>Criticidade</th>
        <th>Data de Vencimento</th>
        <th>Dias Restantes</th> 
        <th>Status</th>
    </tr>
    """
    alertas_df = alertas_df.sort_values(by="dias_restantes")
    for _, row in alertas_df.iterrows():
        data_formatada = row["data_vencimento"].strftime("%d/%m/%Y")
        dias = row["dias_restantes"]

        # ===== STATUS =====
        if dias < 0:
            status = "VENCIDO"
        elif dias <= 7:
            status = "URGENTE"
        elif dias <= 15:
            status = "ATENÇÃO"
        else:
            status = "OK"

        # ===== COR =====
        if dias < 0:
            classe = "urgente"
        elif dias <= 7:
            classe = "urgente"
        elif dias <= 15:
            classe = "atencao"
        else:
            classe = row["criticidade"].lower()

        html += f"""
        <tr class="{classe}">
            <td>{row['outorgante']}</td>
            <td>{row['tipo']}</td>
            <td>{row['criticidade']}</td>
            <td>{data_formatada}</td>
            <td>{row['dias_restantes']}</td>
            <td>{row['status']}</td>
        </tr>
        """

    html += """
    </table>

    <p style="margin-top:20px;">
    Este é um alerta automático de vencimento de procurações.
    </p>

    </body>
    </html>
    """

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = "michele.pedrino@gmail.com"
    mail.Subject = "🚨 ALERTA - Procurações"
    mail.HTMLBody = html

    mail.Send()

    print("Email enviado com sucesso!")