import pandas as pd
from datetime import datetime
import os
from flask import jsonify, make_response

# ============================
# Funções utilitárias
# ============================

def normalize(s):
    """Normaliza nomes e seriais (remove espaços, NT-, lowercase)"""
    if not isinstance(s, str):
        return ""
    return s.strip().lower().replace(" ", "").replace("nt-", "", 1)

# ============================
# Carregamento da planilha
# ============================

def carregar_planilha(upload_dir):
    """Carrega a planilha base do estoque (lansweeper_export.xlsx)"""
    path = os.path.join(upload_dir, "lansweeper_export.xlsx")
    if not os.path.exists(path):
        print("[ ! ] Nenhum arquivo base encontrado. Aguardando upload via interface web.")
        return None

    try:
        df = pd.read_excel(path, engine="openpyxl")
        cols = ["AssetName", "AssetTypename", "Model", "Statename", "Serialnumber", "custom7"]
        df = df[cols]
        return df
    except Exception as e:
        print(f"[ X ] Erro ao carregar planilha: {e}")
        return None

# ============================
# Processamento de bipagem (1 linha por leitura)
# ============================

def processar_bipagem(entrada, df, session):
    """Processa uma entrada (AssetName ou SerialNumber)"""
    if df is None:
        return jsonify({"ok": False, "output": "[ ! ] Nenhuma planilha carregada ainda."})

    status_permitidos = ["Broken", "Active", "Reservado", "Stock", "Old", "In repair"]
    status_recon = ["Broken", "Reservado", "Stock", "Old", "In repair"]

    log = session.get("bipados", [])
    entradas = {normalize(x["AssetName"]) for x in log if "AssetName" in x}
    entradas |= {normalize(x["SerialNumber"]) for x in log if "SerialNumber" in x}

    entrada_norm = normalize(entrada)

    # Verifica duplicata
    if entrada_norm in entradas:
        return jsonify({"ok": True, "output": f"[ ! ] {entrada} já registrado. Ignorando..."})

    # Busca o equipamento na planilha base
    mask = (df["AssetName"].apply(normalize) == entrada_norm) | \
           (df["Serialnumber"].apply(normalize) == entrada_norm)
    equipamento = df[mask]

    registro = {
        "AssetName": "N/A",
        "SerialNumber": "N/A",
        "Statename": "N/A",
        "Resultado": "Não encontrado"
    }

    if not equipamento.empty:
        eq = equipamento.iloc[0]
        registro["AssetName"] = str(eq["AssetName"])
        registro["SerialNumber"] = str(eq["Serialnumber"])
        registro["Statename"] = str(eq["Statename"])

        if eq["Statename"] in status_permitidos or eq["Statename"] in status_recon:
            registro["Resultado"] = "OK"
            cor = "[ OK ]"
        else:
            registro["Resultado"] = "Não permitido"
            cor = "[ ! ]"

        output = f"{cor} {eq['AssetName']} | Serial: {eq['Serialnumber']} | Status: {eq['Statename']}"
    else:
        output = f"[ X ] {entrada} não encontrado"

    log.append(registro)
    session["bipados"] = log
    return jsonify({"ok": True, "output": output})

# ============================
# Geração e comparação do relatório final
# ============================

def gerar_relatorio_final(session):
    """Gera o relatório Excel, compara com a base e retorna para download"""
    log = session.get("bipados", [])
    if not log:
        return jsonify({"ok": False, "output": "[ ! ] Nenhum item bipado."})

    df_stock = pd.DataFrame(log)
    colunas_finais = ["AssetName", "SerialNumber", "Statename", "Resultado"]
    for col in colunas_finais:
        if col not in df_stock.columns:
            df_stock[col] = "N/A"
    df_stock = df_stock[colunas_finais]

    # Gera o arquivo de estoque
    timestamp = datetime.now().strftime("%Y-%m-%d_%H-%M")
    output_dir = "outputs"
    os.makedirs(output_dir, exist_ok=True)
    stock_file = os.path.join(output_dir, f"stock_{timestamp}.xlsx")
    df_stock.to_excel(stock_file, index=False)

    # ============================
    # COMPARAÇÃO AUTOMÁTICA COM A BASE
    # ============================
    base_path = os.path.join("uploads", "lansweeper_export.xlsx")
    faltantes_texto = ""
    if os.path.exists(base_path):
        try:
            df_base = pd.read_excel(base_path, engine="openpyxl")
            df_base["AssetName_norm"] = df_base["AssetName"].apply(normalize)
            df_base["Serialnumber_norm"] = df_base["Serialnumber"].apply(normalize)
            df_stock["AssetName_norm"] = df_stock["AssetName"].apply(normalize)
            df_stock["Serialnumber_norm"] = df_stock["SerialNumber"].apply(normalize)

            status_relevantes = ["Broken", "Reservado", "Old", "In repair"]
            df_filtrado = df_base[df_base["Statename"].isin(status_relevantes)]

            faltando = df_filtrado[
                ~df_filtrado["AssetName_norm"].isin(df_stock["AssetName_norm"]) &
                ~df_filtrado["Serialnumber_norm"].isin(df_stock["Serialnumber_norm"])
            ]

            if not faltando.empty:
                faltantes_texto = "\n".join([
                    f"[ X ] {row['AssetName']} | Serial: {row['Serialnumber']} | Status: {row['Statename']}"
                    for _, row in faltando.iterrows()
                ])
                faltando_path = os.path.join(output_dir, f"faltando_{timestamp}.xlsx")
                faltando[["AssetName", "Serialnumber", "Statename", "Model", "custom7"]].to_excel(faltando_path, index=False)
            else:
                faltantes_texto = "[ OK ] Nenhum equipamento faltando no estoque físico."
        except Exception as e:
            faltantes_texto = f"[ X ] Erro ao comparar estoque: {e}"
    else:
        faltantes_texto = "[ ! ] Base lansweeper_export.xlsx não encontrada."

    # Limpa sessão
    session.pop("bipados", None)

    # ============================
    # ENVIO DO ARQUIVO EXCEL (DOWNLOAD)
    # ============================
    with open(stock_file, "rb") as f:
        data = f.read()

    response = make_response(data)
    response.headers.set("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    response.headers.set("Content-Disposition", f"attachment; filename=stock_{timestamp}.xlsx")
    response.headers.set("Content-Length", len(data))
    response.headers["X-Faltantes"] = faltantes_texto.replace("\n", " | ")

    return response
