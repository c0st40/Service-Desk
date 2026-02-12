import pandas as pd
from docxtpl import DocxTemplate
from babel.dates import format_date
import datetime
import unicodedata
import os

def safe_str(value):
    return str(value) if pd.notnull(value) else "N/A"

def remove_acentos(txt):
    if isinstance(txt, str):
        return ''.join(c for c in unicodedata.normalize('NFD', txt)
                       if unicodedata.category(c) != 'Mn')
    return txt

def gerar_termo(tipo_termo, nome_usuario, serial_equipamento, motivo_input=None):
    try:
        upload_dir = "uploads"
        users_path = os.path.join(upload_dir, "ADMPReport.xlsx")
        equip_path = os.path.join(upload_dir, "lansweeper_export.xlsx")

        if not os.path.exists(users_path) or not os.path.exists(equip_path):
            return {"success": False, "msg": "Arquivos de planilha não encontrados. Faça o upload antes de gerar o termo."}

        # ================================
        # Definição do tipo de termo
        # ================================
        if tipo_termo == "entrega":
            modelo_docx = "modelo_entrega.docx"
            motivo = "Entrega de equipamento"

        elif tipo_termo == "devolucao":
            modelo_docx = "modelo_devolucao.docx"

            # Normaliza o motivo_input para evitar erros de capitalização
            motivo_input = str(motivo_input or "").strip().lower()
            print(f"[DEBUG] Motivo recebido: '{motivo_input}'")  # <-- log para depuração

            if motivo_input == "troca":
                motivo = "Troca de equipamento"
            elif motivo_input == "desligamento":
                motivo = "Desligamento"
            elif motivo_input == "" or motivo_input is None:
                motivo = "Devolução de equipamento"
            else:
                return {"success": False, "msg": f"Motivo inválido: '{motivo_input}' (esperado: troca/desligamento)."}

        else:
            return {"success": False, "msg": "Tipo inválido (entrega/devolucao)."}

        # ================================
        # Carrega planilhas
        # ================================
        df_users = pd.read_excel(users_path, sheet_name="All Users", engine="openpyxl", skiprows=6)
        df_equip = pd.read_excel(equip_path)

        nome_usuario_normalizado = remove_acentos(nome_usuario).strip().lower()
        usuario = df_users[
            df_users['Full Name'].apply(lambda x: remove_acentos(str(x)).strip().lower()) == nome_usuario_normalizado
        ]
        equipamento = df_equip[df_equip['Serialnumber'].astype(str).str.strip().str.lower() == serial_equipamento.strip().lower()]

        if usuario.empty:
            return {"success": False, "msg": f"Usuário '{nome_usuario}' não encontrado."}
        if equipamento.empty:
            return {"success": False, "msg": f"Equipamento '{serial_equipamento}' não encontrado."}

        # ================================
        # Extrai informações do usuário e equipamento
        # ================================
        nome = safe_str(usuario.iloc[0]['Full Name'])
        funcao = safe_str(usuario.iloc[0]['Title'])
        area = safe_str(usuario.iloc[0]['Department'])
        descricao = safe_str(equipamento.iloc[0]['AssetTypename'])
        modelo = safe_str(equipamento.iloc[0]['Model'])
        serial = safe_str(equipamento.iloc[0]['Serialnumber'])
        patrimonio_raw = equipamento.iloc[0].get('custom7', "N/A")

        if pd.notnull(patrimonio_raw):
            try:
                patrimonio = str(int(float(patrimonio_raw)))
            except (ValueError, TypeError):
                patrimonio = str(patrimonio_raw)
        else:
            patrimonio = "N/A"

        data = format_date(datetime.date.today(), "d 'de' MMMM 'de' yyyy", locale='pt_BR')

        # ================================
        # Gera o documento Word
        # ================================
        doc = DocxTemplate(modelo_docx)
        context = {
            'nome': nome,
            'funcao': funcao,
            'area': area,
            'motivo': motivo,
            'descricao': descricao,
            'modelo': modelo,
            'serial': serial,
            'patrimonio': patrimonio,
            'data': data
        }

        output_dir = "outputs"
        os.makedirs(output_dir, exist_ok=True)
        tipo_texto = "de_entrega" if tipo_termo == "entrega" else "de_devolucao"
        safe_nome = nome.replace(" ", "_")
        file_name = os.path.join(output_dir, f"Termo_{tipo_texto}_{safe_nome}.docx")

        doc.render(context)
        doc.save(file_name)

        print(f"[INFO] Termo gerado com sucesso: {file_name} — Motivo: {motivo}")  # log útil para verificação
        return {"success": True, "arquivo": file_name}

    except Exception as e:
        return {"success": False, "msg": str(e)}
