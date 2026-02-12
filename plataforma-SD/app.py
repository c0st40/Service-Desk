from flask import Flask, render_template, request, jsonify, send_file, session
import os
from gerador import gerar_termo
from script_stock import carregar_planilha, processar_bipagem, gerar_relatorio_final
from script_qr import gerar_qr_codes


import warnings
warnings.filterwarnings("ignore", category=UserWarning, module="docxcompose")

# Cria pastas se não existirem
os.makedirs("uploads", exist_ok=True)
os.makedirs("outputs", exist_ok=True)

# Cria arquivos falsos para evitar FileNotFoundError
open(os.path.join("uploads", "lansweeper_export.xlsx"), "a").close()
open(os.path.join("uploads", "ADMPReport.xlsx"), "a").close()


app = Flask(__name__)
app.secret_key = "alt-cyberpunk"
UPLOAD_DIR = "uploads"
OUTPUT_DIR = "outputs"
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(OUTPUT_DIR, exist_ok=True)

# ==============================
# Página inicial
# ==============================
@app.route("/")
def index():
    return render_template("index.html")

# ==============================
# GERADOR DE TERMOS
# ==============================
@app.route("/termos")
def termos():
    return render_template("termos.html")

@app.route("/upload", methods=["POST"])
def upload_termos():
    try:
        admp = request.files["admp"]
        lansweeper = request.files["lansweeper"]
        admp.save(os.path.join(UPLOAD_DIR, "ADMPReport.xlsx"))
        lansweeper.save(os.path.join(UPLOAD_DIR, "lansweeper_export.xlsx"))
        return jsonify({"msg": "✅ Planilhas enviadas com sucesso!"})
    except Exception as e:
        return jsonify({"msg": f"Erro: {str(e)}"}), 500

@app.route("/gerar", methods=["POST"])
def gerar():
    data = request.get_json()
    result = gerar_termo(data["tipo"], data["nome"], data["serial"], data.get("motivo"))
    if not result["success"]:
        return jsonify({"msg": result["msg"]}), 400
    return send_file(result["arquivo"], as_attachment=True)

# ==============================
# CONTROLE DE ESTOQUE (BIPAGEM)
# ==============================
@app.route("/estoque")
def estoque():
    return render_template("estoque.html")

@app.route("/upload_estoque", methods=["POST"])
def upload_estoque():
    file = request.files.get("xlsx")
    if not file:
        return jsonify({"ok": False, "msg": "Nenhum arquivo enviado."})
    file.save(os.path.join(UPLOAD_DIR, "lansweeper_export.xlsx"))
    session["bipados"] = []
    return jsonify({"ok": True, "msg": "Planilha carregada com sucesso!"})

@app.route("/cmd", methods=["POST"])
def cmd():
    entrada = request.json.get("entrada", "").strip()
    if entrada.lower() == "sair":
        return gerar_relatorio_final(session)
    planilha = carregar_planilha(UPLOAD_DIR)
    return processar_bipagem(entrada, planilha, session)

# ==============================
# GERADOR DE QR CODES
# ==============================
@app.route("/qr")
def qr_page():
    return render_template("qr.html")

@app.route("/gerar_qr", methods=["POST"])
def gerar_qr():
    file = request.files.get("arquivo")
    if not file:
        return jsonify({"ok": False, "msg": "Nenhum arquivo enviado."}), 400
    return gerar_qr_codes(file)



# ==============================
# EXECUÇÃO
# ==============================
if __name__ == "__main__":
    print("🚀 Servidor iniciado em http://0.0.0.0:8000 (LAN habilitada)")
    app.run(host="0.0.0.0", port=8000)
