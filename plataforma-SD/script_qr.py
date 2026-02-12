import io
import zipfile
import qrcode
from PIL import Image, ImageDraw, ImageFont
from flask import jsonify, send_file

def gerar_qr_codes(file):
    """Gera QR Codes a partir de um arquivo .txt e retorna o ZIP (rápido e compatível com todas as versões do Pillow)."""
    try:
        # Lê o conteúdo do arquivo TXT enviado
        conteudo = file.read().decode("utf-8", errors="ignore")
        tags = [linha.strip() for linha in conteudo.splitlines() if linha.strip()]

        if not tags:
            return jsonify({"ok": False, "msg": "[ ! ] Nenhuma Service Tag encontrada."})

        # Fonte padrão (sem depender de arial.ttf)
        FONT = ImageFont.load_default()

        # Cria o ZIP diretamente em memória
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for tag in tags:
                # Gera o QR Code
                qr = qrcode.QRCode(
                    version=None,
                    error_correction=qrcode.constants.ERROR_CORRECT_M,
                    box_size=10,
                    border=4,
                )
                qr.add_data(tag)
                qr.make(fit=True)
                img_qr = qr.make_image(fill_color="black", back_color="white").convert("RGB")

                # Calcula posição do texto de forma compatível com todas as versões do Pillow
                largura, altura = img_qr.size
                altura_total = altura + 60
                img_final = Image.new("RGB", (largura, altura_total), "white")
                img_final.paste(img_qr, (0, 0))
                draw = ImageDraw.Draw(img_final)

                try:
                    # Pillow >= 10 usa textbbox()
                    bbox = draw.textbbox((0, 0), tag, font=FONT)
                    text_w = bbox[2] - bbox[0]
                    text_h = bbox[3] - bbox[1]
                except AttributeError:
                    # Pillow < 10 ainda possui textsize()
                    text_w, text_h = draw.textsize(tag, font=FONT)

                pos_x = (largura - text_w) // 2
                pos_y = altura + 10
                draw.text((pos_x, pos_y), tag, fill="black", font=FONT)

                # Salva a imagem em memória e adiciona ao ZIP
                img_bytes = io.BytesIO()
                img_final.save(img_bytes, format="PNG")
                img_bytes.seek(0)
                zipf.writestr(f"{tag}.png", img_bytes.getvalue())

        zip_buffer.seek(0)
        return send_file(
            zip_buffer,
            as_attachment=True,
            download_name="qrcodes.zip",
            mimetype="application/zip"
        )

    except Exception as e:
        return jsonify({"ok": False, "msg": f"[ X ] Erro: {e}"})
