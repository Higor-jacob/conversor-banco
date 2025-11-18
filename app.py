import os
import re
import sqlite3
import pandas as pd
from flask import Flask, render_template, request, jsonify, send_from_directory
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)
app.config["UPLOAD_FOLDER"] = "uploads"
app.config["OUTPUT_FOLDER"] = "outputs"

# Criar pastas
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)


# ----------------------------------------------------------
#  🔍 DETECTAR TABELAS EM ARQUIVO SQL (MySQL/Postgres)
# ----------------------------------------------------------
def detectar_tabelas_sql(conteudo):
    tabelas = {}

    # Pegar CREATE TABLE
    regex_create = r"CREATE TABLE\s+`?(\w+)`?\s*\((.*?)\);"
    matches = re.findall(regex_create, conteudo, re.S | re.I)

    for tabela, corpo in matches:
        colunas = []
        for linha in corpo.split("\n"):
            linha = linha.strip()
            if (not linha 
                or linha.upper().startswith(("PRIMARY", "UNIQUE", "KEY", "INDEX", "CONSTRAINT"))):
                continue

            m = re.match(r"`?(\w+)`?\s", linha)
            if m:
                colunas.append(m.group(1))

        tabelas[tabela] = {"colunas": colunas, "linhas": []}

    return tabelas


# ----------------------------------------------------------
#  Converter .db → Excel
# ----------------------------------------------------------
def converter_db_para_excel(caminho_db, caminho_excel):
    conn = sqlite3.connect(caminho_db)

    tables = pd.read_sql_query(
        "SELECT name FROM sqlite_master WHERE type='table';", conn
    )["name"].tolist()

    writer = pd.ExcelWriter(caminho_excel, engine="xlsxwriter")

    for tabela in tables:
        df = pd.read_sql_query(f"SELECT * FROM {tabela}", conn)
        df.to_excel(writer, sheet_name=tabela[:31], index=False)

    writer.close()
    conn.close()


# ----------------------------------------------------------
#  Converter .sql → Excel
# ----------------------------------------------------------
def converter_sql_para_excel(caminho_sql, caminho_excel):
    with open(caminho_sql, "r", encoding="utf-8", errors="ignore") as f:
        conteudo = f.read()

    tabelas = detectar_tabelas_sql(conteudo)

    # Capturar INSERT na forma: INSERT INTO tabela (...) VALUES (...)
    regex_insert = r"INSERT INTO\s+`?(\w+)`?(?:\s*\((.*?)\))?\s*VALUES\s*(\(.*?\));"
    inserts = re.findall(regex_insert, conteudo, re.S | re.I)

    for tabela, colunas_insert, bloco_valores in inserts:
        if tabela not in tabelas:
            tabelas[tabela] = {"colunas": [], "linhas": []}

        # Se INSERT define colunas explicitamente
        if colunas_insert:
            colunas_do_insert = [c.strip(" `") for c in colunas_insert.split(",")]
        else:
            colunas_do_insert = []

        # Pegar todas as tuplas VALUES(...)
        tuplas = re.findall(r"\((.*?)\)", bloco_valores, re.S)

        for t in tuplas:
            # Separar valores mantendo integridade
            valores = [v.strip().strip("'").strip('"') for v in t.split(",")]

            # CASO 1 - INSERT especificou colunas
            if colunas_do_insert:
                # Garantir que colunas existam na tabela
                for col in colunas_do_insert:
                    if col not in tabelas[tabela]["colunas"]:
                        tabelas[tabela]["colunas"].append(col)

                # Preencher linha com None
                linha_completa = [None] * len(tabelas[tabela]["colunas"])

                for i, col in enumerate(colunas_do_insert):
                    idx = tabelas[tabela]["colunas"].index(col)
                    linha_completa[idx] = valores[i] if i < len(valores) else None

                tabelas[tabela]["linhas"].append(linha_completa)
                continue

            # CASO 2 - INSERT NÃO definiu colunas
            # Ajustar número de colunas/v alores
            if len(valores) > len(tabelas[tabela]["colunas"]):
                faltam = len(valores) - len(tabelas[tabela]["colunas"])
                for i in range(faltam):
                    tabelas[tabela]["colunas"].append(f"col_auto_{len(tabelas[tabela]['colunas'])+1}")

            if len(valores) < len(tabelas[tabela]["colunas"]):
                valores.extend([None] * (len(tabelas[tabela]["colunas"]) - len(valores)))

            tabelas[tabela]["linhas"].append(valores)

    # GERAR EXCEL
    writer = pd.ExcelWriter(caminho_excel, engine="xlsxwriter")
    for tabela, dados in tabelas.items():
        df = pd.DataFrame(dados["linhas"], columns=dados["colunas"])
        df.to_excel(writer, sheet_name=tabela[:31], index=False)

    writer.close()


# ----------------------------------------------------------
#  📤 ROTA UPLOAD → GERA EXCEL → RETORNA URL JSON
# ----------------------------------------------------------
@app.route("/upload", methods=["POST"])
def upload():
    if "file" not in request.files:
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    arquivo = request.files["file"]

    if arquivo.filename == "":
        return jsonify({"error": "Arquivo inválido"}), 400

    nome = secure_filename(arquivo.filename)
    caminho_entrada = os.path.join(app.config["UPLOAD_FOLDER"], nome)
    arquivo.save(caminho_entrada)

    nome_excel = nome.rsplit(".", 1)[0] + ".xlsx"
    caminho_excel = os.path.join(app.config["OUTPUT_FOLDER"], nome_excel)

    try:
        if nome.lower().endswith(".db"):
            converter_db_para_excel(caminho_entrada, caminho_excel)
        elif nome.lower().endswith(".sql"):
            converter_sql_para_excel(caminho_entrada, caminho_excel)
        else:
            return jsonify({"error": "Formato não suportado"}), 400

        return jsonify({
            "download_url": f"/download/{nome_excel}"
        })

    except Exception as e:
        return jsonify({"error": str(e)}), 500


# ----------------------------------------------------------
#  📥 ROTA PARA DOWNLOAD DIRETO DO EXCEL
# ----------------------------------------------------------
@app.route("/download/<nome>")
def download(nome):
    return send_from_directory(app.config["OUTPUT_FOLDER"], nome, as_attachment=True)


# ----------------------------------------------------------
#  Página inicial
# ----------------------------------------------------------
@app.route("/")
def index():
    return render_template("index.html")


# ----------------------------------------------------------
#  RUN
# ----------------------------------------------------------
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)