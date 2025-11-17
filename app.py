from flask import (
    Flask,
    request,
    redirect,
    url_for,
    render_template,
    jsonify,
    session,
    flash,
    get_flashed_messages,
    send_file,
)
import os
import uuid
import re
import locale
from datetime import datetime
from functools import wraps
from io import BytesIO

import pandas as pd
import xlrd  # Para ler arquivos .xls, se necessário
from werkzeug.utils import secure_filename
from openpyxl import load_workbook, Workbook  # Usado para trabalhar com XLSX
from openpyxl.utils import get_column_letter  # Para obter a coluna em letra
from openpyxl.cell import MergedCell  # Para identificar células mescladas


# ==========================================================
#  CONFIGURAÇÃO BÁSICA DA APLICAÇÃO
# ==========================================================

# Tenta definir a localidade para formatação de datas em português
try:
    locale.setlocale(locale.LC_TIME, "pt_BR.UTF-8")
except locale.Error:
    pass

app = Flask(__name__)
app.secret_key = "sua_chave_secreta"  # Altere para uma chave segura
ACCESS_TOKEN = "minha_senha"  # Token de acesso

app.config["UPLOAD_FOLDER"] = "uploads"
ALLOWED_EXTENSIONS = {".jpg", ".jpeg", ".png", ".bmp", ".gif"}

# Cria os diretórios necessários, se não existirem
os.makedirs("static/fotos", exist_ok=True)
os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)

# Caminho do arquivo CSV relativo ao diretório do script
CSV_PATH = os.path.join(os.path.dirname(__file__), "uploads", "escolas.csv")

# Variável global para armazenar os dados do CSV
escolas_df = None


# ==========================================================
#  FUNÇÕES AUXILIARES – ESCOLAS (escolas.csv)
# ==========================================================

def carregar_escolas():
    """Carrega o CSV de escolas em um DataFrame global."""
    global escolas_df
    if os.path.exists(CSV_PATH):
        try:
            escolas_df = pd.read_csv(CSV_PATH, encoding="latin1", sep=";")
            print(f"[INFO] Arquivo {CSV_PATH} carregado com sucesso.")
        except Exception as e:
            escolas_df = None
            print(f"[ERRO] Falha ao carregar {CSV_PATH}: {e}")
    else:
        escolas_df = None
        print(f"[ERRO] Arquivo {CSV_PATH} não encontrado.")


def get_escolas_df():
    """Garante que o DataFrame de escolas está carregado."""
    global escolas_df
    if escolas_df is None or escolas_df.empty:
        print("[INFO] Recarregando arquivo escolas.csv...")
        carregar_escolas()
    return escolas_df


@app.before_request
def inicializar_escolas():
    """Garante que escolas.csv está carregado antes de cada requisição."""
    if escolas_df is None or (isinstance(escolas_df, pd.DataFrame) and escolas_df.empty):
        carregar_escolas()


@app.route("/escolas/search")
def escolas_search():
    """Endpoint para o Select2 buscar escolas no CSV."""
    df = get_escolas_df()
    query = request.args.get("q", "").lower().strip()
    results = []

    if df is not None and not df.empty and query:
        # Filtra usando pandas (assumindo coluna 3 = nome da escola)
        df_filtered = df[df.iloc[:, 3].str.lower().str.contains(query, na=False)]

        # Limita a 50 resultados para não sobrecarregar
        df_filtered = df_filtered.head(50)

        for _, row in df_filtered.iterrows():
            nome = str(row[3]).strip()
            municipio = str(row[2]).strip()
            uf = str(row[1]).strip()
            text = f"{nome} - {municipio}/{uf}"
            results.append({"id": nome, "text": text})

    return jsonify(results)


# Carrega o CSV na inicialização do sistema
carregar_escolas()


# ==========================================================
#  BLUEPRINTS
# ==========================================================

from confere import confere_bp

app.register_blueprint(confere_bp, url_prefix="/confere")


# ==========================================================
#  HELPERS GERAIS
# ==========================================================

def allowed_file(filename):
    _, ext = os.path.splitext(filename)
    return ext.lower() in ALLOWED_EXTENSIONS


def login_required(f):
    @wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get("logged_in"):
            return redirect(url_for("login_route", next=request.url))
        return f(*args, **kwargs)

    return decorated_function


def set_merged_cell_value(ws, cell_coord, value):
    """
    Atualiza o valor de uma célula mesclada em uma planilha openpyxl
    preservando a mesclagem.
    """
    cell = ws[cell_coord]
    if isinstance(cell, MergedCell):
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                range_str = str(merged_range)
                ws.unmerge_cells(range_str)
                min_col, min_row, _, _ = merged_range.bounds
                top_left_coord = f"{get_column_letter(min_col)}{min_row}"
                ws[top_left_coord] = value
                ws.merge_cells(range_str)
                return
    ws[cell_coord] = value


def convert_xls_to_xlsx(file_like):
    """
    Converte um arquivo XLS (file-like) para um Workbook do openpyxl.
    """
    book_xlrd = xlrd.open_workbook(file_contents=file_like.read())
    wb = Workbook()
    # Remove a planilha padrão criada pelo openpyxl, se houver
    if "Sheet" in wb.sheetnames and len(book_xlrd.sheet_names()) > 0:
        std = wb.active
        wb.remove(std)

    for sheet_name in book_xlrd.sheet_names():
        sheet_xlrd = book_xlrd.sheet_by_name(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        for row in range(sheet_xlrd.nrows):
            for col in range(sheet_xlrd.ncols):
                ws.cell(row=row + 1, column=col + 1, value=sheet_xlrd.cell_value(row, col))

    return wb


def load_workbook_model(file):
    """
    Abre o arquivo do modelo XLSX (ou XLS convertendo-o para XLSX)
    preservando a formatação.
    """
    ext = os.path.splitext(file.filename)[1].lower()
    file.seek(0)
    if ext == ".xlsx":
        return load_workbook(file, data_only=False)
    elif ext == ".xls":
        content = file.read()
        return convert_xls_to_xlsx(BytesIO(content))
    else:
        raise ValueError("Formato de arquivo não suportado para o quadro modelo.")


def data_extenso(dt):
    """Retorna a data por extenso em português."""
    meses = [
        "janeiro",
        "fevereiro",
        "março",
        "abril",
        "maio",
        "junho",
        "julho",
        "agosto",
        "setembro",
        "outubro",
        "novembro",
        "dezembro",
    ]
    return f"{dt.day} de {meses[dt.month - 1]} de {dt.year}"


def is_valid_plano(val):
    """Avalia se o valor do Plano de Ação é considerado válido."""
    if val is None:
        return False
    s = str(val).strip()
    return s not in ["", "-", "0", "#REF"]


# ==========================================================
#  CARTEIRINHAS – GERAÇÃO DE HTML
# ==========================================================

def gerar_html_carteirinhas(arquivo_excel):
    # Lê a planilha do Fundamental
    planilha = pd.read_excel(arquivo_excel, sheet_name="LISTA CORRIDA")

    dados = planilha[["RM", "NOME", "DATA NASC.", "RA", "SAI SOZINHO?", "SÉRIE", "HORÁRIO"]].copy()
    dados["RM"] = dados["RM"].fillna(0).astype(int)

    # Filtra apenas alunos com RM válido
    registros_validos = []
    for _, row in dados.iterrows():
        rm_str = str(row["RM"]).strip()
        if not rm_str or rm_str == "0":
            continue
        registros_validos.append(row)

    alunos_sem_fotos_list = []
    alunos = []

    allowed_exts = [".jpg", ".jpeg", ".png", ".bmp", ".gif"]

    for row in registros_validos:
        nome = row["NOME"]
        data_nasc = row["DATA NASC."]
        serie = row["SÉRIE"]
        horario = row["HORÁRIO"]

        # Data de nascimento
        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors="coerce")
                if pd.notna(data_nasc):
                    data_nasc_str = data_nasc.strftime("%d/%m/%Y")
                else:
                    data_nasc_str = "Desconhecida"
            except Exception:
                data_nasc_str = "Desconhecida"
        else:
            data_nasc_str = "Desconhecida"

        ra = row["RA"]

        # Sai sozinho?
        sai_sozinho_raw = str(row["SAI SOZINHO?"]).strip().upper()
        if sai_sozinho_raw in ("SIM", "S", "YES", "Y"):
            classe_cor = "verde"
            status_texto = "Sai sozinho"
            status_icon = "&#10003;"
        else:
            classe_cor = "vermelho"
            status_texto = "Não sai sozinho"
            status_icon = "&#9888;"

        # Foto
        foto_url = None
        for ext in allowed_exts:
            caminho_foto = f"static/fotos/{row['RM']}{ext}"
            if os.path.exists(caminho_foto):
                foto_url = f"/static/fotos/{row['RM']}{ext}"
                break

        if not foto_url:
            alunos_sem_fotos_list.append(
                {
                    "rm": int(row["RM"]),
                    "nome": nome,
                    "serie": serie,
                }
            )

        alunos.append(
            {
                "rm": int(row["RM"]),
                "nome": nome,
                "data_nasc": data_nasc_str,
                "ra": ra,
                "serie": serie,
                "horario": horario,
                "classe_cor": classe_cor,
                "status_texto": status_texto,
                "status_icon": status_icon,
                "foto_url": foto_url,
            }
        )

    # Paginação: 6 carteirinhas por página
    pages = []
    for i in range(0, len(alunos), 6):
        pages.append(alunos[i : i + 6])

    total_sem_foto = len(alunos_sem_fotos_list)

    return render_template(
        "gerar_carteirinhas.html",
        pages=pages,
        alunos_sem_foto=alunos_sem_fotos_list,
        total_sem_foto=total_sem_foto,
    )


# ==========================================================
#  DECLARAÇÕES – GERAÇÃO HTML (SINGULAR)
# ==========================================================

from datetime import datetime
import pandas as pd, re
from flask import session

escolas_df = None

def gerar_declaracao_escolar(
    file_path,
    rm,
    tipo,
    file_path2=None,
    deve_historico=False,
    unidade_anterior=None,
):
    """
    Gera o HTML de uma declaração escolar (Escolaridade, Transferência ou Conclusão)
    tanto para Fundamental quanto EJA, de acordo com session['declaracao_tipo'].

    file_path  -> caminho/arquivo padrão da lista piloto (salvo em sessão/ao entrar no sistema)
    file_path2 -> caminho/arquivo opcional, usado quando o usuário reenviar a lista
                  (por exemplo, após o servidor free acordar). SE informado, TERÁ PRIORIDADE.
    """
    global escolas_df

    # Se um segundo caminho foi informado (lista reenviada), ele tem prioridade.
    effective_path = file_path2 if file_path2 is not None else file_path
    if file_path2 is not None:
        print("[DEBUG] gerar_declaracao_escolar: usando file_path2 =", effective_path)
    else:
        print("[DEBUG] gerar_declaracao_escolar: usando file_path  =", effective_path)

    
    # ------------------------------------------------------
    # 1) CARREGAMENTO DOS DADOS DO ALUNO (FUNDAMENTAL x EJA)
    # ------------------------------------------------------
    if session.get("declaracao_tipo") != "EJA":
        # ---------- FUNDAMENTAL ----------
        planilha = pd.read_excel(effective_path, sheet_name="LISTA CORRIDA")
        planilha.columns = [c.strip().upper() for c in planilha.columns]

        def format_rm(x):
            try:
                return str(int(float(x)))
            except Exception:
                return str(x)

        planilha["RM_str"] = planilha["RM"].apply(format_rm)

        try:
            rm_num = str(int(float(rm)))
        except Exception:
            rm_num = str(rm)

        aluno = planilha[planilha["RM_str"] == rm_num]
        if aluno.empty:
            return None

        row = aluno.iloc[0]

        semestre_texto = ""  # Fundamental não usa semestre
        nome = row["NOME"]
        serie = row["SÉRIE"]

        if isinstance(serie, str):
            # transforma "5ºA" em "5º ano A"
            serie = re.sub(r"(\d+º)([A-Za-z])", r"\1 ano \2", serie)

        data_nasc = row["DATA NASC."]
        ra = row["RA"]
        horario = row.get("HORÁRIO", "Desconhecido")

        if pd.isna(horario) or not str(horario).strip():
            horario = "Desconhecido"
        else:
            horario = str(horario).strip()

        ra_label = "RA"

        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors="coerce")
                if pd.notna(data_nasc):
                    data_nasc = data_nasc.strftime("%d/%m/%Y")
                else:
                    data_nasc = "Desconhecida"
            except Exception:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"

    else:
        # ---------- EJA ----------
        df = pd.read_excel(effective_path, sheet_name=0, header=None, skiprows=1)
        df.columns = [str(c).strip().upper() for c in df.columns]

        # RM (coluna 2)
        df["RM_str"] = df.iloc[:, 2].apply(
            lambda x: str(int(x)) if pd.notna(x) and float(x) != 0 else ""
        )
        # Nome (coluna 3)
        df["NOME"] = df.iloc[:, 3]
        # Nascimento (coluna 6)
        df["NASC."] = df.iloc[:, 6]

        def get_ra(row_local):
            try:
                val = row_local.iloc[7]
                if pd.isna(val) or float(val) == 0:
                    return row_local.iloc[8]
                else:
                    return val
            except Exception:
                return row_local.iloc[7]

        df["RA"] = df.apply(get_ra, axis=1)
        # Série (coluna 0)
        df["SÉRIE"] = df.iloc[:, 0]

        try:
            rm_num = str(int(float(rm)))
        except Exception:
            rm_num = str(rm)

        aluno = df[df["RM_str"] == rm_num]
        if aluno.empty:
            return None

        row = aluno.iloc[0]

        # Semestre (quando existir na planilha)
        if len(row) > 29:
            semestre = row.iloc[29]
            semestre_texto = str(semestre).strip() if pd.notna(semestre) else ""
        else:
            semestre_texto = ""

        nome = row["NOME"]
        serie = row["SÉRIE"]
        if isinstance(serie, str):
            serie = re.sub(r"(\d+º)([A-Za-z])", r"\1 ano \2", serie)

        data_nasc = row["NASC."]
        ra = row["RA"]
        original_ra = row.iloc[7]

        # Se RA for vazio / 0, trata como RG
        if pd.isna(original_ra) or (
            isinstance(original_ra, (int, float)) and float(original_ra) == 0
        ):
            ra_label = "RG"
        else:
            ra_label = "RA"

        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors="coerce")
                data_nasc = (
                    data_nasc.strftime("%d/%m/%Y") if pd.notna(data_nasc) else "Desconhecida"
                )
            except Exception:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"

    # ------------------------------------------------------
    # 2) DATA POR EXTENSO
    # ------------------------------------------------------
    now = datetime.now()
    meses = {
        1: "janeiro",
        2: "fevereiro",
        3: "março",
        4: "abril",
        5: "maio",
        6: "junho",
        7: "julho",
        8: "agosto",
        9: "setembro",
        10: "outubro",
        11: "novembro",
        12: "dezembro",
    }
    mes = meses[now.month].capitalize()
    data_extenso = f"Praia Grande, {now.day:02d} de {mes} de {now.year}"

    additional_css = """
.print-button {
  background-color: #283E51;
  color: #fff;
  border: none;
  padding: 10px 20px;
  border-radius: 5px;
  cursor: pointer;
  margin-top: 20px;
}
.print-button:hover {
  background-color: #1d2d3a;
}
"""

    # ------------------------------------------------------
    # 3) MONTAGEM DO TEXTO DA DECLARAÇÃO
    # ------------------------------------------------------
    declaracao_text = ""

    if tipo == "Escolaridade":
        titulo = "Declaração de Escolaridade"
        if session.get("declaracao_tipo") == "EJA":
                declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, "
                f"encontra-se regularmente matriculado(a) no segmento de "
                f"<strong><u>Educação de Jovens e Adultos (EJA)</u></strong> da "
                f"E.M José Padin Mouta, cursando atualmente o(a) "
                f"<strong><u>{serie}</u></strong>."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, "
                f"encontra-se regularmente matriculado(a) na "
                f"E.M José Padin Mouta, cursando atualmente o(a) "
                f"<strong><u>{serie}</u></strong> no horário de aula: "
                f"<strong><u>{horario}</u></strong>."
            )

    elif tipo == "Transferencia":
        titulo = "Declaração de Transferência"
        if session.get("declaracao_tipo") == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, matriculado(a) no segmento de "
                f"<strong><u>Educação de Jovens e Adultos (EJA)</u></strong> da "
                f"E.M José Padin Mouta, solicitou transferência desta unidade escolar "
                f"na data de hoje, estando apto(a) a cursar o(a) "
                f"<strong><u>{serie}</u></strong>."
            )
        else:
            serie_mod = re.sub(r"^(\d+º).*", r"\1 ano", str(serie))
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) responsável do(a) "
                f"aluno(a) <strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, compareceu a nossa "
                f"unidade escolar e solicitou transferência na data de hoje, "
                f"o aluno está apto(a) a cursar o(a) "
                f"<strong><u>{serie_mod}</u></strong>."
            )

    elif tipo == "Conclusão":
        titulo = "Declaração de Conclusão"

        if session.get("declaracao_tipo") == "EJA":
            # Próxima série/etapa
            mapping = {
                "1ª SÉRIE E.F": "2ª SÉRIE E.F",
                "2ª SÉRIE E.F": "3ª SÉRIE E.F",
                "3ª SÉRIE E.F": "4ª SÉRIE E.F",
                "4ª SÉRIE E.F": "5ª SÉRIE E.F",
                "5ª SÉRIE E.F": "6ª SÉRIE E.F",
                "6ª SÉRIE E.F": "7ª SÉRIE E.F",
                "7ª SÉRIE E.F": "8ª SÉRIE E.F",
                "8ª SÉRIE E.F": "1ª SÉRIE E.M",
                "1ª SÉRIE E.M": "2ª SÉRIE E.M",
                "2ª SÉRIE E.M": "3ª SÉRIE E.M",
                "3ª SÉRIE E.M": "ENSINO SUPERIOR",
            }
            series_text = mapping.get(str(serie).upper(), "a série subsequente")

            # Se tiver semestre na planilha, só acrescenta como complemento
            semestre_parte = (
                f", no <strong><u>{semestre_texto}</u></strong>"
                if semestre_texto
                else ""
            )

            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do {ra_label} "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu com êxito o(a) "
                f"<strong><u>{serie}</u></strong>{semestre_parte} no segmento de "
                f"<strong><u>Educação de Jovens e Adultos (EJA)</u></strong> da "
                f"E.M José Padin Mouta, estando apto(a) a ingressar no(na) "
                f"<strong><u>{series_text}</u></strong>."
            )
        else:
            match = re.search(r"(\d+)º\s*ano", str(serie))
            next_year = int(match.group(1)) + 1 if match else None
            series_text = f"{next_year}º ano" if next_year else "a série subsequente"

            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu com êxito o(a) "
                f"<strong><u>{serie}</u></strong>, estando apto(a) a ingressar no(na) "
                f"<strong><u>{series_text}</u></strong>."
            )

    # ------------------------------------------------------
    # 4) OBSERVAÇÕES / HISTÓRICO / BOLSA FAMÍLIA
    # ------------------------------------------------------
    valor_bolsa = str(row.get("BOLSA FAMILIA", "")).strip().upper()

    if deve_historico or (valor_bolsa == "SIM" and tipo != "Escolaridade"):
        declaracao_text += "<br><br><strong>Observações:</strong><br>"
        declaracao_text += (
            '<label class="checkbox-label" '
            "style='display:block;text-align:justify;font-size:14px;'>"
        )

        # Histórico escolar pendente
        if deve_historico:
            declaracao_text += '<span class="warning-icon">&#9888;</span> '
            declaracao_text += (
                "O aluno deve o histórico escolar da unidade anterior:<br><br>"
            )

            if unidade_anterior:
                unidade_anterior = " ".join(str(unidade_anterior).strip().split())
                esc_df = None
                if escolas_df is not None:
                    try:
                        esc_df = escolas_df[
                            escolas_df.iloc[:, 3].str.upper()
                            == unidade_anterior.upper()
                        ]
                    except Exception:
                        esc_df = None

                if esc_df is not None and not esc_df.empty:
                    unidade_nome = esc_df.iloc[0, 3]
                    inep = esc_df.iloc[0, 4]
                    municipio = esc_df.iloc[0, 2]
                    uf = esc_df.iloc[0, 1]
                    declaracao_text += (
                        "<div style='font-size:14px;'>"
                        f"<strong>Unidade:</strong> {unidade_nome}<br>"
                        f"<strong>INEP:</strong> {inep}<br>"
                        f"<strong>Cidade:</strong> {municipio}<br>"
                        f"<strong>Estado:</strong> {uf}<br><br>"
                        "</div>"
                    )
                else:
                    declaracao_text += (
                        f"<strong>Unidade:</strong> {unidade_anterior}<br><br>"
                    )

            declaracao_text += (
                "Após sua entrega, o documento será confeccionado em "
                "até 30 dias úteis.<br><br>"
            )

        # Bolsa Família (quando não for só escolaridade)
        if valor_bolsa == "SIM" and tipo != "Escolaridade":
            declaracao_text += (
                '<img src="/static/logos/bolsa_familia.jpg" '
                'alt="Bolsa Família" '
                'style="width:28px;vertical-align:middle;margin-right:5px;">'
                "O aluno é beneficiário do Programa Bolsa Família."
            )

        declaracao_text += "</label>"

    # ------------------------------------------------------
    # 5) HTML FINAL (DECLARAÇÃO ÚNICA)
    # ------------------------------------------------------
    base_template = f"""<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>{titulo} - E.M José Padin Mouta</title>
  <style>
    @page {{
      margin: 0;
    }}
    html, body {{
      margin: 0;
      padding: 0.5cm;
      font-family: 'Montserrat', sans-serif;
      font-size: 16px;
      line-height: 1.5;
      color: #333;
    }}
    .header {{
      text-align: center;
      border-bottom: 2px solid #283E51;
      padding-bottom: 5px;
      margin-bottom: 10px;
    }}
    .header h1 {{
      margin: 0;
      font-size: 24px;
      text-transform: uppercase;
      color: #283E51;
    }}
    .header p {{
      margin: 3px 0;
      font-size: 16px;
    }}
    .date {{
      text-align: right;
      font-size: 16px;
      margin-bottom: 10px;
    }}
    .content {{
      text-align: justify;
      margin-bottom: 10px;
      padding: 0 2cm;
      box-sizing: border-box;
      hyphens: auto;
      word-wrap: break-word;
      overflow-wrap: break-word;
    }}
    .content p {{
      white-space: normal !important;
      word-break: break-word !important;
      overflow-wrap: break-word !important;
      hyphens: auto !important;
    }}
    .signature {{
      text-align: center;
      margin: 0;
      padding: 0;
    }}
    .signature .line {{
      height: 1px;
      background-color: #333;
      width: 60%;
      margin: 0 auto 5px auto;
    }}
    .footer {{
      text-align: center;
      border-top: 2px solid #283E51;
      padding-top: 5px;
      margin: 0;
      font-size: 14px;
      color: #555;
    }}
    .warning-icon {{
      display: inline-block;
      width: 18px;
      height: 18px;
      color: red;
      font-weight: bold;
      font-size: 18px;
      line-height: 18px;
      vertical-align: middle;
      user-select: none;
    }}

    @media print {{
      .no-print {{ display: none !important; }}
      body {{
        margin: 0;
        padding: 1.5cm 1.5cm;
        font-size: 14px;
        font-family: 'Montserrat', sans-serif;
        color: #000;
      }}
      .declaration-bottom {{
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
      }}
      .date {{
        margin: 1cm 0 1cm 0;
        text-align: right;
        hyphens: none !important;
      }}
      .content {{
        margin: 0;
        padding: 0;
        text-align: justify !important;
        hyphens: none !important;
        white-space: normal !important;
        word-wrap: break-word !important;
        overflow-wrap: break-word !important;
      }}
      .content p {{
        margin: 0 0 1em 0;
        text-align: justify !important;
        hyphens: none !important;
      }}
      body, .content, .content p {{
        width: auto !important;
        max-width: none !important;
      }}
    }}

    .content, .content p, .date {{
      hyphens: none !important;
      word-break: normal !important;
      overflow-wrap: normal !important;
    }}

    {additional_css}

    .checkbox-label {{
      display: flex;
      align-items: center;
      gap: 8px;
      text-align: left !important;
      font-size: 14px;
      margin-top: 8px;
      margin-bottom: 8px;
      flex-wrap: wrap;
    }}

    header {{
      background: linear-gradient(90deg, #283E51, #4B79A1);
      color: #fff;
      padding: 20px;
      text-align: center;
      border-bottom: 3px solid #1d2d3a;
      border-radius: 0 0 15px 15px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
  </style>
</head>
<body>
  <div class="declaration-container">
    <div class="header">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <img src="/static/logos/escola.png" alt="Escola Logo" style="height:80px;">
        <div>
          <h1>Secretaria de Educação</h1>
          <p>E.M José Padin Mouta</p>
          <p>Município da Estância Balneária de Praia Grande</p>
          <p>Estado de São Paulo</p>
        </div>
        <img src="/static/logos/municipio.png" alt="Município Logo" style="height:80px;">
      </div>
    </div>
    <div class="date">
      <p>{data_extenso}</p>
    </div>
    <div class="content">
      <h2 style="text-align:center;text-transform:uppercase;color:#283E51;">
        {titulo}
      </h2>
      <p>{declaracao_text}</p>
    </div>
    <div class="declaration-bottom">
      <div class="signature">
        <div class="line"></div>
        <p>Luciana Rocha Augustinho</p>
        <p>Diretora da Unidade Escolar</p>
      </div>
      <div class="footer">
        <p>Rua: Bororós, nº 150, Vila Tupi, Praia Grande - SP, CEP: 11703-390</p>
        <p>Telefone: 3496-5321 | E-mail: em.padin@praiagrande.sp.gov.br</p>
      </div>
    </div>
  </div>
  <div class="no-print" style="text-align:center;margin-top:20px;">
    <button onclick="window.print()" class="print-button">Imprimir Declaração</button>
  </div>
</body>
</html>
"""
    return base_template


from datetime import datetime

def gerar_declaracao_personalizada(dados):
    """
    Gera o HTML de declarações personalizadas (Conclusão, Matrícula cancelada
    ou Não Comparecimento - NCOM), utilizando o mesmo layout das demais
    declarações.

    Espera um dicionário `dados` com:
      - segmento: 'Fundamental' ou 'EJA' (ou campo segmento_personalizado)
      - nome_aluno
      - data_nascimento (YYYY-MM-DD ou DD/MM/YYYY)
      - ra
      - tipo_declaracao: 'Conclusao', 'MatriculaCancelada', 'NCOM'
    e campos adicionais conforme o tipo.
    """

    # Helpers internos
    def _get_str(key, default=""):
        return (dados.get(key) or default).strip()

    def _normalizar_semestre(*keys):
        """
        Procura o semestre em várias chaves possíveis e devolve o primeiro
        valor não vazio, já "stripped".
        """
        for k in keys:
            v = dados.get(k)
            if v is None:
                continue
            s = str(v).strip()
            if s:
                return s
        return ""

    nome = _get_str("nome_aluno")
    ra = _get_str("ra")
    data_nasc_raw = _get_str("data_nascimento")

    # aceita tanto 'segmento' quanto 'segmento_personalizado'
    seg_raw = dados.get("segmento") or dados.get("segmento_personalizado") or "Fundamental"
    seg_norm = str(seg_raw).strip().lower()
    if seg_norm in ("fundamental", "fund", "ef", "ensino fundamental"):
        segmento = "Fundamental"
    else:
        segmento = "EJA"

    # Normaliza data de nascimento para DD/MM/AAAA
    data_nasc = "Desconhecida"
    if data_nasc_raw:
        for fmt in ("%Y-%m-%d", "%d/%m/%Y"):
            try:
                dt = datetime.strptime(data_nasc_raw, fmt)
                data_nasc = dt.strftime("%d/%m/%Y")
                break
            except Exception:
                continue

    # Rótulo do segmento e preposição correta (do / da)
    if segmento == "Fundamental":
        segmento_label = "Ensino Fundamental"
        prep_segmento = "do"
    else:
        segmento_label = "Educação de Jovens e Adultos (EJA)"
        prep_segmento = "da"

    # Aceita tanto 'tipo_declaracao' quanto 'tipo_declaracao_personalizada'
    tipo_decl_raw = dados.get("tipo_declaracao") or dados.get("tipo_declaracao_personalizada")
    tipo_decl = (tipo_decl_raw or "").strip().lower()

    declaracao_text = ""
    titulo = ""

    # ------------------------------------------------------
    # 1) MONTAGEM DO TEXTO DA DECLARAÇÃO
    # ------------------------------------------------------
    if tipo_decl in ("conclusao", "conclusão"):
        titulo = "Declaração de Conclusão"
        ano_serie = _get_str("ano_serie_concluida")
        ano_conclusao = _get_str("ano_conclusao")

        # Pode vir 'sim'/'nao', True/False, 'on', etc.
        deve_hist_val_raw = dados.get("deve_historico_unidade")
        deve_hist_str = str(deve_hist_val_raw or "").strip().lower()
        deve_hist_unidade = deve_hist_str in ("sim", "1", "true", "on")

        if segmento == "Fundamental":
            # FUNDAMENTAL – período letivo anual, sem semestre
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu o(a) "
                f"<strong><u>{ano_serie}</u></strong> {prep_segmento} "
                f"<strong><u>{segmento_label}</u></strong>, no ano letivo de "
                f"<strong><u>{ano_conclusao}</u></strong>, nesta unidade escolar."
            )
        else:
            # EJA – usa semestre na conclusão
            semestre_conclusao = _normalizar_semestre(
                "semestre_conclusao",        # radio da tela de conclusão
                "semestre_conclusao_opcao",  # fallback caso mude o name no HTML
                "semestre_matricula",        # fallback se reusar semestre da matrícula
                "semestre_matricula_opcao",
            )

            if semestre_conclusao:
                periodo_conclusao = (
                    f"no <strong><u>{semestre_conclusao}</u></strong> do ano de "
                    f"<strong><u>{ano_conclusao}</u></strong>"
                )
            else:
                periodo_conclusao = (
                    f"no ano letivo de <strong><u>{ano_conclusao}</u></strong>"
                )

            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, concluiu o(a) "
                f"<strong><u>{ano_serie}</u></strong> no segmento de "
                f"<strong><u>Educação de Jovens e Adultos (EJA)</u></strong>, "
                f"{periodo_conclusao}, nesta unidade escolar."
            )

        if deve_hist_unidade:
            declaracao_text += (
                " Ressalta-se que consta, junto a esta unidade escolar, "
                "pendência de histórico escolar referente ao(à) aluno(a) citado(a)."
            )

    elif tipo_decl in ("matriculacancelada", "matricula cancelada", "matricula_cancelada"):
        titulo = "Declaração de Matrícula Cancelada"
        ano_serie = _get_str("ano_serie_matricula")
        ano_matricula = _get_str("ano_matricula")
        semestre_matricula = _normalizar_semestre(
            "semestre_matricula",
            "semestre_matricula_opcao",
        )

        # Só EJA usa semestre na frase; Fundamental fica só com o ano
        if segmento == "EJA" and semestre_matricula:
            periodo_matricula = (
                f"no <strong><u>{semestre_matricula}</u></strong> do ano de "
                f"<strong><u>{ano_matricula}</u></strong>"
            )
        else:
            periodo_matricula = (
                f"no ano de <strong><u>{ano_matricula}</u></strong>"
            )

        if segmento == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, esteve matriculado(a) no(a) "
                f"<strong><u>{ano_serie}</u></strong> no segmento de "
                f"<strong><u>Educação de Jovens e Adultos (EJA)</u></strong>, "
                f"{periodo_matricula}, nesta unidade escolar, tendo sua matrícula cancelada."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, esteve matriculado(a) no(a) "
                f"<strong><u>{ano_serie}</u></strong> {prep_segmento} "
                f"<strong><u>{segmento_label}</u></strong>, {periodo_matricula}, "
                "nesta unidade escolar, tendo sua matrícula cancelada."
            )

    elif tipo_decl == "ncom":
        titulo = "Declaração de Não Comparecimento (NCOM)"
        ano_serie = _get_str("ano_serie_vaga")
        ano_ref = _get_str("ano_referencia_ncom")
        semestre_ref = _normalizar_semestre(
            "semestre_referencia_ncom",
            "semestre_referencia",
        )

        # Semestre só faz sentido na EJA; Fundamental só ano
        if segmento == "EJA" and semestre_ref:
            periodo_ref = (
                f"para o <strong><u>{semestre_ref}</u></strong> do ano de "
                f"<strong><u>{ano_ref}</u></strong>"
            )
        else:
            periodo_ref = (
                f"para o ano de <strong><u>{ano_ref}</u></strong>"
            )

        if segmento == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, teve vaga destinada ao(à) "
                f"<strong><u>{ano_serie}</u></strong> no segmento de "
                f"<strong><u>Educação de Jovens e Adultos (EJA)</u></strong>, "
                f"{periodo_ref} nesta unidade escolar. Todavia, o(a) aluno(a) "
                "não compareceu à unidade escolar, sendo considerado(a) NCOM – "
                "Não Comparecimento, motivo pelo qual a vaga foi cancelada nesta "
                "unidade escolar."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) "
                f"<strong><u>{nome}</u></strong>, portador(a) do RA "
                f"<strong><u>{ra}</u></strong>, nascido(a) em "
                f"<strong><u>{data_nasc}</u></strong>, teve vaga destinada ao(à) "
                f"<strong><u>{ano_serie}</u></strong> {prep_segmento} "
                f"<strong><u>{segmento_label}</u></strong>, {periodo_ref} "
                "nesta unidade escolar. Todavia, o(a) aluno(a) não compareceu à unidade "
                "escolar, sendo considerado(a) NCOM – Não Comparecimento, motivo pelo qual "
                "a vaga foi cancelada nesta unidade escolar."
            )

    else:
        # Tipo desconhecido
        return None

    # ------------------------------------------------------
    # 2) DATA POR EXTENSO
    # ------------------------------------------------------
    now = datetime.now()
    meses = {
        1: "janeiro",
        2: "fevereiro",
        3: "março",
        4: "abril",
        5: "maio",
        6: "junho",
        7: "julho",
        8: "agosto",
        9: "setembro",
        10: "outubro",
        11: "novembro",
        12: "dezembro",
    }
    mes = meses[now.month].capitalize()
    data_extenso = f"Praia Grande, {now.day:02d} de {mes} de {now.year}"

    additional_css = """
.print-button {
  background-color: #283E51;
  color: #fff;
  border: none;
  padding: 10px 20px;
  border-radius: 5px;
  cursor: pointer;
  margin-top: 20px;
}
.print-button:hover {
  background-color: #1d2d3a;
}
"""

    # ------------------------------------------------------
    # 3) HTML FINAL (MESMO LAYOUT DAS OUTRAS DECLARAÇÕES)
    # ------------------------------------------------------
    base_template = f"""<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>{titulo} - E.M José Padin Mouta</title>
  <style>
    @page {{
      margin: 0;
    }}
    html, body {{
      margin: 0;
      padding: 0.5cm;
      font-family: 'Montserrat', sans-serif;
      font-size: 16px;
      line-height: 1.5;
      color: #333;
    }}
    .header {{
      text-align: center;
      border-bottom: 2px solid #283E51;
      padding-bottom: 5px;
      margin-bottom: 10px;
    }}
    .header h1 {{
      margin: 0;
      font-size: 24px;
      text-transform: uppercase;
      color: #283E51;
    }}
    .header p {{
      margin: 3px 0;
      font-size: 16px;
    }}
    .date {{
      text-align: right;
      font-size: 16px;
      margin-bottom: 10px;
    }}
    .content {{
      text-align: justify;
      margin-bottom: 10px;
      padding: 0 2cm;
      box-sizing: border-box;
      hyphens: auto;
      word-wrap: break-word;
      overflow-wrap: break-word;
    }}
    .content p {{
      white-space: normal !important;
      word-break: break-word !important;
      overflow-wrap: break-word !important;
      hyphens: auto !important;
    }}
    .signature {{
      text-align: center;
      margin: 0;
      padding: 0;
    }}
    .signature .line {{
      height: 1px;
      background-color: #333;
      width: 60%;
      margin: 0 auto 5px auto;
    }}
    .footer {{
      text-align: center;
      border-top: 2px solid #283E51;
      padding-top: 5px;
      margin: 0;
      font-size: 14px;
      color: #555;
    }}
    .warning-icon {{
      display: inline-block;
      width: 18px;
      height: 18px;
      color: red;
      font-weight: bold;
      font-size: 18px;
      line-height: 18px;
      vertical-align: middle;
      user-select: none;
    }}

    @media print {{
      .no-print {{ display: none !important; }}
      body {{
        margin: 0;
        padding: 1.5cm 1.5cm;
        font-size: 14px;
        font-family: 'Montserrat', sans-serif;
        color: #000;
      }}
      .declaration-bottom {{
        position: fixed;
        bottom: 0;
        left: 0;
        width: 100%;
      }}
      .date {{
        margin: 1cm 0 1cm 0;
        text-align: right;
        hyphens: none !important;
      }}
      .content {{
        margin: 0;
        padding: 0;
        text-align: justify !important;
        hyphens: none !important;
        white-space: normal !important;
        word-wrap: break-word !important;
        overflow-wrap: break-word !important;
      }}
      .content p {{
        margin: 0 0 1em 0;
        text-align: justify !important;
        hyphens: none !important;
      }}
      body, .content, .content p {{
        width: auto !important;
         max-width: none !important;
      }}
    }}

    .content, .content p, .date {{
      hyphens: none !important;
      word-break: normal !important;
      overflow-wrap: normal !important;
    }}

    {additional_css}

    .checkbox-label {{
      display: flex;
      align-items: center;
      gap: 8px;
      text-align: left !important;
      font-size: 14px;
      margin-top: 8px;
      margin-bottom: 8px;
      flex-wrap: wrap;
    }}

    header {{
      background: linear-gradient(90deg, #283E51, #4B79A1);
      color: #fff;
      padding: 20px;
      text-align: center;
      border-bottom: 3px solid #1d2d3a;
      border-radius: 0 0 15px 15px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }}
  </style>
</head>
<body>
  <div class="declaration-container">
    <div class="header">
      <div style="display:flex;justify-content:space-between;align-items:center;">
        <img src="/static/logos/escola.png" alt="Escola Logo" style="height:80px;">
        <div>
          <h1>Secretaria de Educação</h1>
          <p>E.M José Padin Mouta</p>
          <p>Município da Estância Balneária de Praia Grande</p>
          <p>Estado de São Paulo</p>
        </div>
        <img src="/static/logos/municipio.png" alt="Município Logo" style="height:80px;">
      </div>
    </div>
    <div class="date">
      <p>{data_extenso}</p>
    </div>
    <div class="content">
      <h2 style="text-align:center;text-transform:uppercase;color:#283E51;">
        {titulo}
      </h2>
      <p>{declaracao_text}</p>
    </div>
    <div class="declaration-bottom">
      <div class="signature">
        <div class="line"></div>
        <p>Luciana Rocha Augustinho</p>
        <p>Diretora da Unidade Escolar</p>
      </div>
      <div class="footer">
        <p>Rua: Bororós, nº 150, Vila Tupi, Praia Grande - SP, CEP: 11703-390</p>
        <p>Telefone: 3496-5321 | E-mail: em.padin@praiagrande.sp.gov.br</p>
      </div>
    </div>
  </div>
  <div class="no-print" style="text-align:center;margin-top:20px;">
    <button onclick="window.print()" class="print-button">Imprimir Declaração</button>
  </div>
</body>
</html>
"""
    return base_template

# ==========================================================
#  AUTENTICAÇÃO / DASHBOARD
# ==========================================================

@app.route("/login", methods=["GET", "POST"])
def login_route():
    error = None
    if request.method == "POST":
        token = request.form.get("token")
        if token == ACCESS_TOKEN:
            session["logged_in"] = True
            if "lista_fundamental" not in session or "lista_eja" not in session:
                return redirect(url_for("upload_listas"))
            return redirect(url_for("dashboard"))
        else:
            error = "Token inválido. Tente novamente."

    return render_template("login.html", error=error)


@app.route("/logout")
def logout_route():
    session.clear()
    return redirect(url_for("login_route"))


@app.route("/upload_listas", methods=["GET", "POST"])
@login_required
def upload_listas():
    if request.method == "POST":
        fundamental_file = request.files.get("lista_fundamental")
        eja_file = request.files.get("lista_eja")

        if not fundamental_file or fundamental_file.filename == "":
            flash("Selecione a Lista Piloto - REGULAR - 2025", "error")
            return redirect(url_for("upload_listas"))

        if not eja_file or eja_file.filename == "":
            flash("Selecione a Lista Piloto - EJA - 1º SEM - 2025", "error")
            return redirect(url_for("upload_listas"))

        fundamental_filename = secure_filename(
            f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename
        )
        eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_" + eja_file.filename)

        fundamental_path = os.path.join(app.config["UPLOAD_FOLDER"], fundamental_filename)
        eja_path = os.path.join(app.config["UPLOAD_FOLDER"], eja_filename)

        fundamental_file.save(fundamental_path)
        eja_file.save(eja_path)

        session["lista_fundamental"] = fundamental_path
        session["lista_eja"] = eja_path

        flash("Listas carregadas com sucesso.", "success")
        return redirect(url_for("dashboard"))

    return render_template("upload_listas.html")


@app.route("/", methods=["GET"])
@login_required
def dashboard():
    return render_template("dashboard.html")


# ==========================================================
#  CARTEIRINHAS – ROTA PRINCIPAL
# ==========================================================

@app.route("/carteirinhas", methods=["GET", "POST"])
@login_required
def carteirinhas():
    if request.method == "POST":
        file_path = None

        # Se veio um arquivo novo no POST, salva e guarda na sessão
        if "excel_file" in request.files and request.files["excel_file"].filename != "":
            file = request.files["excel_file"]
            filename = secure_filename(file.filename)
            unique_filename = f"carteirinhas_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], unique_filename)
            file.save(file_path)

            # guarda para reutilizar depois sem precisar reenviar
            session["lista_fundamental"] = file_path
        else:
            # tenta reaproveitar a última lista usada
            file_path = session.get("lista_fundamental")

        if not file_path or not os.path.exists(file_path):
            flash("Nenhum arquivo selecionado. Envie a lista piloto do Fundamental.", "info")
            return redirect(url_for("carteirinhas"))

        flash("Gerando carteirinhas. Aguarde...", "info")
        html_result = gerar_html_carteirinhas(file_path)
        return html_result

    # GET – tela de upload / gerenciamento de fotos
    return render_template("carteirinhas.html")


# ==========================================================
#  DECLARAÇÕES – CONCLUSÃO 5º ANO (LOTE)
# ==========================================================

@app.route("/declaracao/conclusao_5ano")
@login_required
def declaracao_conclusao_5ano():
    """
    Gera TODAS as declarações de CONCLUSÃO dos alunos de 5º ano (Fundamental)
    em um único HTML com quebras de página para impressão/PDF.
    Usa o template declaracao_conclusao_5ano.html.
    """
    if session.get("declaracao_tipo") != "Fundamental":
        flash(
            "As declarações em lote de 5º ano estão disponíveis apenas para o Fundamental.",
            "error",
        )
        return redirect(url_for("declaracao_tipo"))

    # Prioriza o caminho utilizado na última declaração singular,
    # mas também aceita a lista_fundamental
    file_path = session.get("declaracao_excel") or session.get("lista_fundamental")

    # Se o servidor “acordou” e o arquivo não existe mais, força o usuário a reenviar a lista
    if not file_path or not os.path.exists(file_path):
        flash(
            "Arquivo Excel do Fundamental não encontrado. "
            "Anexe a lista piloto novamente pela tela de declarações.",
            "error",
        )
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    # Lê a LISTA CORRIDA
    planilha = pd.read_excel(file_path, sheet_name="LISTA CORRIDA")
    planilha.columns = [c.strip().upper() for c in planilha.columns]

    def format_rm(x):
        try:
            return str(int(float(x)))
        except Exception:
            return str(x)

    planilha["RM_str"] = planilha["RM"].apply(format_rm)

    registros = []

    for _, row in planilha.iterrows():
        rm_str = str(row.get("RM_str", "")).strip()
        if rm_str in ("", "0"):
            continue

        serie_raw = str(row.get("SÉRIE", "")).strip()
        if not serie_raw:
            continue

        # Apenas 5º ano (5ºA, 5º B, 5º ano A, etc.)
        if "5º" not in serie_raw and "5°" not in serie_raw:
            continue

        nome = str(row.get("NOME", "")).strip()
        ra = str(row.get("RA", "")).strip()

        data_nasc_val = row.get("DATA NASC.")
        if pd.notna(data_nasc_val):
            try:
                data_nasc_dt = pd.to_datetime(data_nasc_val, errors="coerce")
                data_nasc = (
                    data_nasc_dt.strftime("%d/%m/%Y")
                    if pd.notna(data_nasc_dt)
                    else "Desconhecida"
                )
            except Exception:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"

        horario = str(row.get("HORÁRIO", "")).strip()
        if not horario:
            horario = "Desconhecido"

        # Formata série para "5º ano A"
        serie_fmt = serie_raw
        try:
            serie_fmt = re.sub(r"(\d+º)\s*([A-Za-z])", r"\1 ano \2", serie_fmt)
        except Exception:
            pass

        # Série subsequente (6º ano)
        series_text = "a série subsequente"
        m = re.search(r"(\d+)º", serie_fmt)
        if m:
            try:
                next_year = int(m.group(1)) + 1
                series_text = f"{next_year}º ano"
            except Exception:
                pass

        # Bolsa Família
        valor_bolsa = str(row.get("BOLSA FAMILIA", "")).strip().upper()

        declaracao_text = (
            f"Declaro, para os devidos fins, que o(a) aluno(a) "
            f"<strong><u>{nome}</u></strong>, "
            f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em "
            f"<strong><u>{data_nasc}</u></strong>, concluiu com êxito o "
            f"<strong><u>{serie_fmt}</u></strong>, estando apto(a) a ingressar no "
            f"<strong><u>{series_text}</u></strong>."
        )

        if valor_bolsa == "SIM":
            declaracao_text += "<br><br><strong>Observações:</strong><br>"
            declaracao_text += (
                '<label class="checkbox-label" '
                'style="display: block; text-align: justify; font-size:14px;">'
            )
            declaracao_text += (
                f'<img src="{url_for("static", filename="logos/bolsa_familia.jpg")}" '
                'alt="Bolsa Família" '
                'style="width:28px; vertical-align:middle; margin-right:5px;">'
                "O aluno é beneficiário do Programa Bolsa Família."
            )
            declaracao_text += "</label>"

        registros.append(
            {
                "nome": nome,
                "ra": ra,
                "data_nasc": data_nasc,
                "serie_fmt": serie_fmt,
                "series_text": series_text,
                "texto": declaracao_text,
            }
        )

    if not registros:
        flash("Nenhum aluno de 5º ano encontrado na lista piloto.", "error")
        return redirect(url_for("declaracao_tipo", segmento="Fundamental"))

    now = datetime.now()
    meses = {
        1: "janeiro",
        2: "fevereiro",
        3: "março",
        4: "abril",
        5: "maio",
        6: "junho",
        7: "julho",
        8: "agosto",
        9: "setembro",
        10: "outubro",
        11: "novembro",
        12: "dezembro",
    }
    mes = meses[now.month].capitalize()
    data_extenso = f"Praia Grande, {now.day:02d} de {mes} de {now.year}"
    titulo = "Declaração de Conclusão"

    return render_template(
        "declaracao_conclusao_5ano.html",
        registros=registros,
        data_extenso=data_extenso,
        titulo=titulo,
        total=len(registros),
    )


# ==========================================================
#  DECLARAÇÕES – TELA ÚNICA (Fundamental / EJA / Personalizada)
# ==========================================================

@app.route("/declaracao/tipo", methods=["GET", "POST"])
@login_required
def declaracao_tipo():
    """
    Tela única para:
    - escolher segmento (Fundamental / EJA / Personalizado)
    - anexar lista piloto (se ainda não houver) – apenas Fundamental/EJA
    - escolher aluno e tipo de declaração (Fundamental/EJA) OU
      preencher dados manuais (Personalizada)
    - gerar a declaração singular
    """

    # --------------------------------------
    # POST: GERAR DECLARAÇÃO
    # --------------------------------------
    if request.method == "POST":
        # Verifica se é fluxo normal (Fundamental/EJA) ou fluxo personalizado
        modo_declaracao = request.form.get("modo_declaracao")

        # -------------------------------
        # FLUXO: DECLARAÇÃO PERSONALIZADA
        # -------------------------------
        if modo_declaracao == "personalizada":
            segmento_pers = request.form.get("segmento_personalizado")
            if segmento_pers not in ("Fundamental", "EJA"):
                flash(
                    "Selecione o segmento (Ensino Fundamental ou EJA) na declaração personalizada.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

            nome_aluno = (request.form.get("nome_aluno") or "").strip()
            data_nascimento = request.form.get("data_nascimento")
            ra = (request.form.get("ra") or "").strip()
            tipo_pers = request.form.get("tipo_declaracao_personalizada")

            if (
                not nome_aluno
                or not data_nascimento
                or not ra
                or tipo_pers not in ("Conclusao", "MatriculaCancelada", "NCOM")
            ):
                flash(
                    "Preencha todos os dados obrigatórios da declaração personalizada.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

            dados_personalizados = {
                "segmento": segmento_pers,
                "nome_aluno": nome_aluno,
                "data_nascimento": data_nascimento,
                "ra": ra,
                "tipo_declaracao": tipo_pers,
            }

            # Campos específicos de cada tipo
            if tipo_pers == "Conclusao":
                ano_serie_concluida = (request.form.get("ano_serie_concluida") or "").strip()
                ano_conclusao = (request.form.get("ano_conclusao") or "").strip()
                deve_hist_unidade = request.form.get("deve_historico_unidade")
                semestre_conclusao = (request.form.get("semestre_conclusao") or "").strip()

                # validação básica
                campos_invalidos = (
                    not ano_serie_concluida
                    or not ano_conclusao
                    or deve_hist_unidade not in ("Sim", "Não")
                )

                # para EJA, o semestre é obrigatório
                if segmento_pers == "EJA" and not semestre_conclusao:
                    campos_invalidos = True

                if campos_invalidos:
                    flash(
                        "Preencha todos os campos da declaração personalizada de conclusão "
                        "(para EJA é obrigatório informar o semestre).",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

                dados_personalizados.update(
                    {
                        "ano_serie_concluida": ano_serie_concluida,
                        "ano_conclusao": ano_conclusao,
                        "deve_historico_unidade": (deve_hist_unidade == "Sim"),
                        "semestre_conclusao": semestre_conclusao,
                    }
                )

            elif tipo_pers == "MatriculaCancelada":
                ano_serie_matricula = (request.form.get("ano_serie_matricula") or "").strip()
                ano_matricula = (request.form.get("ano_matricula") or "").strip()
                semestre_matricula = (request.form.get("semestre_matricula") or "").strip()

                if not ano_serie_matricula or not ano_matricula or not semestre_matricula:
                    flash(
                        "Preencha todos os campos da declaração de matrícula cancelada.",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

                dados_personalizados.update(
                    {
                        "ano_serie_matricula": ano_serie_matricula,
                        "ano_matricula": ano_matricula,
                        "semestre_matricula": semestre_matricula,
                    }
                )

            elif tipo_pers == "NCOM":
                ano_serie_vaga = (request.form.get("ano_serie_vaga") or "").strip()
                ano_referencia_ncom = (request.form.get("ano_referencia_ncom") or "").strip()
                semestre_referencia_ncom = (
                    request.form.get("semestre_referencia_ncom") or ""
                ).strip()

                if not ano_serie_vaga or not ano_referencia_ncom:
                    flash(
                        "Preencha todos os campos obrigatórios da declaração de Não Comparecimento (NCOM).",
                        "error",
                    )
                    return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

                dados_personalizados.update(
                    {
                        "ano_serie_vaga": ano_serie_vaga,
                        "ano_referencia_ncom": ano_referencia_ncom,
                        "semestre_referencia_ncom": semestre_referencia_ncom,
                    }
                )

            # Gera a declaração personalizada
            declaracao_html = gerar_declaracao_personalizada(dados_personalizados)

            if declaracao_html is None:
                flash(
                    "Não foi possível gerar a declaração personalizada. Verifique os dados informados.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento="Personalizado"))

            return declaracao_html

        # -------------------------------
        # FLUXO NORMAL: FUNDAMENTAL / EJA
        # -------------------------------
        segmento = request.form.get("segmento_escolhido")
        if segmento not in ("Fundamental", "EJA"):
            flash("Selecione se a declaração é do Fundamental ou EJA antes de gerar.", "error")
            return redirect(url_for("declaracao_tipo"))

        # Campos básicos já lidos aqui para podermos tratar o caso "upload apenas"
        rm = (request.form.get("rm") or "").strip()
        tipo = (request.form.get("tipo") or "").strip()
        deve_historico_str = request.form.get("deve_historico")

        unidade_select = (request.form.get("unidade_anterior_select") or "").strip()
        unidade_manual = (request.form.get("unidade_anterior_manual") or "").strip()
        unidade_anterior = unidade_select or unidade_manual

        # Upload ou reaproveita lista da sessão
        file_path = None
        excel_file = request.files.get("excel_file")
        novo_upload = excel_file is not None and excel_file.filename

        if novo_upload:
            # Novo envio de lista piloto (usado inclusive para “recuperar” após hibernação)
            filename = secure_filename(excel_file.filename)
            unique_filename = f"declaracao_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config["UPLOAD_FOLDER"], unique_filename)
            excel_file.save(file_path)

            if segmento == "Fundamental":
                session["lista_fundamental"] = file_path
            else:
                session["lista_eja"] = file_path
        else:
            # Reaproveita o caminho já salvo na sessão
            if segmento == "Fundamental":
                file_path = session.get("lista_fundamental")
            else:
                file_path = session.get("lista_eja")

        # Se ainda assim não houver arquivo válido, não há o que fazer
        if not file_path or not os.path.exists(file_path):
            flash(
                "Nenhuma lista piloto encontrada para este segmento. Anexe o arquivo em Excel.",
                "error",
            )
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        # Atualiza infos de declaração na sessão (usadas em outras rotas, ex: 5º ano)
        session["declaracao_tipo"] = segmento
        session["declaracao_excel"] = file_path

        # --------------------------------------------------
        # CASO ESPECIAL: usuário acabou de ENVIAR APENAS A LISTA
        # (sem RM / tipo) – típico após o servidor “acordar”
        # --------------------------------------------------
        if novo_upload and (not rm or not tipo):
            flash(
                "Lista piloto carregada com sucesso. "
                "Agora selecione o aluno e o tipo de declaração.",
                "success",
            )
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        # --------------------------------------------------
        # A partir daqui, fluxo normal de geração de declaração
        # --------------------------------------------------
        if not rm or not tipo:
            flash("Escolha o aluno e o tipo de declaração.", "error")
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        # Valida histórico quando for Transferência ou Conclusão
        if tipo in ("Transferencia", "Conclusão"):
            if deve_historico_str not in ("sim", "nao"):
                flash("Por favor, responda se o aluno deve o histórico escolar.", "error")
                return redirect(url_for("declaracao_tipo", segmento=segmento))

            if deve_historico_str == "sim" and not unidade_anterior:
                flash(
                    "Informe a unidade escolar anterior para a qual o aluno deve o histórico.",
                    "error",
                )
                return redirect(url_for("declaracao_tipo", segmento=segmento))

            deve_historico = deve_historico_str == "sim"
        else:
            deve_historico = False
            unidade_anterior = ""

        declaracao_html = gerar_declaracao_escolar(
            file_path=file_path,
            rm=rm,
            tipo=tipo,
            deve_historico=deve_historico,
            unidade_anterior=unidade_anterior,
        )

        if declaracao_html is None:
            flash("Aluno não encontrado na lista piloto.", "error")
            return redirect(url_for("declaracao_tipo", segmento=segmento))

        return declaracao_html

    # --------------------------------------
    # GET: EXIBE A TELA
    # --------------------------------------
    segmento = request.args.get("segmento")
    if segmento not in ("Fundamental", "EJA", "Personalizado"):
        segmento = None

    alunos = []
    tem_lista = False

    if segmento == "Fundamental":
        file_path = session.get("lista_fundamental")
        if file_path and os.path.exists(file_path):
            tem_lista = True
            session["declaracao_tipo"] = "Fundamental"
            session["declaracao_excel"] = file_path

            planilha = pd.read_excel(file_path, sheet_name="LISTA CORRIDA")

            def format_rm(x):
                try:
                    return str(int(float(x)))
                except Exception:
                    return str(x)

            planilha["RM_str"] = planilha["RM"].apply(format_rm)
            alunos_df = (
                planilha[planilha["RM_str"] != "0"][["RM_str", "NOME"]].drop_duplicates()
            )

            alunos = [
                {"rm": row["RM_str"], "nome": row["NOME"]}
                for _, row in alunos_df.iterrows()
            ]

    elif segmento == "EJA":
        file_path = session.get("lista_eja")
        if file_path and os.path.exists(file_path):
            tem_lista = True
            session["declaracao_tipo"] = "EJA"
            session["declaracao_excel"] = file_path

            df = pd.read_excel(file_path, sheet_name=0, header=None, skiprows=1)
            df["RM_str"] = df.iloc[:, 2].apply(
                lambda x: str(int(x)) if pd.notna(x) and float(x) != 0 else ""
            )
            df["NOME"] = df.iloc[:, 3]
            alunos_df = df[df["RM_str"] != ""][["RM_str", "NOME"]].drop_duplicates()

            alunos = [
                {"rm": row["RM_str"], "nome": row["NOME"]}
                for _, row in alunos_df.iterrows()
            ]

    # segmento == "Personalizado": não há lista piloto nem alunos pré-carregados
    dashboard_url = url_for("dashboard")
    conclusao_5ano_url = url_for("declaracao_conclusao_5ano")

    return render_template(
        "declaracao_tipo.html",
        segmento=segmento,
        tem_lista=tem_lista,
        alunos=alunos,
        dashboard_url=dashboard_url,
        conclusao_5ano_url=conclusao_5ano_url,
    )

# ==========================================================
#  UPLOAD DE FOTOS (CARTEIRINHAS)
# ==========================================================

@app.route("/upload_foto", methods=["POST"])
@login_required
def upload_foto():
    rm = (request.form.get("rm") or "").strip()
    if not rm:
        flash("RM não fornecido.", "error")
        return redirect(url_for("carteirinhas"))

    file = request.files.get("foto_file")
    if not file or file.filename == "":
        flash("Nenhuma foto selecionada.", "error")
        return redirect(url_for("carteirinhas"))

    if not allowed_file(file.filename):
        flash("Formato de imagem não permitido. Envie JPG, PNG, GIF ou BMP.", "error")
        return redirect(url_for("carteirinhas"))

    original_filename = secure_filename(file.filename)
    _, ext = os.path.splitext(original_filename)

    fotos_dir = os.path.join("static", "fotos")
    os.makedirs(fotos_dir, exist_ok=True)

    new_filename = secure_filename(f"{rm}{ext.lower()}")
    file_path = os.path.join(fotos_dir, new_filename)
    file.save(file_path)

    flash("Foto anexada com sucesso.", "success")
    return redirect(url_for("carteirinhas"))


@app.route("/upload_multiplas_fotos", methods=["POST"])
@login_required
def upload_multiplas_fotos():
    rms = request.form.getlist("rm[]")
    files = request.files.getlist("foto_file[]")

    if not files:
        flash("Nenhuma foto enviada.", "error")
        return redirect(url_for("carteirinhas"))

    fotos_dir = os.path.join("static", "fotos")
    os.makedirs(fotos_dir, exist_ok=True)

    total_salvas = 0
    total_ignoradas = 0

    for rm, file in zip(rms, files):
        rm = (rm or "").strip()

        if not rm or not file or file.filename == "" or not allowed_file(file.filename):
            total_ignoradas += 1
            continue

        original_filename = secure_filename(file.filename)
        _, ext = os.path.splitext(original_filename)
        new_filename = secure_filename(f"{rm}{ext.lower()}")
        file_path = os.path.join(fotos_dir, new_filename)
        file.save(file_path)
        total_salvas += 1

    if total_salvas:
        msg = f"Foto(s) anexada(s) com sucesso: {total_salvas} arquivo(s)."
        if total_ignoradas:
            msg += f" {total_ignoradas} arquivo(s) foram ignorados por RM ou formato inválido."
        flash(msg, "success")
    else:
        flash("Nenhuma foto válida foi enviada.", "error")

    return redirect(url_for("carteirinhas"))


@app.route("/upload_inline_foto", methods=["POST"])
@login_required
def upload_inline_foto():
    file = request.files.get("foto_file")
    rm = (request.form.get("rm") or "").strip()

    if not rm:
        return jsonify({"error": "RM não fornecido"}), 400

    if not file or file.filename == "":
        return jsonify({"error": "Nenhum arquivo enviado"}), 400

    if not allowed_file(file.filename):
        return jsonify({"error": "Formato de imagem não permitido"}), 400

    original_filename = secure_filename(file.filename)
    _, ext = os.path.splitext(original_filename)

    fotos_dir = os.path.join("static", "fotos")
    os.makedirs(fotos_dir, exist_ok=True)

    new_filename = secure_filename(f"{rm}{ext.lower()}")
    file_path = os.path.join(fotos_dir, new_filename)
    file.save(file_path)

    return jsonify({"url": f"/static/fotos/{new_filename}", "message": "Foto anexada com sucesso"}), 200


# ==========================================================
#  QUADROS – MENU PRINCIPAL
# ==========================================================

@app.route("/quadros")
@login_required
def quadros():
    return render_template("quadros.html")


# ==========================================================
#  QUADRO – INCLUSÃO
# ==========================================================

@app.route("/quadros/inclusao", methods=["GET", "POST"])
@login_required
def quadros_inclusao():
    if request.method == "POST":
        # Atualiza as listas na sessão (Regular e EJA)
        fundamental_file = request.files.get("lista_fundamental")
        eja_file = request.files.get("lista_eja")

        if fundamental_file and fundamental_file.filename != "":
            fundamental_filename = secure_filename(
                f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename
            )
            fundamental_path = os.path.join(app.config["UPLOAD_FOLDER"], fundamental_filename)
            fundamental_file.save(fundamental_path)
            session["lista_fundamental"] = fundamental_path

        if eja_file and eja_file.filename != "":
            eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_" + eja_file.filename)
            eja_path = os.path.join(app.config["UPLOAD_FOLDER"], eja_filename)
            eja_file.save(eja_path)
            session["lista_eja"] = eja_path

        # Carrega as listas piloto
        df_fundamental = None
        df_eja = None

        if session.get("lista_fundamental"):
            try:
                with open(session["lista_fundamental"], "rb") as f_fund:
                    df_fundamental = pd.read_excel(f_fund, sheet_name="LISTA CORRIDA")
            except Exception:
                flash("Erro ao ler a Lista Piloto Fundamental.", "error")
                return redirect(url_for("quadros_inclusao"))

        if session.get("lista_eja"):
            try:
                with open(session["lista_eja"], "rb") as f_eja:
                    df_eja = pd.read_excel(f_eja, sheet_name="LISTA CORRIDA")
            except Exception:
                flash("Erro ao ler a Lista Piloto EJA.", "error")
                return redirect(url_for("quadros_inclusao"))

        if df_fundamental is None and df_eja is None:
            flash("Nenhuma lista piloto disponível.", "error")
            return redirect(url_for("quadros_inclusao"))

        # Abre o modelo
        model_path = os.path.join("modelos", "Quadro de Alunos com Deficiência - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo de Inclusão não encontrado.", "error")
            return redirect(url_for("quadros_inclusao"))

        try:
            with open(model_path, "rb") as f:
                wb = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo de inclusão: {str(e)}", "error")
            return redirect(url_for("quadros_inclusao"))

        ws = wb.active
        set_merged_cell_value(ws, "C2", "E.M. José Padin Mouta")
        set_merged_cell_value(ws, "C3", "Luciana Rocha Augustinho")
        set_merged_cell_value(ws, "H3", "Ana Carolina Valencio da Silva Rodrigues")
        set_merged_cell_value(ws, "K3", "Ana Paula Rodrigues de Assis Santos")
        set_merged_cell_value(ws, "C4", "Rafael Marques Lima")
        set_merged_cell_value(ws, "H4", "Rita de Cassia de Andrade")
        set_merged_cell_value(ws, "K4", "Ana Paula Rodrigues de Assis Santos")
        set_merged_cell_value(ws, "P4", datetime.now().strftime("%d/%m/%Y"))

        start_row = 7
        current_row = start_row

        # Processa alunos da Lista Piloto Regular (Fundamental)
        if df_fundamental is not None:
            if len(df_fundamental.columns) < 25:
                flash("O arquivo da Lista Piloto Fundamental não possui colunas suficientes.", "error")
                return redirect(url_for("quadros_inclusao"))

            inclusion_col_fund = df_fundamental.columns[13]
            for _, row in df_fundamental.iterrows():
                if not str(row["RM"]).strip() or str(row["RM"]).strip() == "0":
                    continue

                if str(row[inclusion_col_fund]).strip().lower() == "sim":
                    # Coluna X (índice 23)
                    valor_coluna_x = row[df_fundamental.columns[23]]

                    col_a_val = str(row[df_fundamental.columns[0]]).strip()
                    match = re.match(r"(\d+º).*?([A-Za-z])$", col_a_val)
                    if match:
                        turma = match.group(2)
                    else:
                        turma = ""

                    horario = str(row[df_fundamental.columns[10]]).strip()
                    if "08h" in horario and "12h" in horario:
                        periodo = "MANHÃ"
                    elif horario == "13h30 às 17h30":
                        periodo = "TARDE"
                    elif horario == "19h00 às 23h00":
                        periodo = "NOITE"
                    else:
                        periodo = ""

                    nome_aluno = str(row[df_fundamental.columns[3]]).strip()
                    data_nasc = row[df_fundamental.columns[5]]
                    if pd.notna(data_nasc):
                        try:
                            data_nasc = pd.to_datetime(data_nasc, errors="coerce")
                            if pd.notna(data_nasc):
                                data_nasc = data_nasc.strftime("%d/%m/%Y")
                            else:
                                data_nasc = "Desconhecida"
                        except Exception:
                            data_nasc = "Desconhecida"
                    else:
                        data_nasc = "Desconhecida"

                    professor = str(row[df_fundamental.columns[14]]).strip()
                    plano = str(row[df_fundamental.columns[15]]).strip()
                    aee = (
                        str(row[df_fundamental.columns[16]]).strip()
                        if len(df_fundamental.columns) > 16
                        else ""
                    )
                    deficiencia = (
                        str(row[df_fundamental.columns[17]]).strip()
                        if len(df_fundamental.columns) > 17
                        else ""
                    )
                    observacoes = (
                        str(row[df_fundamental.columns[18]]).strip()
                        if len(df_fundamental.columns) > 18
                        else ""
                    )
                    cadeira = (
                        str(row[df_fundamental.columns[19]]).strip()
                        if len(df_fundamental.columns) > 19
                        else ""
                    )

                    # Coluna N: coluna U (índice 20)
                    valor_coluna_n = row[df_fundamental.columns[20]]
                    # Coluna O: coluna Y (índice 24)
                    valor_coluna_o = row[df_fundamental.columns[24]]

                    ws.cell(row=current_row, column=2, value=valor_coluna_x)
                    ws.cell(row=current_row, column=3, value=turma)
                    ws.cell(row=current_row, column=4, value=periodo)
                    ws.cell(row=current_row, column=5, value=horario)
                    ws.cell(row=current_row, column=6, value=nome_aluno)
                    ws.cell(row=current_row, column=7, value=data_nasc)
                    ws.cell(row=current_row, column=8, value=professor)
                    ws.cell(row=current_row, column=9, value=plano)
                    ws.cell(row=current_row, column=10, value=aee)
                    ws.cell(row=current_row, column=11, value=deficiencia)
                    ws.cell(row=current_row, column=12, value=observacoes)
                    ws.cell(row=current_row, column=13, value=cadeira)
                    ws.cell(row=current_row, column=14, value=valor_coluna_n)
                    ws.cell(row=current_row, column=15, value=valor_coluna_o)
                    current_row += 1

        # Processa alunos da Lista Piloto EJA
        if df_eja is not None:
            if len(df_eja.columns) < 29:
                flash("O arquivo da Lista Piloto EJA não possui colunas suficientes.", "error")
                return redirect(url_for("quadros_inclusao"))

            inclusion_col_eja = df_eja.columns[17]
            for _, row in df_eja.iterrows():
                if not str(row["RM"]).strip() or str(row["RM"]).strip() == "0":
                    continue

                if str(row[inclusion_col_eja]).strip().lower() == "sim":
                    # Coluna AB (índice 27) – nível/ano
                    valor_coluna_ab = row[df_eja.columns[27]]

                    turma = "A"
                    periodo = "NOITE"
                    horario = str(row[df_eja.columns[15]]).strip()
                    nome_aluno = str(row[df_eja.columns[3]]).strip()
                    data_nasc = row[df_eja.columns[6]]
                    if pd.notna(data_nasc):
                        try:
                            data_nasc = pd.to_datetime(data_nasc, errors="coerce")
                            if pd.notna(data_nasc):
                                data_nasc = data_nasc.strftime("%d/%m/%Y")
                            else:
                                data_nasc = "Desconhecida"
                        except Exception:
                            data_nasc = "Desconhecida"
                    else:
                        data_nasc = "Desconhecida"

                    professor = str(row[df_eja.columns[18]]).strip()
                    plano = str(row[df_eja.columns[19]]).strip()
                    aee = (
                        str(row[df_eja.columns[20]]).strip()
                        if len(df_eja.columns) > 20
                        else ""
                    )
                    deficiencia = (
                        str(row[df_eja.columns[21]]).strip()
                        if len(df_eja.columns) > 21
                        else ""
                    )
                    observacoes = (
                        str(row[df_eja.columns[22]]).strip()
                        if len(df_eja.columns) > 22
                        else ""
                    )
                    # Coluna M: coluna X (índice 23)
                    cadeira = row[df_eja.columns[23]]

                    # Coluna N: coluna Y (índice 24)
                    valor_coluna_n = row[df_eja.columns[24]]
                    # Coluna O: coluna AC (índice 28)
                    valor_coluna_o = row[df_eja.columns[28]]

                    ws.cell(row=current_row, column=2, value=valor_coluna_ab)
                    ws.cell(row=current_row, column=3, value=turma)
                    ws.cell(row=current_row, column=4, value=periodo)
                    ws.cell(row=current_row, column=5, value=horario)
                    ws.cell(row=current_row, column=6, value=nome_aluno)
                    ws.cell(row=current_row, column=7, value=data_nasc)
                    ws.cell(row=current_row, column=8, value=professor)
                    ws.cell(row=current_row, column=9, value=plano)
                    ws.cell(row=current_row, column=10, value=aee)
                    ws.cell(row=current_row, column=11, value=deficiencia)
                    ws.cell(row=current_row, column=12, value=observacoes)
                    ws.cell(row=current_row, column=13, value=cadeira)
                    ws.cell(row=current_row, column=14, value=valor_coluna_n)
                    ws.cell(row=current_row, column=15, value=valor_coluna_o)
                    current_row += 1

        output = BytesIO()
        wb.save(output)
        output.seek(0)
        meses = {
            1: "janeiro",
            2: "fevereiro",
            3: "março",
            4: "abril",
            5: "maio",
            6: "junho",
            7: "julho",
            8: "agosto",
            9: "setembro",
            10: "outubro",
            11: "novembro",
            12: "dezembro",
        }
        mes = meses[datetime.now().month].capitalize()
        filename = f"Quadro de Inclusão - {mes} - E.M José Padin Mouta.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # GET – simplesmente renderiza a página de upload
    return render_template("quadros_inclusao.html")


# ==========================================================
#  QUADRO – ATENDIMENTO MENSAL
# ==========================================================

@app.route('/quadros/atendimento_mensal', methods=['GET', 'POST'])
@login_required
def quadro_atendimento_mensal():
    if request.method == 'POST':
        fundamental_file = request.files.get('lista_fundamental')
        eja_file = request.files.get('lista_eja')

        # FUNDAMENTAL – atualiza arquivo da sessão, se enviado
        if fundamental_file and fundamental_file.filename != '':
            filename = secure_filename(fundamental_file.filename)
            unique_filename = f"atendimento_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            fundamental_file.save(file_path)
            session['lista_fundamental'] = file_path

        # EJA – atualiza arquivo da sessão, se enviado
        if eja_file and eja_file.filename != '':
            filename = secure_filename(eja_file.filename)
            unique_filename = f"atendimento_eja_{uuid.uuid4().hex}_{filename}"
            file_path_eja = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            eja_file.save(file_path_eja)
            session['lista_eja'] = file_path_eja

        # Usa a última lista FUNDAMENTAL salva em sessão, se existir
        file_path = session.get('lista_fundamental')
        if not file_path or not os.path.exists(file_path):
            flash("Nenhum arquivo da Lista Piloto FUNDAMENTAL disponível.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        # Carrega modelo
        model_path = os.path.join("modelos", "Quadro de Atendimento Mensal - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo Atendimento Mensal não encontrado.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        try:
            with open(model_path, "rb") as f:
                wb_modelo = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo de atendimento mensal: {str(e)}", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        # Se o modelo tiver mais de uma planilha, pega a segunda, senão a primeira
        if len(wb_modelo.worksheets) > 1:
            ws_modelo = wb_modelo.worksheets[1]
        else:
            ws_modelo = wb_modelo.active

        # Cabeçalho fixo
        set_merged_cell_value(ws_modelo, "B5", "E.M José Padin Mouta")
        set_merged_cell_value(ws_modelo, "C6", "Rafael Fernando da Silva")
        set_merged_cell_value(ws_modelo, "B7", "46034")

        current_month = datetime.now().strftime("%m")
        # Mantido /2025 conforme seu código original
        set_merged_cell_value(ws_modelo, "A13", f"{current_month}/2025")

        # Lê lista piloto FUNDAMENTAL
        try:
            with open(file_path, "rb") as lista_file:
                wb_lista = load_workbook(lista_file, data_only=True)
        except Exception:
            flash("Erro ao ler o arquivo da Lista Piloto FUNDAMENTAL.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        sheet_name = None
        for name in wb_lista.sheetnames:
            if name.strip().lower() == "total de alunos":
                sheet_name = name
                break

        if not sheet_name:
            flash("A aba 'Total de Alunos' não foi encontrada na Lista Piloto FUNDAMENTAL.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        ws_total = wb_lista[sheet_name]

        # Preenche blocos do modelo com dados da lista piloto FUNDAMENTAL
        for r, source_row in zip(range(55, 57), range(13, 15)):
            value_B = ws_total.cell(row=source_row, column=7).value
            value_C = ws_total.cell(row=source_row, column=8).value
            set_merged_cell_value(ws_modelo, f"B{r}", value_B)
            set_merged_cell_value(ws_modelo, f"C{r}", value_C)
            set_merged_cell_value(ws_modelo, f"D{r}", f"=B{r}+C{r}")

        for r, source_row in zip(range(57, 61), range(15, 19)):
            value_B = ws_total.cell(row=source_row, column=7).value
            value_C = ws_total.cell(row=source_row, column=8).value
            set_merged_cell_value(ws_modelo, f"B{r}", value_B)
            set_merged_cell_value(ws_modelo, f"C{r}", value_C)
            set_merged_cell_value(ws_modelo, f"D{r}", f"=B{r}+C{r}")

        for r, source_row in zip(range(73, 80), range(20, 27)):
            value_B = ws_total.cell(row=source_row, column=7).value
            value_C = ws_total.cell(row=source_row, column=8).value
            set_merged_cell_value(ws_modelo, f"B{r}", value_B)
            set_merged_cell_value(ws_modelo, f"C{r}", value_C)
            set_merged_cell_value(ws_modelo, f"D{r}", f"=B{r}+C{r}")

        for r, source_row in zip(range(91, 98), range(28, 35)):
            value_B = ws_total.cell(row=source_row, column=7).value
            value_C = ws_total.cell(row=source_row, column=8).value
            set_merged_cell_value(ws_modelo, f"B{r}", value_B)
            set_merged_cell_value(ws_modelo, f"C{r}", value_C)
            set_merged_cell_value(ws_modelo, f"D{r}", f"=B{r}+C{r}")

        # Campos específicos FUNDAMENTAL
        value_R20 = ws_total.cell(row=37, column=9).value  # I37
        set_merged_cell_value(ws_modelo, "R20", value_R20)

        set_merged_cell_value(ws_modelo, "R24", "-")

        value_R28 = ws_total.cell(row=39, column=9).value  # I39
        set_merged_cell_value(ws_modelo, "R28", value_R28)

        set_merged_cell_value(ws_modelo, "B37", ws_total.cell(row=6, column=7).value)   # G6
        set_merged_cell_value(ws_modelo, "B38", ws_total.cell(row=7, column=7).value)   # G7
        set_merged_cell_value(ws_modelo, "B39", ws_total.cell(row=8, column=7).value)   # G8
        set_merged_cell_value(ws_modelo, "B40", ws_total.cell(row=9, column=7).value)   # G9
        set_merged_cell_value(ws_modelo, "B41", ws_total.cell(row=10, column=7).value)  # G10
        set_merged_cell_value(ws_modelo, "B42", ws_total.cell(row=11, column=7).value)  # G11

        set_merged_cell_value(ws_modelo, "C37", ws_total.cell(row=6, column=8).value)   # H6
        set_merged_cell_value(ws_modelo, "C38", ws_total.cell(row=7, column=8).value)   # H7
        set_merged_cell_value(ws_modelo, "C39", ws_total.cell(row=8, column=8).value)   # H8
        set_merged_cell_value(ws_modelo, "C40", ws_total.cell(row=9, column=8).value)   # H9
        set_merged_cell_value(ws_modelo, "C41", ws_total.cell(row=10, column=8).value)  # H10
        set_merged_cell_value(ws_modelo, "C42", ws_total.cell(row=11, column=8).value)  # H11

        # ---- EJA (usa arquivo salvo em sessão) ----
        eja_path = session.get('lista_eja')
        if not eja_path or not os.path.exists(eja_path):
            flash("Arquivo da Lista Piloto EJA não encontrado.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        try:
            with open(eja_path, "rb") as f_eja:
                wb_eja = load_workbook(f_eja, data_only=True)
        except Exception as e:
            flash(f"Erro ao ler a Lista Piloto EJA: {str(e)}", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        sheet_name_eja = None
        for name in wb_eja.sheetnames:
            if name.strip().lower() == "total de alunos":
                sheet_name_eja = name
                break

        if not sheet_name_eja:
            flash("A aba 'Total de Alunos' não foi encontrada na Lista Piloto EJA.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        ws_total_eja = wb_eja[sheet_name_eja]

        set_merged_cell_value(ws_modelo, "L19", ws_total_eja.cell(row=6, column=5).value)  # E6
        set_merged_cell_value(ws_modelo, "L20", ws_total_eja.cell(row=7, column=5).value)  # E7
        set_merged_cell_value(ws_modelo, "L21", ws_total_eja.cell(row=8, column=5).value)  # E8
        set_merged_cell_value(ws_modelo, "L22", ws_total_eja.cell(row=9, column=5).value)  # E9

        set_merged_cell_value(ws_modelo, "M19", ws_total_eja.cell(row=6, column=6).value)  # F6
        set_merged_cell_value(ws_modelo, "M20", ws_total_eja.cell(row=7, column=6).value)  # F7
        set_merged_cell_value(ws_modelo, "M21", ws_total_eja.cell(row=8, column=6).value)  # F8
        set_merged_cell_value(ws_modelo, "M22", ws_total_eja.cell(row=9, column=6).value)  # F9

        set_merged_cell_value(ws_modelo, "L27", ws_total_eja.cell(row=11, column=5).value)  # E11
        set_merged_cell_value(ws_modelo, "L28", ws_total_eja.cell(row=12, column=5).value)  # E12
        set_merged_cell_value(ws_modelo, "L29", ws_total_eja.cell(row=13, column=5).value)  # E13
        set_merged_cell_value(ws_modelo, "L30", ws_total_eja.cell(row=14, column=5).value)  # E14

        set_merged_cell_value(ws_modelo, "M27", ws_total_eja.cell(row=11, column=6).value)  # F11
        set_merged_cell_value(ws_modelo, "M28", ws_total_eja.cell(row=12, column=6).value)  # F12
        set_merged_cell_value(ws_modelo, "M29", ws_total_eja.cell(row=13, column=6).value)  # F13
        set_merged_cell_value(ws_modelo, "M30", ws_total_eja.cell(row=14, column=6).value)  # F14

        set_merged_cell_value(ws_modelo, "L35", ws_total_eja.cell(row=16, column=5).value)  # E16
        set_merged_cell_value(ws_modelo, "L36", ws_total_eja.cell(row=17, column=5).value)  # E17
        set_merged_cell_value(ws_modelo, "L37", ws_total_eja.cell(row=18, column=5).value)  # E18

        set_merged_cell_value(ws_modelo, "M35", ws_total_eja.cell(row=16, column=6).value)  # F16
        set_merged_cell_value(ws_modelo, "M36", ws_total_eja.cell(row=17, column=6).value)  # F17
        set_merged_cell_value(ws_modelo, "M37", ws_total_eja.cell(row=18, column=6).value)  # F18

        set_merged_cell_value(ws_modelo, "R32", ws_total_eja.cell(row=20, column=7).value)  # G20
        set_merged_cell_value(ws_modelo, "R24", "-")

        # Gera arquivo em memória
        output = BytesIO()
        wb_modelo.save(output)
        output.seek(0)

        filename = f"Quadro de Atendimento Mensal - {datetime.now().strftime('%d%m')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

    # GET – renderiza a tela padronizada
    return render_template('quadro_atendimento_mensal.html')

# ==========================================================
#  QUADRO – TRANSFERÊNCIAS
# ==========================================================

@app.route("/quadros/transferencias", methods=["GET", "POST"])
@login_required
def quadro_transferencias():
    if request.method == "POST":
        period_start_str = request.form.get("period_start")
        period_end_str = request.form.get("period_end")
        responsavel = request.form.get("responsavel")

        fundamental_file = request.files.get("lista_fundamental")
        eja_file = request.files.get("lista_eja")

        if not period_start_str or not period_end_str or not responsavel:
            flash("Por favor, preencha todos os campos.", "error")
            return redirect(url_for("quadro_transferencias"))

        # Salva/atualiza a Lista Piloto Fundamental, se enviada
        if fundamental_file and fundamental_file.filename != "":
            fundamental_filename = secure_filename(
                f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename
            )
            fundamental_path = os.path.join(app.config["UPLOAD_FOLDER"], fundamental_filename)
            fundamental_file.save(fundamental_path)
            session["lista_fundamental"] = fundamental_path
        else:
            fundamental_path = session.get("lista_fundamental")
            if not fundamental_path or not os.path.exists(fundamental_path):
                flash("Lista Piloto Fundamental não encontrada.", "error")
                return redirect(url_for("quadro_transferencias"))

        # Salva/atualiza a Lista Piloto EJA, se enviada (opcional)
        if eja_file and eja_file.filename != "":
            eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_" + eja_file.filename)
            eja_path = os.path.join(app.config["UPLOAD_FOLDER"], eja_filename)
            eja_file.save(eja_path)
            session["lista_eja"] = eja_path
        else:
            eja_path = session.get("lista_eja")
            # EJA é opcional

        try:
            period_start = datetime.strptime(period_start_str, "%Y-%m-%d")
            period_end = datetime.strptime(period_end_str, "%Y-%m-%d")
        except Exception:
            flash("Formato de data inválido.", "error")
            return redirect(url_for("quadro_transferencias"))

        # ---- PARTE 1: FUNDAMENTAL ----
        try:
            df_fundamental = pd.read_excel(fundamental_path, sheet_name="LISTA CORRIDA")
        except Exception as e:
            flash(f"Erro ao ler a Lista Piloto Fundamental: {str(e)}", "error")
            return redirect(url_for("quadro_transferencias"))

        motivo_map = {
            "Dentro da Rede": "Dentro da rede",
            "Rede Estadual": "Dentro da rede",
            "Litoral": "Mudança de Municipio",
            "Mudança de Municipio": "Mudança de Municipio",
            "São Paulo": "Mudança de Municipio",
            "ABCD": "Mudança de Municipio",
            "Interior": "Mudança de Municipio",
            "Outros Estados": "Mudança de estado",
            "Particular": "Mudança para Escola Particular",
            "País": "Mudança de País",
        }

        transfer_records = []
        col_V_index = 21  # índice 0-based da coluna V

        for _, row in df_fundamental.iterrows():
            if len(row) < 9:
                continue

            obs_value = str(row.iloc[8]) if len(row) > 8 else ""
            motivo_raw = ""
            if len(row) > col_V_index:
                motivo_raw = str(row.iloc[col_V_index]).strip()
            else:
                motivo_raw = ""

            motivo_w = ""
            if len(row) > 22:
                motivo_w = str(row.iloc[22]).strip()

            match = re.search(r"(TE)\s*(\d{1,2}/\d{1,2})", obs_value)
            if match:
                te_date_str = match.group(2)  # dd/mm
                te_date_full_str = f"{te_date_str}/{period_start.year}"
                try:
                    te_date = datetime.strptime(te_date_full_str, "%d/%m/%Y")
                except Exception:
                    continue

                if period_start <= te_date <= period_end:
                    nome = str(row.iloc[3])
                    dn_val = row.iloc[5]
                    dn_str = ""
                    if pd.notna(dn_val):
                        try:
                            dn_dt = pd.to_datetime(dn_val, errors="coerce")
                            if pd.notna(dn_dt):
                                dn_str = dn_dt.strftime("%d/%m/%y")
                            else:
                                dn_str = ""
                        except Exception:
                            dn_str = ""

                    ra = str(row.iloc[6])
                    situacao = "Parcial"
                    breda = "Não"
                    nivel_classe = str(row.iloc[0])
                    tipo_field = "TE"

                    if motivo_raw in motivo_map:
                        reason_final = motivo_map[motivo_raw]
                    else:
                        reason_final = motivo_raw

                    if motivo_w:
                        reason_final = f"{reason_final} ({motivo_w})"

                    remanejamento = "-"
                    data_te = te_date.strftime("%d/%m/%Y")

                    record = {
                        "nome": nome,
                        "dn": dn_str,
                        "ra": ra,
                        "situacao": situacao,
                        "breda": breda,
                        "nivel_classe": nivel_classe,
                        "tipo": tipo_field,
                        "observacao": reason_final,
                        "remanejamento": remanejamento,
                        "data": data_te,
                    }
                    transfer_records.append(record)

        # ---- PARTE 2: EJA (TE / MC / MCC) ----
        if eja_path and os.path.exists(eja_path):
            try:
                df_eja = pd.read_excel(eja_path, sheet_name="LISTA CORRIDA")
            except Exception as e:
                flash(f"Erro ao ler a Lista Piloto EJA: {str(e)}", "error")
                return redirect(url_for("quadro_transferencias"))

            for _, row in df_eja.iterrows():
                if len(row) < 11:
                    continue

                col_k_value = str(row.iloc[10]).strip() if len(row) > 10 else ""
                if not col_k_value:
                    continue

                match_eja = re.search(
                    r"(TE|MC|MCC)\s*(\d{1,2}/\d{1,2})", col_k_value, re.IGNORECASE
                )
                if match_eja:
                    tipo_str = match_eja.group(1).upper()
                    date_str = match_eja.group(2)
                    eja_date_full = f"{date_str}/{period_start.year}"
                    try:
                        eja_date = datetime.strptime(eja_date_full, "%d/%m/%Y")
                    except Exception:
                        continue

                    if period_start <= eja_date <= period_end:
                        nome = str(row.iloc[3])
                        dn_val = row.iloc[6]
                        dn_str = ""
                        if pd.notna(dn_val):
                            try:
                                dn_dt = pd.to_datetime(dn_val, errors="coerce")
                                if pd.notna(dn_dt):
                                    dn_str = dn_dt.strftime("%d/%m/%Y")
                            except Exception:
                                dn_str = ""

                        ra_val = row.iloc[7]
                        if pd.isna(ra_val) or (
                            isinstance(ra_val, (int, float)) and float(ra_val) == 0
                        ):
                            ra_val = row.iloc[8]

                        situacao = "Parcial"
                        breda = "Não"
                        nivel_classe = str(row.iloc[0])
                        tipo_field = tipo_str

                        if tipo_field in ["MC", "MCC"]:
                            obs_final = "Desistencia"
                        else:
                            part_z = str(row.iloc[25]).strip() if len(row) > 25 else ""
                            part_aa = str(row.iloc[26]).strip() if len(row) > 26 else ""
                            if part_aa:
                                obs_final = f"{part_z} ({part_aa})".strip()
                            else:
                                obs_final = part_z

                        remanejamento = "-"
                        data_te = eja_date.strftime("%d/%m/%Y")

                        record = {
                            "nome": nome,
                            "dn": dn_str,
                            "ra": str(ra_val),
                            "situacao": situacao,
                            "breda": breda,
                            "nivel_classe": nivel_classe,
                            "tipo": tipo_field,
                            "observacao": obs_final,
                            "remanejamento": remanejamento,
                            "data": data_te,
                        }
                        transfer_records.append(record)

        if not transfer_records:
            flash("Nenhum registro de TE/MC/MCC encontrado no período especificado.", "error")
            return redirect(url_for("quadro_transferencias"))

        model_path = os.path.join("modelos", "Quadro Informativo - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo de Quadro Informativo (Transferências) não encontrado.", "error")
            return redirect(url_for("quadro_transferencias"))

        try:
            with open(model_path, "rb") as f:
                wb = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo: {str(e)}", "error")
            return redirect(url_for("quadro_transferencias"))

        ws = wb.active

        set_merged_cell_value(ws, "B9", responsavel)
        set_merged_cell_value(ws, "J9", datetime.now().strftime("%d/%m/%Y"))

        start_row = 12
        current_row = start_row

        for record in transfer_records:
            set_merged_cell_value(ws, f"A{current_row}", record["nome"])
            set_merged_cell_value(ws, f"B{current_row}", record["dn"])
            set_merged_cell_value(ws, f"C{current_row}", record["ra"])
            set_merged_cell_value(ws, f"D{current_row}", record["situacao"])
            set_merged_cell_value(ws, f"E{current_row}", record["breda"])
            set_merged_cell_value(ws, f"F{current_row}", record["nivel_classe"])
            set_merged_cell_value(ws, f"G{current_row}", record["tipo"])
            set_merged_cell_value(ws, f"H{current_row}", record["observacao"])
            set_merged_cell_value(ws, f"I{current_row}", record["remanejamento"])
            set_merged_cell_value(ws, f"J{current_row}", record["data"])
            current_row += 1

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = (
            f"Quadro_de_Transferencias_"
            f"{period_start.strftime('%d%m')}_{period_end.strftime('%d%m')}.xlsx"
        )
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # GET – exibe o formulário estilizado
    return render_template("quadro_transferencias.html")

# ==========================================================
#  QUADRO – QUANTITATIVO MENSAL (Fundamental)
# ==========================================================

@app.route("/quadros/quantitativo_mensal", methods=["GET", "POST"])
@login_required
def quadro_quantitativo_mensal():
    if request.method == "POST":
        period_start_str = request.form.get("period_start")
        period_end_str = request.form.get("period_end")
        responsavel = request.form.get("responsavel")

        if not period_start_str or not period_end_str or not responsavel:
            flash("Preencha todos os campos obrigatórios.", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        try:
            period_start = datetime.strptime(period_start_str, "%Y-%m-%d")
            period_end = datetime.strptime(period_end_str, "%Y-%m-%d")
        except Exception:
            flash("Formato de data inválido.", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        fundamental_file = request.files.get("lista_fundamental")
        if fundamental_file and fundamental_file.filename != "":
            fundamental_filename = secure_filename(
                f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename
            )
            fundamental_path = os.path.join(app.config["UPLOAD_FOLDER"], fundamental_filename)
            fundamental_file.save(fundamental_path)
            session["lista_fundamental"] = fundamental_path
        else:
            fundamental_path = session.get("lista_fundamental")
            if not fundamental_path or not os.path.exists(fundamental_path):
                flash("Arquivo da Lista Piloto Fundamental não encontrado.", "error")
                return redirect(url_for("quadro_quantitativo_mensal"))

        try:
            df_fundamental = pd.read_excel(fundamental_path, sheet_name="LISTA CORRIDA")
        except Exception as e:
            flash(f"Erro ao ler a Lista Piloto Fundamental: {str(e)}", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        model_path = os.path.join("modelos", "Quadro Quantitativo Mensal - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo de Quadro Quantitativo Mensal não encontrado.", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        try:
            with open(model_path, "rb") as f:
                wb = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo: {str(e)}", "error")
            return redirect(url_for("quadro_quantitativo_mensal"))

        ws = wb.active

        mapping = {
            "2º": {
                "Dentro da Rede": "K14",
                "Rede Estadual": "K15",
                "Litoral": "K16",
                "São Paulo": "K17",
                "ABCD": "K18",
                "Interior": "K19",
                "Outros Estados": "K20",
                "Particular": "K21",
                "País": "K22",
                "Sem Informação": "K23",
            },
            "3º": {
                "Dentro da Rede": "L14",
                "Rede Estadual": "L15",
                "Litoral": "L16",
                "São Paulo": "L17",
                "ABCD": "L18",
                "Interior": "L19",
                "Outros Estados": "L20",
                "Particular": "L21",
                "País": "L22",
                "Sem Informação": "L23",
            },
            "4º": {
                "Dentro da Rede": "M14",
                "Rede Estadual": "M15",
                "Litoral": "M16",
                "São Paulo": "M17",
                "ABCD": "M18",
                "Interior": "M19",
                "Outros Estados": "M20",
                "Particular": "M21",
                "País": "M22",
                "Sem Informação": "M23",
            },
            "5º": {
                "Dentro da Rede": "N14",
                "Rede Estadual": "N15",
                "Litoral": "N16",
                "São Paulo": "N17",
                "ABCD": "N18",
                "Interior": "N19",
                "Outros Estados": "N20",
                "Particular": "N21",
                "País": "N22",
                "Sem Informação": "N23",
            },
        }

        # Inicializa contadores
        for _, tipos in mapping.items():
            for cell_addr in tipos.values():
                current_val = ws[cell_addr].value
                if current_val is None or not isinstance(current_val, (int, float)):
                    set_merged_cell_value(ws, cell_addr, 0)

        for _, row in df_fundamental.iterrows():
            if len(row) < 9:
                continue

            col_I_val = str(row.iloc[8]).strip() if pd.notna(row.iloc[8]) else ""
            if "TE" not in col_I_val:
                continue

            match = re.search(r"TE\s*([0-9]{1,2}/[0-9]{1,2})", col_I_val)
            if not match:
                continue

            te_date_str = match.group(1)
            te_date_full_str = f"{te_date_str}/{period_start.year}"
            try:
                te_date = datetime.strptime(te_date_full_str, "%d/%m/%Y")
            except Exception:
                continue

            if not (period_start <= te_date <= period_end):
                continue

            serie_value = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
            series_key = None
            if "2º" in serie_value:
                series_key = "2º"
            elif "3º" in serie_value:
                series_key = "3º"
            elif "4º" in serie_value:
                series_key = "4º"
            elif "5º" in serie_value:
                series_key = "5º"
            else:
                continue

            tipo_te = ""
            if len(row) > 21 and pd.notna(row.iloc[21]):
                tipo_te = str(row.iloc[21]).strip()
            if not tipo_te:
                tipo_te = "Sem Informação"

            if series_key in mapping and tipo_te in mapping[series_key]:
                cell_addr = mapping[series_key][tipo_te]
                current_count = ws[cell_addr].value
                if not isinstance(current_count, (int, float)):
                    current_count = 0
                set_merged_cell_value(ws, cell_addr, current_count + 1)

        set_merged_cell_value(ws, "B3", responsavel)
        set_merged_cell_value(
            ws,
            "D3",
            f"{period_start.strftime('%d/%m/%Y')} a {period_end.strftime('%d/%m/%Y')}",
        )

        meses = {
            1: "Janeiro",
            2: "Fevereiro",
            3: "Março",
            4: "Abril",
            5: "Maio",
            6: "Junho",
            7: "Julho",
            8: "Agosto",
            9: "Setembro",
            10: "Outubro",
            11: "Novembro",
            12: "Dezembro",
        }
        current_month = meses[datetime.now().month]
        current_year = datetime.now().year
        set_merged_cell_value(ws, "A8", f"{current_month}/{current_year}")

        output = BytesIO()
        wb.save(output)
        output.seek(0)

        filename = (
            f"Quadro_Quantitativo_Fundamental_"
            f"{period_start.strftime('%d%m')}_{period_end.strftime('%d%m')}.xlsx"
        )
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # GET: exibe o formulário
    return render_template("quadro_quantitativo_mensal.html")

# ==========================================================
#  QUADRO QUANTITATIVO DE INCLUSÃO – REGULAR + EJA
# ==========================================================

SETORES = ["Financeiro", "Recursos Humanos", "Administrativo", "Marketing", "TI"]


@app.route("/quantinclusao", methods=["GET", "POST"])
def quantinclusao():
    if request.method == "POST":
        reg_file = request.files.get("lista_regular")
        eja_file = request.files.get("lista_eja")
        responsavel = request.form.get("responsavel")

        if not reg_file or reg_file.filename == "":
            flash("Selecione o arquivo da Lista Piloto Regular.", "error")
            return redirect(url_for("quantinclusao"))

        if not eja_file or eja_file.filename == "":
            flash("Selecione o arquivo da Lista Piloto EJA.", "error")
            return redirect(url_for("quantinclusao"))

        if not responsavel or responsavel.strip() == "":
            flash("Informe o Responsável pelo preenchimento.", "error")
            return redirect(url_for("quantinclusao"))

        reg_filename = secure_filename(f"regular_{uuid.uuid4().hex}_{reg_file.filename}")
        eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_{eja_file.filename}")
        reg_path = os.path.join(app.config["UPLOAD_FOLDER"], reg_filename)
        eja_path = os.path.join(app.config["UPLOAD_FOLDER"], eja_filename)
        reg_file.save(reg_path)
        eja_file.save(eja_path)

        # Regular – Total de Alunos + LISTA CORRIDA
        try:
            wb_reg = load_workbook(reg_path, data_only=True)
            ws_total_reg = wb_reg["Total de Alunos"]
            ws_lista_reg = wb_reg["LISTA CORRIDA"]
        except Exception as e:
            flash(f"Erro ao ler o arquivo Regular: {str(e)}", "error")
            return redirect(url_for("quantinclusao"))

        reg_map = {
            "D13": ws_total_reg["O6"].value,
            "D17": ws_total_reg["O7"].value,
            "D21": ws_total_reg["O8"].value,
            "D25": ws_total_reg["O9"].value,
            "D29": ws_total_reg["O10"].value,
            "D33": ws_total_reg["O11"].value,
            "D41": ws_total_reg["O13"].value,
            "D45": ws_total_reg["O14"].value,
            "H13": ws_total_reg["O15"].value,
            "H17": ws_total_reg["O16"].value,
            "H21": ws_total_reg["O17"].value,
            "H25": ws_total_reg["O18"].value,
            "H29": ws_total_reg["O20"].value,
            "H33": ws_total_reg["O21"].value,
            "H37": ws_total_reg["O22"].value,
            "H45": ws_total_reg["O24"].value,
            "L13": ws_total_reg["O25"].value,
            "L17": ws_total_reg["O26"].value,
            "L21": ws_total_reg["O28"].value,
            "L25": ws_total_reg["O29"].value,
            "L29": ws_total_reg["O30"].value,
            "L33": ws_total_reg["O31"].value,
            "L37": ws_total_reg["O32"].value,
            "L41": ws_total_reg["O33"].value,
            "L45": ws_total_reg["O34"].value,
        }

        # Contadores para séries (Regular)
        count_2A = count_2B = count_2C = count_2D = count_2E = count_2F = 0
        count_3A = count_3B = count_3C = count_3D = count_3E = count_3F = 0
        count_4A = count_4B = count_4C = count_4D = count_4E = count_4F = count_4G = 0
        count_5A = count_5B = count_5C = count_5D = count_5E = count_5F = count_5G = 0

        series_list = [
            "2ºA",
            "2ºB",
            "2ºC",
            "2ºD",
            "2ºE",
            "2ºF",
            "2ºG",
            "3ºA",
            "3ºB",
            "3ºC",
            "3ºD",
            "3ºE",
            "3ºF",
            "4ºA",
            "4ºB",
            "4ºC",
            "4ºD",
            "4ºE",
            "4ºF",
            "4ºG",
            "5ºA",
            "5ºB",
            "5ºC",
            "5ºD",
            "5ºE",
            "5ºF",
            "5ºG",
        ]
        unique_names = {serie: set() for serie in series_list}

        for row in ws_lista_reg.iter_rows(min_row=2, values_only=True):
            serie = str(row[0]).strip() if row[0] is not None else ""
            plano = row[15] if len(row) > 15 else None

            if serie in unique_names and is_valid_plano(plano):
                unique_names[serie].add(str(plano).strip())

            if serie == "2ºA" and is_valid_plano(plano):
                count_2A += 1
            elif serie == "2ºB" and is_valid_plano(plano):
                count_2B += 1
            elif serie == "2ºC" and is_valid_plano(plano):
                count_2C += 1
            elif serie == "2ºD" and is_valid_plano(plano):
                count_2D += 1
            elif serie == "2ºE" and is_valid_plano(plano):
                count_2E += 1
            elif serie == "2ºF" and is_valid_plano(plano):
                count_2F += 1
            elif serie == "3ºA" and is_valid_plano(plano):
                count_3A += 1
            elif serie == "3ºB" and is_valid_plano(plano):
                count_3B += 1
            elif serie == "3ºC" and is_valid_plano(plano):
                count_3C += 1
            elif serie == "3ºD" and is_valid_plano(plano):
                count_3D += 1
            elif serie == "3ºE" and is_valid_plano(plano):
                count_3E += 1
            elif serie == "3ºF" and is_valid_plano(plano):
                count_3F += 1
            elif serie == "4ºA" and is_valid_plano(plano):
                count_4A += 1
            elif serie == "4ºB" and is_valid_plano(plano):
                count_4B += 1
            elif serie == "4ºC" and is_valid_plano(plano):
                count_4C += 1
            elif serie == "4ºD" and is_valid_plano(plano):
                count_4D += 1
            elif serie == "4ºE" and is_valid_plano(plano):
                count_4E += 1
            elif serie == "4ºF" and is_valid_plano(plano):
                count_4F += 1
            elif serie == "4ºG" and is_valid_plano(plano):
                count_4G += 1
            elif serie == "5ºA" and is_valid_plano(plano):
                count_5A += 1
            elif serie == "5ºB" and is_valid_plano(plano):
                count_5B += 1
            elif serie == "5ºC" and is_valid_plano(plano):
                count_5C += 1
            elif serie == "5ºD" and is_valid_plano(plano):
                count_5D += 1
            elif serie == "5ºE" and is_valid_plano(plano):
                count_5E += 1
            elif serie == "5ºF" and is_valid_plano(plano):
                count_5F += 1
            elif serie == "5ºG" and is_valid_plano(plano):
                count_5G += 1

        # EJA – Total de Alunos
        try:
            wb_eja = load_workbook(eja_path, data_only=True)
            ws_total_eja = wb_eja["Total de Alunos"]
        except Exception as e:
            flash(f"Erro ao ler o arquivo EJA: {str(e)}", "error")
            return redirect(url_for("quantinclusao"))

        eja_map = {
            "D53": ws_total_eja["M10"].value,
            "D57": (ws_total_eja["M11"].value or 0) + (ws_total_eja["M12"].value or 0),
            "D61": ws_total_eja["M13"].value,
            "H53": ws_total_eja["M14"].value,
            "H57": ws_total_eja["M16"].value,
            "H61": ws_total_eja["M17"].value,
            "L53": ws_total_eja["M18"].value,
        }

        try:
            ws_lista_eja = wb_eja["LISTA CORRIDA"]
        except Exception as e:
            flash(f"Erro ao ler a aba LISTA CORRIDA no arquivo EJA: {str(e)}", "error")
            return redirect(url_for("quantinclusao"))

        series_ef_group1 = {"1ª SÉRIE E.F", "2ª SÉRIE E.F", "3ª SÉRIE E.F", "4ª SÉRIE E.F"}
        series_ef_group2 = {"5ª SÉRIE E.F", "6ª SÉRIE E.F"}
        series_ef_group3 = {"7ª SÉRIE E.F"}
        series_ef_group4 = {"8ª SÉRIE E.F"}
        series_em_group1 = {"1ª SÉRIE E.M"}
        series_em_group2 = {"2ª SÉRIE E.M"}
        series_em_group3 = {"3ª SÉRIE E.M"}

        d54_count = 0
        unique_d55 = set()
        d58_count = 0
        unique_d59 = set()
        d62_count = 0
        unique_d63 = set()
        h54_count = 0
        unique_h55 = set()
        h58_count = 0
        unique_h59 = set()
        h62_count = 0
        unique_h63 = set()
        l54_count = 0
        unique_l55 = set()

        for row in ws_lista_eja.iter_rows(min_row=2, values_only=True):
            serie = str(row[0]).strip() if row[0] is not None else ""
            nome = row[19] if len(row) > 19 else None

            if serie in series_ef_group1:
                if is_valid_plano(nome):
                    d54_count += 1
                    unique_d55.add(str(nome).strip())

            if serie in series_ef_group2:
                if is_valid_plano(nome):
                    d58_count += 1
                    unique_d59.add(str(nome).strip())

            if serie in series_ef_group3:
                if is_valid_plano(nome):
                    d62_count += 1
                    unique_d63.add(str(nome).strip())

            if serie in series_ef_group4:
                if is_valid_plano(nome):
                    h54_count += 1
                    unique_h55.add(str(nome).strip())

            if serie in series_em_group1:
                if is_valid_plano(nome):
                    h58_count += 1
                    unique_h59.add(str(nome).strip())

            if serie in series_em_group2:
                if is_valid_plano(nome):
                    h62_count += 1
                    unique_h63.add(str(nome).strip())

            if serie in series_em_group3:
                if is_valid_plano(nome):
                    l54_count += 1
                    unique_l55.add(str(nome).strip())

        model_path = os.path.join("modelos", "Quadro Quantitativo de Inclusão - Modelo.xlsx")
        try:
            wb_model = load_workbook(model_path)
        except Exception as e:
            flash(f"Erro ao abrir o modelo de inclusão: {str(e)}", "error")
            return redirect(url_for("quantinclusao"))

        ws_model = wb_model.active

        # Regular fixo
        for cell_addr, value in reg_map.items():
            ws_model[cell_addr] = value

        # EJA fixo
        for cell_addr, value in eja_map.items():
            ws_model[cell_addr] = value

        # Contagem simples Regular
        ws_model["D14"] = count_2A
        ws_model["D18"] = count_2B
        ws_model["D22"] = count_2C
        ws_model["D26"] = count_2D
        ws_model["D30"] = count_2E
        ws_model["D34"] = count_2F

        ws_model["D42"] = count_3A
        ws_model["D46"] = count_3B

        ws_model["H14"] = count_3C
        ws_model["H18"] = count_3D
        ws_model["H22"] = count_3E
        ws_model["H26"] = count_3F

        ws_model["H30"] = count_4A
        ws_model["H34"] = count_4B
        ws_model["H38"] = count_4C
        ws_model["H42"] = count_4D
        ws_model["H46"] = count_4E

        ws_model["L14"] = count_4F
        ws_model["L18"] = count_4G

        ws_model["L22"] = count_5A
        ws_model["L26"] = count_5B
        ws_model["L30"] = count_5C
        ws_model["L34"] = count_5D
        ws_model["L38"] = count_5E
        ws_model["L42"] = count_5F
        ws_model["L46"] = count_5G

        # Nomes únicos Regular
        ws_model["D15"] = len(unique_names["2ºA"])
        ws_model["D19"] = len(unique_names["2ºB"])
        ws_model["D23"] = len(unique_names["2ºC"])
        ws_model["D27"] = len(unique_names["2ºD"])
        ws_model["D31"] = len(unique_names["2ºE"])
        ws_model["D35"] = len(unique_names["2ºF"])
        ws_model["D39"] = len(unique_names["2ºG"])

        ws_model["D43"] = len(unique_names["3ºA"])
        ws_model["D47"] = len(unique_names["3ºB"])

        ws_model["H15"] = len(unique_names["3ºC"])
        ws_model["H19"] = len(unique_names["3ºD"])
        ws_model["H23"] = len(unique_names["3ºE"])
        ws_model["H27"] = len(unique_names["3ºF"])

        ws_model["H31"] = len(unique_names["4ºA"])
        ws_model["H35"] = len(unique_names["4ºB"])
        ws_model["H39"] = len(unique_names["4ºC"])
        ws_model["H43"] = len(unique_names["4ºD"])
        ws_model["H47"] = len(unique_names["4ºE"])

        ws_model["L15"] = len(unique_names["4ºF"])
        ws_model["L19"] = len(unique_names["4ºG"])

        ws_model["L23"] = len(unique_names["5ºA"])
        ws_model["L27"] = len(unique_names["5ºB"])
        ws_model["L31"] = len(unique_names["5ºC"])
        ws_model["L35"] = len(unique_names["5ºD"])
        ws_model["L39"] = len(unique_names["5ºE"])
        ws_model["L43"] = len(unique_names["5ºF"])
        ws_model["L47"] = len(unique_names["5ºG"])

        # Dados do EJA (LISTA CORRIDA)
        ws_model["H41"] = ws_total_reg["O23"].value

        ws_model["D54"] = d54_count
        ws_model["D55"] = len(unique_d55)
        ws_model["D58"] = d58_count
        ws_model["D59"] = len(unique_d59)
        ws_model["D62"] = d62_count
        ws_model["D63"] = len(unique_d63)

        ws_model["H54"] = h54_count
        ws_model["H55"] = len(unique_h55)
        ws_model["H58"] = h58_count
        ws_model["H59"] = len(unique_h59)
        ws_model["H62"] = h62_count
        ws_model["H63"] = len(unique_h63)

        ws_model["L54"] = l54_count
        ws_model["L55"] = len(unique_l55)

        # Informações adicionais
        meses = {
            1: "JANEIRO",
            2: "FEVEREIRO",
            3: "MARÇO",
            4: "ABRIL",
            5: "MAIO",
            6: "JUNHO",
            7: "JULHO",
            8: "AGOSTO",
            9: "SETEMBRO",
            10: "OUTUBRO",
            11: "NOVEMBRO",
            12: "DEZEMBRO",
        }
        current_date = datetime.now()
        ws_model["B4"] = f"{meses[current_date.month]}/{current_date.year}"
        ws_model["C8"] = responsavel.strip()
        ws_model["K8"] = current_date.strftime("%d/%m/%Y")

        output = BytesIO()
        wb_model.save(output)
        output.seek(0)
        filename = f"Quadro_Quantitativo_de_Inclusao_{datetime.now().strftime('%d%m%Y')}.xlsx"
        return send_file(
            output,
            as_attachment=True,
            download_name=filename,
            mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # GET
    return render_template("quantinclusao.html")

# ==========================================================
#  MAIN
# ==========================================================

if __name__ == "__main__":
    app.run(debug=True)
