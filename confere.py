import os
from flask import Blueprint, Flask, request, render_template, flash, redirect, url_for, jsonify
import pandas as pd
import pdfplumber
import unicodedata

# Cria o blueprint
confere_bp = Blueprint('confere', __name__, template_folder='templates')

# Definições e configurações do subsistema
UPLOAD_FOLDER = os.path.join(os.path.abspath(os.path.dirname(__file__)), "uploads")

# Configuração do upload
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

# Variável global para armazenar o caminho do arquivo Excel atual
current_excel_file = None

# Mapeamento de séries para os intervalos de linhas
mapeamento = {
    '2ºA': (0, 50),
    '2ºB': (50, 100),
    '2ºC': (100, 150),
    '2ºD': (150, 200),
    '2ºE': (200, 250),
    '2ºF': (250, 300),
    '3ºA': (300, 350),
    '3ºB': (350, 400),
    '3ºC': (400, 450),
    '3ºD': (450, 500),
    '3ºE': (500, 550),
    '3ºF': (550, 600),
    '4ºA': (600, 650),
    '4ºB': (650, 700),
    '4ºC': (700, 750),
    '4ºD': (750, 800),
    '4ºE': (800, 850),
    '4ºF': (850, 900),
    '4ºG': (900, 950),
    '5ºA': (950, 1000),
    '5ºB': (1000, 1050),
    '5ºC': (1050, 1100),
    '5ºD': (1100, 1150),
    '5ºE': (1150, 1200),
    '5ºF': (1200, 1250),
    '5ºG': (1250, 1300),
}

def normalize_str(s):
    """Normaliza a string removendo acentuação, espaços extras e convertendo para minúsculas."""
    s = s.lower()
    s = unicodedata.normalize('NFKD', s).encode('ASCII', 'ignore').decode('utf-8')
    s = " ".join(s.split())
    return s

def obter_dados_serie(serie, excel_filepath):
    """
    Extrai os dados do Excel referentes à série selecionada.
    Lê a aba "LISTA CORRIDA" do arquivo Excel e filtra o intervalo mapeado.
    """
    try:
        df = pd.read_excel(excel_filepath, sheet_name="LISTA CORRIDA", usecols='D,I', header=0)
    except Exception as e:
        return f"Erro ao ler o arquivo Excel: {str(e)}"
    
    if df.shape[1] < 2:
        return "As colunas necessárias não foram encontradas no arquivo Excel."
    
    if serie not in mapeamento:
        return f"Série '{serie}' não encontrada no mapeamento."
    
    inicio, fim = mapeamento[serie]
    df_serie = df.iloc[inicio:fim].copy()
    df_serie.columns = ['Nome', 'OBS']
    df_serie = df_serie[df_serie['Nome'].notna()]
    df_serie = df_serie[~df_serie['Nome'].astype(str).str.strip().isin(['0', '#REF#'])]
    df_serie['OBS'] = df_serie['OBS'].apply(lambda x: '-' if str(x).strip() == '0' else x)
    return df_serie

def obter_dados_pdf(file):
    """
    Extrai a tabela do PDF enviado pelo usuário.
    """
    all_rows = []
    header = None
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            table = page.extract_table()
            if table:
                if header is None:
                    header = table[0]
                for row in table[1:]:
                    all_rows.append(row)
    if not all_rows or header is None:
        return None
    df_pdf = pd.DataFrame(all_rows, columns=header)
    df_pdf.columns = [col.strip() for col in df_pdf.columns]
    df_pdf.rename(columns={
        'Nome do Aluno': 'Nome',
        'Situação': 'Situacao',
        'Data Movimentação': 'DataMovimentacao'
    }, inplace=True)
    df_pdf = df_pdf[df_pdf['Nome'].notna()]
    df_pdf = df_pdf[~df_pdf['Nome'].astype(str).str.strip().isin(['0', '#REF#'])]
    return df_pdf

def comparar_listas(df_excel, df_pdf):
    """
    Compara as listas do Excel e do PDF para identificar divergências.
    """
    def has_TE(obs):
        if pd.isna(obs):
            return False
        return 'te' in normalize_str(str(obs))
    
    def has_BXTR_or_TRAN(sit):
        if pd.isna(sit):
            return False
        sit_norm = normalize_str(str(sit))
        return sit_norm in ['bxtr', 'tran']
    
    def has_REMP(obs):
        if pd.isna(obs):
            return False
        return 'rem p' in normalize_str(str(obs))
    
    def has_REMA(sit):
        if pd.isna(sit):
            return False
        return 'rema' in normalize_str(str(sit))
    
    df_excel = df_excel.copy()
    df_pdf = df_pdf.copy()
    df_excel['nome_norm'] = df_excel['Nome'].apply(normalize_str)
    df_pdf['nome_norm'] = df_pdf['Nome'].apply(normalize_str)
    
    divergencias = []
    df_merged = pd.merge(df_excel, df_pdf, on='nome_norm', suffixes=('_excel', '_pdf'))
    for _, row in df_merged.iterrows():
        te_excel = has_TE(row['OBS'])
        bxtr_tran_pdf = has_BXTR_or_TRAN(row['Situacao'])
        if te_excel and not bxtr_tran_pdf:
            reason = "Aluno transferido na lista piloto mas não transferido no SED"
            divergencias.append({
                "Nome": row['Nome_excel'],
                "OBS (Excel)": row['OBS'],
                "Situacao (PDF)": row['Situacao'],
                "Divergência": reason
            })
        elif bxtr_tran_pdf and not te_excel:
            reason = "Aluno transferido no SED mas não transferido na lista piloto"
            divergencias.append({
                "Nome": row['Nome_excel'],
                "OBS (Excel)": row['OBS'],
                "Situacao (PDF)": row['Situacao'],
                "Divergência": reason
            })
        remp_excel = has_REMP(row['OBS'])
        rema_pdf = has_REMA(row['Situacao'])
        if remp_excel and not rema_pdf:
            reason = "Aluno remanejado na lista piloto mas não remanejado no SED"
            divergencias.append({
                "Nome": row['Nome_excel'],
                "OBS (Excel)": row['OBS'],
                "Situacao (PDF)": row['Situacao'],
                "Divergência": reason
            })
        elif rema_pdf and not remp_excel:
            reason = "Aluno remanejado no SED mas não remanejado na lista piloto"
            divergencias.append({
                "Nome": row['Nome_excel'],
                "OBS (Excel)": row['OBS'],
                "Situacao (PDF)": row['Situacao'],
                "Divergência": reason
            })
    
    excel_lookup = df_excel.set_index('nome_norm').to_dict(orient='index')
    pdf_lookup = df_pdf.set_index('nome_norm').to_dict(orient='index')
    
    set_excel = set(df_excel['nome_norm'])
    set_pdf = set(df_pdf['nome_norm'])
    
    for nome in set_pdf - set_excel:
        info_pdf = pdf_lookup.get(nome, {})
        divergencias.append({
            "Nome": info_pdf.get('Nome', nome),
            "OBS (Excel)": "-",
            "Situacao (PDF)": "-",
            "Divergência": "Aluno presente no SED mas não na lista piloto"
        })
    
    for nome in set_excel - set_pdf:
        info_excel = excel_lookup.get(nome, {})
        divergencias.append({
            "Nome": info_excel.get('Nome', nome),
            "OBS (Excel)": "-",
            "Situacao (PDF)": "-",
            "Divergência": "Aluno presente na lista piloto mas não no SED"
        })
    
    if divergencias:
        return pd.DataFrame(divergencias).sort_values(by="Nome")
    else:
        return None

@confere_bp.route('/upload_excel', methods=['POST'])
def upload_excel():
    global current_excel_file
    if 'listaExcel' not in request.files:
        return jsonify({"success": False, "message": "Nenhum arquivo Excel enviado."})
    file_excel = request.files['listaExcel']
    if file_excel.filename == '':
        return jsonify({"success": False, "message": "Nenhum arquivo selecionado."})
    excel_path = os.path.join(UPLOAD_FOLDER, file_excel.filename)
    file_excel.save(excel_path)
    current_excel_file = excel_path
    return jsonify({"success": True, "message": "Arquivo Excel carregado com sucesso!"})

@confere_bp.route('/', methods=['GET', 'POST'])
def index():
    global current_excel_file
    dados_excel = None
    dados_pdf = None
    divergencias = None
    error_excel = None
    selected_series = None

    if request.method == 'POST':
        selected_series = request.form.get('serie')
        if current_excel_file is None:
            flash("Nenhum arquivo Excel foi carregado. Por favor, anexe um arquivo Excel.", "danger")
        else:
            result = obter_dados_serie(selected_series, current_excel_file)
            if isinstance(result, str):
                error_excel = result
            else:
                dados_excel = result

        if 'listaPDF' in request.files:
            file_pdf = request.files['listaPDF']
            if file_pdf.filename != '':
                dados_pdf = obter_dados_pdf(file_pdf)
        if dados_excel is not None and dados_pdf is not None:
            divergencias = comparar_listas(dados_excel, dados_pdf)
    
    return render_template('index.html', 
                           dados_excel=dados_excel, 
                           dados_pdf=dados_pdf, 
                           divergencias=divergencias,
                           error_excel=error_excel,
                           selected_series=selected_series)

# Permite rodar o subsistema de forma independente para testes
if __name__ == '__main__':
    app = Flask(__name__)
    app.secret_key = 'sua_chave_secreta_aqui'
    app.register_blueprint(confere_bp, url_prefix='/confere')
    app.run(debug=True)
