from flask import Flask, request, redirect, url_for, render_template_string, jsonify, session, flash, send_file, render_template, send_from_directory
import pandas as pd
import os
import uuid
import re
from datetime import datetime
from io import BytesIO
from werkzeug.utils import secure_filename
from functools import wraps
import locale
import xlrd  # Para ler arquivos .xls, se necessário
from openpyxl import load_workbook, Workbook  # Usado para trabalhar com XLSX
from openpyxl.utils import get_column_letter  # Para obter a coluna em letra
from openpyxl.cell import MergedCell  # Para identificar células mescladas
from urllib.parse import unquote
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4

# Tenta definir a localidade para formatação de datas em português
try:
    locale.setlocale(locale.LC_TIME, 'pt_BR.UTF-8')
except locale.Error:
    pass

app = Flask(__name__)
app.secret_key = 'sua_chave_secreta'  # Altere para uma chave segura
ACCESS_TOKEN = "minha_senha"  # Token de acesso

app.config['UPLOAD_FOLDER'] = 'uploads'
ALLOWED_EXTENSIONS = {'.jpg', '.jpeg', '.png', '.bmp', '.gif'}

# Caminho do arquivo CSV relativo ao diretório do script
CSV_PATH = os.path.join(os.path.dirname(__file__), "uploads", "escolas.csv")

# Variável global para armazenar os dados do CSV
escolas_df = None

# Função para carregar o arquivo de escolas
def carregar_escolas():
    global escolas_df
    if os.path.exists(CSV_PATH):
        try:
            # Usando encoding 'latin1' (ou 'cp1252')
            escolas_df = pd.read_csv(CSV_PATH, encoding="latin1", sep=";")
            print(f"[INFO] Arquivo {CSV_PATH} carregado com sucesso.")
        except Exception as e:
            print(f"[ERRO] Falha ao carregar {CSV_PATH}: {e}")
    else:
        print(f"[ERRO] Arquivo {CSV_PATH} não encontrado.")

# Função para garantir que os dados estão disponíveis
def get_escolas_df():
    global escolas_df
    if escolas_df is None or escolas_df.empty:
        print("[INFO] Recarregando arquivo escolas.csv...")
        carregar_escolas()
    return escolas_df

@app.route('/escolas/search')
def escolas_search():
    df = get_escolas_df()
    query = request.args.get('q', '').lower().strip()
    results = []

    if df is not None and not df.empty and query:
        # Filtra usando pandas
        df_filtered = df[df.iloc[:, 3].str.lower().str.contains(query, na=False)]

        # Limita a 50 resultados para não sobrecarregar
        df_filtered = df_filtered.head(50)

        for idx, row in df_filtered.iterrows():
            nome = str(row[3]).strip()
            municipio = str(row[2]).strip()
            uf = str(row[1]).strip()
            text = f"{nome} - {municipio}/{uf}"
            results.append({
                "id": nome,   # ou algum identificador único
                "text": text
            })

    return jsonify(results)

@app.before_request
def inicializar_escolas():
    if escolas_df is None or escolas_df.empty:
        carregar_escolas()

# Cria os diretórios necessários, se não existirem
if not os.path.exists('static/fotos'):
    os.makedirs('static/fotos')
if not os.path.exists(app.config['UPLOAD_FOLDER']):
    os.makedirs(app.config['UPLOAD_FOLDER'])

# Carrega o CSV na inicialização do sistema
carregar_escolas()

# Importa e registra o blueprint do confere.py
from confere import confere_bp
app.register_blueprint(confere_bp, url_prefix='/confere')

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

# Função para atualizar valor em célula mesclada (mantém a mesclagem)
def set_merged_cell_value(ws, cell_coord, value):
    cell = ws[cell_coord]
    if isinstance(cell, MergedCell):
        # Procura o intervalo mesclado que contém a célula
        for merged_range in ws.merged_cells.ranges:
            if cell_coord in merged_range:
                range_str = str(merged_range)
                ws.unmerge_cells(range_str)
                # Obtém a célula superior esquerda do intervalo mesclado
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
    # Remover a planilha padrão criada pelo openpyxl, se houver
    if "Sheet" in wb.sheetnames and len(book_xlrd.sheet_names()) > 0:
        std = wb.active
        wb.remove(std)

    for sheet_name in book_xlrd.sheet_names():
        sheet_xlrd = book_xlrd.sheet_by_name(sheet_name)
        ws = wb.create_sheet(title=sheet_name)
        for row in range(sheet_xlrd.nrows):
            for col in range(sheet_xlrd.ncols):
                ws.cell(row=row+1, column=col+1, value=sheet_xlrd.cell_value(row, col))

    return wb


def load_workbook_model(file):
    """
    Abre o arquivo do modelo XLSX (ou XLS convertendo-o para XLSX) preservando toda a formatação.
    """
    ext = os.path.splitext(file.filename)[1].lower()
    file.seek(0)     
    if ext == '.xlsx':
        return load_workbook(file, data_only=False)
    elif ext == '.xls':
        content = file.read()
        return convert_xls_to_xlsx(BytesIO(content))
    else:
        raise ValueError("Formato de arquivo não suportado para o quadro modelo.")


def gerar_html_carteirinhas(arquivo_excel):
    planilha = pd.read_excel(arquivo_excel, sheet_name='LISTA CORRIDA')
    dados = planilha[['RM', 'NOME', 'DATA NASC.', 'RA', 'SAI SOZINHO?', 'SÉRIE', 'HORÁRIO']]
    dados['RM'] = dados['RM'].fillna(0).astype(int)

    alunos_sem_fotos_list = []
    html_content = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>E.M José Padin Mouta - Carteirinhas</title>
  <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
  <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
  <style>
    /* Estilos CSS para carteirinhas */
    body {
      font-family: 'Montserrat', sans-serif;
      margin: 0;
      padding: 20px;
      background-color: #f4f4f4;
      display: flex;
      flex-direction: column;
      align-items: center;
    }
    #search-container {
      margin-top: 10px;
    }
    #localizarAluno {
      padding: 0.2cm;
      font-size: 0.3cm;
      width: 3.5cm;
    }
    .carteirinhas-container {
      width: 100%;
      max-width: 1100px;
    }
    .page {
      margin-bottom: 40px;
      position: relative;
    }
    .page-number {
      text-align: center;
      font-size: 0.3cm;
      font-weight: 600;
      color: #333;
      margin-bottom: 0.2cm;
    }
    .cards-grid {
      display: grid;
      grid-template-columns: repeat(2, 1fr);
      gap: 0.2cm;
      justify-items: center;
    }
    .borda-pontilhada {
      border: 0.05cm dotted #ccc;
      padding: 0.1cm;
      position: relative;
    }
    .borda-pontilhada::after {
      content: "✂️";
      position: absolute;
      top: -0.35cm;
      right: -0.30cm;
      font-size: 0.3cm;
      color: #2196F3;
    }
    input {
      width: 100%;
      padding: 0.2cm;
      margin: 0.1cm 0;
      border: 0.05cm solid #ccc;
      border-radius: 0.2cm;
      box-sizing: border-box;
      font-size: 0.3cm;
    }
    input:focus {
      border-color: #008CBA;
      box-shadow: 0 0 0.2cm rgba(0, 140, 186, 0.5);
      outline: none;
    }
    .carteirinha {
      background-color: #fff;
      border-radius: 0.3cm;
      box-shadow: 0 0.1cm 0.2cm rgba(0,0,0,0.1);
      overflow: hidden;
      display: flex;
      flex-direction: column;
      width: 6.0cm;
      height: 9.0cm;
      padding: 0.2cm;
      position: relative;
      border: 0.05cm solid #2196F3;
    }
    .escola {
      font-size: 0.35cm;
      font-weight: 600;
      color: #2196F3;
      margin-bottom: 0.1cm;
      text-align: center;
      text-transform: uppercase;
      letter-spacing: 0.05cm;
      margin-top: 0.1cm;
      white-space: nowrap;
    }
    .foto {
      width: 1.8cm;
      height: 1.8cm;
      margin-bottom: 0.1cm;
      border-radius: 50%;
      object-fit: cover;
      margin-left: auto;
      margin-right: auto;
      border: 0.1cm solid #2196F3;
      cursor: pointer;
    }
    .info {
      display: flex;
      flex-direction: column;
      align-items: flex-start;
      text-align: left;
      margin-left: 0.1cm;
      margin-bottom: 0.1cm;
      font-size: 0.3cm;
      color: #333;
    }
    .info div, .info span { margin: 0.08cm 0; }
    .info .titulo {
      font-weight: 600;
      color: #2196F3;
      text-transform: uppercase;
      letter-spacing: 0.02cm;
    }
    .info .descricao { color: #555; }
    .linha-nome { display: flex; align-items: center; gap: 0.1cm; }
    .linha, .linha-ra, .linha-horario, .linha-rm { display: flex; flex-direction: row; align-items: center; gap: 0.2cm; }
    .status {
      padding: 0.2cm;
      font-weight: 600;
      border-radius: 0.2cm;
      color: #fff;
      text-transform: uppercase;
      margin-bottom: 0.1cm;
      display: flex;
      justify-content: center;
      align-items: center;
      height: 0.6cm;
      min-width: 1.5cm;
      text-align: center;
    }
    .verde { background-color: #81C784; }
    .vermelho { background-color: #E57373; }
    .ano {
      position: absolute;
      bottom: 0.2cm;
      left: 0;
      right: 0;
      text-align: center;
      font-size: 0.4cm;
      font-weight: 600;
      color: #2196F3;
    }
    #loading-overlay {
      position: fixed;
      top: 0; left: 0; right: 0; bottom: 0;
      background: rgba(0, 0, 0, 0.5);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 9999;
    }
    #cards-success {
      display: none;
      position: fixed;
      top: 10px;
      left: 50%;
      transform: translateX(-50%);
      background: #d4edda;
      color: #155724;
      padding: 0.2cm;
      border-radius: 0.2cm;
      z-index: 10000;
    }
    .no-print { }
    @media print {
      .no-print { display: none !important; }
      body {
        margin: 0;
        padding: 0;
        font-size: 16px;
        background-color: #fff !important;
      }
      .page {
        page-break-after: always;
      }
    }
    .imprimir-carteirinhas {
      position: fixed;
      bottom: 0.5cm;
      right: 0.5cm;
      background-color: #2196F3;
      color: #fff;
      padding: 0.2cm 0.4cm;
      font-size: 0.3cm;
      border-radius: 0.2cm;
      cursor: pointer;
      box-shadow: 0 0.1cm 0.2cm rgba(0,0,0,0.2);
    }
    .imprimir-pagina {
      background-color: #FF5722;
      color: #fff;
      padding: 0.2cm 0.4cm;
      font-size: 0.3cm;
      border-radius: 0.2cm;
      cursor: pointer;
      margin: 0.2cm auto;
      display: block;
    }
    .imprimir-pagina:hover {
      background-color: #FF7043;
    }
    .alunos-sem-fotos-btn {
      background-color: #4B79A1;
      color: #fff;
      border: none;
      padding: 0.2cm 0.5cm;
      font-size: 0.3cm;
      border-radius: 0.2cm;
      cursor: pointer;
      margin-bottom: 0.2cm;
    }
    .alunos-sem-fotos-btn:hover {
      background-color: #3a5d78;
    }
    #relatorio-container {
      display: none;
      position: fixed;
      top: 10%;
      left: 50%;
      transform: translateX(-50%);
      width: 80%;
      max-height: 80%;
      overflow-y: auto;
      background: #fff;
      border: 1px solid #ccc;
      border-radius: 10px;
      box-shadow: 0 0 10px rgba(0,0,0,0.5);
      z-index: 10000;
      padding: 20px;
    }
    #relatorio-container h2 {
      text-align: center;
      margin-top: 0;
    }
    #relatorio-container table {
      width: 100%;
      border-collapse: collapse;
    }
    #relatorio-container th, #relatorio-container td {
      border: 1px solid #ccc;
      padding: 8px;
      text-align: left;
    }
    #relatorio-container button.close-relatorio {
      float: right;
      font-size: 1.2em;
      border: none;
      background: none;
      cursor: pointer;
    }
    header {
      background: linear-gradient(90deg, #283E51, #4B79A1);
      color: #fff;
      padding: 20px;
      text-align: center;
      border-bottom: 3px solid #1d2d3a;
      border-radius: 0 0 15px 15px;
      box-shadow: 0 4px 6px rgba(0,0,0,0.1);
    }
  </style>
</head>
<body>
  <div id="loading-overlay">
    <div style="text-align: center; color: white;">
      <div class="spinner-border" role="status">
        <span class="sr-only">Carregando...</span>
      </div>
      <p>Carregando carteirinhas...</p>
    </div>
  </div>
  <div id="cards-success">Carteirinhas geradas com sucesso</div>
  <div class="carteirinhas-container">
    <div class="no-print" style="margin-bottom: 10px;">
      <button class="alunos-sem-fotos-btn" onclick="mostrarRelatorioAlunosSemFotos()">Alunos sem fotos</button>
      <button class="imprimir-carteirinhas" onclick="imprimirCarteirinhas()">Imprimir Carteirinhas</button>
    </div>
    <div id="search-container" class="no-print">
      <input type="text" id="localizarAluno" placeholder="Localizar Aluno">
    </div>
"""
    contador = 0
    num_pagina = 1
    html_content += '<div class="page"><div class="page-number">Página ' + str(num_pagina) + '</div>'
    html_content += '<button class="imprimir-pagina no-print" onclick="imprimirPagina(this)">Imprimir Página</button>'
    html_content += '<div class="cards-grid">'
    for _, row in dados.iterrows():
        if not str(row['RM']).strip() or str(row['RM']).strip() == "0":
            continue
        nome = row['NOME']
        data_nasc = row['DATA NASC.']
        serie = row['SÉRIE']
        horario = row['HORÁRIO']
        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors='coerce')
                if pd.notna(data_nasc):
                    data_nasc = data_nasc.strftime('%d/%m/%Y')
                else:
                    data_nasc = "Desconhecida"
            except Exception:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"
        ra = row['RA']
        sai_sozinho = row['SAI SOZINHO?']
        if sai_sozinho == 'Sim':
            classe_cor = 'verde'
            status_texto = "Sai Sozinho"
        else:
            classe_cor = 'vermelho'
            status_texto = "Não Sai Sozinho"
        allowed_exts = ['.jpg', '.jpeg', '.png', '.bmp', '.gif']
        found_photo = None
        for ext in allowed_exts:
            caminho_foto = f'static/fotos/{row["RM"]}{ext}'
            if os.path.exists(caminho_foto):
                found_photo = f"/static/fotos/{row['RM']}{ext}"
                break
        if not found_photo:
            alunos_sem_fotos_list.append({
                'rm': row['RM'],
                'nome': nome,
                'serie': serie
            })
        if found_photo:
            foto_tag = f'<img src="{found_photo}" alt="Foto" class="foto uploadable" data-rm="{row["RM"]}">'
        else:
            foto_tag = f'''
            <div class="foto uploadable" data-rm="{row["RM"]}" style="display: flex; flex-direction: column; align-items: center; justify-content: center;">
              <span style="font-size:0.8cm; opacity:0.5; color: grey; margin-bottom: 0.1cm;">&#128247;</span>
              <small style="font-size:0.2cm; opacity:0.5; color: grey;">Anexe uma foto</small>
            </div>
            '''
        hidden_input = f'<input type="file" class="inline-upload" data-rm="{row["RM"]}" style="display:none;" accept="image/*">'
        html_content += f"""
      <div class="borda-pontilhada">
        <div class="carteirinha">
          <div class="escola">E.M José Padin Mouta</div>
          {foto_tag}
          {hidden_input}
          <div class="info">
            <div class="linha-nome">
              <span class="titulo">Nome:</span>
              <span class="descricao">{nome}</span>
            </div>
            <div class="linha-rm">
              <span class="titulo">RM:</span>
              <span class="descricao">{row['RM']}</span>
            </div>
            <div class="linha">
              <div class="titulo">Série:</div>
              <div class="descricao">{serie}</div>
            </div>
            <div class="linha">
              <div class="titulo">Data Nasc.:</div>
              <div class="descricao">{data_nasc}</div>
            </div>
            <div class="linha-ra">
              <span class="titulo">RA:</span>
              <span class="descricao">{ra}</span>
            </div>
            <div class="linha-horario">
              <div class="titulo">Horário:</div>
              <div class="descricao">{horario}</div>
            </div>
          </div>
          <div class="status {classe_cor}">{status_texto}</div>
          <div class="ano">2025</div>
        </div>
      </div>
"""
        contador += 1
        if contador % 4 == 0:
            html_content += '</div></div>'
            if contador < len(dados):
                num_pagina += 1
                html_content += '<div class="page"><div class="page-number">Página ' + str(num_pagina) + '</div>'
                html_content += '<button class="imprimir-pagina no-print" onclick="imprimirPagina(this)">Imprimir Página</button>'
                html_content += '<div class="cards-grid">'
    if contador % 4 != 0:
        html_content += '</div></div>'
    relatorio_linhas = ""
    for aluno in alunos_sem_fotos_list:
        relatorio_linhas += f"<tr><td>{aluno['rm']}</td><td>{aluno['nome']}</td><td>{aluno['serie']}</td></tr>"
    html_content += f"""
  </div>
  <div id="relatorio-container" class="no-print">
    <button class="close-relatorio" onclick="fecharRelatorio()">&times;</button>
    <h2>Alunos sem Fotos</h2>
    <table>
      <thead>
        <tr>
          <th>RM</th>
          <th>Nome</th>
          <th>Série</th>
        </tr>
      </thead>
      <tbody>
        {relatorio_linhas}
      </tbody>
    </table>
  </div>
<script>
// Função para confirmar o envio se a declaração for do tipo Transferencia
function confirmDeclaration() {{
    var tipo = document.getElementById('tipo').value;
    if (tipo === "Transferencia") {{
        return confirm("Você está gerando uma declaração de transferência, essa é a declaração correta a ser gerada?");
    }}
    return true;
}}

function showLoading() {{
    var existingOverlay = document.getElementById('loading-overlay');
    if (existingOverlay) {{
      existingOverlay.remove();
    }}

    var loadingOverlay = document.createElement('div');
    loadingOverlay.id = 'loading-overlay';
    loadingOverlay.style.position = 'fixed';
    loadingOverlay.style.top = '0';
    loadingOverlay.style.left = '0';
    loadingOverlay.style.right = '0';
    loadingOverlay.style.bottom = '0';
    loadingOverlay.style.background = 'rgba(0,0,0,0.5)';
    loadingOverlay.style.display = 'flex';
    loadingOverlay.style.alignItems = 'center';
    loadingOverlay.style.justifyContent = 'center';
    loadingOverlay.style.zIndex = '9999';

    // AQUI é um string normal; não substituímos nada dentro da string
    loadingOverlay.innerHTML = 
      `<div style="text-align: center; color: white; font-family: Arial, sans-serif;">
        <svg width="3.0cm" height="4.5cm" viewBox="0 0 6.0 9.0" xmlns="http://www.w3.org/2000/svg">
          <rect x="0.3" y="0.3" width="5.4" height="8.4" rx="0.3" ry="0.3" stroke="white" stroke-width="0.1" fill="none" />
          <rect id="badge-fill" x="0.3" y="8.7" width="5.4" height="0" rx="0.3" ry="0.3" fill="white" />
        </svg>
        <p id="loading-text" style="margin-top: 0.2cm;">Gerando carteirinhas...</p>
      </div>`;

    document.body.appendChild(loadingOverlay);

    let fillHeight = 0;
    const maxHeight = 8.4; 
    function animateBadge() {{
      fillHeight += 0.2;
      if (fillHeight > maxHeight) {{
        fillHeight = maxHeight;
        clearInterval(interval);
      }}
      const badgeFill = document.getElementById('badge-fill');
      badgeFill.setAttribute('y', 8.7 - fillHeight);
      badgeFill.setAttribute('height', fillHeight);
    }}

    var interval = setInterval(animateBadge, 100);
    loadingOverlay.dataset.animationId = interval;
}}

// Chama showLoading() imediatamente
showLoading();

// Quando a janela terminar de carregar
window.onload = function() {{
    var overlay = document.getElementById('loading-overlay');
    if (overlay) {{
      var animationId = Number(overlay.dataset.animationId);
      clearInterval(animationId);
      overlay.style.display = 'none';
    }}
    var cardsMsg = document.getElementById('cards-success');
    if (cardsMsg) {{
      cardsMsg.style.display = 'block';
      cardsMsg.innerHTML = 'Carteirinhas geradas com sucesso!';
      setTimeout(function() {{
        cardsMsg.style.display = 'none';
      }}, 3000);
    }}
}};

// Imprimir todas as carteirinhas
function imprimirCarteirinhas() {{
    window.print();
}}

// Imprimir só a página em que o botão foi clicado
function imprimirPagina(botao) {{
    let pagina = botao.closest('.page');
    let todasPaginas = document.querySelectorAll('.page');
    todasPaginas.forEach(p => {{
      if (p !== pagina) {{
        p.style.display = 'none';
      }}
    }});
    setTimeout(() => {{
      window.print();
      // Restaura a visibilidade
      todasPaginas.forEach(p => {{
        p.style.display = '';
      }});
    }}, 100);
}}

// Exibir relatório de Alunos sem fotos
function mostrarRelatorioAlunosSemFotos() {{
    document.getElementById('relatorio-container').style.display = 'block';
}}

// Fechar relatório de Alunos sem fotos
function fecharRelatorio() {{
    document.getElementById('relatorio-container').style.display = 'none';
}}

// Filtro de busca pelo nome do aluno
document.getElementById('localizarAluno').addEventListener('keyup', function() {{
    var filtro = this.value.toLowerCase();
    var cards = document.querySelectorAll('.borda-pontilhada');
    cards.forEach(function(card) {{
      var nomeElem = card.querySelector('.linha-nome .descricao');
      if (nomeElem) {{
        var nome = nomeElem.textContent.toLowerCase();
        if (nome.indexOf(filtro) > -1) {{
          card.style.display = '';
        }} else {{
          card.style.display = 'none';
        }}
      }}
    }});
}});

var flashTimeout = null;
document.addEventListener('DOMContentLoaded', function() {{
    document.querySelectorAll('.uploadable').forEach(function(element) {{
      element.addEventListener('click', function() {{
        var rm = element.getAttribute('data-rm');
        var input = document.querySelector('.inline-upload[data-rm="'+rm+'"]');
        if(input) {{
          input.click();
        }}
      }});
    }});
    
    document.querySelectorAll('.inline-upload').forEach(function(input) {{
      input.addEventListener('change', function() {{
        var file = input.files[0];
        if(file) {{
          var rm = input.getAttribute('data-rm');
          var formData = new FormData();
          formData.append('rm', rm);
          formData.append('foto_file', file);
          
          fetch('/upload_inline_foto', {{
            method: 'POST',
            body: formData
          }})
          .then(response => response.json())
          .then(data => {{
            if(data.url) {{
              var uploadable = document.querySelector('.uploadable[data-rm="'+rm+'"]');
              if(uploadable.tagName.toLowerCase() === 'img') {{
                uploadable.src = data.url;
              }} else {{
                var img = document.createElement('img');
                img.src = data.url;
                img.alt = "Foto";
                img.className = "foto uploadable";
                img.setAttribute('data-rm', rm);
                uploadable.parentNode.replaceChild(img, uploadable);
              }}
              var msgDiv = document.getElementById('upload-success');
              if(!msgDiv) {{
                msgDiv = document.createElement('div');
                msgDiv.id = 'upload-success';
                msgDiv.style.position = 'fixed';
                msgDiv.style.top = '0.2cm';
                msgDiv.style.right = '0.2cm';
                msgDiv.style.backgroundColor = '#d4edda';
                msgDiv.style.color = '#155724';
                msgDiv.style.padding = '0.2cm';
                msgDiv.style.borderRadius = '0.2cm';
                document.body.appendChild(msgDiv);
              }}
              msgDiv.style.display = 'block';
              msgDiv.innerHTML = data.message;
              if(flashTimeout) {{
                clearTimeout(flashTimeout);
              }}
              flashTimeout = setTimeout(function() {{
                msgDiv.style.display = 'none';
              }}, 3000);
            }} else {{
              alert("Erro ao fazer upload: " + (data.error || "Erro desconhecido"));
            }}
          }})
          .catch(error => {{
            console.error('Erro:', error);
            alert("Erro no upload da foto.");
          }});
        }}
      }});
    }});
}});
</script>
</body>
</html>
"""
    return render_template_string(html_content)

from datetime import datetime 
import pandas as pd
import re
from flask import session

def gerar_declaracao_escolar(file_path, rm, tipo, file_path2=None, deve_historico=False, unidade_anterior=None):
    # Verifica o tipo de declaração salvo na sessão para distinguir Regular e EJA
    if session.get('declaracao_tipo') != "EJA":
        # Parte Regular - leitura da planilha "LISTA CORRIDA"
        planilha = pd.read_excel(file_path, sheet_name='LISTA CORRIDA')
        planilha.columns = [c.strip().upper() for c in planilha.columns]  # normaliza colunas

        def format_rm(x):
            try:
                return str(int(float(x)))
            except:
                return str(x)

        planilha['RM_str'] = planilha['RM'].apply(format_rm)
        try:
            rm_num = str(int(float(rm)))
        except:
            rm_num = str(rm)

        aluno = planilha[planilha['RM_str'] == rm_num]
        if aluno.empty:
            return None

        row = aluno.iloc[0]

        semestre_texto = ""
        nome = row['NOME']
        serie = row['SÉRIE']
        if isinstance(serie, str):
            serie = re.sub(r"(\d+º)([A-Za-z])", r"\1 ano \2", serie)
        data_nasc = row['DATA NASC.']
        ra = row['RA']
        horario = row.get('HORÁRIO', 'Desconhecido')
        if pd.isna(horario) or not str(horario).strip():
            horario = "Desconhecido"
        else:
            horario = str(horario).strip()

        ra_label = "RA"
        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors='coerce')
                if pd.notna(data_nasc):
                    data_nasc = data_nasc.strftime('%d/%m/%Y')
                else:
                    data_nasc = "Desconhecida"
            except Exception:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"

    else:
        # Parte EJA - leitura da planilha na primeira aba, sem cabeçalho e com skiprows=1
        df = pd.read_excel(file_path, sheet_name=0, header=None, skiprows=1)
        df.columns = [str(c).strip().upper() for c in df.columns]  # normaliza colunas

        df['RM_str'] = df.iloc[:, 2].apply(lambda x: str(int(x)) if pd.notna(x) and float(x) != 0 else "")
        df['NOME'] = df.iloc[:, 3]
        df['NASC.'] = df.iloc[:, 6]

        def get_ra(row):
            try:
                val = row.iloc[7]
                if pd.isna(val) or float(val) == 0:
                    return row.iloc[8]
                else:
                    return val
            except:
                return row.iloc[7]

        df['RA'] = df.apply(get_ra, axis=1)
        df['SÉRIE'] = df.iloc[:, 0]

        try:
            rm_num = str(int(float(rm)))
        except:
            rm_num = str(rm)

        aluno = df[df['RM_str'] == rm_num]
        if aluno.empty:
            return None

        row = aluno.iloc[0]

        if len(row) > 29:
            semestre = row.iloc[29]
            semestre_texto = str(semestre).strip() if pd.notna(semestre) else ""
        else:
            semestre_texto = ""

        nome = row['NOME']
        serie = row['SÉRIE']
        if isinstance(serie, str):
            serie = re.sub(r"(\d+º)([A-Za-z])", r"\1 ano \2", serie)
        data_nasc = row['NASC.']
        ra = row['RA']
        original_ra = row.iloc[7]
        ra_label = "RG" if pd.isna(original_ra) or (isinstance(original_ra, (int, float)) and float(original_ra) == 0) else "RA"

        if pd.notna(data_nasc):
            try:
                data_nasc = pd.to_datetime(data_nasc, errors='coerce')
                data_nasc = data_nasc.strftime('%d/%m/%Y') if pd.notna(data_nasc) else "Desconhecida"
            except Exception:
                data_nasc = "Desconhecida"
        else:
            data_nasc = "Desconhecida"

    # Data por extenso
    now = datetime.now()
    meses = {
        1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
        7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"
    }
    mes = meses[now.month].capitalize()
    data_extenso = f"Praia Grande, {now.day:02d} de {mes} de {now.year}"

    additional_css = '''
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
'''

# --- Montagem do texto da declaração --- 
    # --- Montagem do texto da declaração --- 
    declaracao_text = ""
    if tipo == "Escolaridade":
        titulo = "Declaração de Escolaridade"
        if session.get('declaracao_tipo') == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
                f"portador(a) do {ra_label} <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
                f"encontra-se regularmente matriculado(a) na E.M José Padin Mouta, cursando atualmente o "
                f"<strong><u>{serie}</u></strong>."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
                f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
                f"encontra-se regularmente matriculado(a) na E.M José Padin Mouta, cursando atualmente o "
                f"<strong><u>{serie}</u></strong> no horário de aula: <strong><u>{horario}</u></strong>."
            )

    elif tipo == "Transferencia":
        titulo = "Declaração de Transferência"
        if session.get('declaracao_tipo') == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
                f"portador(a) do {ra_label} <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
                f"solicitou transferência de nossa unidade escolar na data de hoje, estando apto(a) a cursar o "
                f"<strong><u>{serie}</u></strong>."
            )
        else:
            serie_mod = re.sub(r"^(\d+º).*", r"\1 ano", serie)
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) responsável do(a) aluno(a) <strong><u>{nome}</u></strong>, "
                f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
                f"compareceu a nossa unidade escolar e solicitou transferência na data de hoje, o aluno está apto(a) a cursar o "
                f"<strong><u>{serie_mod}</u></strong>."
            )

    elif tipo == "Conclusão":
        titulo = "Declaração de Conclusão"
        if session.get('declaracao_tipo') == "EJA":
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
                "3ª SÉRIE E.M": "ENSINO SUPERIOR"
            }
            series_text = mapping.get(serie, "a série subsequente")
        else:
            match = re.search(r"(\d+)º\s*ano", serie)
            next_year = int(match.group(1)) + 1 if match else None
            series_text = f"{next_year}º ano" if next_year else "a série subsequente"

        if session.get('declaracao_tipo') == "EJA":
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
                f"portador(a) do {ra_label} <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
                f"concluiu com êxito o <strong><u>{serie}</u></strong>{f', no <strong><u>{semestre_texto}</strong></u>' if semestre_texto else ''}, "
                f"estando apto(a) a ingressar no <strong><u>{series_text}</u></strong>."
            )
        else:
            declaracao_text = (
                f"Declaro, para os devidos fins, que o(a) aluno(a) <strong><u>{nome}</u></strong>, "
                f"portador(a) do RA <strong><u>{ra}</u></strong>, nascido(a) em <strong><u>{data_nasc}</u></strong>, "
                f"concluiu com êxito o <strong><u>{serie}</u></strong>, estando apto(a) a ingressar no "
                f"<strong><u>{series_text}</u></strong>."
            )

    # --- Observações e Bolsa Família ---
    valor_bolsa = str(row.get('BOLSA FAMILIA', '')).strip().upper()
    if deve_historico or (valor_bolsa == "SIM" and tipo != "Escolaridade"):
        declaracao_text += "<br><br><strong>Observações:</strong><br>"
        declaracao_text += '<label class="checkbox-label" style="display: block; text-align: justify;">'

        # Histórico, se houver
        if deve_historico:
            declaracao_text += '<span class="warning-icon">&#9888;</span> '
            declaracao_text += "O aluno deve o histórico escolar da unidade anterior:<br><br>"
            if unidade_anterior:
                unidade_anterior = ' '.join(unidade_anterior.strip().split())
                esc_df = escolas_df[escolas_df.iloc[:, 3].str.upper() == unidade_anterior.upper()] if escolas_df is not None else None
                if esc_df is not None and not esc_df.empty:
                    unidade_nome = esc_df.iloc[0, 3]
                    inep = esc_df.iloc[0, 4]
                    municipio = esc_df.iloc[0, 2]
                    uf = esc_df.iloc[0, 1]
                    declaracao_text += (
                        f"<strong>Unidade:</strong> {unidade_nome}<br>"
                        f"<strong>INEP:</strong> {inep}<br>"
                        f"<strong>Cidade:</strong> {municipio}<br>"
                        f"<strong>Estado:</strong> {uf}<br><br>"
                    )
                else:
                    # Caso manual, mostra diretamente
                    declaracao_text += f"<strong>Unidade:</strong> {unidade_anterior}<br><br>"
            declaracao_text += "Após sua entrega, o documento será confeccionado em até 30 dias úteis.<br><br>"

        # Bolsa Família, se houver
        if valor_bolsa == "SIM" and tipo != "Escolaridade":
            declaracao_text += '<img src="/static/logos/bolsa_familia.jpg" alt="Bolsa Família" style="width:28px; vertical-align:middle; margin-right:5px;">'
            declaracao_text += "O aluno é beneficiário do Programa Bolsa Família."

        declaracao_text += "</label>"
    base_template = f'''<!doctype html>
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
      padding: 0 2cm;  /* aumentar margem lateral para evitar texto colado */
      box-sizing: border-box;
      hyphens: auto; /* ativa hifenização automática */
      word-wrap: break-word;
      overflow-wrap: break-word;
    }}
    .content p {{
      white-space: normal !important;
      word-break: break-word !important; /* evita quebra no meio da palavra sem hífen */
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
    padding: 1.5cm 1.5cm;  /* margens laterais reduzidas para 1.5cm */
    font-size: 16px;
    font-family: 'Montserrat', sans-serif;
    color: #000;
  }}

  .declaration-bottom {{
    position: fixed;   /* fica fixo na página impressa */
    bottom: 0;         /* colado ao fundo da página */
    left: 0;
    width: 100%;
    text-align: center; /* centraliza se quiser */
  }}

  .date {{
    margin: 1cm 0 1cm 0;  /* margem topo reduzida para 1cm */
    text-align: right;
    hyphens: none !important; /* sem hifenização na data */
  }}

  .content {{
    margin: 0;
    padding: 0;
    text-align: justify !important;
    hyphens: none !important; /* desativa hifenização */
    white-space: normal !important;
    word-wrap: break-word !important;
    overflow-wrap: break-word !important;
  }}

  .content p {{
    margin: 0 0 1em 0;
    text-align: justify !important;
    hyphens: none !important; /* desativa hifenização */
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

/* CSS para checkbox alinhado */
.checkbox-label {{
  display: flex;
  align-items: center;
  gap: 8px;
  text-align: left !important;
  font-size: 16px;
  margin-top: 8px;
  margin-bottom: 8px;
  flex-wrap: wrap;
}}

.checkbox-label input[type="checkbox"] {{
  width: 18px;
  height: 18px;
  margin: 0;
  vertical-align: middle;
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
      <div style="display: flex; justify-content: space-between; align-items: center;">
        <img src="/static/logos/escola.png" alt="Escola Logo" style="height: 80px;">
        <div>
          <h1>Secretaria de Educação</h1>
          <p>E.M José Padin Mouta</p>
          <p>Município da Estância Balneária de Praia Grande</p>
          <p>Estado de São Paulo</p>
        </div>
        <img src="/static/logos/municipio.png" alt="Município Logo" style="height: 80px;">
      </div>
    </div>
    <div class="date">
      <p>{data_extenso}</p>
    </div>
    <div class="content">
      <h2 style="text-align: center; text-transform: uppercase; color: #283E51;">{titulo}</h2>
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
  <div class="no-print" style="text-align: center; margin-top: 20px;">
    <button onclick="window.print()" class="print-button">Imprimir Declaração</button>
  </div>
</body>
</html>
'''

    return base_template

@app.route('/login', methods=['GET', 'POST'])
def login_route():
    error = None
    if request.method == 'POST':
        token = request.form.get('token')
        if token == ACCESS_TOKEN:
            session['logged_in'] = True
            if 'lista_fundamental' not in session or 'lista_eja' not in session:
                return redirect(url_for('upload_listas'))
            return redirect(url_for('dashboard'))
        else:
            error = "Token inválido. Tente novamente."

    login_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>Login - Acesso Restrito</title>
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <style>
          body {
            background: linear-gradient(135deg, #283E51, #4B79A1);
            font-family: 'Montserrat', sans-serif;
            margin: 0;
            padding: 0;
            display: flex;
            flex-direction: column;
            min-height: 100vh;
          }
          header {
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          }
          main {
            flex: 1;
            display: flex;
            align-items: center;
            justify-content: center;
            padding: 20px;
          }
          .container-login {
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            width: 100%;
            max-width: 400px;
          }
          .container-login h2 {
            margin-bottom: 20px;
            font-weight: 600;
            color: #283E51;
          }
          .btn-primary {
            background-color: #283E51;
            border: none;
          }
          .btn-primary:hover {
            background-color: #1d2d3a;
          }
          footer {
            background-color: #424242;
            color: #fff;
            text-align: center;
            padding: 10px;
          }
          .error {
            color: #ff0000;
            margin-top: 15px;
          }
        </style>
      </head>
      <body>
        <header>
          <h1>E.M José Padin Mouta - Secretaria</h1>
        </header>
        <main class="container">
          <div class="container-login">
            <h2 class="text-center">Acesso Restrito</h2>
            <form method="POST">
              <div class="form-group">
                <input type="password" name="token" class="form-control" placeholder="Digite o token de acesso" required>
              </div>
              <button type="submit" class="btn btn-primary btn-block mt-3">Entrar</button>
            </form>
            {% if error %}
              <div class="text-center">
                <p class="error">{{ error }}</p>
              </div>
            {% endif %}
          </div>
        </main>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
      </body>
    </html>
    '''
    return render_template_string(login_html, error=error)


@app.route('/logout')
def logout_route():
    session.clear()
    return redirect(url_for('login_route'))

# Nova rota para upload prévio das listas piloto
@app.route('/upload_listas', methods=['GET', 'POST'])
@login_required
def upload_listas():
    if request.method == 'POST':
        fundamental_file = request.files.get('lista_fundamental')
        eja_file = request.files.get('lista_eja')

        if not fundamental_file or fundamental_file.filename == '':
            flash("Selecione a Lista Piloto - REGULAR - 2025", "error")
            return redirect(url_for('upload_listas'))

        if not eja_file or eja_file.filename == '':
            flash("Selecione a Lista Piloto - EJA - 1º SEM - 2025", "error")
            return redirect(url_for('upload_listas'))

        fundamental_filename = secure_filename(f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename)
        eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_" + eja_file.filename)

        fundamental_path = os.path.join(app.config['UPLOAD_FOLDER'], fundamental_filename)
        eja_path = os.path.join(app.config['UPLOAD_FOLDER'], eja_filename)

        fundamental_file.save(fundamental_path)
        eja_file.save(eja_path)

        session['lista_fundamental'] = fundamental_path
        session['lista_eja'] = eja_path

        flash("Listas carregadas com sucesso.", "success")
        return redirect(url_for('dashboard'))

    upload_listas_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Upload de Listas Piloto</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
      body {
          background: #eef2f3;
          font-family: 'Montserrat', sans-serif;
      }
      header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
      }
      .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
      }
      .btn-primary {
          background-color: #283E51;
          border: none;
      }
      .btn-primary:hover {
          background-color: #1d2d3a;
      }
      footer {
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
      }
      </style>
    </head>
    <body>
      <header>
        <h1>Upload de Listas Piloto</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="lista_fundamental">Selecione a Lista Piloto - Regular:</label>
            <input type="file" class="form-control-file" name="lista_fundamental" id="lista_fundamental" accept=".xlsx, .xls" required>
          </div>
          <div class="form-group">
            <label for="lista_eja">Selecione a Lista Piloto - EJA:</label>
            <input type="file" class="form-control-file" name="lista_eja" id="lista_eja" accept=".xlsx, .xls" required>
          </div>
          <button type="submit" class="btn btn-primary">Carregar Listas</button>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_listas_html)

@app.route('/', methods=['GET'])
@login_required
def dashboard():
    dashboard_html = '''
<!doctype html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8">
    <title>E.M José Padin Mouta - Secretaria</title>
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap e Font Awesome -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <!-- Google Fonts -->
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
    <style>
      body {
        background: #eef2f3;
        font-family: 'Montserrat', sans-serif;
        margin-bottom: 60px; /* Espaço para o rodapé fixo */
      }
      header {
        background: linear-gradient(90deg, #283E51, #4B79A1);
        color: #fff;
        padding: 20px;
        text-align: center;
        border-bottom: 3px solid #1d2d3a;
        border-radius: 0 0 15px 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
      }
      .container-dashboard {
        background: #fff;
        padding: 40px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        margin: 40px auto;
        max-width: 900px;
      }
      .option-row {
        display: block;
      }
      .option-card {
        border: 1px solid #e0e0e0;
        border-radius: 10px;
        padding: 20px;
        cursor: pointer;
        background: #fff;
        margin-bottom: 20px;
        transition: transform 0.2s, box-shadow 0.2s;
      }
      .option-card:hover {
        transform: scale(1.02);
        box-shadow: 0 8px 16px rgba(0,0,0,0.1);
      }
      .option-content {
        flex: 1;
      }
      .option-icon {
        width: 60px;
        text-align: center;
        margin-right: 20px;
      }
      .option-icon i {
        font-size: 2rem;
        color: #283E51;
      }
      .option-card h2 {
        margin: 0 0 10px 0;
        font-size: 1.25rem;
        color: #283E51;
      }
      .option-card p {
        margin: 0;
        color: #555;
        font-size: 1rem;
      }
      .logout-container {
        text-align: center;
        margin-top: 30px;
      }
      .btn-logout {
        background-color: #dc3545;
        color: #fff;
        padding: 10px 25px;
        border: none;
        border-radius: 5px;
        font-size: 1rem;
        text-decoration: none;
        transition: background-color 0.3s;
      }
      .btn-logout:hover {
        background-color: #c82333;
      }
      footer {
        background-color: #424242;
        color: #fff;
        text-align: center;
        padding: 10px;
        position: fixed;
        bottom: 0;
        width: 100%;
      }
    </style>
  </head>
  <body>
    <header>
      <h1>E.M José Padin Mouta - Secretaria</h1>
    </header>
<div class="container container-dashboard"> 
  <div class="option-row">
    <!-- 1. Declaração Escolar -->
    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('declaracao_tipo') }}'">
      <div class="option-icon">
        <i class="fas fa-file-alt"></i>
      </div>
      <div class="option-content">
        <h2>Declaração Escolar</h2>
        <p>Gerar declaração escolar.</p>
      </div>
    </div>
    <!-- 2. Carteirinhas -->
    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('carteirinhas') }}'">
      <div class="option-icon">
        <i class="fas fa-id-card"></i>
      </div>
      <div class="option-content">
        <h2>Carteirinhas</h2>
        <p>Gerar carteirinhas para os alunos.</p>
      </div>
    </div>
    <!-- 3. Quadros -->
    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('quadros') }}'">
      <div class="option-icon">
        <i class="fas fa-chalkboard-teacher"></i>
      </div>
      <div class="option-content">
        <h2>Quadros</h2>
        <p>Gerar quadros para os alunos.</p>
      </div>
    </div>
    <!-- 4. Conferir Listas -->
    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('confere.index') }}'">
      <div class="option-icon">
        <i class="fas fa-check-circle"></i>
      </div>
      <div class="option-content">
        <h2>Conferir Listas</h2>
        <p>Acessar a conferência de listas.</p>
      </div>
    </div>
    <!--
    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('oficio') }}'">
      <div class="option-icon">
        <i class="fas fa-envelope"></i>
      </div>
      <div class="option-content">
        <h2>Ofício</h2>
        <p>Gerar ofícios.</p>
      </div>
    </div>

    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('acumulo') }}'">
      <div class="option-icon">
        <i class="fas fa-briefcase"></i>
      </div>
      <div class="option-content">
        <h2>Acúmulo de Cargo</h2>
        <p>Gerar formulário de acúmulo de cargo.</p>
      </div>
    </div>

    <div class="option-card d-flex align-items-center" onclick="window.location.href='{{ url_for('documentos') }}'">
      <div class="option-icon">
        <i class="fas fa-folder-open"></i>
      </div>
      <div class="option-content">
        <h2>Documentos</h2>
        <p>Documentos importantes por segmento.</p>
      </div>
    </div>
    -->
  </div>
  <div class="logout-container">
    <a href="{{ url_for('logout_route') }}" class="btn-logout">
      <i class="fas fa-sign-out-alt"></i> Logout
    </a>
  </div>
</div>
    <footer>
      Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
    </footer>
    <!-- Scripts do Bootstrap -->
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
  </body>
</html>


    '''
    return render_template_string(dashboard_html)


@app.route('/carteirinhas', methods=['GET', 'POST'])
@login_required
def carteirinhas():
    if request.method == 'POST':
        file = None
        if 'excel_file' in request.files and request.files['excel_file'].filename != '':
            file = request.files['excel_file']
            filename = secure_filename(file.filename)
            unique_filename = f"carteirinhas_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(file_path)
            session['lista_fundamental'] = file_path
            file = open(file_path, 'rb')
        else:
            file_path = session.get('lista_fundamental')
            if file_path and os.path.exists(file_path):
                file = open(file_path, 'rb')
        if not file:
            return "Nenhum arquivo selecionado", 400
        flash("Gerando carteirinhas. Aguarde...", "info")
        html_result = gerar_html_carteirinhas(file)
        file.close()
        return html_result
    carteirinhas_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>E.M José Padin Mouta - Carteirinhas</title>
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <style>
          body {
            background: #eef2f3;
            font-family: 'Montserrat', sans-serif;
          }
          header {
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          }
          .container-upload {
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            margin: 40px auto;
            max-width: 800px;
          }
          h2 {
            color: #283E51;
            font-weight: 600;
          }
          .btn-primary {
            background-color: #283E51;
            border: none;
          }
          .btn-primary:hover {
            background-color: #1d2d3a;
          }
          .btn-secondary {
            background-color: #4B79A1;
            border: none;
          }
          .btn-secondary:hover {
            background-color: #3a5d78;
          }
          .logout-container {
            text-align: center;
            margin-top: 20px;
          }
          .btn-logout {
            background-color: #dc3545;
            color: #fff;
            padding: 10px 20px;
            border: none;
            border-radius: 5px;
            font-size: 16px;
            text-decoration: none;
            transition: background-color 0.3s;
          }
          .btn-logout:hover {
            background-color: #c82333;
          }
          /* Botão Voltar ao Dashboard */
          .btn-voltar {
            display: inline-block;
            padding: 10px 20px;
            font-size: 16px;
            font-weight: 600;
            font-family: 'Montserrat', sans-serif;
            color: #fff;
            background-color: #4B79A1;
            border: none;
            border-radius: 5px;
            text-decoration: none;
            transition: background-color 0.3s;
            margin-top: 20px;
          }
          .btn-voltar:hover {
            background-color: #3a5d78;
          }
          footer {
            background-color: #424242;
            color: #fff;
            text-align: center;
            padding: 10px;
            position: fixed;
            bottom: 0;
            width: 100%;
          }
          #multi-upload-section {
            margin-top: 20px;
            border: 1px solid #ccc;
            padding: 20px;
            border-radius: 8px;
            background-color: #f9f9f9;
          }
          .multi-upload-group {
            margin-bottom: 15px;
          }
          #flash-messages {
            position: relative;
            top: 10px;
            left: 50%;
            transform: translateX(-50%);
            z-index: 10000;
          }
        </style>
      </head>
      <body>
        <header>
          <h1 class="mb-0">E.M José Padin Mouta - Carteirinhas</h1>
        </header>
        <div class="container container-upload">
          {% with messages = get_flashed_messages(with_categories=true) %}
            {% if messages %}
              <div id="flash-messages">
                {% for category, message in messages %}
                  <div class="alert alert-{{ 'success' if category == 'success' else 'info' }}" role="alert">{{ message }}</div>
                {% endfor %}
              </div>
            {% endif %}
          {% endwith %}
          <h2 class="mb-4">Envie a lista piloto (Excel)</h2>
          <form method="POST" enctype="multipart/form-data" onsubmit="showLoading()">
            <div class="form-group">
              <label for="excel_file">Selecione a Lista do Fundamental (opcional caso tenha anexado no inicio do sistema):</label>
              <input type="file" class="form-control-file" name="excel_file" id="excel_file" accept=".xlsx, .xls">
            </div>
            <button type="submit" class="btn btn-primary">Gerar Carteirinhas</button>
          </form>
          <hr>
          <h2 class="mb-4">Upload da Foto</h2>
          <form method="POST" action="/upload_foto" enctype="multipart/form-data">
            <div class="form-group">
              <label>RM do Aluno:</label>
              <input type="text" class="form-control" name="rm" placeholder="Digite o RM">
            </div>
            <div class="form-group">
              <input type="file" class="form-control-file" name="foto_file" accept="image/*">
            </div>
            <button type="submit" class="btn btn-secondary">Enviar Foto</button>
          </form>
          <hr>
          <h2 class="mb-4">Upload de Múltiplas Fotos</h2>
          <button type="button" class="btn btn-secondary" id="show-multi-upload">Enviar múltiplas fotos</button>
          <div id="multi-upload-section" style="display: none;">
            <form method="POST" action="/upload_multiplas_fotos" enctype="multipart/form-data" id="multi-upload-form">
              <div id="multi-upload-fields">
                <div class="multi-upload-group">
                  <div class="form-group">
                    <label>RM do Aluno:</label>
                    <input type="text" class="form-control" name="rm[]" placeholder="Digite o RM">
                  </div>
                  <div class="form-group">
                    <input type="file" class="form-control-file" name="foto_file[]" accept="image/*">
                  </div>
                </div>
              </div>
              <button type="button" class="btn btn-info" id="add-more" style="margin-top:10px;">Adicionar outra foto</button>
              <button type="submit" class="btn btn-primary" style="margin-top:10px;">Enviar Fotos</button>
            </form>
          </div>
          <!-- Botão Voltar ao Dashboard -->
          <div class="text-center">
            <a href="{{ url_for('dashboard') }}" class="btn-voltar">Voltar ao Dashboard</a>
          </div>
        </div>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
        <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
        <script>
          setTimeout(function(){
            var flashDiv = document.getElementById('flash-messages');
            if(flashDiv){
              flashDiv.style.display = 'none';
            }
          }, 3000);

          function showLoading() {
            var existingOverlay = document.getElementById('loading-overlay');
            if (existingOverlay) {
              existingOverlay.remove();
            }
            var loadingOverlay = document.createElement('div');
            loadingOverlay.id = 'loading-overlay';
            loadingOverlay.style.position = 'fixed';
            loadingOverlay.style.top = '0';
            loadingOverlay.style.left = '0';
            loadingOverlay.style.right = '0';
            loadingOverlay.style.bottom = '0';
            loadingOverlay.style.background = 'rgba(0,0,0,0.5)';
            loadingOverlay.style.display = 'flex';
            loadingOverlay.style.alignItems = 'center';
            loadingOverlay.style.justifyContent = 'center';
            loadingOverlay.style.zIndex = '9999';
            loadingOverlay.innerHTML = 
              `<div style="text-align: center; color: white; font-family: Arial, sans-serif;">
                <svg width="3.0cm" height="4.5cm" viewBox="0 0 6.0 9.0" xmlns="http://www.w3.org/2000/svg">
                  <rect x="0.3" y="0.3" width="5.4" height="8.4" rx="0.3" ry="0.3" stroke="white" stroke-width="0.1" fill="none" />
                  <rect id="badge-fill" x="0.3" y="8.7" width="5.4" height="0" rx="0.3" ry="0.3" fill="white" />
                </svg>
                <p id="loading-text" style="margin-top: 0.2cm;">Gerando carteirinhas...</p>
              </div>`;
            document.body.appendChild(loadingOverlay);
            let fillHeight = 0;
            const maxHeight = 8.4;
            const interval = setInterval(() => {
              fillHeight += 0.2;
              if (fillHeight > maxHeight) {
                fillHeight = maxHeight;
                clearInterval(interval);
              }
              const badgeFill = document.getElementById('badge-fill');
              badgeFill.setAttribute('y', 8.7 - fillHeight);
              badgeFill.setAttribute('height', fillHeight);
            }, 100);
          }

          document.getElementById('show-multi-upload').addEventListener('click', function() {
            var section = document.getElementById('multi-upload-section');
            if(section.style.display === 'none') {
              section.style.display = 'block';
            } else {
              section.style.display = 'none';
            }
          });
          document.getElementById('add-more').addEventListener('click', function() {
            var container = document.getElementById('multi-upload-fields');
            var group = document.createElement('div');
            group.className = 'multi-upload-group';
            group.innerHTML = 
              `<div class="form-group">
                <label>RM do Aluno:</label>
                <input type="text" class="form-control" name="rm[]" placeholder="Digite o RM">
              </div>
              <div class="form-group">
                <input type="file" class="form-control-file" name="foto_file[]" accept="image/*">
              </div>`;
            container.appendChild(group);
          });
        </script>
      </body>
    </html>
    '''
    return render_template_string(carteirinhas_html)


# Rota para Declaração Escolar para DECLARAÇÃO FUNDAMENTAL (uma lista piloto)
@app.route('/declaracao/upload', methods=['GET', 'POST'])
@login_required
def declaracao_upload():
    if request.method == 'POST':
        file = None
        if 'excel_file' in request.files and request.files['excel_file'].filename != '':
            file = request.files['excel_file']
            filename = secure_filename(file.filename)
            unique_filename = f"declaracao_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(file_path)
            session['lista_fundamental'] = file_path
        else:
            file_path = session.get('lista_fundamental')
            if file_path and os.path.exists(file_path):
                file = open(file_path, 'rb')

        if not file:
            flash("Nenhum arquivo enviado.", "error")
            return redirect(url_for('declaracao_upload'))

        session['declaracao_excel'] = file_path
        session['declaracao_tipo'] = "Fundamental"

        if hasattr(file, 'close'):
            file.close()

        return redirect(url_for('declaracao_select'))

    upload_form = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Declaração Escolar - Fundamental</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {
          background: #eef2f3;
          font-family: 'Montserrat', sans-serif;
        }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
          text-align: center;
        }
        .btn-primary {
          background-color: #283E51;
          border: none;
        }
        .btn-primary:hover {
          background-color: #1d2d3a;
        }
        footer {
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
        }
      </style>
    </head>
    <body>
      <header>
        <h1>Declaração Escolar - Fundamental</h1>
      </header>
      <div class="container container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="excel_file">Selecione a lista piloto do Fundamental:</label>
            <input type="file" class="form-control-file" name="excel_file" id="excel_file" accept=".xlsx, .xls">
          </div>
          <button type="submit" class="btn btn-primary">Anexar Lista do Fundamental</button>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_form)

# Nova rota para Declaração EJA – upload da lista EJA
@app.route('/declaracao/upload_eja', methods=['GET', 'POST'])
@login_required
def declaracao_upload_eja():
    if request.method == 'POST':
        file = None
        if 'excel_file' in request.files and request.files['excel_file'].filename != '':
            file = request.files['excel_file']
            filename = secure_filename(file.filename)
            unique_filename = f"declaracao2_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            file.save(file_path)
            session['lista_eja'] = file_path
        else:
            file_path = session.get('lista_eja')
            if file_path and os.path.exists(file_path):
                file = open(file_path, 'rb')

        if not file:
            flash("Nenhum arquivo enviado.", "error")
            return redirect(url_for('declaracao_upload_eja'))

        session['declaracao_excel'] = file_path  # Para EJA, usamos o mesmo nome de sessão
        session['declaracao_tipo'] = "EJA"

        if hasattr(file, 'close'):
            file.close()

        return redirect(url_for('declaracao_select'))

    upload_form = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Declaração Escolar - EJA</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {
          background: #eef2f3;
          font-family: 'Montserrat', sans-serif;
        }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
          text-align: center;
        }
        .btn-primary {
          background-color: #283E51;
          border: none;
        }
        .btn-primary:hover {
          background-color: #1d2d3a;
        }
        footer {
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
        }
      </style>
    </head>
    <body>
      <header>
        <h1>Declaração Escolar - EJA</h1>
      </header>
      <div class="container container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="excel_file">Selecione a lista piloto da EJA em Excel:</label>
            <input type="file" class="form-control-file" name="excel_file" id="excel_file" accept=".xlsx, .xls">
          </div>
          <button type="submit" class="btn btn-primary">Anexar Lista EJA</button>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_form)

@app.route('/declaracao/select', methods=['GET', 'POST'])
@login_required
def declaracao_select():
    file_path = session.get('declaracao_excel')
    if not file_path or not os.path.exists(file_path):
        flash("Arquivo Excel não encontrado. Por favor, anexe a lista piloto.", "error")
        if session.get('declaracao_tipo') == "EJA":
            return redirect(url_for('declaracao_upload_eja'))
        else:
            return redirect(url_for('declaracao_upload'))

    # Carrega dados dos alunos
    if session.get('declaracao_tipo') == "EJA":
        df = pd.read_excel(file_path, sheet_name=0, header=None, skiprows=1)
        df['RM_str'] = df.iloc[:, 2].apply(lambda x: str(int(x)) if pd.notna(x) and float(x) != 0 else "")
        df['NOME'] = df.iloc[:, 3]
        df['NASC.'] = df.iloc[:, 6]

        def get_ra(row):
            try:
                val = row.iloc[7]
                if pd.isna(val) or float(val) == 0:
                    return row.iloc[8]
                else:
                    return val
            except:
                return row.iloc[7]

        df['RA'] = df.apply(get_ra, axis=1)
        df['SÉRIE'] = df.iloc[:, 0]
        alunos = df[df['RM_str'] != ""][['RM_str', 'NOME']].drop_duplicates()
    else:
        planilha = pd.read_excel(file_path, sheet_name='LISTA CORRIDA')

        def format_rm(x):
            try:
                return str(int(float(x)))
            except:
                return str(x)

        planilha['RM_str'] = planilha['RM'].apply(format_rm)
        alunos = planilha[planilha['RM_str'] != "0"][['RM_str', 'NOME']].drop_duplicates()

    # Gera HTML das opções de alunos
    options_html = ""
    for _, row in alunos.iterrows():
        rm_str = row['RM_str']
        nome = row['NOME']
        options_html += f'<option value="{rm_str}">{rm_str} - {nome}</option>'

    if request.method == 'POST':
        rm = request.form.get('rm')
        tipo = request.form.get('tipo')
        deve_historico_str = request.form.get('deve_historico')
        # Seleção ou input manual
        unidade_select = request.form.get('unidade_anterior_select', '').strip()
        unidade_manual = request.form.get('unidade_anterior_manual', '').strip()
        unidade_anterior = unidade_select if unidade_select else unidade_manual

        if not rm or not tipo:
            flash("Escolha o aluno e o tipo de declaração.", "error")
            return redirect(url_for('declaracao_select'))

        # Validação para Transferência e Conclusão
        if tipo in ["Transferencia", "Conclusão"]:
            if deve_historico_str not in ['sim', 'nao']:
                flash("Por favor, responda se o aluno deve o histórico escolar.", "error")
                return redirect(url_for('declaracao_select'))
            if deve_historico_str == 'sim' and not unidade_anterior:
                flash("Por favor, informe a unidade escolar anterior.", "error")
                return redirect(url_for('declaracao_select'))
            deve_historico = (deve_historico_str == 'sim')
        else:
            deve_historico = False
            unidade_anterior = ""

        declaracao_html = gerar_declaracao_escolar(
            file_path, rm, tipo,
            deve_historico=deve_historico,
            unidade_anterior=unidade_anterior
        )

        if declaracao_html is None:
            flash("Aluno não encontrado.", "error")
            return redirect(url_for('declaracao_select'))

        return declaracao_html

    # HTML do formulário com Select2 e input manual
    select_form = f'''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Declaração Escolar - Seleção de Aluno</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/css/select2.min.css" rel="stylesheet" />
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body {{ background: #eef2f3; font-family: 'Montserrat', sans-serif; }}
        header {{
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff; padding: 20px; text-align: center;
          border-bottom: 3px solid #1d2d3a; border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }}
        .container-form {{ background: #fff; padding: 40px; border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15); margin: 40px auto; max-width: 600px; }}
        .btn-primary {{ background-color: #283E51; border: none; }}
        .btn-primary:hover {{ background-color: #1d2d3a; }}
        footer {{ background-color: #424242; color: #fff; text-align: center; padding: 10px;
          position: fixed; bottom: 0; width: 100%; }}
        #historico-container {{ display: none; margin-top: 15px; border: 1px solid #ccc;
          padding: 15px; border-radius: 8px; background: #f9f9f9; }}
        #unidade-anterior-container {{ display: none; margin-top: 15px; }}
        #btn-container {{ margin-top: 20px; }}
      </style>
    </head>
    <body>
      <header>
        <h1>E.M José Padin Mouta - Declaração Escolar</h1>
        <p>Escolha o aluno para realizar a declaração</p>
      </header>
      <div class="container container-form">
        <form method="POST" onsubmit="return validarFormulario();">
          <div class="form-group">
            <label for="rm">Aluno:</label>
            <select class="form-control" id="rm" name="rm" required>
              <option value="">Selecione</option>
              {options_html}
            </select>
          </div>

          <div class="form-group">
            <label for="tipo">Tipo de Declaração:</label>
            <select class="form-control" id="tipo" name="tipo" required>
              <option value="">Selecione</option>
              <option value="Escolaridade">Declaração de Escolaridade</option>
              <option value="Transferencia">Declaração de Transferência</option>
              <option value="Conclusão">Declaração de Conclusão</option>
            </select>
          </div>

          <div id="historico-container">
            <label>Deve Histórico Escolar? <span style="color:red;">*</span></label><br>
            <div class="form-check form-check-inline">
              <input class="form-check-input" type="radio" name="deve_historico" id="historico_sim" value="sim">
              <label class="form-check-label" for="historico_sim">Sim</label>
            </div>
            <div class="form-check form-check-inline">
              <input class="form-check-input" type="radio" name="deve_historico" id="historico_nao" value="nao">
              <label class="form-check-label" for="historico_nao">Não</label>
            </div>
          </div>

          <div id="unidade-anterior-container" class="form-group">
            <label for="unidade_anterior">Nome da unidade escolar anterior: <span style="color:red;">*</span></label>
            
            <!-- Select AJAX -->
            <select class="form-control" id="unidade_anterior" name="unidade_anterior_select"></select>
            <!-- Input manual -->
            <input type="text" id="unidade_anterior_manual" name="unidade_anterior_manual" class="form-control mt-2"
              placeholder="Ou digite o nome da unidade">
          </div>

          <div id="btn-container">
            <button type="submit" class="btn btn-primary">Gerar Declaração</button>
          </div>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>

      <script src="https://code.jquery.com/jquery-3.5.1.min.js"></script>
      <script src="https://cdnjs.cloudflare.com/ajax/libs/select2/4.0.13/js/select2.min.js"></script>
      <script>
      $(document).ready(function() {{
          // Select2 para alunos
          $('#rm').select2({{
              placeholder: "Selecione o aluno",
              allowClear: true,
              width: '100%'
          }});

          // Select2 AJAX para unidades
          $('#unidade_anterior').select2({{
              placeholder: "Selecione ou busque a escola",
              allowClear: true,
              width: '100%',
              ajax: {{
                  url: '/escolas/search',
                  dataType: 'json',
                  delay: 250,
                  data: function(params) {{
                      return {{ q: params.term }};
                  }},
                  processResults: function(data) {{
                      var results = data.map(function(item) {{
                          return {{ id: item.id, text: item.text }};
                      }});
                      return {{ results: results }};
                  }},
                  cache: true
              }},
              minimumInputLength: 1
          }});

          // Histórico e unidade anterior
          $('#tipo').on('change', function() {{
              if ($(this).val() === 'Transferencia' || $(this).val() === 'Conclusão') {{
                  $('#historico-container').show();
              }} else {{
                  $('#historico-container').hide();
                  $('#unidade-anterior-container').hide();
                  $('input[name="deve_historico"]').prop('checked', false);
                  $('#unidade_anterior').val(null).trigger('change');
                  $('#unidade_anterior_manual').val('');
              }}
          }});

          $('input[name="deve_historico"]').on('change', function() {{
              if ($(this).val() === 'sim') {{
                  $('#unidade-anterior-container').show();
              }} else {{
                  $('#unidade-anterior-container').hide();
                  $('#unidade_anterior').val(null).trigger('change');
                  $('#unidade_anterior_manual').val('');
              }}
          }});
      }});

      function validarFormulario() {{
          var tipo = document.getElementById('tipo').value;
          if (tipo === 'Transferencia' || tipo === 'Conclusão') {{
              var radios = document.getElementsByName('deve_historico');
              var marcado = false;
              for (var i = 0; i < radios.length; i++) {{
                  if (radios[i].checked) {{ marcado = true; break; }}
              }}
              if (!marcado) {{
                  alert('Por favor, responda se o aluno deve o histórico escolar.');
                  return false;
              }}
              if (document.getElementById('historico_sim').checked) {{
                  var unidade_select = $('#unidade_anterior').val();
                  var unidade_manual = $('#unidade_anterior_manual').val().trim();
                  if (!unidade_select && !unidade_manual) {{
                      alert('Por favor, informe a unidade escolar anterior.');
                      return false;
                  }}
              }}
              return confirm("Você está gerando uma declaração de transferência ou conclusão, essa é a declaração correta a ser gerada?");
          }}
          return true;
      }}
      </script>
    </body>
    </html>
    '''
    return render_template_string(select_form)


@app.route('/declaracao/tipo', methods=['GET', 'POST'])
@login_required
def declaracao_tipo():
    if request.method == 'POST':
        tipo = request.form.get('tipo')
        if tipo == 'Fundamental':
            return redirect(url_for('declaracao_upload'))
        elif tipo == 'EJA':
            return redirect(url_for('declaracao_upload_eja'))

    form_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
         <meta charset="utf-8">
         <title>E.M José Padin Mouta - Declaração Escolar</title>
         <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
         <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
         <style>
         body {
             background: #eef2f3;
             font-family: 'Montserrat', sans-serif;
         }
         header {
             background: linear-gradient(90deg, #283E51, #4B79A1);
             color: #fff;
             padding: 20px;
             text-align: center;
             border-bottom: 3px solid #1d2d3a;
         }
         .container-form {
             background: #fff;
             padding: 40px;
             border-radius: 10px;
             box-shadow: 0 4px 12px rgba(0,0,0,0.15);
             margin: 40px auto;
             max-width: 600px;
         }
         .btn-primary {
             background-color: #283E51;
             border: none;
         }
         .btn-primary:hover {
             background-color: #1d2d3a;
         }
         /* Estilo para o botão Voltar ao Dashboard */
         .btn-voltar {
             display: inline-block;
             padding: 10px 20px;
             font-size: 16px;
             font-weight: 600;
             font-family: 'Montserrat', sans-serif;
             color: #fff;
             background-color: #4B79A1;
             border: none;
             border-radius: 5px;
             text-decoration: none;
             transition: background-color 0.3s;
             margin-top: 20px;
         }
         .btn-voltar:hover {
             background-color: #3a5d78;
         }
         footer {
             background-color: #424242;
             color: #fff;
             text-align: center;
             padding: 10px;
             position: fixed;
             bottom: 0;
             width: 100%;
         }
         </style>
    </head>
    <body>
         <header>
             <h1>E.M José Padin Mouta - Declaração Escolar</h1>
         </header>
         <div class="container-form">
             <form method="POST">
                 <div class="form-group">
                     <label for="tipo">Selecione o tipo de declaração:</label>
                     <select class="form-control" id="tipo" name="tipo" required>
                         <option value="">Selecione</option>
                         <option value="Fundamental">Declaração Fundamental</option>
                         <option value="EJA">Declaração EJA</option>
                     </select>
                 </div>
                 <button type="submit" class="btn btn-primary">Continuar</button>
             </form>
             <!-- Botão Voltar ao Dashboard -->
             <div class="text-center">
               <a href="{{ url_for('dashboard') }}" class="btn-voltar">Voltar ao Dashboard</a>
             </div>
         </div>
         <footer>
             Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
         </footer>
    </body>
    </html>
    '''
    return render_template_string(form_html)


@app.route('/upload_foto', methods=['POST'])
def upload_foto():
    if 'foto_file' not in request.files:
        return "Nenhum arquivo de foto enviado", 400

    rm = request.form.get('rm')
    if not rm:
        return "RM não fornecido", 400

    file = request.files['foto_file']
    if file.filename == '':
        return "Nenhuma foto selecionada", 400

    if not allowed_file(file.filename):
        return "Formato de imagem não permitido", 400
    original_filename = secure_filename(file.filename)
    _, ext = os.path.splitext(original_filename)
    new_filename = secure_filename(f"{rm}{ext.lower()}")
    file_path = os.path.join('static', 'fotos', new_filename)
    file.save(file_path)

    flash("Foto anexada com sucesso", "success")
    return redirect(url_for('carteirinhas'))


@app.route('/upload_multiplas_fotos', methods=['POST'])
def upload_multiplas_fotos():
    rms = request.form.getlist("rm[]")
    files = request.files.getlist("foto_file[]")
    if not files:
        return "Nenhuma foto enviada", 400
    for rm, file in zip(rms, files):
        if file.filename == '' or not rm or not allowed_file(file.filename):
            continue

        original_filename = secure_filename(file.filename)
        _, ext = os.path.splitext(original_filename)
        new_filename = secure_filename(f"{rm}{ext.lower()}")
        file_path = os.path.join('static', 'fotos', new_filename)
        file.save(file_path)

    flash("Foto(s) anexada(s) com sucesso", "success")
    return redirect(url_for('carteirinhas'))


@app.route('/upload_inline_foto', methods=['POST'])
def upload_inline_foto():
    if 'foto_file' not in request.files:
        return jsonify({'error': 'Nenhum arquivo enviado'}), 400

    rm = request.form.get('rm')
    if not rm:
        return jsonify({'error': 'RM não fornecido'}), 400

    file = request.files['foto_file']
    if file.filename == '' or not allowed_file(file.filename):
        return jsonify({'error': 'Formato de imagem não permitido'}), 400

    original_filename = secure_filename(file.filename)
    _, ext = os.path.splitext(original_filename)
    new_filename = secure_filename(f"{rm}{ext.lower()}")
    file_path = os.path.join('static', 'fotos', new_filename)
    file.save(file_path)

    return jsonify({'url': f"/static/fotos/{new_filename}", 'message': "Foto anexada com sucesso"})


@app.route('/quadros')
@login_required
def quadros():
    quadros_html = '''
    <!doctype html>
    <html lang="pt-br">
      <head>
        <meta charset="utf-8">
        <title>E.M José Padin Mouta - Quadros</title>
        <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
        <!-- Bootstrap, Font Awesome e Google Fonts -->
        <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
        <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
        <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
        <style>
          body {
            background: #eef2f3;
            font-family: 'Montserrat', sans-serif;
            margin-bottom: 60px; /* Espaço para o rodapé fixo */
          }
          header {
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
          }
          .container-menu {
            margin: 40px auto;
            max-width: 900px;
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          }
          .option-row {
            display: block;
          }
          .option-card {
            border: 1px solid #e0e0e0;
            border-radius: 10px;
            padding: 20px;
            background: #fff;
            margin-bottom: 20px;
            transition: transform 0.2s, box-shadow 0.2s;
            cursor: pointer;
            display: flex;
            align-items: center;
          }
          .option-card:hover {
            transform: scale(1.02);
            box-shadow: 0 8px 16px rgba(0,0,0,0.1);
          }
          .option-icon {
            width: 60px;
            text-align: center;
            margin-right: 20px;
          }
          .option-icon i {
            font-size: 2rem;
            color: #283E51;
          }
          .option-content {
            flex: 1;
          }
          .option-content h2 {
            margin: 0 0 10px 0;
            font-size: 1.25rem;
            color: #283E51;
          }
          .option-content p {
            margin: 0;
            font-size: 1rem;
            color: #555;
          }
          .btn-voltar {
            display: inline-block;
            padding: 10px 20px;
            font-size: 16px;
            font-weight: 600;
            color: #fff;
            background-color: #4B79A1;
            border: none;
            border-radius: 5px;
            text-decoration: none;
            transition: background-color 0.3s;
          }
          .btn-voltar:hover {
            background-color: #3a5d78;
          }
          footer {
            background-color: #424242;
            color: #fff;
            text-align: center;
            padding: 10px;
            position: fixed;
            bottom: 0;
            width: 100%;
          }
        </style>
      </head>
      <body>
        <header>
          <h1>E.M José Padin Mouta - Quadros</h1>
        </header>
        <div class="container-menu">
          <div class="option-row">
            <div class="option-card" onclick="window.location.href='{{ url_for('quadros_inclusao') }}'">
              <div class="option-icon">
                <i class="fas fa-user-plus"></i>
              </div>
              <div class="option-content">
                <h2>Inclusão</h2>
                <p>Gerar quadro de inclusão.</p>
              </div>
            </div>
            <!-- Nova opção adicionada -->
            <div class="option-card" onclick="window.location.href='{{ url_for('quantinclusao') }}'">
              <div class="option-icon">
                <i class="fas fa-chart-pie"></i>
              </div>
              <div class="option-content">
                <h2>Quantitativo de Inclusão</h2>
                <p>Gerar quadro quantitativo de inclusão.</p>
              </div>
            </div>
            <div class="option-card" onclick="window.location.href='{{ url_for('quadro_atendimento_mensal') }}'">
              <div class="option-icon">
                <i class="fas fa-calendar-alt"></i>
              </div>
              <div class="option-content">
                <h2>Atendimento Mensal</h2>
                <p>Gerar quadro de atendimento mensal.</p>
              </div>
            </div>
            <div class="option-card" onclick="window.location.href='{{ url_for('quadro_transferencias') }}'">
              <div class="option-icon">
                <i class="fas fa-exchange-alt"></i>
              </div>
              <div class="option-content">
                <h2>Transferências</h2>
                <p>Gerar quadro de transferências.</p>
              </div>
            </div>
            <div class="option-card" onclick="window.location.href='{{ url_for('quadro_quantitativo_mensal') }}'">
              <div class="option-icon">
                <i class="fas fa-chart-bar"></i>
              </div>
              <div class="option-content">
                <h2>Quantitativo Mensal</h2>
                <p>Gerar quadro quantitativo mensal.</p>
              </div>
            </div>
          </div>
          <div class="text-center mt-4">
            <a href="{{ url_for('dashboard') }}" class="btn-voltar">Voltar ao Dashboard</a>
          </div>
        </div>
        <footer>
          Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
        </footer>
        <!-- Scripts do Bootstrap -->
        <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
      </body>
    </html>
    '''
    return render_template_string(quadros_html)
# Rota para Quadro de Inclusão (com upload opcional para duas listas)
@app.route('/quadros/inclusao', methods=['GET', 'POST'])
@login_required
def quadros_inclusao():
    if request.method == 'POST':
        # Atualiza as listas na sessão (Regular e EJA)
        fundamental_file = request.files.get('lista_fundamental')
        eja_file = request.files.get('lista_eja')

        if fundamental_file and fundamental_file.filename != '':
            fundamental_filename = secure_filename(f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename)
            fundamental_path = os.path.join(app.config['UPLOAD_FOLDER'], fundamental_filename)
            fundamental_file.save(fundamental_path)
            session['lista_fundamental'] = fundamental_path

        if eja_file and eja_file.filename != '':
            eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_" + eja_file.filename)
            eja_path = os.path.join(app.config['UPLOAD_FOLDER'], eja_filename)
            eja_file.save(eja_path)
            session['lista_eja'] = eja_path

        # Carrega as listas piloto
        df_fundamental = None
        df_eja = None

        if session.get('lista_fundamental'):
            try:
                with open(session['lista_fundamental'], 'rb') as f_fund:
                    df_fundamental = pd.read_excel(f_fund, sheet_name="LISTA CORRIDA")
            except Exception:
                flash("Erro ao ler a Lista Piloto Fundamental.", "error")
                return redirect(url_for('quadros_inclusao'))

        if session.get('lista_eja'):
            try:
                with open(session['lista_eja'], 'rb') as f_eja:
                    df_eja = pd.read_excel(f_eja, sheet_name="LISTA CORRIDA")
            except Exception:
                flash("Erro ao ler a Lista Piloto EJA.", "error")
                return redirect(url_for('quadros_inclusao'))

        if df_fundamental is None and df_eja is None:
            flash("Nenhuma lista piloto disponível.", "error")
            return redirect(url_for('quadros_inclusao'))

        # Abre o modelo
        model_path = os.path.join("modelos", "Quadro de Alunos com Deficiência - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo de Inclusão não encontrado.", "error")
            return redirect(url_for('quadros_inclusao'))

        try:
            with open(model_path, "rb") as f:
                wb = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo de inclusão: {str(e)}", "error")
            return redirect(url_for('quadros_inclusao'))

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
        # Verifica se há pelo menos 25 colunas (coluna Y é a 25ª)
        if df_fundamental is not None:
            if len(df_fundamental.columns) < 25:
                flash("O arquivo da Lista Piloto Fundamental não possui colunas suficientes.", "error")
                return redirect(url_for('quadros_inclusao'))

            inclusion_col_fund = df_fundamental.columns[13]
            for idx, row in df_fundamental.iterrows():
                if not str(row['RM']).strip() or str(row['RM']).strip() == "0":
                    continue

                if str(row[inclusion_col_fund]).strip().lower() == "sim":
                    # Obtém o valor da coluna X (índice 23)
                    valor_coluna_x = row[df_fundamental.columns[23]]
                    
                    # Processa os demais dados conforme o processamento original
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
                            data_nasc = pd.to_datetime(data_nasc, errors='coerce')
                            if pd.notna(data_nasc):
                                data_nasc = data_nasc.strftime('%d/%m/%Y')
                            else:
                                data_nasc = "Desconhecida"
                        except:
                            data_nasc = "Desconhecida"
                    else:
                        data_nasc = "Desconhecida"
                    
                    professor = str(row[df_fundamental.columns[14]]).strip()
                    plano = str(row[df_fundamental.columns[15]]).strip()
                    aee = str(row[df_fundamental.columns[16]]).strip() if len(df_fundamental.columns) > 16 else ""
                    deficiencia = str(row[df_fundamental.columns[17]]).strip() if len(df_fundamental.columns) > 17 else ""
                    observacoes = str(row[df_fundamental.columns[18]]).strip() if len(df_fundamental.columns) > 18 else ""
                    cadeira = str(row[df_fundamental.columns[19]]).strip() if len(df_fundamental.columns) > 19 else ""
                    
                    # Coluna N: recebe o valor da coluna U (índice 20)
                    valor_coluna_n = row[df_fundamental.columns[20]]
                    # Coluna O: recebe o valor da coluna Y (índice 24)
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

        # Processa alunos da Lista Piloto EJA com novo mapeamento
        # Verifica se há pelo menos 29 colunas (coluna AC é a 29ª)
        if df_eja is not None:
            if len(df_eja.columns) < 29:
                flash("O arquivo da Lista Piloto EJA não possui colunas suficientes.", "error")
                return redirect(url_for('quadros_inclusao'))

            inclusion_col_eja = df_eja.columns[17]
            for idx, row in df_eja.iterrows():
                if not str(row['RM']).strip() or str(row['RM']).strip() == "0":
                    continue

                if str(row[inclusion_col_eja]).strip().lower() == "sim":
                    # Obtém o valor da coluna AB (índice 27) para "nível e ano"
                    valor_coluna_ab = row[df_eja.columns[27]]
                    
                    # Processa os demais dados conforme o mapeamento atual
                    turma = "A"
                    periodo = "NOITE"
                    horario = str(row[df_eja.columns[15]]).strip()
                    nome_aluno = str(row[df_eja.columns[3]]).strip()
                    data_nasc = row[df_eja.columns[6]]
                    if pd.notna(data_nasc):
                        try:
                            data_nasc = pd.to_datetime(data_nasc, errors='coerce')
                            if pd.notna(data_nasc):
                                data_nasc = data_nasc.strftime('%d/%m/%Y')
                            else:
                                data_nasc = "Desconhecida"
                        except:
                            data_nasc = "Desconhecida"
                    else:
                        data_nasc = "Desconhecida"
                    
                    professor = str(row[df_eja.columns[18]]).strip()
                    plano = str(row[df_eja.columns[19]]).strip()
                    aee = str(row[df_eja.columns[20]]).strip() if len(df_eja.columns) > 20 else ""
                    deficiencia = str(row[df_eja.columns[21]]).strip() if len(df_eja.columns) > 21 else ""
                    observacoes = str(row[df_eja.columns[22]]).strip() if len(df_eja.columns) > 22 else ""
                    # Aqui, a coluna M do modelo receberá o valor da coluna X (índice 23) da lista piloto EJA
                    cadeira = row[df_eja.columns[23]]
                    
                    # Coluna N: recebe o valor da coluna Y (índice 24)
                    valor_coluna_n = row[df_eja.columns[24]]
                    # Coluna O: recebe o valor da coluna AC (índice 28)
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
        meses = {1: "janeiro", 2: "fevereiro", 3: "março", 4: "abril", 5: "maio", 6: "junho",
                 7: "julho", 8: "agosto", 9: "setembro", 10: "outubro", 11: "novembro", 12: "dezembro"}
        mes = meses[datetime.now().month].capitalize()
        filename = f"Quadro de Inclusão - {mes} - E.M José Padin Mouta.xlsx"
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    upload_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>E.M José Padin Mouta - Quadro de Inclusão</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form { background: #fff; padding: 40px; border-radius: 10px; box-shadow: 0 4px 12px rgba(0,0,0,0.15); margin: 40px auto; max-width: 600px; }
        .btn-primary { background-color: #283E51; border: none; }
        .btn-primary:hover { background-color: #1d2d3a; }
        footer { background-color: #424242; color: #fff; text-align: center; padding: 10px; position: fixed; bottom: 0; width: 100%; }
      </style>
    </head>
    <body>
      <header>
        <h1>E.M José Padin Mouta - Quadro de Inclusão</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="lista_fundamental">Selecione a Lista Piloto - FUNDAMENTAL (Excel):</label>
            <input type="file" class="form-control-file" name="lista_fundamental" id="lista_fundamental" accept=".xlsx, .xls">
          </div>
          <div class="form-group">
            <label for="lista_eja">Selecione a Lista Piloto - EJA (Excel):</label>
            <input type="file" class="form-control-file" name="lista_eja" id="lista_eja" accept=".xlsx, .xls">
          </div>
          <button type="submit" class="btn btn-primary">Gerar Quadro de Inclusão</button>
        </form>
        <br>
        <a href="{{ url_for('quadros') }}">Voltar para Quadros</a>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_html)


# Rota para Quadro de Atendimento Mensal – com upload opcional para duas listas
@app.route('/quadros/atendimento_mensal', methods=['GET', 'POST'])
@login_required
def quadro_atendimento_mensal():
    if request.method == 'POST':
        fundamental_file = request.files.get('lista_fundamental')
        eja_file = request.files.get('lista_eja')

        if fundamental_file and fundamental_file.filename != '':
            filename = secure_filename(fundamental_file.filename)
            unique_filename = f"atendimento_{uuid.uuid4().hex}_{filename}"
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            fundamental_file.save(file_path)
            session['lista_fundamental'] = file_path

        if eja_file and eja_file.filename != '':
            filename = secure_filename(eja_file.filename)
            unique_filename = f"atendimento_eja_{uuid.uuid4().hex}_{filename}"
            file_path_eja = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)
            eja_file.save(file_path_eja)
            session['lista_eja'] = file_path_eja

        file_path = session.get('lista_fundamental')
        if file_path and os.path.exists(file_path):
            lista_file = open(file_path, 'rb')
        else:
            lista_file = None

        if not lista_file:
            flash("Nenhum arquivo enviado.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

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

        # Se o modelo tiver mais de uma planilha, pega a segunda (ou a primeira, se só tiver uma)
        if len(wb_modelo.worksheets) > 1:
            ws_modelo = wb_modelo.worksheets[1]
        else:
            ws_modelo = wb_modelo.active

        set_merged_cell_value(ws_modelo, "B5", "E.M José Padin Mouta")
        set_merged_cell_value(ws_modelo, "C6", "Rafael Fernando da Silva")
        set_merged_cell_value(ws_modelo, "B7", "46034")
        current_month = datetime.now().strftime("%m")
        set_merged_cell_value(ws_modelo, "A13", f"{current_month}/2025")

        try:
            lista_file.seek(0)
            wb_lista = load_workbook(lista_file, data_only=True)
        except Exception:
            flash("Erro ao ler o arquivo da lista piloto.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        sheet_name = None
        for name in wb_lista.sheetnames:
            if name.strip().lower() == "total de alunos":
                sheet_name = name
                break

        if not sheet_name:
            flash("A aba 'Total de Alunos' não foi encontrada na lista piloto.", "error")
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

        # Preenchimento dos campos específicos
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

        # ---- NOVA ALTERAÇÃO PARA EJA ----
        eja_path = session.get('lista_eja')
        if not eja_path or not os.path.exists(eja_path):
            flash("Arquivo da Lista Piloto EJA não encontrado.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        with open(eja_path, 'rb') as f_eja:
            wb_eja = load_workbook(f_eja, data_only=True)

        sheet_name_eja = None
        for name in wb_eja.sheetnames:
            if name.strip().lower() == "total de alunos":
                sheet_name_eja = name
                break

        if not sheet_name_eja:
            flash("A aba 'Total de Alunos' não foi encontrada na Lista Piloto EJA.", "error")
            return redirect(url_for('quadro_atendimento_mensal'))

        ws_total_eja = wb_eja[sheet_name_eja]

        set_merged_cell_value(ws_modelo, "L19", ws_total_eja.cell(row=6, column=5).value)   # E6
        set_merged_cell_value(ws_modelo, "L20", ws_total_eja.cell(row=7, column=5).value)   # E7
        set_merged_cell_value(ws_modelo, "L21", ws_total_eja.cell(row=8, column=5).value)   # E8
        set_merged_cell_value(ws_modelo, "L22", ws_total_eja.cell(row=9, column=5).value)   # E9

        set_merged_cell_value(ws_modelo, "M19", ws_total_eja.cell(row=6, column=6).value)   # F6
        set_merged_cell_value(ws_modelo, "M20", ws_total_eja.cell(row=7, column=6).value)   # F7
        set_merged_cell_value(ws_modelo, "M21", ws_total_eja.cell(row=8, column=6).value)   # F8
        set_merged_cell_value(ws_modelo, "M22", ws_total_eja.cell(row=9, column=6).value)   # F9

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
        # ---- FIM NOVA ALTERAÇÃO ----

        output = BytesIO()
        wb_modelo.save(output)
        output.seek(0)
        filename = f"Quadro de Atendimento Mensal - {datetime.now().strftime('%d%m')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    upload_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Quadro de Atendimento Mensal - E.M José Padin Mouta</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
        }
        .btn-primary {
          background-color: #283E51;
          border: none;
        }
        .btn-primary:hover {
          background-color: #1d2d3a;
        }
        footer {
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
        }
      </style>
    </head>
    <body>
      <header>
        <h1>Quadro de Atendimento Mensal</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="lista_fundamental">Selecione a Lista Piloto - FUNDAMENTAL (Excel):</label>
            <input type="file" class="form-control-file" name="lista_fundamental" id="lista_fundamental" accept=".xlsx, .xls">
          </div>
          <div class="form-group">
            <label for="lista_eja">Selecione a Lista Piloto - EJA (Excel):</label>
            <input type="file" class="form-control-file" name="lista_eja" id="lista_eja" accept=".xlsx, .xls">
          </div>
          <button type="submit" class="btn btn-primary">Gerar Quadro de Atendimento Mensal</button>
        </form>
        <br>
        <a href="{{ url_for('quadros') }}">Voltar para Quadros</a>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(upload_html)

# Rota para o Quadro de Transferências (AGORA TAMBÉM LENDO EJA, COL K, E MAPEANDO CAMPOS DIFERENTES)
@app.route('/quadros/transferencias', methods=['GET', 'POST'])
@login_required
def quadro_transferencias():
    if request.method == 'POST':
        # Obtém dados do formulário
        period_start_str = request.form.get('period_start')
        period_end_str = request.form.get('period_end')
        responsavel = request.form.get('responsavel')

        fundamental_file = request.files.get('lista_fundamental')
        eja_file = request.files.get('lista_eja')

        if not period_start_str or not period_end_str or not responsavel:
            flash("Por favor, preencha todos os campos.", "error")
            return redirect(url_for('quadro_transferencias'))

        # Salva/atualiza a Lista Piloto Fundamental, se enviada
        if fundamental_file and fundamental_file.filename != '':
            fundamental_filename = secure_filename(f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename)
            fundamental_path = os.path.join(app.config['UPLOAD_FOLDER'], fundamental_filename)
            fundamental_file.save(fundamental_path)
            session['lista_fundamental'] = fundamental_path
        else:
            fundamental_path = session.get('lista_fundamental')
            if not fundamental_path or not os.path.exists(fundamental_path):
                flash("Lista Piloto Fundamental não encontrada.", "error")
                return redirect(url_for('quadro_transferencias'))

        # Salva/atualiza a Lista Piloto EJA, se enviada
        if eja_file and eja_file.filename != '':
            eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_" + eja_file.filename)
            eja_path = os.path.join(app.config['UPLOAD_FOLDER'], eja_filename)
            eja_file.save(eja_path)
            session['lista_eja'] = eja_path
        else:
            eja_path = session.get('lista_eja')
            # Se não existir, não dá erro (caso o user não vá usar a EJA).

        try:
            period_start = datetime.strptime(period_start_str, "%Y-%m-%d")
            period_end = datetime.strptime(period_end_str, "%Y-%m-%d")
        except Exception:
            flash("Formato de data inválido.", "error")
            return redirect(url_for('quadro_transferencias'))

        # ---- PARTE 1: Lê a lista piloto FUNDAMENTAL e extrai TE XX/XX dentro do range
        try:
            df_fundamental = pd.read_excel(fundamental_path, sheet_name="LISTA CORRIDA")
        except Exception as e:
            flash(f"Erro ao ler a Lista Piloto Fundamental: {str(e)}", "error")
            return redirect(url_for('quadro_transferencias'))

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
            "País": "Mudança de País"
        }

        transfer_records = []
        col_V_index = 21  # suposto índice da coluna V (0-based)

        for idx, row in df_fundamental.iterrows():
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
                except:
                    continue

                if period_start <= te_date <= period_end:
                    nome = str(row.iloc[3])
                    dn_val = row.iloc[5]
                    dn_str = ""
                    if pd.notna(dn_val):
                        try:
                            dn_dt = pd.to_datetime(dn_val, errors='coerce')
                            if pd.notna(dn_dt):
                                dn_str = dn_dt.strftime('%d/%m/%y')
                            else:
                                dn_str = ""
                        except:
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

                    # concatena com valor de W, se houver
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
                        "data": data_te
                    }
                    transfer_records.append(record)

        # PARTE 2: Processamento da Lista Piloto EJA (código corrigido com parênteses)
        if eja_path and os.path.exists(eja_path):
            try:
                # Lê a aba "LISTA CORRIDA" da lista piloto EJA
                df_eja = pd.read_excel(eja_path, sheet_name="LISTA CORRIDA")
            except Exception as e:
                flash(f"Erro ao ler a Lista Piloto EJA: {str(e)}", "error")
                return redirect(url_for('quadro_transferencias'))

            # Percorre cada linha da EJA procurando na coluna K registros com o padrão TE, MC ou MCC
            for idx, row in df_eja.iterrows():
                if len(row) < 11:
                    continue

                # Coluna K (índice 10)
                col_k_value = str(row.iloc[10]).strip() if len(row) > 10 else ""
                if not col_k_value:
                    continue

                # Usa re.search para encontrar o padrão em qualquer parte do texto
                match_eja = re.search(r"(TE|MC|MCC)\s*(\d{1,2}/\d{1,2})", col_k_value, re.IGNORECASE)
                if match_eja:
                    tipo_str = match_eja.group(1).upper()  # Tipo: TE, MC ou MCC
                    date_str = match_eja.group(2)           # Data: dd/mm
                    eja_date_full = f"{date_str}/{period_start.year}"
                    try:
                        eja_date = datetime.strptime(eja_date_full, "%d/%m/%Y")
                    except:
                        continue

                    # Se a data estiver dentro do intervalo informado
                    if period_start <= eja_date <= period_end:
                        # Mapeamento dos campos conforme solicitado:
                        # - Nome: coluna D (índice 3)
                        nome = str(row.iloc[3])
                        # - D.N.: coluna G (índice 6), formatando a data se possível
                        dn_val = row.iloc[6]
                        dn_str = ""
                        if pd.notna(dn_val):
                            try:
                                dn_dt = pd.to_datetime(dn_val, errors='coerce')
                                if pd.notna(dn_dt):
                                    dn_str = dn_dt.strftime('%d/%m/%Y')
                            except:
                                dn_str = ""
                        # - R.A.: coluna H (índice 7); se for 0 ou estiver vazio, buscar na coluna I (índice 8)
                        ra_val = row.iloc[7]
                        if pd.isna(ra_val) or (isinstance(ra_val, (int, float)) and float(ra_val) == 0):
                            ra_val = row.iloc[8]
                        # Campos fixos
                        situacao = "Parcial"
                        breda = "Não"
                        # - Nível / Classe e Turma: coluna A (índice 0)
                        nivel_classe = str(row.iloc[0])
                        # - Tipo: a palavra capturada (TE, MC ou MCC)
                        tipo_field = tipo_str
                        # - Motivo/Observação:
                        if tipo_field in ["MC", "MCC"]:
                            obs_final = "Desistencia"
                        else:
                            # Para TE, concatena o conteúdo das colunas Z (índice 25) e AA (índice 26)
                            part_z = str(row.iloc[25]).strip() if len(row) > 25 else ""
                            part_aa = str(row.iloc[26]).strip() if len(row) > 26 else ""
                            if part_aa:
                                obs_final = f"{part_z} ({part_aa})".strip()
                            else:
                                obs_final = part_z
                        # - Remanejamento: sempre "-"
                        remanejamento = "-"
                        # - Data: a data extraída, formatada em dd/mm/YYYY
                        data_te = eja_date.strftime("%d/%m/%Y")

                        # Cria o registro para o Quadro Informativo
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
                            "data": data_te
                        }
                        transfer_records.append(record)

        if not transfer_records:
            flash("Nenhum registro de TE/MC/MCC encontrado no período especificado.", "error")
            return redirect(url_for('quadro_transferencias'))

        model_path = os.path.join("modelos", "Quadro Informativo - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo de Quadro Informativo (Transferências) não encontrado.", "error")
            return redirect(url_for('quadro_transferencias'))

        try:
            with open(model_path, "rb") as f:
                wb = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo: {str(e)}", "error")
            return redirect(url_for('quadro_transferencias'))

        ws = wb.active

        set_merged_cell_value(ws, "B9", responsavel)
        set_merged_cell_value(ws, "J9", datetime.now().strftime("%d/%m/%Y"))

        start_row = 12
        current_row = start_row

        # Preenche cada linha do Quadro Informativo
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

        filename = f"Quadro_de_Transferencias_{period_start.strftime('%d%m')}_{period_end.strftime('%d%m')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    form_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>E.M José Padin Mouta - Quadro de Transferências</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
        }
        .btn-primary {
          background-color: #283E51;
          border: none;
        }
        .btn-primary:hover {
          background-color: #1d2d3a;
        }
        footer {
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
        }
      </style>
    </head>
    <body>
      <header>
        <h1>E.M José Padin Mouta - Quadro de Transferências</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="period_start">Data Inicial da Semana:</label>
            <input type="date" class="form-control" name="period_start" id="period_start" required>
          </div>
          <div class="form-group">
            <label for="period_end">Data Final da Semana:</label>
            <input type="date" class="form-control" name="period_end" id="period_end" required>
          </div>
          <div class="form-group">
            <label for="responsavel">Responsável pelas Informações:</label>
            <input type="text" class="form-control" name="responsavel" id="responsavel" required>
          </div>
          <div class="form-group">
            <label for="lista_fundamental">Selecione a Lista Piloto - FUNDAMENTAL (Excel):</label>
            <input type="file" class="form-control-file" name="lista_fundamental" id="lista_fundamental" accept=".xlsx, .xls">
          </div>
          <div class="form-group">
            <label for="lista_eja">Selecione a Lista Piloto - EJA (Excel):</label>
            <input type="file" class="form-control-file" name="lista_eja" id="lista_eja" accept=".xlsx, .xls">
          </div>
          <button type="submit" class="btn btn-primary">Gerar Quadro de Transferências</button>
        </form>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(form_html)


@app.route('/quadros/quantitativo_mensal', methods=['GET', 'POST'])
@login_required
def quadro_quantitativo_mensal():
    if request.method == 'POST':
        # Recupera os dados do formulário
        period_start_str = request.form.get('period_start')
        period_end_str = request.form.get('period_end')
        responsavel = request.form.get('responsavel')
        
        if not period_start_str or not period_end_str or not responsavel:
            flash("Preencha todos os campos obrigatórios.", "error")
            return redirect(url_for('quadro_quantitativo_mensal'))
        
        try:
            period_start = datetime.strptime(period_start_str, "%Y-%m-%d")
            period_end = datetime.strptime(period_end_str, "%Y-%m-%d")
        except Exception:
            flash("Formato de data inválido.", "error")
            return redirect(url_for('quadro_quantitativo_mensal'))
        
        # Trata o upload do arquivo da Lista Piloto Fundamental
        fundamental_file = request.files.get('lista_fundamental')
        if fundamental_file and fundamental_file.filename != '':
            fundamental_filename = secure_filename(f"fundamental_{uuid.uuid4().hex}_" + fundamental_file.filename)
            fundamental_path = os.path.join(app.config['UPLOAD_FOLDER'], fundamental_filename)
            fundamental_file.save(fundamental_path)
        else:
            fundamental_path = session.get('lista_fundamental')
            if not fundamental_path or not os.path.exists(fundamental_path):
                flash("Arquivo da Lista Piloto Fundamental não encontrado.", "error")
                return redirect(url_for('quadro_quantitativo_mensal'))
        
        try:
            # Carrega a aba "LISTA CORRIDA" da Lista Piloto Fundamental
            df_fundamental = pd.read_excel(fundamental_path, sheet_name="LISTA CORRIDA")
        except Exception as e:
            flash(f"Erro ao ler a Lista Piloto Fundamental: {str(e)}", "error")
            return redirect(url_for('quadro_quantitativo_mensal'))
        
        # Carrega o modelo de Quadro Quantitativo Mensal
        model_path = os.path.join("modelos", "Quadro Quantitativo Mensal - Modelo.xlsx")
        if not os.path.exists(model_path):
            flash("Modelo de Quadro Quantitativo Mensal não encontrado.", "error")
            return redirect(url_for('quadro_quantitativo_mensal'))
        
        try:
            with open(model_path, "rb") as f:
                wb = load_workbook(f, data_only=False)
        except Exception as e:
            flash(f"Erro ao ler o modelo: {str(e)}", "error")
            return redirect(url_for('quadro_quantitativo_mensal'))
        
        ws = wb.active
        
        # Mapeamento: para cada série (coluna A) e TIPO TE (coluna V) define a célula do modelo a ser incrementada
        mapping = {
            '2º': {
                "Dentro da Rede": "K14",
                "Rede Estadual": "K15",
                "Litoral": "K16",
                "São Paulo": "K17",
                "ABCD": "K18",
                "Interior": "K19",
                "Outros Estados": "K20",
                "Particular": "K21",
                "País": "K22",
                "Sem Informação": "K23"
            },
            '3º': {
                "Dentro da Rede": "L14",
                "Rede Estadual": "L15",
                "Litoral": "L16",
                "São Paulo": "L17",
                "ABCD": "L18",
                "Interior": "L19",
                "Outros Estados": "L20",
                "Particular": "L21",
                "País": "L22",
                "Sem Informação": "L23"
            },
            '4º': {
                "Dentro da Rede": "M14",
                "Rede Estadual": "M15",
                "Litoral": "M16",
                "São Paulo": "M17",
                "ABCD": "M18",
                "Interior": "M19",
                "Outros Estados": "M20",
                "Particular": "M21",
                "País": "M22",
                "Sem Informação": "M23"
            },
            '5º': {
                "Dentro da Rede": "N14",
                "Rede Estadual": "N15",
                "Litoral": "N16",
                "São Paulo": "N17",
                "ABCD": "N18",
                "Interior": "N19",
                "Outros Estados": "N20",
                "Particular": "N21",
                "País": "N22",
                "Sem Informação": "N23"
            }
        }
        
        # Inicializa as células de contagem com zero, usando a função auxiliar para células mescladas
        for serie, tipos in mapping.items():
            for tipo, cell_addr in tipos.items():
                current_val = ws[cell_addr].value
                if current_val is None or not isinstance(current_val, (int, float)):
                    set_merged_cell_value(ws, cell_addr, 0)
        
        # Processa cada registro da Lista Piloto Fundamental
        for idx, row in df_fundamental.iterrows():
            if len(row) < 9:
                continue
            
            # Verifica se na coluna I (índice 8) há indicação de transferência ("TE")
            col_I_val = str(row.iloc[8]).strip() if pd.notna(row.iloc[8]) else ""
            if "TE" not in col_I_val:
                continue
            
            # Extrai a data logo após "TE" (formato "dd/mm")
            match = re.search(r"TE\s*([0-9]{1,2}/[0-9]{1,2})", col_I_val)
            if not match:
                continue
            
            te_date_str = match.group(1)
            te_date_full_str = f"{te_date_str}/{period_start.year}"
            try:
                te_date = datetime.strptime(te_date_full_str, "%d/%m/%Y")
            except Exception:
                continue
            
            # Verifica se a data está dentro do intervalo informado
            if not (period_start <= te_date <= period_end):
                continue
            
            # Extrai a série (coluna A, índice 0)
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
                continue  # Se a série não for reconhecida, ignora
            
            # Extrai o TIPO TE (coluna V, índice 21)
            tipo_te = ""
            if len(row) > 21 and pd.notna(row.iloc[21]):
                tipo_te = str(row.iloc[21]).strip()
            if not tipo_te:
                tipo_te = "Sem Informação"
            
            # Se a combinação (série, TIPO TE) estiver mapeada, incrementa a célula correspondente
            if series_key in mapping and tipo_te in mapping[series_key]:
                cell_addr = mapping[series_key][tipo_te]
                current_count = ws[cell_addr].value
                if not isinstance(current_count, (int, float)):
                    current_count = 0
                set_merged_cell_value(ws, cell_addr, current_count + 1)
        
        # Preenche informações adicionais no modelo:
        # Atualiza o responsável e o período informado
        set_merged_cell_value(ws, "B3", responsavel)
        set_merged_cell_value(ws, "D3", f"{period_start.strftime('%d/%m/%Y')} a {period_end.strftime('%d/%m/%Y')}")
        # Atualiza a célula A8 com o mês/ano atual (ex: Março/2025)
        meses = {
            1: "Janeiro", 2: "Fevereiro", 3: "Março", 4: "Abril", 5: "Maio",
            6: "Junho", 7: "Julho", 8: "Agosto", 9: "Setembro", 10: "Outubro",
            11: "Novembro", 12: "Dezembro"
        }
        current_month = meses[datetime.now().month]
        current_year = datetime.now().year
        set_merged_cell_value(ws, "A8", f"{current_month}/{current_year}")
        
        # Prepara o arquivo para download
        output = BytesIO()
        wb.save(output)
        output.seek(0)
        
        filename = f"Quadro_Quantitativo_Fundamental_{period_start.strftime('%d%m')}_{period_end.strftime('%d%m')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    
    # GET: exibe o formulário para entrada dos dados
    form_html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>E.M José Padin Mouta - Quadro Quantitativo Mensal</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
        body { background: #eef2f3; font-family: 'Montserrat', sans-serif; }
        header {
          background: linear-gradient(90deg, #283E51, #4B79A1);
          color: #fff;
          padding: 20px;
          text-align: center;
          border-bottom: 3px solid #1d2d3a;
          border-radius: 0 0 15px 15px;
          box-shadow: 0 4px 6px rgba(0,0,0,0.1);
        }
        .container-form {
          background: #fff;
          padding: 40px;
          border-radius: 10px;
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
          margin: 40px auto;
          max-width: 600px;
        }
        .btn-primary {
          background-color: #283E51;
          border: none;
        }
        .btn-primary:hover {
          background-color: #1d2d3a;
        }
        footer {
          background-color: #424242;
          color: #fff;
          text-align: center;
          padding: 10px;
          position: fixed;
          bottom: 0;
          width: 100%;
        }
      </style>
    </head>
    <body>
      <header>
        <h1>E.M José Padin Mouta - Quadro Quantitativo Mensal</h1>
      </header>
      <div class="container-form">
        <form method="POST" enctype="multipart/form-data">
          <div class="form-group">
            <label for="period_start">Data Inicial:</label>
            <input type="date" class="form-control" name="period_start" id="period_start" required>
          </div>
          <div class="form-group">
            <label for="period_end">Data Final:</label>
            <input type="date" class="form-control" name="period_end" id="period_end" required>
          </div>
          <div class="form-group">
            <label for="responsavel">Responsável pelas informações:</label>
            <input type="text" class="form-control" name="responsavel" id="responsavel" required>
          </div>
          <div class="form-group">
            <label for="lista_fundamental">Selecione a Lista Piloto - FUNDAMENTAL (Excel):</label>
            <input type="file" class="form-control-file" name="lista_fundamental" id="lista_fundamental" accept=".xlsx, .xls">
          </div>
          <button type="submit" class="btn btn-primary">Gerar Quadro Quantitativo</button>
        </form>
        <br>
        <a href="{{ url_for('quadros') }}">Voltar para Quadros</a>
      </div>
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(form_html)


@app.route('/documentos/estagio/<path:filename>')
@login_required
def estagio_documento(filename):
    # Essa rota serve os arquivos diretamente da pasta especificada
    estagio_path = r"C:\Users\Neto\Desktop\Projetos\Em uso\carteirinhas\modelos\estagio"
    return send_from_directory(estagio_path, filename)

@app.route('/documentos')
@login_required
def documentos():
    # Define os diretórios de cada segmento
    base_dir = os.path.join('static', 'documentos')
    matricula_dir = os.path.join(base_dir, 'matricula')
    # Para Estágio, usaremos a pasta definida externamente:
    estagio_dir = r"C:\Users\Neto\Desktop\Projetos\Em uso\carteirinhas\modelos\estagio"
    atas_dir = os.path.join(base_dir, 'atas')
    prontuario_dir = os.path.join(base_dir, 'prontuarios')
    pagamentos_dir = os.path.join(base_dir, 'pagamentos')  # Novo segmento

    # Lista os arquivos de cada diretório (se existir)
    matricula_files = os.listdir(matricula_dir) if os.path.exists(matricula_dir) else []
    estagio_files = os.listdir(estagio_dir) if os.path.exists(estagio_dir) else []
    atas_files = os.listdir(atas_dir) if os.path.exists(atas_dir) else []
    prontuario_files = os.listdir(prontuario_dir) if os.path.exists(prontuario_dir) else []
    pagamentos_files = os.listdir(pagamentos_dir) if os.path.exists(pagamentos_dir) else []

    # Converte os nomes dos arquivos em URLs para acesso via navegador
    matricula_files = [f"/static/documentos/matricula/{file}" for file in matricula_files]
    estagio_files = [url_for('estagio_documento', filename=file) for file in estagio_files]
    atas_files = [f"/static/documentos/atas/{file}" for file in atas_files]
    prontuario_files = [f"/static/documentos/prontuarios/{file}" for file in prontuario_files]
    pagamentos_files = [f"/static/documentos/pagamentos/{file}" for file in pagamentos_files]

    html = '''
    <!doctype html>
    <html lang="pt-br">
    <head>
      <meta charset="utf-8">
      <title>Documentos - E.M José Padin Mouta</title>
      <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
      <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
      <style>
         body {
            background: #eef2f3;
            font-family: 'Montserrat', sans-serif;
         }
         header {
            background: linear-gradient(90deg, #283E51, #4B79A1);
            color: #fff;
            padding: 20px;
            text-align: center;
            border-bottom: 3px solid #1d2d3a;
            border-radius: 0 0 15px 15px;
            box-shadow: 0 4px 6px rgba(0,0,0,0.1);
         }
         .container-dashboard {
            background: #fff;
            padding: 40px;
            border-radius: 10px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.15);
            margin: 40px auto;
            max-width: 800px;
         }
         .card {
             margin-bottom: 20px;
         }
         .card-header {
             background: #f8f9fa;
             border-bottom: 1px solid #dee2e6;
         }
         .card-header h5 {
             margin: 0;
         }
         .btn-secondary {
             background-color: #4B79A1;
             color: #fff;
             border: none;
             padding: 10px 20px;
             border-radius: 5px;
             transition: background-color 0.3s;
         }
         .btn-secondary:hover {
             background-color: #3a5d78;
         }
         footer {
            background-color: #424242;
            color: #fff;
            text-align: center;
            padding: 10px;
            position: fixed;
            bottom: 0;
            width: 100%;
         }
         /* Estilos customizados para o modal */
         .modal-dialog {
             max-width: 800px; /* Aumenta a largura do container do modal */
         }
         .modal-body ul li {
             margin-bottom: 10px; /* Espaçamento entre os itens da lista */
         }
         .modal-body ul li ul li {
             margin-bottom: 5px;  /* Espaçamento entre os subtópicos */
         }
      </style>
    </head>
    <body>
      <header>
        <h1>E.M José Padin Mouta - Documentos</h1>
      </header>
      <div class="container container-dashboard">
        <h1 class="mb-4">Documentos</h1>
        <div id="accordion">
          <!-- Segmento Matrícula -->
          <div class="card">
            <div class="card-header" id="headingMatricula">
              <h5 class="mb-0">
                <button class="btn btn-link collapsed" type="button" data-toggle="collapse" data-target="#collapseMatricula" aria-expanded="false" aria-controls="collapseMatricula">
                  Matrícula
                </button>
              </h5>
            </div>
            <div id="collapseMatricula" class="collapse" aria-labelledby="headingMatricula" data-parent="#accordion">
              <div class="card-body">
                {% if matricula_files %}
                  <ul>
                    {% for file in matricula_files %}
                      <li><a href="{{ file }}" target="_blank">{{ unquote(file.split('/')[-1]) }}</a></li>
                    {% endfor %}
                  </ul>
                {% else %}
                  <p>Nenhum documento de Matrícula.</p>
                {% endif %}
                <p>
                  <a href="#" data-toggle="modal" data-target="#modalMatricula">Leia os procedimentos para Matrícula</a>
                </p>
              </div>
            </div>
          </div>
          <!-- Segmento Estágio -->
          <div class="card">
            <div class="card-header" id="headingEstagio">
              <h5 class="mb-0">
                <button class="btn btn-link collapsed" type="button" data-toggle="collapse" data-target="#collapseEstagio" aria-expanded="false" aria-controls="collapseEstagio">
                  Estágio
                </button>
              </h5>
            </div>
            <div id="collapseEstagio" class="collapse" aria-labelledby="headingEstagio" data-parent="#accordion">
              <div class="card-body">
                {% if estagio_files %}
                  <ul>
                    {% for file in estagio_files %}
                      <li><a href="{{ file }}" target="_blank">{{ unquote(file.split('/')[-1]) }}</a></li>
                    {% endfor %}
                  </ul>
                {% else %}
                  <p>Nenhum documento de Estágio.</p>
                {% endif %}
                <p>
                  <a href="#" data-toggle="modal" data-target="#modalEstagio">Leia os procedimentos para Estágio</a>
                </p>
              </div>
            </div>
          </div>
          <!-- Segmento Atas -->
          <div class="card">
            <div class="card-header" id="headingAtas">
              <h5 class="mb-0">
                <button class="btn btn-link collapsed" type="button" data-toggle="collapse" data-target="#collapseAtas" aria-expanded="false" aria-controls="collapseAtas">
                  Atas
                </button>
              </h5>
            </div>
            <div id="collapseAtas" class="collapse" aria-labelledby="headingAtas" data-parent="#accordion">
              <div class="card-body">
                {% if atas_files %}
                  <ul>
                    {% for file in atas_files %}
                      <li><a href="{{ file }}" target="_blank">{{ unquote(file.split('/')[-1]) }}</a></li>
                    {% endfor %}
                  </ul>
                {% else %}
                  <p>Nenhum documento de Atas.</p>
                {% endif %}
              </div>
            </div>
          </div>
          <!-- Segmento Prontuários -->
          <div class="card">
            <div class="card-header" id="headingProntuarios">
              <h5 class="mb-0">
                <button class="btn btn-link collapsed" type="button" data-toggle="collapse" data-target="#collapseProntuarios" aria-expanded="false" aria-controls="collapseProntuarios">
                  Prontuários
                </button>
              </h5>
            </div>
            <div id="collapseProntuarios" class="collapse" aria-labelledby="headingProntuarios" data-parent="#accordion">
              <div class="card-body">
                {% if prontuario_files %}
                  <ul>
                    {% for file in prontuario_files %}
                      <li><a href="{{ file }}" target="_blank">{{ unquote(file.split('/')[-1]) }}</a></li>
                    {% endfor %}
                  </ul>
                {% else %}
                  <p>Nenhum documento de Prontuários.</p>
                {% endif %}
              </div>
            </div>
          </div>
          <!-- Segmento Pagamentos -->
          <div class="card">
            <div class="card-header" id="headingPagamentos">
              <h5 class="mb-0">
                <button class="btn btn-link collapsed" type="button" data-toggle="collapse" data-target="#collapsePagamentos" aria-expanded="false" aria-controls="collapsePagamentos">
                  Pagamentos
                </button>
              </h5>
            </div>
            <div id="collapsePagamentos" class="collapse" aria-labelledby="headingPagamentos" data-parent="#accordion">
              <div class="card-body">
                {% if pagamentos_files %}
                  <ul>
                    {% for file in pagamentos_files %}
                      <li><a href="{{ file }}" target="_blank">{{ unquote(file.split('/')[-1]) }}</a></li>
                    {% endfor %}
                  </ul>
                {% else %}
                  <p>Nenhum documento de Pagamentos.</p>
                {% endif %}
              </div>
            </div>
          </div>
        </div>
        <br>
        <a href="{{ url_for('dashboard') }}" class="btn btn-secondary">Voltar ao Dashboard</a>
      </div>
      
      <!-- Modal para Procedimentos de Matrícula -->
      <div class="modal fade" id="modalMatricula" tabindex="-1" role="dialog" aria-labelledby="modalMatriculaLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered" role="document">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title" id="modalMatriculaLabel">Procedimento para Realização de Matrícula</h5>
              <button type="button" class="close" data-dismiss="modal" aria-label="Fechar">
                <span aria-hidden="true">&times;</span>
              </button>
            </div>
            <div class="modal-body">
              <ul>
                <li>
                  Realizar a conferência dos documentos entregues pelo responsável, verificando obrigatoriamente:
                  <ul>
                    <li>Certidão de Nascimento do aluno;</li>
                    <li>CPF do aluno;</li>
                    <li>RG do responsável;</li>
                    <li>Comprovante de residência;</li>
                    <li>Carteira de vacinação;</li>
                    <li>Cartão do SUS;</li>
                    <li>Declaração original de transferência emitida pela escola de origem;</li>
                    <li>2 fotos 3x4 (não obrigatórias no ato da matrícula).</li>
                  </ul>
                </li>
                <li>
                  Após confirmar que toda documentação está correta, realizar a impressão da ficha cadastral a partir do sistema Gestão, atribuindo ao aluno o número de Registro de Matrícula (RM) e preenchendo adequadamente o arquivo de RM;
                </li>
                <li>
                  Providenciar a impressão do Kit de Matrícula e entregá-lo ao responsável para preenchimento. Após devolução, verificar se o responsável assinou todos os campos exigidos;
                </li>
                <li>
                  Informar claramente ao responsável sobre o horário das aulas, o nome do professor responsável, série do aluno e demais procedimentos escolares pertinentes. Em seguida, liberar o responsável;
                </li>
                <li>
                  Preparar o prontuário do aluno, preenchendo-o integralmente com todas as informações pertinentes, e deixá-lo na mesa do secretário para arquivamento.
                </li>
              </ul>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-secondary" data-dismiss="modal">Fechar</button>
            </div>
          </div>
        </div>
      </div>
      
      <!-- Modal para Procedimentos de Estágio -->
      <div class="modal fade" id="modalEstagio" tabindex="-1" role="dialog" aria-labelledby="modalEstagioLabel" aria-hidden="true">
        <div class="modal-dialog modal-dialog-centered" role="document">
          <div class="modal-content">
            <div class="modal-header">
              <h5 class="modal-title" id="modalEstagioLabel">Procedimento para Recebimento e Cadastro de Estagiário</h5>
              <button type="button" class="close" data-dismiss="modal" aria-label="Fechar">
                <span aria-hidden="true">&times;</span>
              </button>
            </div>
            <div class="modal-body">
              <p>Receber o estagiário, que deverá apresentar o encaminhamento. <br>Solicitar que o mesmo preencha integralmente a Ficha de Cadastro - Estágio não remunerado disponível em nosso sistema.</p>
              <p>Após o preenchimento completo da ficha cadastral, dispensar o estagiário informando que será encaminhado um e-mail para a SEDUC solicitando autorização e que o contato será feito posteriormente para informar sobre o deferimento.</p>
              <p>Enviar e-mail para <a href="mailto:seduc.legislacao3@praiagrande.sp.gov.br">seduc.legislacao3@praiagrande.sp.gov.br</a> com os seguintes dados coletados na ficha cadastral:</p>
              <ul>
                <li>Nome completo;</li>
                <li>RG;</li>
                <li>CPF;</li>
                <li>Data de nascimento;</li>
                <li>Curso;</li>
                <li>Semestre atual;</li>
                <li>Horário pretendido para o estágio.</li>
              </ul>
              <p>Aguardar a resposta da SEDUC por e-mail com o deferimento ou indeferimento do estágio solicitado. <br> Assim que obtiver resposta, entrar em contato com o estagiário para informá-lo se está autorizado a iniciar o estágio ou não.</p>
            </div>
            <div class="modal-footer">
              <button type="button" class="btn btn-secondary" data-dismiss="modal">Fechar</button>
            </div>
          </div>
        </div>
      </div>
      
      <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
      <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
      <!-- Footer -->
      <footer>
        Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
      </footer>
    </body>
    </html>
    '''
    return render_template_string(html, 
                                  matricula_files=matricula_files, 
                                  estagio_files=estagio_files, 
                                  atas_files=atas_files,
                                  prontuario_files=prontuario_files,
                                  pagamentos_files=pagamentos_files,
                                  unquote=unquote)


# Lista de setores disponíveis para o dropdown
SETORES = ["Financeiro", "Recursos Humanos", "Administrativo", "Marketing", "TI"]

def data_extenso(dt):
    """Retorna a data por extenso em português."""
    meses = [
        "janeiro", "fevereiro", "março", "abril", "maio", "junho",
        "julho", "agosto", "setembro", "outubro", "novembro", "dezembro"
    ]
    return f"{dt.day} de {meses[dt.month-1]} de {dt.year}"

# Template para o formulário de criação do ofício
oficio_form_template = """
<!doctype html>
<html lang="pt-br">
  <head>
    <meta charset="utf-8">
    <title>E.M José Padin Mouta - Gerador de Ofício</title>
    <meta name="viewport" content="width=device-width, initial-scale=1, shrink-to-fit=no">
    <!-- Bootstrap, Font Awesome e Google Fonts -->
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/5.15.4/css/all.min.css">
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
    <style>
      body {
        background: #eef2f3;
        font-family: 'Montserrat', sans-serif;
        margin-bottom: 60px;
      }
      header {
        background: linear-gradient(90deg, #283E51, #4B79A1);
        color: #fff;
        padding: 20px;
        text-align: center;
        border-bottom: 3px solid #1d2d3a;
        border-radius: 0 0 15px 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
      }
      .container-menu {
        margin: 40px auto;
        max-width: 900px;
        background: #fff;
        padding: 40px;
        border-radius: 10px;
        box-shadow: 0 4px 12px rgba(0,0,0,0.15);
      }
      .form-group {
        margin-bottom: 20px;
      }
      label {
        font-weight: 600;
      }
      .btn-submit {
        background-color: #4B79A1;
        color: #fff;
        border: none;
        padding: 10px 20px;
        border-radius: 5px;
        transition: background-color 0.3s;
      }
      .btn-submit:hover {
        background-color: #3a5d78;
      }
      .error {
        color: red;
      }
      .btn-voltar {
        display: inline-block;
        padding: 10px 20px;
        font-size: 16px;
        font-weight: 600;
        color: #fff;
        background-color: #4B79A1;
        border: none;
        border-radius: 5px;
        text-decoration: none;
        transition: background-color 0.3s;
      }
      .btn-voltar:hover {
        background-color: #3a5d78;
      }
      footer {
        background-color: #424242;
        color: #fff;
        text-align: center;
        padding: 10px;
        position: fixed;
        bottom: 0;
        width: 100%;
      }
    </style>
  </head>
  <body>
    <header>
      <h1>E.M José Padin Mouta - Gerador de Ofício</h1>
    </header>
    <div class="container-menu">
      {% with messages = get_flashed_messages(category_filter=["error"]) %}
        {% if messages %}
          <div class="alert alert-danger" role="alert">
            {% for msg in messages %}
              <p>{{ msg }}</p>
            {% endfor %}
          </div>
        {% endif %}
      {% endwith %}
      <form method="POST">
        <div class="form-group">
          <label for="destinatario">Destinatário:</label>
          <input type="text" id="destinatario" name="destinatario" class="form-control" required>
        </div>
        <div class="form-group">
          <label for="setor">Setor:</label>
          <select id="setor" name="setor" class="form-control" required>
            {% for setor in setores %}
              <option value="{{ setor }}">{{ setor }}</option>
            {% endfor %}
          </select>
        </div>
        <div class="form-group">
          <label for="corpo">Corpo do Ofício:</label>
          <textarea id="corpo" name="corpo" rows="6" class="form-control" required></textarea>
        </div>
        <div class="form-group">
          <label for="oficio_num">Número do Ofício (3 dígitos):</label>
          <input type="text" id="oficio_num" name="oficio_num" pattern="\\d{3}" placeholder="Ex: 001" class="form-control" required>
        </div>
        <button type="submit" class="btn btn-submit">Gerar Ofício</button>
      </form>
      <div class="mt-4">
      </div>
    </div>
    <footer>
      Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
    </footer>
    <!-- Scripts do Bootstrap -->
    <script src="https://code.jquery.com/jquery-3.5.1.slim.min.js"></script>
    <script src="https://cdn.jsdelivr.net/npm/bootstrap@4.5.2/dist/js/bootstrap.bundle.min.js"></script>
  </body>
</html>
"""

# Template para exibição do ofício gerado (HTML para impressão)
oficio_result_template = """
<!DOCTYPE html>
<html lang="pt-br">
<head>
  <meta charset="utf-8">
  <title>{{ titulo }} - E.M José Padin Mouta</title>
  <style>
    /* Estilos para visualização na tela como uma folha A4 */
    @media screen {
      body {
        background: #eee;
      }
      .declaration-container {
        width: 210mm;
        min-height: 297mm;
        margin: 20px auto;
        padding: 20mm;
        background: white;
      }
      .no-print {
        width: 210mm;
        margin: 10px auto 0 auto;
        text-align: right;
      }
      .no-print button, .no-print a {
        background-color: #283E51;
        color: #fff;
        border: none;
        padding: 10px 20px;
        border-radius: 4px;
        cursor: pointer;
        font-size: 12pt;
        margin-right: 10px;
        text-decoration: none;
      }
      .no-print button:hover, .no-print a:hover {
        background-color: #1e2c3d;
      }
    }
    
    /* Margens para impressão: topo 1cm, demais 2.5cm */
    @page { margin: 1cm 2.5cm 2.5cm 2.5cm; }
    
    html, body {
      margin: 0;
      padding: 0;
      font-family: Arial, sans-serif;
      font-size: 12pt;
      line-height: 1.5;
      color: #333;
    }
    
    /* Cabeçalho com informações centralizadas e logo à direita */
    .header {
      position: relative;
      text-align: center;
      border-bottom: 2px solid #283E51;
      padding-bottom: 5px;
      margin-bottom: 40px;
    }
    
    .header .header-text h1 {
      margin: 0;
      font-size: 16pt;
      text-transform: uppercase;
      color: #283E51;
    }
    
    .header .header-text p {
      margin: 3px 0;
      font-size: 12pt;
    }
    
    /* Logo posicionado à direita */
    .header .logo {
      position: absolute;
      top: 0;
      right: 0;
    }
    
    /* Data posicionada no topo direito */
    .date {
      text-align: right;
      font-size: 12pt;
      margin-bottom: 10px;
    }
    
    .content {
      text-align: justify;
      margin-bottom: 10px;
    }

    /* Número do ofício, alinhado à esquerda */
    .oficio-number {
      text-align: left;
      font-size: 14pt;
      font-weight: bold;
      margin-bottom: 10px;
    }
    
    .signature {
      text-align: center;
      margin-top: 0;
    }
    
    .signature .line {
      height: 1px;
      background-color: #333;
      width: 60%;
      margin: 0 auto 5px auto;
    }
    
    .footer {
      text-align: center;
      border-top: 2px solid #283E51;
      padding-top: 5px;
      margin: 0;
      font-size: 10pt;
      color: #555;
    }
    
    /* Estilos aplicados apenas na impressão */
    @media print {
      .no-print { display: none !important; }
      .header { margin-top: 0; margin-bottom: 60px; }
      .signature { margin-top: 7cm; }
      .footer { margin-top: 1cm; }
    }
    
    {{ additional_css|default('') }}
  </style>
</head>
<body>
  <!-- Botões de Imprimir e Download PDF Editável posicionados acima da folha -->
  <div class="no-print">
    <button onclick="window.print()">Imprimir Ofício</button>
    <a href="{{ url_for('oficio_pdf', destinatario=destinatario, setor=setor, corpo=corpo, oficio_num=oficio_num, data_extenso=data_extenso, titulo=titulo) }}">Download PDF Editável</a>
  </div>
  
  <div class="declaration-container">
    <!-- Cabeçalho com informações centralizadas e logo à direita -->
    <div class="header">
      <div class="header-text">
        <h1>Secretaria de Educação</h1>
        <p>E.M José Padin Mouta</p>
        <p>Município da Estância Balneária de Praia Grande</p>
        <p>Estado de São Paulo</p>
      </div>
      <div class="logo">
        <img src="/static/logos/municipio.png" alt="Município Logo" style="height: 80px;">
      </div>
    </div>
    
    <!-- Data exibida no topo direito -->
    <div class="date">
      <p>{{ data_extenso }}</p>
    </div>
    
    <div class="content">
      <div class="oficio-number">
        <p>OFÍCIO Nº {{ oficio_num }}/<span id="currentYear">{{ current_year|default("2025") }}</span></p>
      </div>
      <p><strong>Destinatário:</strong> {{ destinatario }}</p>
      <p><strong>Setor:</strong> {{ setor }}</p>
      <p>{{ corpo }}</p>
    </div>
    
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

  <script>
    // Atualiza o ano automaticamente, se necessário
    document.addEventListener("DOMContentLoaded", function() {
      var currentYearElement = document.getElementById("currentYear");
      if (currentYearElement && (!currentYearElement.textContent || currentYearElement.textContent === "2025")) {
        currentYearElement.textContent = new Date().getFullYear();
      }
    });
  </script>
</body>
</html>
"""

def generate_editable_pdf(destinatario, setor, corpo, oficio_num, data_extenso_text, titulo):
    """
    Gera um PDF editável com AcroForm que tenta replicar a formatação do ofício impresso,
    com os rótulos "Destinatário:" e "Setor:" em negrito, e permitindo maior quantidade de
    caracteres no campo "Corpo do Ofício".
    """
    from reportlab.pdfgen import canvas
    from reportlab.lib.pagesizes import A4
    from io import BytesIO

    buffer = BytesIO()
    c = canvas.Canvas(buffer, pagesize=A4)
    width, height = A4

    # Margens aproximadas
    left_margin = 70      # ~2.5 cm
    right_margin = 70     # ~2.5 cm
    top_margin = 28       # ~1 cm
    bottom_margin = 70    # ~2.5 cm

    # --- Cabeçalho ---
    header_top = height - top_margin
    header_bottom = header_top - 80  # área do cabeçalho
    c.setFont("Helvetica-Bold", 16)
    c.drawCentredString(width/2, header_top - 20, "Secretaria de Educação")
    c.setFont("Helvetica", 12)
    c.drawCentredString(width/2, header_top - 40, "E.M José Padin Mouta")
    c.drawCentredString(width/2, header_top - 55, "Município da Estância Balneária de Praia Grande")
    c.drawCentredString(width/2, header_top - 70, "Estado de São Paulo")
    c.line(left_margin, header_bottom, width - right_margin, header_bottom)
    
    # --- Logo ---
    try:
        logo_width = 80
        logo_height = 60
        logo_x = width - right_margin - logo_width
        logo_y = header_top - logo_height
        c.drawImage("static/logos/municipio.png", logo_x, logo_y,
                    width=logo_width, height=logo_height,
                    mask='auto', preserveAspectRatio=True)
    except Exception:
        pass

    # --- Data e Número do Ofício ---
    c.setFont("Helvetica", 12)
    c.drawRightString(width - right_margin, header_bottom - 20, data_extenso_text)
    c.setFont("Helvetica-Bold", 14)
    year = data_extenso_text.split()[-1]
    of_text = f"OFÍCIO Nº {oficio_num}/{year}"
    c.drawString(left_margin, header_bottom - 50, of_text)

    # --- Campos Editáveis ---
    current_y = header_bottom - 90  # espaçamento após o número do ofício
    field_height = 20

    # Destinatário: rótulo (negrito) e campo
    label1 = "Destinatário:"
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left_margin, current_y, label1)
    label1_width = c.stringWidth(label1, "Helvetica-Bold", 12)
    field_x = left_margin + label1_width + 10  # campo imediatamente após o rótulo
    field_y = current_y - 8
    c.acroForm.textfield(
        name='destinatario', tooltip='Destinatário',
        x=field_x, y=field_y,
        width=300, height=field_height,
        borderWidth=0, forceBorder=False,
        fontName="Helvetica", fontSize=12,
        value=destinatario
    )
    current_y -= 30

    # Setor: rótulo (negrito) e campo
    label2 = "Setor:"
    c.setFont("Helvetica-Bold", 12)
    c.drawString(left_margin, current_y, label2)
    label2_width = c.stringWidth(label2, "Helvetica-Bold", 12)
    field_x = left_margin + label2_width + 10
    field_y = current_y - 8
    c.acroForm.textfield(
        name='setor', tooltip='Setor',
        x=field_x, y=field_y,
        width=300, height=field_height,
        borderWidth=0, forceBorder=False,
        fontName="Helvetica", fontSize=12,
        value=setor
    )
    current_y -= 30

    # Corpo do Ofício: somente o campo editável (sem rótulo)
    corpo_field_height = 200
    # Adicionado maxlen=0 para permitir quantidade ilimitada de caracteres
    c.acroForm.textfield(
        name='corpo', tooltip='Corpo do Ofício',
        x=left_margin, y=current_y - corpo_field_height,
        width=width - left_margin - right_margin, height=corpo_field_height,
        borderWidth=0, forceBorder=False,
        fontName="Helvetica", fontSize=12,
        value=corpo, fieldFlags=4096, maxlen=0
    )
    current_y = current_y - corpo_field_height - 30

    # --- Assinatura ---
    # Área de assinatura
    sig_top = bottom_margin + 170
    short_line_width = 250  # largura do traço superior
    c.line((width - short_line_width) / 2, sig_top, (width + short_line_width) / 2, sig_top)
    c.setFont("Helvetica", 12)
    c.drawCentredString(width/2, sig_top - 15, "Luciana Rocha Augustinho")
    c.drawCentredString(width/2, sig_top - 30, "Diretora da Unidade Escolar")
    # Linha divisória abaixo da assinatura (posição mantida conforme ajuste anterior)
    divider_y = bottom_margin + 40
    c.line(left_margin, divider_y, width - right_margin, divider_y)

    # --- Rodapé ---
    c.setFont("Helvetica", 10)
    footer_y1 = bottom_margin - 20
    footer_y2 = bottom_margin - 35
    c.drawCentredString(width/2, footer_y1, "Rua: Bororós, nº 150, Vila Tupi, Praia Grande - SP, CEP: 11703-390")
    c.drawCentredString(width/2, footer_y2, "Telefone: 3496-5321 | E-mail: em.padin@praiagrande.sp.gov.br")

    c.showPage()
    c.save()
    buffer.seek(0)
    return buffer

@app.route('/oficio', methods=["GET", "POST"])
def oficio():
    if request.method == "POST":
        destinatario = request.form.get("destinatario")
        setor = request.form.get("setor")
        corpo = request.form.get("corpo")
        oficio_num = request.form.get("oficio_num")

        # Validação: o número do ofício deve ter exatamente 3 dígitos
        if not re.fullmatch(r"\d{3}", oficio_num):
            flash("Número do ofício deve ter exatamente 3 dígitos.", "error")
            return redirect(url_for("oficio"))

        hoje = datetime.today()
        data_por_extenso = data_extenso(hoje)
        titulo = f"Ofício nº {oficio_num}"

        return render_template_string(
            oficio_result_template,
            destinatario=destinatario,
            setor=setor,
            corpo=corpo,
            oficio_num=oficio_num,
            data_extenso=data_por_extenso,
            titulo=titulo,
            current_year=hoje.year,
            additional_css=""
        )
    # Método GET: renderiza o formulário com a lista de setores
    return render_template_string(oficio_form_template, setores=SETORES)

@app.route('/oficio/pdf')
def oficio_pdf():
    """
    Rota que gera e retorna o PDF editável do ofício.
    Os dados são passados via query string.
    """
    destinatario = request.args.get("destinatario")
    setor = request.args.get("setor")
    corpo = request.args.get("corpo")
    oficio_num = request.args.get("oficio_num")
    data_extenso_text = request.args.get("data_extenso")
    titulo = request.args.get("titulo")
    if not all([destinatario, setor, corpo, oficio_num, data_extenso_text, titulo]):
        flash("Dados insuficientes para gerar PDF editável.", "error")
        return redirect(url_for("oficio"))
    pdf_buffer = generate_editable_pdf(destinatario, setor, corpo, oficio_num, data_extenso_text, titulo)
    return send_file(pdf_buffer, as_attachment=True, download_name="oficio_editavel.pdf", mimetype="application/pdf")

# Predefinições (cada cargo e suas horas) - mesmo dicionário fornecido 
CARGOS = {
    "PROFESSOR RECREACIONISTA": {
        "total": 30,
        "aluno": 20,
        "htpi_pres": 2,
        "htpi_online": 3,
        "htpi_total": 5,
        "htpc": 2,
        "hal": 3
    },
    "PROFESSOR ADJUNTO I": {
        "total": 30,
        "aluno": 20,
        "htpi_pres": 2,
        "htpi_online": 3,
        "htpi_total": 5,
        "htpc": 2,
        "hal": 3
    },
    "PROFESSOR I": {
        "total": 27,
        "aluno": 18,
        "htpi_pres": 2,
        "htpi_online": 2,
        "htpi_total": 4,
        "htpc": 2,
        "hal": 3
    },
    "PROFESSOR II": {
        "total": 27,
        "aluno": 18,
        "htpi_pres": 2,
        "htpi_online": 2,
        "htpi_total": 4,
        "htpc": 2,
        "hal": 3
    },
    "PROFESSOR III": {
        "total": 30,
        "aluno": 20,
        "htpi_pres": 2,
        "htpi_online": 3,
        "htpi_total": 5,
        "htpc": 2,
        "hal": 3
    },
    "PROFESSOR IV": {
        "total": 15,
        "aluno": 10,
        "htpi_pres": 1,
        "htpi_online": 1,
        "htpi_total": 2,
        "htpc": 2,
        "hal": 1
    }
}

FORM_HTML = """
<!doctype html>
<html lang="pt-br">
  <head>
    <meta charset="UTF-8">
    <title>Acumulo de Cargo</title>
    <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
    <style>
      body {
        background: #eef2f3;
        font-family: 'Montserrat', sans-serif;
        margin: 0;
        padding: 0;
        display: flex;
        flex-direction: column;
        min-height: 100vh;
      }
      header {
        background: linear-gradient(90deg, #283E51, #4B79A1);
        color: #fff;
        padding: 20px;
        text-align: center;
        border-bottom: 3px solid #1d2d3a;
        border-radius: 0 0 15px 15px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.1);
      }
      main {
        flex: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        padding: 20px;
      }
      .container {
        background: #fff;
        padding: 20px;
        margin-top: 30px;
        border-radius: 8px;
        max-width: 700px;
        width: 100%;
      }
      h1 {
        margin-bottom: 20px;
      }
      .form-day {
        border: 1px solid #ccc;
        padding: 10px;
        margin-bottom: 10px;
        border-radius: 5px;
      }
      .form-day h5 {
        margin-bottom: 8px;
      }
      .day-row {
        display: flex;
        flex-wrap: wrap;
        margin-bottom: 5px;
      }
      .day-row label {
        width: 120px;
        margin-right: 10px;
      }
      .day-row input {
        width: 150px;
        margin-right: 10px;
      }
      .btn-primary {
        background-color: #283E51;
        border: none;
      }
      .btn-primary:hover {
        background-color: #1d2d3a;
      }
      footer {
        background-color: #424242;
        color: #fff;
        text-align: center;
        padding: 10px;
      }
    </style>
  </head>
  <body>
    <header>
      <h1>E.M José Padin Mouta - Secretaria</h1>
    </header>
    <main>
      <div class="container">
        <h1>Declaração de Acúmulo de Cargo</h1>
        <form method="POST" id="acumuloForm">
          <!-- Dados Pessoais -->
          <div class="form-group">
            <label for="nome">Nome do Professor(a):</label>
            <input type="text" class="form-control" id="nome" name="nome" required>
          </div>
          <div class="form-group">
            <label for="rg">RG:</label>
            <input type="text" class="form-control" id="rg" name="rg" required>
          </div>
          <div class="form-group">
            <label for="cargo">Selecione o Cargo:</label>
            <select id="cargo" name="cargo" class="form-control" required>
              <option value="">Selecione</option>
              <option value="PROFESSOR RECREACIONISTA">PROFESSOR RECREACIONISTA</option>
              <option value="PROFESSOR ADJUNTO I">PROFESSOR ADJUNTO I</option>
              <option value="PROFESSOR I">PROFESSOR I</option>
              <option value="PROFESSOR II">PROFESSOR II</option>
              <option value="PROFESSOR III">PROFESSOR III</option>
              <option value="PROFESSOR IV">PROFESSOR IV</option>
            </select>
          </div>
          <div class="form-group">
            <label for="disciplina">Disciplina (ex: Polivalente, Matemática...):</label>
            <select class="form-control" id="disciplina" name="disciplina" required>
              <option value="">Selecione</option>
              <option value="Administração">Administração</option>
              <option value="Artes">Artes</option>
              <option value="Biologia">Biologia</option>
              <option value="Ciências">Ciências</option>
              <option value="Ciências Sociais">Ciências Sociais</option>
              <option value="Educação Ambiental">Educação Ambiental</option>
              <option value="Educação Especial">Educação Especial</option>
              <option value="Educação Física">Educação Física</option>
              <option value="Filosofia">Filosofia</option>
              <option value="Física">Física</option>
              <option value="Geografia">Geografia</option>
              <option value="História">História</option>
              <option value="Inglês">Inglês</option>
              <option value="Informática">Informática</option>
              <option value="Língua Portuguesa">Língua Portuguesa</option>
              <option value="Literatura">Literatura</option>
              <option value="Matemática">Matemática</option>
              <option value="Música">Música</option>
              <option value="Polivalente">Polivalente</option>
              <option value="Química">Química</option>
              <option value="Sociologia">Sociologia</option>
              <option value="Tecnologia">Tecnologia</option>
            </select>
          </div>

          <!-- Horários de Trabalho -->
          <!-- Os rótulos aqui foram abreviados -->
          <div class="form-day">
            <h5>Segunda-Feira</h5>
            <div class="day-row">
              <label>Com Aluno:</label>
              <input type="time" name="mon_aluno_ini">
              <input type="time" name="mon_aluno_fim">
            </div>
            <div class="day-row">
              <label>HTPC:</label>
              <input type="time" name="mon_htpc_ini">
              <input type="time" name="mon_htpc_fim">
            </div>
            <div class="day-row">
              <label>HTPI:</label>
              <input type="time" name="mon_hti_ini">
              <input type="time" name="mon_hti_fim">
            </div>
          </div>

          <div class="form-day">
            <h5>Terça-Feira</h5>
            <div class="day-row">
              <label>Com Aluno:</label>
              <input type="time" name="tue_aluno_ini">
              <input type="time" name="tue_aluno_fim">
            </div>
            <div class="day-row">
              <label>HTPC:</label>
              <input type="time" name="tue_htpc_ini">
              <input type="time" name="tue_htpc_fim">
            </div>
            <div class="day-row">
              <label>HTPI:</label>
              <input type="time" name="tue_hti_ini">
              <input type="time" name="tue_hti_fim">
            </div>
          </div>

          <div class="form-day">
            <h5>Quarta-Feira</h5>
            <div class="day-row">
              <label>Com Aluno:</label>
              <input type="time" name="wed_aluno_ini">
              <input type="time" name="wed_aluno_fim">
            </div>
            <div class="day-row">
              <label>HTPC:</label>
              <input type="time" name="wed_htpc_ini">
              <input type="time" name="wed_htpc_fim">
            </div>
            <div class="day-row">
              <label>HTPI:</label>
              <input type="time" name="wed_hti_ini">
              <input type="time" name="wed_hti_fim">
            </div>
          </div>

          <div class="form-day">
            <h5>Quinta-Feira</h5>
            <div class="day-row">
              <label>Com Aluno:</label>
              <input type="time" name="thu_aluno_ini">
              <input type="time" name="thu_aluno_fim">
            </div>
            <div class="day-row">
              <label>HTPC:</label>
              <input type="time" name="thu_htpc_ini">
              <input type="time" name="thu_htpc_fim">
            </div>
            <div class="day-row">
              <label>HTPI:</label>
              <input type="time" name="thu_hti_ini">
              <input type="time" name="thu_hti_fim">
            </div>
          </div>

          <div class="form-day">
            <h5>Sexta-Feira</h5>
            <div class="day-row">
              <label>Com Aluno:</label>
              <input type="time" name="fri_aluno_ini">
              <input type="time" name="fri_aluno_fim">
            </div>
            <div class="day-row">
              <label>HTPC:</label>
              <input type="time" name="fri_htpc_ini">
              <input type="time" name="fri_htpc_fim">
            </div>
            <div class="day-row">
              <label>HTPI:</label>
              <input type="time" name="fri_hti_ini">
              <input type="time" name="fri_hti_fim">
            </div>
          </div>

          <button type="submit" class="btn btn-primary">Gerar Declaração</button>
        </form>
      </div>
    </main>
    <footer>
      Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
    </footer>
    
    <!-- Script para propagar automaticamente o horário "Com Aluno" da Segunda-Feira para os demais dias -->
    <script>
      document.addEventListener("DOMContentLoaded", function() {
        var cargoSelect = document.querySelector("select[name='cargo']");
        var monAlunoIni = document.querySelector("input[name='mon_aluno_ini']");
        var monAlunoFim = document.querySelector("input[name='mon_aluno_fim']");
        var otherAlunoIniFields = document.querySelectorAll("input[name$='_aluno_ini']:not([name='mon_aluno_ini'])");
        var otherAlunoFimFields = document.querySelectorAll("input[name$='_aluno_fim']:not([name='mon_aluno_fim'])");
        var allowedCargos = ["PROFESSOR RECREACIONISTA", "PROFESSOR ADJUNTO I", "PROFESSOR I", "PROFESSOR II"];
  
        function propagateAlunoSchedule() {
          if (cargoSelect.value && allowedCargos.includes(cargoSelect.value.toUpperCase()) &&
              monAlunoIni.value && monAlunoFim.value) {
            otherAlunoIniFields.forEach(function(field) {
              field.value = monAlunoIni.value;
            });
            otherAlunoFimFields.forEach(function(field) {
              field.value = monAlunoFim.value;
            });
          }
        }
  
        // Propaga quando os campos de segunda-feira forem alterados
        monAlunoIni.addEventListener("change", propagateAlunoSchedule);
        monAlunoFim.addEventListener("change", propagateAlunoSchedule);
        // E também quando o cargo for alterado
        cargoSelect.addEventListener("change", propagateAlunoSchedule);
      });
    </script>
  </body>
</html>
"""

# HTML do documento final (Declaração) - layout horizontal (landscape)
# A primeira página permanece inalterada; a segunda página apresenta duas cópias da declaração,
# cada uma com seu próprio cabeçalho (logo menor + "Declaração de Acúmulo")
# e os campos organizados conforme o layout abaixo:
#
#   ( ) Acumulo cargo/função:
#
#   Federal:
#       Local: __________________________
#       Cargo/Função: __________________________
#
#   Estadual:
#       Local: __________________________
#       Cargo/Função: __________________________
#
#   Municipal:
#       Local: __________________________
#       Cargo/Função: __________________________
#
#   ( ) Não acumulo cargo/função.
#
DECLARACAO_HTML = """
<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="UTF-8">
  <title>Declaração de Acúmulo de Cargo - Gerado</title>
  <style>
    @page { size: landscape; margin: 5mm; }
    body {
      font-family: Arial, sans-serif;
      margin: 0;
      padding: 5mm;
      font-size: 15px;
    }
    /* Primeira página */
    .header { text-align: center; border-bottom: 1px solid #000; padding-bottom: 3px; margin-bottom: 3px; }
    .header .municipio-logo img { max-height: 70px; }
    .header .title { text-transform: uppercase; font-weight: bold; font-size: 1.5em; margin-top: 5px; }
    .subheader { text-align: center; font-size: 0.95em; margin-top: 3px; margin-bottom: 5px; }
    .info-table, .horario-table, .cargo-table {
      width: 100%; border-collapse: collapse; margin-bottom: 5px;
    }
    .info-table td, .cargo-table td { border: 1px solid #000; padding: 2px; text-align: left; }
    /* Estrutura da tabela de horários com título "Atividade Extraclasse" */
    .horario-table th, .horario-table td {
      border: 1px solid #000; padding: 2px; text-align: center;
    }
    .signature { text-align: center; margin-top: 5px; }
    .page-break { page-break-before: always; }
  
    /* Botão de imprimir estilizado no canto superior direito */
    .print-button {
      position: fixed;
      top: 10px;
      right: 10px;
      z-index: 9999;
    }
    .print-button button {
      padding: 8px 12px;
      font-size: 14px;
      background-color: #007bff;
      color: #fff;
      border: none;
      border-radius: 4px;
      cursor: pointer;
    }
    .print-button button:hover {
      background-color: #0056b3;
    }
    @media print {
      .print-button { display: none; }
    }
  
    /* Segunda página: Declaração de Acúmulo (2 cópias) */
    .declaration-container {
      display: flex;
      justify-content: space-between;
      margin-top: 20px;
    }
    .declaration {
      width: 48%;
      border: 1px solid #000;
      padding: 10px;
      box-sizing: border-box;
      font-size: 14px;
    }
    /* Cabeçalho individual com logo menor, centralizado e data reposicionada */
    .decl-header { text-align: center; margin-bottom: 10px; }
    .decl-header .logo { text-align: center; }
    .decl-header .logo img { max-height: 20px; } /* Tamanho normal da logo na tela */
    .decl-header h2 { margin-top: 5px; text-decoration: underline; font-size: 1.2em; }
    .decl-header .declaration-date { text-align: left; margin-top: 5px; font-weight: bold; }
    .declaration p { margin: 5px 0; line-height: 1.4; }
    .section-label { font-weight: bold; margin-top: 10px; }
    .field-row { margin-left: 20px; }
    .field-row span { display: block; }
    /* Mantendo apenas os underscores literais */
    .line { display: inline-block; width: 30%; margin-left: 5px; }
    .footer { margin-top: 40px; text-align: center; }
    .signature-line { margin-top: 30px; border-top: 1px solid #000; width: 60%; margin: auto; }
    .note { font-size: 0.9em; margin-top: 10px; text-align: justify; }
  
    /* CSS para impressão: ajusta somente o logo da segunda página */
    @media print {
      .decl-header .logo img {
         max-height: 10px !important;
         display: block;
         margin: 0 auto;
      }
    }
  </style>
</head>
<body>
  <!-- Primeira página -->
  <div class="header">
    <div class="municipio-logo">
      <img src="/static/logos/cabecalho_municipio.png" alt="Logo do Município">
    </div>
    <div class="title">Declaração de Acúmulo de Cargo</div>
  </div>
  <div class="subheader">
    (Fundamentação Legal: Art. 37, XVI, “a” e “b” da Constituição Federal)
  </div>
  <table class="info-table">
    <tr>
      <td><strong>Unidade:</strong> E.M. JOSÉ PADIN MOUTA</td>
      <td><strong>Fone:</strong> (13) 3496-5321</td>
    </tr>
    <tr>
      <td><strong>Endereço:</strong> Rua Bororós, S/N - Vila Tupi</td>
      <td><strong>Cidade:</strong> Praia Grande - SP</td>
    </tr>
    <tr>
      <td><strong>Nome:</strong> {{ nome }}</td>
      <td><strong>RG:</strong> {{ rg }}</td>
    </tr>
    <tr>
      <td><strong>Cargo:</strong> {{ cargo }}</td>
      <td><strong>Disciplina:</strong> {{ disciplina }}</td>
    </tr>
  </table>
  
  <h3>Horário de Trabalho</h3>
  <table class="horario-table">
    <tr>
      <th rowspan="2">Dia da Semana</th>
      <th rowspan="2">Carga Horária c/ Aluno</th>
      <th colspan="2">Atividade Extraclasse</th>
    </tr>
    <tr>
      <th>Hora de Trabalho Pedagógico Coletivo</th>
      <th>Hora de Trabalho Pedagógico Individual</th>
    </tr>
    <tr>
      <td>Segunda-Feira</td>
      <td>{{ mon_aluno_ini }} - {{ mon_aluno_fim }}</td>
      <td>{{ mon_htpc_ini }} - {{ mon_htpc_fim }}</td>
      <td>{{ mon_hti_ini }} - {{ mon_hti_fim }}</td>
    </tr>
    <tr>
      <td>Terça-Feira</td>
      <td>{{ tue_aluno_ini }} - {{ tue_aluno_fim }}</td>
      <td>{{ tue_htpc_ini }} - {{ tue_htpc_fim }}</td>
      <td>{{ tue_hti_ini }} - {{ tue_hti_fim }}</td>
    </tr>
    <tr>
      <td>Quarta-Feira</td>
      <td>{{ wed_aluno_ini }} - {{ wed_aluno_fim }}</td>
      <td>{{ wed_htpc_ini }} - {{ wed_htpc_fim }}</td>
      <td>{{ wed_hti_ini }} - {{ wed_hti_fim }}</td>
    </tr>
    <tr>
      <td>Quinta-Feira</td>
      <td>{{ thu_aluno_ini }} - {{ thu_aluno_fim }}</td>
      <td>{{ thu_htpc_ini }} - {{ thu_htpc_fim }}</td>
      <td>{{ thu_hti_ini }} - {{ thu_hti_fim }}</td>
    </tr>
    <tr>
      <td>Sexta-Feira</td>
      <td>{{ fri_aluno_ini }} - {{ fri_aluno_fim }}</td>
      <td>{{ fri_htpc_ini }} - {{ fri_htpc_fim }}</td>
      <td>{{ fri_hti_ini }} - {{ fri_hti_fim }}</td>
    </tr>
  </table>
  
  <table class="cargo-table">
    <tr>
      <td><strong>Carga Horária c/ Aluno</strong></td>
      <td>{{ aluno }} horas</td>
    </tr>
    <tr>
      <td><strong>Hora de Trabalho Pedagógico Individual (Presencial e On-line)</strong></td>
      <td>{{ htpi_total }} horas</td>
    </tr>
    <tr>
      <td><strong>Hora de Trabalho Pedagógico Coletivo</strong></td>
      <td>{{ htpc }} horas</td>
    </tr>
    <tr>
      <td><strong>Hora de Atividade Livre (HAL)</strong></td>
      <td>{{ hal }} horas</td>
    </tr>
    <tr>
      <td><strong>Carga Horária Total</strong></td>
      <td>{{ total }} horas</td>
    </tr>
  </table>
  
  <p>
    * A Hora de Atividade Livre – HAL é o período remunerado no qual o docente deve cumpri-la em hora e local de livre.
  </p>
  <p>
    * {{ htpi_online }} Horas de Trabalho Pedagógico Individual – HTPI serão realizados de forma on-line, portanto, não haverá determinação de horário para cumprimento.
  </p>
  <p>
    Declaro, sob pena de responsabilidade, que as informações constantes desta declaração representam a verdade.
  </p>
  <p>
    Praia Grande, {{ data_atual }}.
  </p>
  <div class="signature">
    ______________________________________________<br>
    Assinatura e Carimbo do Diretor(a) da Unidade Escolar
  </div>
  
  <!-- Botão de imprimir estilizado no canto superior direito -->
  <div class="print-button">
    <button onclick="window.print()">Imprimir</button>
  </div>
  
  <!-- Segunda página: Declaração de Acúmulo (2 cópias lado a lado) -->
  <div class="page-break"></div>
  <div class="declaration-container">
    <!-- Cópia 1 -->
    <div class="declaration">
      <div class="decl-header logo">
        <img src="/static/logos/cabecalho_municipio.png" alt="Logo do Município">
        <h2>Declaração de Acúmulo</h2>
        <p class="declaration-date"><strong>Praia Grande, {{ data_atual }}.</strong></p>
      </div>
      <p>Eu, <strong>{{ nome }}</strong>, R.G. <strong>{{ rg }}</strong>, (Cargo/Função que exerce: <strong>{{ cargo }}</strong>), lotado(a) na <strong>E.M. JOSÉ PADIN MOUTA</strong>, declaro para todos os fins que:</p>
      
      <p><strong>(       ) Acumulo cargo/função:</strong></p>
      
      <p class="section-label">Federal: <span style="margin-left:10px;">(       )</span></p>
      <div class="field-row">
        <span class="field-label">Local:</span> <span class="line">______________________________</span>
      </div>
      <div class="field-row">
        <span class="field-label">Cargo/Função:</span> <span class="line">______________________________</span>
      </div>
      
      <p class="section-label">Estadual: <span style="margin-left:10px;">(       )</span></p>
      <div class="field-row">
        <span class="field-label">Local:</span> <span class="line">______________________________</span>
      </div>
      <div class="field-row">
        <span class="field-label">Cargo/Função:</span> <span class="line">______________________________</span>
      </div>
      
      <p class="section-label">Municipal: <span style="margin-left:10px;">(       )</span></p>
      <div class="field-row">
        <span class="field-label">Local:</span> <span class="line">______________________________</span>
      </div>
      <div class="field-row">
        <span class="field-label">Cargo/Função:</span> <span class="line">______________________________</span>
      </div>
      
      <p><strong>(       ) Não acumulo cargo/função.</strong></p>
      <p>Por ser verdade, firmo o presente.</p>
      <div class="footer">
        <div class="signature-line"></div>
        <p>Assinatura</p>
      </div>
      <p class="note">
        Nota: Em caso de acúmulo de cargo/função, anexar declaração de horário do outro local de trabalho assinado pelo responsável e xérox do D.O., exceto o aposentado.
      </p>
    </div>
    
    <!-- Cópia 2 -->
    <div class="declaration">
      <div class="decl-header logo">
        <img src="/static/logos/cabecalho_municipio.png" alt="Logo do Município">
        <h2>Declaração de Acúmulo</h2>
        <p class="declaration-date"><strong>Praia Grande, {{ data_atual }}.</strong></p>
      </div>
      <p>Eu, <strong>{{ nome }}</strong>, R.G. <strong>{{ rg }}</strong>, (Cargo/Função que exerce: <strong>{{ cargo }}</strong>), lotado(a) na <strong>E.M. JOSÉ PADIN MOUTA</strong>, declaro para todos os fins que:</p>
      
      <p><strong>(       ) Acumulo cargo/função:</strong></p>
      
      <p class="section-label">Federal: <span style="margin-left:10px;">(       )</span></p>
      <div class="field-row">
        <span class="field-label">Local:</span> <span class="line">______________________________</span>
      </div>
      <div class="field-row">
        <span class="field-label">Cargo/Função:</span> <span class="line">______________________________</span>
      </div>
      
      <p class="section-label">Estadual: <span style="margin-left:10px;">(       )</span></p>
      <div class="field-row">
        <span class="field-label">Local:</span> <span class="line">______________________________</span>
      </div>
      <div class="field-row">
        <span class="field-label">Cargo/Função:</span> <span class="line">______________________________</span>
      </div>
      
      <p class="section-label">Municipal: <span style="margin-left:10px;">(       )</span></p>
      <div class="field-row">
        <span class="field-label">Local:</span> <span class="line">______________________________</span>
      </div>
      <div class="field-row">
        <span class="field-label">Cargo/Função:</span> <span class="line">______________________________</span>
      </div>
      
      <p><strong>(       ) Não acumulo cargo/função.</strong></p>
      <p>Por ser verdade, firmo o presente.</p>
      <div class="footer">
        <div class="signature-line"></div>
        <p>Assinatura</p>
      </div>
      <p class="note">
        Nota: Em caso de acúmulo de cargo/função, anexar declaração de horário do outro local de trabalho assinado pelo responsável e xérox do D.O., exceto o aposentado.
      </p>
    </div>
  </div>
</body>
</html>
"""

from flask import Flask, request
from datetime import datetime


@app.route("/acumulo", methods=["GET", "POST"])
def acumulo():
    if request.method == "POST":
        # Dados básicos
        nome = request.form.get("nome")
        rg = request.form.get("rg")
        cargo = request.form.get("cargo")
        disciplina = request.form.get("disciplina", "")

        # Horários (não obrigatórios)
        mon_aluno_ini = request.form.get("mon_aluno_ini", "")
        mon_aluno_fim = request.form.get("mon_aluno_fim", "")
        mon_htpc_ini  = request.form.get("mon_htpc_ini", "")
        mon_htpc_fim  = request.form.get("mon_htpc_fim", "")
        mon_hti_ini   = request.form.get("mon_hti_ini", "")
        mon_hti_fim   = request.form.get("mon_hti_fim", "")

        tue_aluno_ini = request.form.get("tue_aluno_ini", "")
        tue_aluno_fim = request.form.get("tue_aluno_fim", "")
        tue_htpc_ini  = request.form.get("tue_htpc_ini", "")
        tue_htpc_fim  = request.form.get("tue_htpc_fim", "")
        tue_hti_ini   = request.form.get("tue_hti_ini", "")
        tue_hti_fim   = request.form.get("tue_hti_fim", "")

        wed_aluno_ini = request.form.get("wed_aluno_ini", "")
        wed_aluno_fim = request.form.get("wed_aluno_fim", "")
        wed_htpc_ini  = request.form.get("wed_htpc_ini", "")
        wed_htpc_fim  = request.form.get("wed_htpc_fim", "")
        wed_hti_ini   = request.form.get("wed_hti_ini", "")
        wed_hti_fim   = request.form.get("wed_hti_fim", "")

        thu_aluno_ini = request.form.get("thu_aluno_ini", "")
        thu_aluno_fim = request.form.get("thu_aluno_fim", "")
        thu_htpc_ini  = request.form.get("thu_htpc_ini", "")
        thu_htpc_fim  = request.form.get("thu_htpc_fim", "")
        thu_hti_ini   = request.form.get("thu_hti_ini", "")
        thu_hti_fim   = request.form.get("thu_hti_fim", "")

        fri_aluno_ini = request.form.get("fri_aluno_ini", "")
        fri_aluno_fim = request.form.get("fri_aluno_fim", "")
        fri_htpc_ini  = request.form.get("fri_htpc_ini", "")
        fri_htpc_fim  = request.form.get("fri_htpc_fim", "")
        fri_hti_ini   = request.form.get("fri_hti_ini", "")
        fri_hti_fim   = request.form.get("fri_hti_fim", "")

        # Propaga os horários "Com Aluno" de Segunda-Feira para os demais dias se o cargo for aplicável.
        if cargo.upper() in ["PROFESSOR RECREACIONISTA", "PROFESSOR ADJUNTO I", "PROFESSOR I", "PROFESSOR II"] and mon_aluno_ini and mon_aluno_fim:
            tue_aluno_ini = wed_aluno_ini = thu_aluno_ini = fri_aluno_ini = mon_aluno_ini
            tue_aluno_fim = wed_aluno_fim = thu_aluno_fim = fri_aluno_fim = mon_aluno_fim

        # Função para formatar o horário (ex: "13:30" -> "13h30")
        def format_time(t):
            return t.replace(":", "h") if t else t

        # Formatar todos os horários
        mon_aluno_ini = format_time(mon_aluno_ini)
        mon_aluno_fim = format_time(mon_aluno_fim)
        mon_htpc_ini  = format_time(mon_htpc_ini)
        mon_htpc_fim  = format_time(mon_htpc_fim)
        mon_hti_ini   = format_time(mon_hti_ini)
        mon_hti_fim   = format_time(mon_hti_fim)

        tue_aluno_ini = format_time(tue_aluno_ini)
        tue_aluno_fim = format_time(tue_aluno_fim)
        tue_htpc_ini  = format_time(tue_htpc_ini)
        tue_htpc_fim  = format_time(tue_htpc_fim)
        tue_hti_ini   = format_time(tue_hti_ini)
        tue_hti_fim   = format_time(tue_hti_fim)

        wed_aluno_ini = format_time(wed_aluno_ini)
        wed_aluno_fim = format_time(wed_aluno_fim)
        wed_htpc_ini  = format_time(wed_htpc_ini)
        wed_htpc_fim  = format_time(wed_htpc_fim)
        wed_hti_ini   = format_time(wed_hti_ini)
        wed_hti_fim   = format_time(wed_hti_fim)

        thu_aluno_ini = format_time(thu_aluno_ini)
        thu_aluno_fim = format_time(thu_aluno_fim)
        thu_htpc_ini  = format_time(thu_htpc_ini)
        thu_htpc_fim  = format_time(thu_htpc_fim)
        thu_hti_ini   = format_time(thu_hti_ini)
        thu_hti_fim   = format_time(thu_hti_fim)

        fri_aluno_ini = format_time(fri_aluno_ini)
        fri_aluno_fim = format_time(fri_aluno_fim)
        fri_htpc_ini  = format_time(fri_htpc_ini)
        fri_htpc_fim  = format_time(fri_htpc_fim)
        fri_hti_ini   = format_time(fri_hti_ini)
        fri_hti_fim   = format_time(fri_hti_fim)

        # Busca os dados do cargo
        cargo_info = CARGOS.get(cargo.upper())
        if not cargo_info:
            return f"Cargo '{cargo}' não encontrado na tabela de horas."

        total = cargo_info["total"]
        aluno = cargo_info["aluno"]
        htpi_pres = cargo_info["htpi_pres"]
        htpi_online = cargo_info["htpi_online"]
        htpi_total = cargo_info["htpi_total"]
        htpc = cargo_info["htpc"]
        hal = cargo_info["hal"]

        # Data por extenso
        hoje = datetime.now()
        meses = [
            "janeiro","fevereiro","março","abril","maio","junho",
            "julho","agosto","setembro","outubro","novembro","dezembro"
        ]
        data_atual = f"{hoje.day} de {meses[hoje.month-1]} de {hoje.year}"

        # Monta o HTML final substituindo variáveis
        html_final = (DECLARACAO_HTML
            .replace("{{ nome }}", nome)
            .replace("{{ rg }}", rg)
            .replace("{{ cargo }}", cargo)
            .replace("{{ disciplina }}", disciplina)
            .replace("{{ total }}", str(total))
            .replace("{{ aluno }}", str(aluno))
            .replace("{{ htpi_total }}", str(htpi_total))
            .replace("{{ htpc }}", str(htpc))
            .replace("{{ hal }}", str(hal))
            .replace("{{ htpi_online }}", str(htpi_online))
            .replace("{{ data_atual }}", data_atual)
  
            .replace("{{ mon_aluno_ini }}", mon_aluno_ini)
            .replace("{{ mon_aluno_fim }}", mon_aluno_fim)
            .replace("{{ mon_htpc_ini }}", mon_htpc_ini)
            .replace("{{ mon_htpc_fim }}", mon_htpc_fim)
            .replace("{{ mon_hti_ini }}", mon_hti_ini)
            .replace("{{ mon_hti_fim }}", mon_hti_fim)
  
            .replace("{{ tue_aluno_ini }}", tue_aluno_ini)
            .replace("{{ tue_aluno_fim }}", tue_aluno_fim)
            .replace("{{ tue_htpc_ini }}", tue_htpc_ini)
            .replace("{{ tue_htpc_fim }}", tue_htpc_fim)
            .replace("{{ tue_hti_ini }}", tue_hti_ini)
            .replace("{{ tue_hti_fim }}", tue_hti_fim)
  
            .replace("{{ wed_aluno_ini }}", wed_aluno_ini)
            .replace("{{ wed_aluno_fim }}", wed_aluno_fim)
            .replace("{{ wed_htpc_ini }}", wed_htpc_ini)
            .replace("{{ wed_htpc_fim }}", wed_htpc_fim)
            .replace("{{ wed_hti_ini }}", wed_hti_ini)
            .replace("{{ wed_hti_fim }}", wed_hti_fim)
  
            .replace("{{ thu_aluno_ini }}", thu_aluno_ini)
            .replace("{{ thu_aluno_fim }}", thu_aluno_fim)
            .replace("{{ thu_htpc_ini }}", thu_htpc_ini)
            .replace("{{ thu_htpc_fim }}", thu_htpc_fim)
            .replace("{{ thu_hti_ini }}", thu_hti_ini)
            .replace("{{ thu_hti_fim }}", thu_hti_fim)
  
            .replace("{{ fri_aluno_ini }}", fri_aluno_ini)
            .replace("{{ fri_aluno_fim }}", fri_aluno_fim)
            .replace("{{ fri_htpc_ini }}", fri_htpc_ini)
            .replace("{{ fri_htpc_fim }}", fri_htpc_fim)
            .replace("{{ fri_hti_ini }}", fri_hti_ini)
            .replace("{{ fri_hti_fim }}", fri_hti_fim)
        )
  
        return html_final
    else:
        return FORM_HTML
    
# Função auxiliar para avaliar se o valor do Plano de Ação é considerado válido
def is_valid_plano(val):
    if val is None:
        return False
    s = str(val).strip()
    return s not in ["", "-", "0", "#REF"]

@app.route('/quantinclusao', methods=["GET", "POST"])
def quantinclusao():
    if request.method == "POST":
        # Obter os arquivos enviados
        reg_file = request.files.get("lista_regular")
        eja_file = request.files.get("lista_eja")
        # Obter o responsável pelo preenchimento
        responsavel = request.form.get("responsavel")
        if not reg_file or reg_file.filename == "":
            flash("Selecione o arquivo da Lista Piloto Regular.", "error")
            return redirect(url_for('quantinclusao'))
        if not eja_file or eja_file.filename == "":
            flash("Selecione o arquivo da Lista Piloto EJA.", "error")
            return redirect(url_for('quantinclusao'))
        if not responsavel or responsavel.strip() == "":
            flash("Informe o Responsável pelo preenchimento.", "error")
            return redirect(url_for('quantinclusao'))
        
        # Salva os arquivos na pasta definida
        reg_filename = secure_filename(f"regular_{uuid.uuid4().hex}_{reg_file.filename}")
        eja_filename = secure_filename(f"eja_{uuid.uuid4().hex}_{eja_file.filename}")
        reg_path = os.path.join(app.config['UPLOAD_FOLDER'], reg_filename)
        eja_path = os.path.join(app.config['UPLOAD_FOLDER'], eja_filename)
        reg_file.save(reg_path)
        eja_file.save(eja_path)
        
        # Carrega o arquivo Regular – espera que contenha as abas "Total de Alunos" e "LISTA CORRIDA"
        try:
            wb_reg = load_workbook(reg_path, data_only=True)
            ws_total_reg = wb_reg["Total de Alunos"]
            ws_lista_reg = wb_reg["LISTA CORRIDA"]
        except Exception as e:
            flash(f"Erro ao ler o arquivo Regular: {str(e)}", "error")
            return redirect(url_for('quantinclusao'))
        
        # Mapeamento do arquivo Regular (células da aba "Total de Alunos")
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
            "H45": ws_total_reg["O24"].value,   # Corrigido: H45 recebe O24
            "L13": ws_total_reg["O25"].value,
            "L17": ws_total_reg["O26"].value,
            "L21": ws_total_reg["O28"].value,
            "L25": ws_total_reg["O29"].value,
            "L29": ws_total_reg["O30"].value,
            "L33": ws_total_reg["O31"].value,
            "L37": ws_total_reg["O32"].value,
            "L41": ws_total_reg["O33"].value,
            "L45": ws_total_reg["O34"].value
        }
        
        # Inicializa os contadores para as séries baseados na aba "LISTA CORRIDA" do Regular
        count_2A = count_2B = count_2C = count_2D = count_2E = count_2F = 0
        count_3A = count_3B = count_3C = count_3D = count_3E = count_3F = 0
        count_4A = count_4B = count_4C = count_4D = count_4E = count_4F = count_4G = 0
        count_5A = count_5B = count_5C = count_5D = count_5E = count_5F = count_5G = 0

        # Dicionário para armazenar conjuntos de nomes únicos por série (LISTA CORRIDA do Regular)
        series_list = ["2ºA", "2ºB", "2ºC", "2ºD", "2ºE", "2ºF", "2ºG",
                       "3ºA", "3ºB", "3ºC", "3ºD", "3ºE", "3ºF",
                       "4ºA", "4ºB", "4ºC", "4ºD", "4ºE", "4ºF", "4ºG",
                       "5ºA", "5ºB", "5ºC", "5ºD", "5ºE", "5ºF", "5ºG"]
        unique_names = {serie: set() for serie in series_list}

        # Processa a aba "LISTA CORRIDA" do Regular
        for row in ws_lista_reg.iter_rows(min_row=2, values_only=True):
            serie = str(row[0]).strip() if row[0] is not None else ""
            plano = row[15] if len(row) > 15 else None

            # Armazena nomes únicos (coluna P) por série
            if serie in unique_names and is_valid_plano(plano):
                unique_names[serie].add(str(plano).strip())

            # Contagem simples conforme série
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
        
        # Carrega o arquivo EJA – espera que contenha a aba "Total de Alunos"
        try:
            wb_eja = load_workbook(eja_path, data_only=True)
            ws_total_eja = wb_eja["Total de Alunos"]
        except Exception as e:
            flash(f"Erro ao ler o arquivo EJA: {str(e)}", "error")
            return redirect(url_for('quantinclusao'))
        
        eja_map = {
            "D53": ws_total_eja["M10"].value,
            "D57": (ws_total_eja["M11"].value or 0) + (ws_total_eja["M12"].value or 0),
            "D61": ws_total_eja["M13"].value,
            "H53": ws_total_eja["M14"].value,
            "H57": ws_total_eja["M16"].value,
            "H61": ws_total_eja["M17"].value,
            "L53": ws_total_eja["M18"].value
        }
        
        # Tenta carregar a aba "LISTA CORRIDA" do EJA para as novas regras
        try:
            ws_lista_eja = wb_eja["LISTA CORRIDA"]
        except Exception as e:
            flash(f"Erro ao ler a aba LISTA CORRIDA no arquivo EJA: {str(e)}", "error")
            return redirect(url_for('quantinclusao'))
        
        # --- NOVAS REGRAS DE PREENCHIMENTO (LISTA CORRIDA EJA) ---
        # H41 do modelo receberá o valor da célula O23 da aba "Total de Alunos" do Regular
        # Definição dos grupos de séries para os cálculos
        series_ef_group1 = {"1ª SÉRIE E.F", "2ª SÉRIE E.F", "3ª SÉRIE E.F", "4ª SÉRIE E.F"}
        series_ef_group2 = {"5ª SÉRIE E.F", "6ª SÉRIE E.F"}
        series_ef_group3 = {"7ª SÉRIE E.F"}
        series_ef_group4 = {"8ª SÉRIE E.F"}
        series_em_group1 = {"1ª SÉRIE E.M"}
        series_em_group2 = {"2ª SÉRIE E.M"}
        series_em_group3 = {"3ª SÉRIE E.M"}
        
        # Inicializa contadores e conjuntos para os cálculos
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
        
        # Itera sobre as linhas da aba "LISTA CORRIDA" do EJA
        for row in ws_lista_eja.iter_rows(min_row=2, values_only=True):
            serie = str(row[0]).strip() if row[0] is not None else ""
            # Coluna T é o índice 19 (A=0, ..., T=19)
            nome = row[19] if len(row) > 19 else None

            # Grupo 1 (D54 e D55): Séries E.F de 1ª a 4ª
            if serie in series_ef_group1:
                if is_valid_plano(nome):
                    d54_count += 1
                    unique_d55.add(str(nome).strip())
            # Grupo 2 (D58 e D59): Séries E.F de 5ª e 6ª
            if serie in series_ef_group2:
                if is_valid_plano(nome):
                    d58_count += 1
                    unique_d59.add(str(nome).strip())
            # Grupo 3 (D62 e D63): Série E.F de 7ª
            if serie in series_ef_group3:
                if is_valid_plano(nome):
                    d62_count += 1
                    unique_d63.add(str(nome).strip())
            # Grupo 4 (H54 e H55): Série E.F de 8ª
            if serie in series_ef_group4:
                if is_valid_plano(nome):
                    h54_count += 1
                    unique_h55.add(str(nome).strip())
            # Grupo 5 (H58 e H59): Série E.M de 1ª
            if serie in series_em_group1:
                if is_valid_plano(nome):
                    h58_count += 1
                    unique_h59.add(str(nome).strip())
            # Grupo 6 (H62 e H63): Série E.M de 2ª
            if serie in series_em_group2:
                if is_valid_plano(nome):
                    h62_count += 1
                    unique_h63.add(str(nome).strip())
            # Grupo 7 (L54 e L55): Série E.M de 3ª
            if serie in series_em_group3:
                if is_valid_plano(nome):
                    l54_count += 1
                    unique_l55.add(str(nome).strip())
        # --- FIM DAS NOVAS REGRAS (LISTA CORRIDA EJA) ---
        
        # Carrega o modelo "Quadro Quantitativo de Inclusão - Modelo.xlsx" da pasta "modelos"
        model_path = os.path.join("modelos", "Quadro Quantitativo de Inclusão - Modelo.xlsx")
        try:
            wb_model = load_workbook(model_path)
        except Exception as e:
            flash(f"Erro ao abrir o modelo de inclusão: {str(e)}", "error")
            return redirect(url_for('quantinclusao'))
        ws_model = wb_model.active
        
        # Preenche os valores fixos do Regular no modelo
        for cell_addr, value in reg_map.items():
            ws_model[cell_addr] = value
        
        # Preenche os valores do EJA no modelo
        for cell_addr, value in eja_map.items():
            ws_model[cell_addr] = value
        
        # Preenche as células de contagem simples (Regular)
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

        # Preenche as células de nomes únicos (Regular)
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
        
        # --- NOVOS PREENCHIMENTOS USANDO DADOS DO EJA ---
        # H41 do modelo recebe O23 da aba "Total de Alunos" do Regular
        ws_model["H41"] = ws_total_reg["O23"].value
        
        # Preenchimentos para a aba LISTA CORRIDA do EJA:
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
        # --- FIM DOS NOVOS PREENCHIMENTOS (LISTA CORRIDA EJA) ---
        
        # --- NOVAS INFORMAÇÕES ADICIONAIS ---
        # Preencher B4 com o mês e ano atual (ex.: ABRIL/2025),
        # C8 com o nome do responsável pelo preenchimento,
        # e K8 com a data atual (ex.: 07/04/2025).
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
            12: "DEZEMBRO"
        }
        current_date = datetime.now()
        ws_model["B4"] = f"{meses[current_date.month]}/{current_date.year}"
        ws_model["C8"] = responsavel.strip()
        ws_model["K8"] = current_date.strftime("%d/%m/%Y")
        # --- FIM DAS INFORMAÇÕES ADICIONAIS ---
        
        # Salva o modelo preenchido em um buffer e envia para download
        output = BytesIO()
        wb_model.save(output)
        output.seek(0)
        filename = f"Quadro_Quantitativo_de_Inclusao_{datetime.now().strftime('%d%m%Y')}.xlsx"
        return send_file(output, as_attachment=True, download_name=filename,
                         mimetype="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    else:
        # Formulário HTML com campo adicional para o responsável pelo preenchimento
        form_html = """
        <!doctype html>
        <html lang="pt-br">
          <head>
            <meta charset="UTF-8">
            <title>Quadro Quantitativo de Inclusão</title>
            <link href="https://fonts.googleapis.com/css2?family=Montserrat:wght@400;600&display=swap" rel="stylesheet">
            <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/bootstrap/4.5.2/css/bootstrap.min.css">
            <style>
              body {
                background: #eef2f3;
                font-family: 'Montserrat', sans-serif;
                margin: 0;
                padding: 0;
                display: flex;
                flex-direction: column;
                min-height: 100vh;
              }
              header {
                background: linear-gradient(90deg, #283E51, #4B79A1);
                color: #fff;
                padding: 20px;
                text-align: center;
                border-bottom: 3px solid #1d2d3a;
                border-radius: 0 0 15px 15px;
                box-shadow: 0 4px 6px rgba(0,0,0,0.1);
              }
              main {
                flex: 1;
                display: flex;
                align-items: center;
                justify-content: center;
                padding: 20px;
              }
              .container {
                background: #fff;
                padding: 20px;
                margin-top: 30px;
                border-radius: 8px;
                max-width: 700px;
                width: 100%;
              }
              h1 {
                margin-bottom: 20px;
              }
              footer {
                background-color: #424242;
                color: #fff;
                text-align: center;
                padding: 10px;
              }
            </style>
          </head>
          <body>
            <header>
              <h1>E.M José Padin Mouta - Quadro Quantitativo de Inclusão</h1>
            </header>
            <main>
              <div class="container">
                <h1>Upload das Listas Piloto</h1>
                <form method="POST" enctype="multipart/form-data">
                  <div class="form-group">
                    <label for="responsavel">Responsável pelo preenchimento:</label>
                    <input type="text" class="form-control" name="responsavel" id="responsavel" placeholder="Digite seu nome" required>
                  </div>
                  <div class="form-group">
                    <label for="lista_regular">Lista Piloto - REGULAR (Excel):</label>
                    <input type="file" class="form-control-file" name="lista_regular" id="lista_regular" accept=".xlsx, .xls" required>
                  </div>
                  <div class="form-group">
                    <label for="lista_eja">Lista Piloto - EJA (Excel):</label>
                    <input type="file" class="form-control-file" name="lista_eja" id="lista_eja" accept=".xlsx, .xls" required>
                  </div>
                  <button type="submit" class="btn btn-primary">Gerar Quadro de Inclusão</button>
                </form>
              </div>
            </main>
            <footer>
              Desenvolvido por Nilson Cruz © 2025. Todos os direitos reservados.
            </footer>
          </body>
        </html>
        """
        return render_template_string(form_html)

if __name__ == "__main__":
    app.run(debug=True)
