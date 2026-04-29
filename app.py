import os
import logging
import socket
import tempfile
import pandas as pd
import io
import re
import json
import uuid
import gc
import pdfplumber
import psycopg2
import psycopg2.extras
from datetime import datetime
from functools import wraps
from flask import Flask, render_template, jsonify, request, redirect, url_for, send_file, flash, session
from werkzeug.utils import secure_filename
from werkzeug.security import check_password_hash, generate_password_hash
from rapidfuzz import process, fuzz
from sqlalchemy import create_engine, text as sa_text

# Railway não suporta IPv6 — força resolução IPv4 em todas as conexões de rede
_orig_getaddrinfo = socket.getaddrinfo
def _getaddrinfo_ipv4(*args, **kwargs):
    responses = _orig_getaddrinfo(*args, **kwargs)
    ipv4 = [r for r in responses if r[0] == socket.AF_INET]
    return ipv4 if ipv4 else responses
socket.getaddrinfo = _getaddrinfo_ipv4

logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s %(levelname)s %(name)s: %(message)s'
)
logger = logging.getLogger(__name__)

# --- Configuração ---
app = Flask(__name__)
app.secret_key = os.environ.get('SECRET_KEY', 'barril-dobrado-dev-key-mude-em-producao')

def _normalise_db_url(url: str) -> str:
    """Normaliza a DATABASE_URL para uso com psycopg2."""
    import re
    url = url.strip()
    # Remove prefixo acidental "DATABASE_URL=..." caso o usuário copie a linha inteira
    if not url.startswith(('postgresql://', 'postgres://')) and '=' in url:
        url = url.split('=', 1)[1].strip()
    url = url.replace('postgres://', 'postgresql://', 1)
    # Remove sslmode da URL — é passado via kwarg no connect()
    url = re.sub(r'[?&]sslmode=[^&]*', '', url).rstrip('?&')
    return url

DATABASE_URL = os.environ.get('DATABASE_URL', '')
_SA_DATABASE_URL = _normalise_db_url(DATABASE_URL) + ('&' if '?' in _normalise_db_url(DATABASE_URL) else '?') + 'sslmode=require' if DATABASE_URL else ''

_SEED_USERNAME = os.environ.get('APP_USERNAME', 'admin')
_SEED_PASSWORD = os.environ.get('APP_PASSWORD', 'barril2025')

app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32 MB (Excel VAR pode ser grande)
PDF_MAX_BYTES = 15 * 1024 * 1024  # 15 MB por PDF

LIMIAR_ALERTA = 0.90
MODULOS_PADRAO = [
    'TOTVS Educacional', 'TOTVS Folha de Pagamento', 'TOTVS Gestão Contábil',
    'TOTVS Gestão de Estoque, Compras e Faturamento', 'TOTVS Gestão de Pessoas',
    'TOTVS Gestão Financeira', 'TOTVS Gestão Fiscal', 'TOTVS Gestão Patrimonial',
    'TOTVS Inteligência de Negócios'
]
_MODULOS_SEM_ESPACO = [re.sub(r'\s+', '', m).lower() for m in MODULOS_PADRAO]


# --- Wrapper de conexão (interface sqlite3-like sobre psycopg2) ---
class DBConn:
    """Wraps psycopg2 connection com interface compatível com sqlite3."""

    def __init__(self, conn):
        self._conn = conn

    def execute(self, sql, params=None):
        """Usa RealDictCursor para acesso por nome de coluna."""
        sql = sql.replace('?', '%s')
        cur = self._conn.cursor(cursor_factory=psycopg2.extras.RealDictCursor)
        cur.execute(sql, params or ())
        return cur

    def commit(self):
        self._conn.commit()

    def close(self):
        self._conn.close()


def get_db_connection() -> DBConn:
    conn = psycopg2.connect(_normalise_db_url(DATABASE_URL), sslmode='require')
    return DBConn(conn)


def get_engine():
    """Engine SQLAlchemy para operações pandas (to_sql, read_sql_query)."""
    return create_engine(_SA_DATABASE_URL)


def _normalizar_modulo(raw):
    """Corrige nomes de módulo fragmentados pelo PDF (ex: 'TOT VS Educ acio nal' → 'TOTVS Educacional')."""
    if not raw:
        return raw
    sem_espaco = re.sub(r'\s+', '', raw).lower()
    try:
        idx = _MODULOS_SEM_ESPACO.index(sem_espaco)
        return MODULOS_PADRAO[idx]
    except ValueError:
        pass
    resultado = process.extractOne(sem_espaco, _MODULOS_SEM_ESPACO, scorer=fuzz.ratio, score_cutoff=70)
    if resultado:
        return MODULOS_PADRAO[_MODULOS_SEM_ESPACO.index(resultado[0])]
    return raw


def setup():
    """Cria tabelas no Postgres e faz seed do admin padrão."""
    if not DATABASE_URL:
        logger.error("DATABASE_URL não configurada. Configure a variável de ambiente.")
        return

    conn = psycopg2.connect(_normalise_db_url(DATABASE_URL), sslmode='require')
    cur = conn.cursor()

    cur.execute('''
        CREATE TABLE IF NOT EXISTS usuarios (
            id SERIAL PRIMARY KEY,
            username TEXT NOT NULL UNIQUE,
            password_hash TEXT NOT NULL,
            nome_completo TEXT NOT NULL DEFAULT '',
            role TEXT NOT NULL DEFAULT 'user',
            ativo SMALLINT NOT NULL DEFAULT 1,
            criado_em TEXT NOT NULL,
            ultimo_acesso TEXT,
            deve_trocar_senha SMALLINT NOT NULL DEFAULT 0
        )
    ''')
    # Migração: adiciona coluna se o banco já existia sem ela
    cur.execute('''
        ALTER TABLE usuarios ADD COLUMN IF NOT EXISTS
        deve_trocar_senha SMALLINT NOT NULL DEFAULT 0
    ''')

    # conteudo BYTEA armazena o arquivo Excel da base VAR no banco (sem depender de disco)
    cur.execute('''
        CREATE TABLE IF NOT EXISTS uploads_historico (
            id SERIAL PRIMARY KEY,
            nome_arquivo_original TEXT NOT NULL,
            nome_arquivo_salvo TEXT NOT NULL,
            timestamp TEXT NOT NULL,
            status TEXT NOT NULL,
            conteudo BYTEA
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS dados_var (
            "ID Funcionalidade" TEXT,
            "Funcionalidade" TEXT,
            "ID Módulo" TEXT,
            "Módulo" TEXT
        )
    ''')

    cur.execute('''
        CREATE TABLE IF NOT EXISTS atividades_log (
            id SERIAL PRIMARY KEY,
            usuario TEXT NOT NULL,
            tipo TEXT NOT NULL,
            descricao TEXT NOT NULL,
            timestamp TEXT NOT NULL
        )
    ''')

    # Substitui a pasta outputs/ — armazena resultados de comparação temporariamente
    cur.execute('''
        CREATE TABLE IF NOT EXISTS resultados_temp (
            result_id UUID PRIMARY KEY,
            dados JSONB NOT NULL,
            criado_em TIMESTAMPTZ DEFAULT NOW()
        )
    ''')

    # Seed: cria admin padrão se não existir
    # try/except evita race condition entre workers gunicorn tentando inserir ao mesmo tempo
    try:
        ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        ph = generate_password_hash(_SEED_PASSWORD, method='pbkdf2:sha256')
        cur.execute(
            '''INSERT INTO usuarios (username, password_hash, nome_completo, role, ativo, criado_em)
               VALUES (%s, %s, %s, %s, %s, %s)
               ON CONFLICT (username) DO NOTHING''',
            (_SEED_USERNAME, ph, 'Administrador', 'admin', 1, ts)
        )
    except Exception:
        pass

    conn.commit()
    cur.close()
    conn.close()


def registrar_atividade(tipo, descricao):
    try:
        conn = get_db_connection()
        ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        usuario = session.get('username', 'admin')
        conn.execute(
            'INSERT INTO atividades_log (usuario, tipo, descricao, timestamp) VALUES (?, ?, ?, ?)',
            (usuario, tipo, descricao, ts)
        )
        conn.commit()
        conn.close()
    except Exception:
        pass


@app.context_processor
def inject_user():
    return dict(
        current_user=session.get('username', ''),
        current_role=session.get('role', 'user'),
        session_login_time=session.get('login_time', '')
    )


setup()


def allowed_file(filename, allowed_extensions={'xlsx', 'xls', 'pdf'}):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in allowed_extensions


# --- AUTENTICAÇÃO ---
def require_login(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        if session.get('deve_trocar_senha'):
            return redirect(url_for('primeiro_acesso'))
        return f(*args, **kwargs)
    return decorated


def require_admin(f):
    @wraps(f)
    def decorated(*args, **kwargs):
        if not session.get('logged_in'):
            return redirect(url_for('login'))
        if session.get('role') != 'admin':
            flash('Acesso negado. Apenas administradores podem acessar esta área.', 'danger')
            return redirect(url_for('home'))
        return f(*args, **kwargs)
    return decorated


@app.route('/login', methods=['GET', 'POST'])
def login():
    if session.get('logged_in'):
        return redirect(url_for('home'))
    if request.method == 'POST':
        username = request.form.get('username', '').strip()
        password = request.form.get('password', '')
        conn = get_db_connection()
        user = conn.execute(
            'SELECT * FROM usuarios WHERE username = ? AND ativo = 1', (username,)
        ).fetchone()
        if user and check_password_hash(user['password_hash'], password):
            ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
            conn.execute('UPDATE usuarios SET ultimo_acesso = ? WHERE id = ?', (ts, user['id']))
            conn.commit()
            conn.close()
            session['logged_in'] = True
            session['username'] = username
            session['role'] = user['role']
            session['nome_completo'] = user['nome_completo'] or username
            session['login_time'] = ts
            session['deve_trocar_senha'] = bool(user['deve_trocar_senha'])
            if user['deve_trocar_senha']:
                return redirect(url_for('primeiro_acesso'))
            return redirect(url_for('home'))
        conn.close()
        flash('Usuário ou senha incorretos.', 'danger')
    return render_template('auth/login.html')


@app.route('/logout')
def logout():
    session.clear()
    return redirect(url_for('login'))


@app.route('/primeiro_acesso', methods=['GET', 'POST'])
def primeiro_acesso():
    if not session.get('logged_in'):
        return redirect(url_for('login'))
    if not session.get('deve_trocar_senha'):
        return redirect(url_for('home'))

    if request.method == 'POST':
        data = request.get_json()
        nova_senha = (data.get('nova_senha') or '').strip()

        if len(nova_senha) < 6:
            return jsonify({'success': False, 'message': 'A senha deve ter no mínimo 6 caracteres.'}), 400

        try:
            ph = generate_password_hash(nova_senha, method='pbkdf2:sha256')
            conn = get_db_connection()
            conn.execute(
                'UPDATE usuarios SET password_hash = ?, deve_trocar_senha = 0 WHERE username = ?',
                (ph, session['username'])
            )
            conn.commit()
            conn.close()
            session['deve_trocar_senha'] = False
            return jsonify({'success': True})
        except Exception as e:
            return jsonify({'success': False, 'message': str(e)}), 500

    return render_template('auth/primeiro_acesso.html')


# --- RESULTADOS TEMPORÁRIOS (substituem a pasta outputs/) ---
def salvar_resultado_temp(dados):
    result_id = str(uuid.uuid4())
    conn = get_db_connection()
    conn.execute(
        'INSERT INTO resultados_temp (result_id, dados) VALUES (?, ?)',
        (result_id, psycopg2.extras.Json(dados, dumps=lambda x: json.dumps(x, ensure_ascii=False, default=str)))
    )
    conn.commit()
    conn.close()
    return result_id


def carregar_resultado_temp(result_id):
    conn = get_db_connection()
    row = conn.execute(
        'SELECT dados FROM resultados_temp WHERE result_id = ?', (result_id,)
    ).fetchone()
    conn.close()
    # psycopg2 desserializa JSONB automaticamente para dict Python
    return row['dados'] if row else None


# --- EXTRAÇÃO DE PDF ---
def extrair_funcionalidades_pdf(filepath):
    try:
        funcionalidades = []
        regex_linha = re.compile(r"\[([\d\.]+)\]\s+(.+)$")

        with pdfplumber.open(filepath) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if text:
                    for linha in text.split('\n'):
                        linha = linha.strip()
                        if linha.startswith("Tabela:") or "The following table" in linha:
                            continue
                        match = regex_linha.search(linha)
                        if match:
                            codigo = match.group(1).strip()
                            nome_func = match.group(2).strip()
                            item_completo = f"[{codigo}] {nome_func}"
                            if len(nome_func) > 2:
                                funcionalidades.append(item_completo)

                page.flush_cache()
                del text
                if i % 5 == 0:
                    gc.collect()

        funcionalidades = list(dict.fromkeys(funcionalidades))
        return pd.DataFrame(funcionalidades, columns=['Funcionalidade'])

    except Exception as e:
        logger.exception("Erro ao ler PDF: %s", str(e))
        return pd.DataFrame()


# --- LÓGICA SoD ---
def consolidar_colunas(df, col_inicio, col_fim):
    try:
        col_fim_real = min(col_fim, len(df.columns))
        return df.iloc[:, col_inicio:col_fim_real].fillna('').astype(str).agg(''.join, axis=1)
    except Exception:
        return pd.Series([""] * len(df), index=df.index)


def analisar_riscos_excel(caminho_arquivo, cenario):
    try:
        df_bruto = pd.read_excel(caminho_arquivo, header=None, sheet_name=0)
    except Exception as e:
        return {'status': 'error', 'message': f"Erro Excel: {e}"}

    idx_relatorio = None
    idx_riscos = None

    col0 = df_bruto.iloc[:, 0].apply(lambda x: '' if pd.isna(x) else str(x))
    for idx, val in col0.items():
        if "Relatório da Análise de ticket do perfil" in val: idx_relatorio = idx
        if "Riscos SoD para perfil" in val: idx_riscos = idx

    if idx_riscos is None:
        return {'status': 'error', 'message': "Aba 'Riscos SoD' não encontrada."}
    if cenario == 'manutencao' and idx_relatorio is None:
        return {'status': 'error', 'message': "Aba 'Relatório' não encontrada."}

    perfil = "Desconhecido"
    if idx_relatorio is not None:
        try:
            m = re.search(r"perfil\s*(.*)", str(df_bruto.iloc[idx_relatorio, 0]), re.IGNORECASE)
            if m: perfil = m.group(1).strip().replace(":", "")
        except Exception:
            pass

    try:
        dados_t2 = df_bruto.iloc[idx_riscos + 2:].copy()
        df_riscos = pd.DataFrame({
            'ID Risco': consolidar_colunas(dados_t2, 0, 2),
            'Descrição Risco': consolidar_colunas(dados_t2, 2, 6),
            'Criticidade': consolidar_colunas(dados_t2, 6, 8),
            'Sistema': consolidar_colunas(dados_t2, 10, 12),
            'Módulo': consolidar_colunas(dados_t2, 12, 14),
            'Funcionalidade': consolidar_colunas(dados_t2, 16, 18),
            'Funcionalidade 2': consolidar_colunas(dados_t2, 20, 22)
        }).dropna(how='all')
    except Exception:
        return {'status': 'error', 'message': "Erro layout Riscos."}

    matriz = df_riscos[['ID Risco', 'Descrição Risco', 'Funcionalidade', 'Funcionalidade 2']].fillna('').to_dict('records')

    if cenario == 'criacao':
        if df_riscos.empty:
            return {'status': 'no_risks', 'perfil': perfil, 'escopo_analisado': [], 'matriz_referencia': matriz}

        modulo = _normalizar_modulo(df_riscos['Módulo'].dropna().iloc[0]) if 'Módulo' in df_riscos.columns and not df_riscos['Módulo'].dropna().empty else ''
        recs = df_riscos.to_dict('records')
        agrupado = {}
        for r in recs:
            k = r['Funcionalidade'] if r['Funcionalidade'] else "Geral"
            if k not in agrupado: agrupado[k] = []
            r['Conflito Com'] = r['Funcionalidade 2']
            agrupado[k].append(r)
        return {'status': 'success', 'data': agrupado, 'perfil': perfil, 'modulo': modulo, 'matriz_referencia': matriz, 'escopo_analisado': []}

    elif cenario == 'manutencao':
        try:
            dados_t1 = df_bruto.iloc[idx_relatorio + 2: idx_riscos - 1].copy()
            df_rel = pd.DataFrame({
                'Sistema': consolidar_colunas(dados_t1, 2, 4),
                'Funcionalidade': consolidar_colunas(dados_t1, 4, 8),
                'Status': consolidar_colunas(dados_t1, 8, 10)
            }).dropna(how='all')
        except Exception:
            return {'status': 'error', 'message': "Erro layout Funcionalidades."}

        df_add = df_rel[df_rel['Status'].astype(str).str.strip() == 'Adicionado']
        escopo = df_add[['Sistema', 'Funcionalidade']].fillna('').to_dict('records')

        if df_add.empty:
            return {'status': 'no_risks', 'message': "Sem itens 'Adicionado'.", 'perfil': perfil, 'escopo_analisado': escopo, 'matriz_referencia': matriz}

        funcs = set(
            str(f) for f in df_add['Funcionalidade'].unique()
            if pd.notna(f) and str(f).strip()
        )
        func_col = df_riscos['Funcionalidade'].fillna('').astype(str)
        func2_col = df_riscos['Funcionalidade 2'].fillna('').astype(str)
        match = df_riscos[func_col.isin(funcs) | func2_col.isin(funcs)].copy()

        if match.empty:
            return {'status': 'no_risks', 'perfil': perfil, 'escopo_analisado': escopo, 'matriz_referencia': matriz}

        modulo = match['Módulo'].dropna().iloc[0] if 'Módulo' in match.columns and not match['Módulo'].dropna().empty else ''
        agrupado = {}
        for r in match.to_dict('records'):
            f1 = str(r.get('Funcionalidade') or '')
            f2 = str(r.get('Funcionalidade 2') or '')
            gatilho, conflito = (f1, f2) if f1 in funcs else (f2, f1)

            if gatilho:
                if gatilho not in agrupado: agrupado[gatilho] = []
                agrupado[gatilho].append({
                    'ID Risco': r['ID Risco'], 'Descrição Risco': r['Descrição Risco'],
                    'Criticidade': r['Criticidade'], 'Sistema': r['Sistema'], 'Conflito Com': conflito
                })

        return {'status': 'success', 'data': agrupado, 'perfil': perfil, 'modulo': modulo, 'escopo_analisado': escopo, 'matriz_referencia': matriz}


# --- EXTRAÇÃO DE PERFIL PDF (formato TOTVS Funcionalidades do Perfil) ---
def extrair_sod_pdf(filepath, cenario):
    perfil = "Desconhecido"
    rel_rows = []
    risco_rows = []

    def clean(s):
        if s is None:
            return ''
        s = re.sub(r'\n+', ' ', str(s)).strip()
        s = re.sub(r'\s+', ' ', s).strip()
        tokens = s.split()
        if len(tokens) > 2 and sum(1 for t in tokens if len(t) == 1) / len(tokens) > 0.5:
            return ''.join(tokens)
        return s

    def find_col(header, *keywords):
        for i, h in enumerate(header):
            hl = h.lower()
            if any(k in hl for k in keywords):
                return i
        return None

    try:
        with pdfplumber.open(filepath) as pdf:
            for pi, page in enumerate(pdf.pages):
                text = page.extract_text() or ''

                if pi == 0:
                    m = re.search(r'Relatório da Análise de ticket do perfil\s+(\S+)', text, re.IGNORECASE)
                    if m:
                        perfil = m.group(1).strip()

                tables = page.extract_tables()
                for table in tables:
                    if not table or len(table) < 2:
                        continue

                    header = [clean(c) for c in table[0]]
                    hj = ' '.join(header).lower()
                    ncols = len(table[0]) if table[0] else 0

                    if ncols == 4 and 'funcionalidade' in hj and 'status' in hj:
                        fi  = find_col(header, 'funcionalidade')
                        si  = find_col(header, 'status')
                        sti = find_col(header, 'sistema')

                        for row in table[1:]:
                            if not row or all(not c for c in row):
                                continue
                            func   = clean(row[fi])  if fi  is not None and fi  < len(row) else ''
                            status = clean(row[si])  if si  is not None and si  < len(row) else ''
                            sis    = clean(row[sti]) if sti is not None and sti < len(row) else ''
                            if func:
                                rel_rows.append({'Sistema': sis, 'Funcionalidade': func, 'Status': status})

                    elif ncols >= 9:
                        for row in table[1:]:
                            if not row or len(row) < 9:
                                continue
                            def g(idx):
                                return clean(row[idx]) if idx < len(row) else ''

                            risco_id  = g(0)
                            descricao = g(1)
                            if not risco_id and not descricao:
                                continue
                            if risco_id and not re.match(r'SOD', risco_id, re.IGNORECASE):
                                continue

                            risco_rows.append({
                                'ID Risco':         risco_id,
                                'Descrição Risco':  descricao,
                                'Criticidade':      g(2),
                                'Sistema':          g(4),
                                'Módulo':           g(5),
                                'Funcionalidade':   g(7),
                                'Funcionalidade 2': g(9) if len(row) > 9 else '',
                            })

                page.flush_cache()

    except Exception as e:
        logger.exception("Erro ao extrair SoD PDF: %s", str(e))
        return {'status': 'error', 'message': f'Erro ao ler PDF: {str(e)}'}

    df_riscos = pd.DataFrame(risco_rows) if risco_rows else pd.DataFrame(
        columns=['ID Risco', 'Descrição Risco', 'Criticidade', 'Sistema', 'Módulo', 'Funcionalidade', 'Funcionalidade 2']
    )
    matriz = df_riscos[['ID Risco', 'Descrição Risco', 'Funcionalidade', 'Funcionalidade 2']].fillna('').to_dict('records')

    modulo = df_riscos['Módulo'].dropna().iloc[0] if 'Módulo' in df_riscos.columns and not df_riscos['Módulo'].dropna().empty else ''

    if cenario == 'criacao':
        if df_riscos.empty:
            return {'status': 'no_risks', 'perfil': perfil, 'modulo': modulo, 'escopo_analisado': [], 'matriz_referencia': []}

        agrupado = {}
        for _, r in df_riscos.iterrows():
            k = r['Funcionalidade'] or 'Geral'
            if k not in agrupado:
                agrupado[k] = []
            agrupado[k].append({
                'ID Risco': r['ID Risco'], 'Descrição Risco': r['Descrição Risco'],
                'Criticidade': r['Criticidade'], 'Sistema': r['Sistema'],
                'Conflito Com': r['Funcionalidade 2']
            })
        return {'status': 'success', 'data': agrupado, 'perfil': perfil, 'modulo': modulo,
                'matriz_referencia': matriz, 'escopo_analisado': []}

    df_rel = pd.DataFrame(rel_rows) if rel_rows else pd.DataFrame(columns=['Sistema', 'Funcionalidade', 'Status'])
    df_add = df_rel[df_rel['Status'].str.strip() == 'Adicionado']
    escopo = df_add[['Sistema', 'Funcionalidade']].fillna('').to_dict('records')

    if df_add.empty:
        return {'status': 'no_risks', 'message': "Sem itens 'Adicionado'.",
                'perfil': perfil, 'modulo': modulo, 'escopo_analisado': escopo, 'matriz_referencia': matriz}

    if df_riscos.empty:
        return {'status': 'no_risks', 'perfil': perfil, 'modulo': modulo, 'escopo_analisado': escopo, 'matriz_referencia': matriz}

    funcs = [str(f) for f in df_add['Funcionalidade'].dropna().unique() if not isinstance(f, float) and str(f).strip()]
    match = df_riscos[df_riscos['Funcionalidade'].isin(funcs) | df_riscos['Funcionalidade 2'].isin(funcs)].copy()

    if match.empty:
        return {'status': 'no_risks', 'perfil': perfil, 'modulo': modulo, 'escopo_analisado': escopo, 'matriz_referencia': matriz}

    modulo_match = _normalizar_modulo(match['Módulo'].dropna().iloc[0]) if 'Módulo' in match.columns and not match['Módulo'].dropna().empty else modulo
    agrupado = {}
    for _, r in match.iterrows():
        f1 = str(r['Funcionalidade']) if not isinstance(r['Funcionalidade'], float) else ''
        f2 = str(r['Funcionalidade 2']) if not isinstance(r['Funcionalidade 2'], float) else ''
        gatilho, conflito = (f1, f2) if f1 in funcs else (f2, f1)
        if gatilho:
            if gatilho not in agrupado:
                agrupado[gatilho] = []
            agrupado[gatilho].append({
                'ID Risco': r['ID Risco'], 'Descrição Risco': r['Descrição Risco'],
                'Criticidade': r['Criticidade'], 'Sistema': r['Sistema'],
                'Conflito Com': conflito
            })

    return {'status': 'success' if agrupado else 'no_risks', 'data': agrupado,
            'perfil': perfil, 'modulo': modulo_match, 'escopo_analisado': escopo, 'matriz_referencia': matriz}


def extrair_perfil_pdf(filepath):
    info = {'codperfil': '', 'modulo': '', 'emitido_por': '', 'data_hora': '', 'ambiente': ''}
    rows = []
    regex_row = re.compile(r'^([\d\.]+)\s+(.+?)\s+([\w_]+)\s+(True|False)\s*$')
    regex_modulo = re.compile(r'Módulo[:\s]+(.+?)(?:\s+Ambiente.*)?$', re.IGNORECASE)
    regex_emitido = re.compile(r'Emitido por[:\s]+(\S+)', re.IGNORECASE)
    regex_data = re.compile(r'(\d{2}/\d{2}/\d{4}\s+\d{2}:\d{2})')
    regex_ambiente = re.compile(r'^(?:RM\s+)?(\w+)\s+Módulo:', re.IGNORECASE)

    try:
        with pdfplumber.open(filepath) as pdf:
            for i, page in enumerate(pdf.pages):
                text = page.extract_text()
                if not text:
                    page.flush_cache()
                    continue

                for linha in text.split('\n'):
                    linha = linha.strip()
                    if not linha:
                        continue

                    if not info['ambiente']:
                        m = regex_ambiente.match(linha)
                        if m:
                            info['ambiente'] = m.group(1).strip()
                    if not info['modulo']:
                        m = regex_modulo.search(linha)
                        if m:
                            val = re.sub(r'\s+Ambiente.*$', '', m.group(1), flags=re.IGNORECASE).strip()
                            info['modulo'] = val
                    if not info['emitido_por']:
                        m = regex_emitido.search(linha)
                        if m:
                            info['emitido_por'] = m.group(1).strip()
                    if not info['data_hora']:
                        m = regex_data.search(linha)
                        if m:
                            info['data_hora'] = m.group(1).strip()

                    m = regex_row.match(linha)
                    if m:
                        codigo = m.group(1).strip()
                        funcionalidade = m.group(2).strip()
                        codperfil = m.group(3).strip()
                        permitido = m.group(4).strip()
                        if len(funcionalidade) > 1:
                            if not info['codperfil'] and codperfil:
                                info['codperfil'] = codperfil
                            rows.append({
                                'codigo': codigo,
                                'funcionalidade': funcionalidade,
                                'codperfil': codperfil,
                                'permitido': permitido
                            })

                page.flush_cache()
                del text
                if i % 5 == 0:
                    gc.collect()

    except Exception as e:
        logger.exception("Erro ao extrair perfil PDF: %s", str(e))

    df = pd.DataFrame(rows, columns=['codigo', 'funcionalidade', 'codperfil', 'permitido'])
    df = df.drop_duplicates(subset=['codigo', 'funcionalidade'])
    return info, df


def _comparar_dfs(df_espelho, df_solicitado):
    codigos_e = set(df_espelho['codigo'].tolist())
    codigos_s = set(df_solicitado['codigo'].tolist())

    faltantes_codigos = codigos_e - codigos_s
    extras_codigos = codigos_s - codigos_e
    comuns_codigos = codigos_e & codigos_s

    faltantes = df_espelho[df_espelho['codigo'].isin(faltantes_codigos)].to_dict('records')
    extras = df_solicitado[df_solicitado['codigo'].isin(extras_codigos)].to_dict('records')
    comuns = df_espelho[df_espelho['codigo'].isin(comuns_codigos)].to_dict('records')

    return faltantes, extras, comuns


# --- ROTAS PRINCIPAIS ---
@app.route('/')
@require_login
def home():
    return render_template('home.html', active_app='home')


@app.route('/validator')
@require_login
def validator():
    conn = get_db_connection()
    historico = conn.execute('SELECT * FROM uploads_historico ORDER BY id DESC').fetchall()
    modulos = []
    is_var_active = False
    try:
        count_var = conn.execute('SELECT COUNT(*) AS cnt FROM dados_var').fetchone()['cnt']
        if count_var > 0:
            is_var_active = True
            engine = get_engine()
            df = pd.read_sql_query('SELECT DISTINCT "Módulo" FROM dados_var', engine)
            modulos = df["Módulo"].dropna().sort_values().tolist()
    except Exception:
        pass

    if not modulos: modulos = MODULOS_PADRAO
    conn.close()

    h_fmt = [
        {
            'id': h['id'],
            'nome_arquivo_original': h['nome_arquivo_original'],
            'timestamp_formatado': h['timestamp'],
            'status': h['status']
        }
        for h in historico
    ]

    return render_template('validator/index.html', historico=h_fmt, modulos=modulos, is_var_active=is_var_active, active_app='validator')


@app.route('/perfil')
@require_login
def perfil():
    conn = get_db_connection()
    atividades = conn.execute(
        'SELECT * FROM atividades_log WHERE usuario = ? ORDER BY id DESC LIMIT 50',
        (session.get('username', ''),)
    ).fetchall()
    conn.close()
    return render_template('perfil/index.html', active_app='perfil', atividades=atividades)


# --- GERENCIAMENTO DE USUÁRIOS (admin only) ---
@app.route('/admin/usuarios')
@require_admin
def admin_usuarios():
    conn = get_db_connection()
    usuarios = conn.execute('SELECT * FROM usuarios ORDER BY role DESC, username ASC').fetchall()
    conn.close()
    return render_template('admin/usuarios.html', active_app='admin', usuarios=usuarios)


@app.route('/admin/usuarios/criar', methods=['POST'])
@require_admin
def admin_criar_usuario():
    data = request.get_json()
    username = (data.get('username') or '').strip().lower()
    nome = (data.get('nome_completo') or '').strip()
    senha = data.get('password') or ''
    role = data.get('role') or 'user'

    if not username or not senha:
        return jsonify({'success': False, 'message': 'Usuário e senha são obrigatórios.'}), 400
    if role not in ('admin', 'user'):
        return jsonify({'success': False, 'message': 'Perfil inválido.'}), 400
    if len(senha) < 6:
        return jsonify({'success': False, 'message': 'A senha deve ter no mínimo 6 caracteres.'}), 400

    try:
        conn = get_db_connection()
        ts = datetime.now().strftime('%d/%m/%Y %H:%M:%S')
        ph = generate_password_hash(senha, method='pbkdf2:sha256')
        conn.execute(
            'INSERT INTO usuarios (username, password_hash, nome_completo, role, ativo, criado_em, deve_trocar_senha) VALUES (?, ?, ?, ?, 1, ?, 1)',
            (username, ph, nome, role, ts)
        )
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': f'Usuário "{username}" criado com sucesso. Ele deverá definir uma senha no primeiro acesso.'})
    except psycopg2.IntegrityError:
        return jsonify({'success': False, 'message': f'O usuário "{username}" já existe.'}), 409
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/admin/usuarios/<int:user_id>/editar', methods=['POST'])
@require_admin
def admin_editar_usuario(user_id):
    data = request.get_json()
    nome = (data.get('nome_completo') or '').strip()
    role = data.get('role') or 'user'
    senha = data.get('password') or ''

    if role not in ('admin', 'user'):
        return jsonify({'success': False, 'message': 'Perfil inválido.'}), 400

    try:
        conn = get_db_connection()
        row = conn.execute('SELECT * FROM usuarios WHERE id = ?', (user_id,)).fetchone()
        if not row:
            conn.close()
            return jsonify({'success': False, 'message': 'Usuário não encontrado.'}), 404

        if row['username'] == session.get('username') and role != 'admin':
            conn.close()
            return jsonify({'success': False, 'message': 'Você não pode remover seu próprio perfil de administrador.'}), 400

        if senha:
            if len(senha) < 6:
                conn.close()
                return jsonify({'success': False, 'message': 'A senha deve ter no mínimo 6 caracteres.'}), 400
            ph = generate_password_hash(senha, method='pbkdf2:sha256')
            conn.execute('UPDATE usuarios SET nome_completo=?, role=?, password_hash=? WHERE id=?', (nome, role, ph, user_id))
        else:
            conn.execute('UPDATE usuarios SET nome_completo=?, role=? WHERE id=?', (nome, role, user_id))

        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Usuário atualizado com sucesso.'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/admin/usuarios/<int:user_id>/toggle', methods=['POST'])
@require_admin
def admin_toggle_usuario(user_id):
    try:
        conn = get_db_connection()
        row = conn.execute('SELECT * FROM usuarios WHERE id = ?', (user_id,)).fetchone()
        if not row:
            conn.close()
            return jsonify({'success': False, 'message': 'Usuário não encontrado.'}), 404
        if row['username'] == session.get('username'):
            conn.close()
            return jsonify({'success': False, 'message': 'Você não pode desativar sua própria conta.'}), 400

        novo_status = 0 if row['ativo'] else 1
        conn.execute('UPDATE usuarios SET ativo = ? WHERE id = ?', (novo_status, user_id))
        conn.commit()
        conn.close()
        label = 'ativado' if novo_status else 'desativado'
        return jsonify({'success': True, 'message': f'Usuário {label} com sucesso.', 'ativo': novo_status})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/admin/usuarios/<int:user_id>/excluir', methods=['POST'])
@require_admin
def admin_excluir_usuario(user_id):
    try:
        conn = get_db_connection()
        row = conn.execute('SELECT * FROM usuarios WHERE id = ?', (user_id,)).fetchone()
        if not row:
            conn.close()
            return jsonify({'success': False, 'message': 'Usuário não encontrado.'}), 404
        if row['username'] == session.get('username'):
            conn.close()
            return jsonify({'success': False, 'message': 'Você não pode excluir sua própria conta.'}), 400
        if row['role'] == 'admin':
            count_admins = conn.execute(
                "SELECT COUNT(*) AS cnt FROM usuarios WHERE role='admin' AND ativo=1"
            ).fetchone()['cnt']
            if count_admins <= 1:
                conn.close()
                return jsonify({'success': False, 'message': 'Não é possível excluir o único administrador ativo.'}), 400

        conn.execute('DELETE FROM usuarios WHERE id = ?', (user_id,))
        conn.commit()
        conn.close()
        return jsonify({'success': True, 'message': 'Usuário excluído com sucesso.'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500


@app.route('/comparador_perfis')
@require_login
def comparador_perfis():
    return render_template('comparador_perfis/index.html', active_app='comparador_perfis')


@app.route('/comparar_perfis', methods=['POST'])
@require_login
def comparar_perfis_route():
    f_espelho = request.files.get('pdf_espelho')
    f_solicitado = request.files.get('pdf_solicitado')

    if not f_espelho or not f_solicitado:
        return jsonify({'erro': 'Envie os dois PDFs para comparar.'}), 400
    if not (f_espelho.filename.lower().endswith('.pdf') and f_solicitado.filename.lower().endswith('.pdf')):
        return jsonify({'erro': 'Apenas arquivos PDF são aceitos.'}), 400

    for f_check in [f_espelho, f_solicitado]:
        f_check.seek(0, 2)
        if f_check.tell() > PDF_MAX_BYTES:
            return jsonify({'erro': f'PDF "{f_check.filename}" muito grande. Limite: 15 MB.'}), 413
        f_check.seek(0)

    fd_e, path_e = tempfile.mkstemp(suffix='.pdf')
    fd_s, path_s = tempfile.mkstemp(suffix='.pdf')
    os.close(fd_e)
    os.close(fd_s)
    f_espelho.save(path_e)
    f_solicitado.save(path_s)

    try:
        info_e, df_e = extrair_perfil_pdf(path_e)
        info_s, df_s = extrair_perfil_pdf(path_s)

        if df_e.empty:
            return jsonify({'erro': 'Não foi possível extrair funcionalidades do PDF espelho. Verifique o formato.'}), 400
        if df_s.empty:
            return jsonify({'erro': 'Não foi possível extrair funcionalidades do PDF solicitado. Verifique o formato.'}), 400

        norm_mod = lambda s: re.sub(r'\s+', ' ', s.lower().strip())
        modulo_e = norm_mod(info_e.get('modulo', ''))
        modulo_s = norm_mod(info_s.get('modulo', ''))
        if modulo_e and modulo_s and modulo_e != modulo_s:
            return jsonify({
                'erro_modulo': True,
                'erro': (
                    f'Os PDFs pertencem a módulos diferentes e não podem ser comparados.\n'
                    f'• Espelho: {info_e.get("modulo") or "não identificado"}\n'
                    f'• Solicitado: {info_s.get("modulo") or "não identificado"}'
                ),
                'modulo_espelho': info_e.get('modulo', ''),
                'modulo_solicitado': info_s.get('modulo', ''),
            }), 422

        norm_amb = lambda s: s.strip().upper()
        amb_e = norm_amb(info_e.get('ambiente', ''))
        amb_s = norm_amb(info_s.get('ambiente', ''))
        aviso_ambiente = None
        if amb_e and amb_s and amb_e != amb_s:
            aviso_ambiente = (
                f'Os PDFs foram extraídos de ambientes diferentes: '
                f'Espelho em "{info_e.get("ambiente")}" e Solicitado em "{info_s.get("ambiente")}". '
                f'Os resultados podem não refletir o estado real de um único ambiente.'
            )

        faltantes, extras, comuns = _comparar_dfs(df_e, df_s)

        resultado = {
            'info_espelho': info_e,
            'info_solicitado': info_s,
            'total_espelho': len(df_e),
            'total_solicitado': len(df_s),
            'faltantes': faltantes,
            'extras': extras,
            'comuns': comuns,
            'aviso_ambiente': aviso_ambiente,
        }
        result_id = salvar_resultado_temp(resultado)
        resultado['result_id'] = result_id

        perf_e = info_e.get('codperfil') or 'espelho'
        perf_s = info_s.get('codperfil') or 'solicitado'
        registrar_atividade(
            'comparacao_perfis',
            f'Comparação de Perfis · {perf_e} vs {perf_s} — {len(faltantes)} faltantes, {len(extras)} extras, {len(comuns)} em comum'
        )

        return jsonify(resultado)

    except Exception as e:
        logger.exception("Erro em /comparar_perfis")
        return jsonify({'erro': f'Erro interno: {str(e)}'}), 500
    finally:
        for p in [path_e, path_s]:
            try:
                os.unlink(p)
            except Exception:
                pass


@app.route('/exportar_comparacao/<result_id>')
@require_login
def exportar_comparacao(result_id):
    resultado = carregar_resultado_temp(result_id)
    if not resultado:
        return jsonify({'erro': 'Resultado não encontrado ou expirado.'}), 404

    try:
        info_e = resultado.get('info_espelho', {})
        info_s = resultado.get('info_solicitado', {})
        faltantes = resultado.get('faltantes', [])
        extras = resultado.get('extras', [])
        comuns = resultado.get('comuns', [])

        perfil_s = info_s.get('codperfil', 'Solicitado')
        perfil_e = info_e.get('codperfil', 'Espelho')
        aba_filtro = request.args.get('aba')

        def to_df(lst):
            if not lst:
                return pd.DataFrame(columns=['Código', 'Funcionalidade', 'Codperfil'])
            df = pd.DataFrame(lst)
            df = df.rename(columns={'codigo': 'Código', 'funcionalidade': 'Funcionalidade', 'codperfil': 'Codperfil'})
            cols = [c for c in ['Código', 'Funcionalidade', 'Codperfil'] if c in df.columns]
            return df[cols]

        df_faltantes = to_df(faltantes)
        df_extras = to_df(extras)
        df_comuns = to_df(comuns)

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            wb = writer.book
            header_fmt = wb.add_format({'bold': True, 'bg_color': '#1e293b', 'font_color': '#ffffff', 'border': 1})
            red_fmt = wb.add_format({'bg_color': '#fee2e2', 'border': 1})
            blue_fmt = wb.add_format({'bg_color': '#dbeafe', 'border': 1})
            green_fmt = wb.add_format({'bg_color': '#dcfce7', 'border': 1})

            def write_sheet(df, sheet_name, row_fmt):
                df.to_excel(writer, sheet_name=sheet_name, index=False, startrow=1, header=False)
                ws = writer.sheets[sheet_name]
                for col_num, col_name in enumerate(df.columns):
                    ws.write(0, col_num, col_name, header_fmt)
                for row_num in range(len(df)):
                    for col_num in range(len(df.columns)):
                        ws.write(row_num + 1, col_num, str(df.iloc[row_num, col_num]), row_fmt)
                ws.set_column('A:A', 18)
                ws.set_column('B:B', 55)
                ws.set_column('C:C', 20)

            if aba_filtro == 'faltantes':
                write_sheet(df_faltantes, f'Faltantes em {perfil_s[:20]}', red_fmt)
                fname = f"Faltantes_{perfil_e}_para_{perfil_s}.xlsx"
            else:
                write_sheet(df_faltantes, f'Faltantes em {perfil_s[:20]}', red_fmt)
                write_sheet(df_extras, f'Extras em {perfil_s[:20]}', blue_fmt)
                write_sheet(df_comuns, 'Em Comum', green_fmt)
                fname = f"Comparacao_{perfil_e}_vs_{perfil_s}.xlsx"

        out.seek(0)
        return send_file(out, as_attachment=True, download_name=secure_filename(fname),
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')

    except Exception as e:
        logger.exception("Erro em /exportar_comparacao")
        return jsonify({'erro': str(e)}), 500


@app.route('/sod_analyzer')
@require_login
def sod_analyzer():
    return render_template('sod_analyzer/index.html', active_app='sod_analyzer')


@app.route('/analisar_sod', methods=['POST'])
@require_login
def analisar_sod():
    file = request.files.get('file')
    if not file or file.filename == '':
        return jsonify({'status': 'error', 'message': 'Nenhum arquivo enviado.'}), 400

    ext = file.filename.rsplit('.', 1)[-1].lower() if '.' in file.filename else ''
    if ext not in ('xlsx', 'xls', 'pdf'):
        return jsonify({'status': 'error', 'message': 'Formato inválido. Envie um arquivo .xlsx, .xls ou .pdf.'}), 400

    cenario = request.form.get('analysis_type', 'manutencao')

    if ext == 'pdf':
        file.seek(0, 2)
        if file.tell() > PDF_MAX_BYTES:
            return jsonify({'status': 'error', 'message': f'PDF muito grande. Limite: 15 MB.'}), 413
        file.seek(0)

    fd, fpath = tempfile.mkstemp(suffix=f'.{ext}')
    os.close(fd)
    file.save(fpath)

    try:
        if ext == 'pdf':
            res = extrair_sod_pdf(fpath, cenario)
        else:
            res = analisar_riscos_excel(fpath, cenario)
    except Exception as e:
        logger.exception("Erro em /analisar_sod")
        res = {'status': 'error', 'message': f'Erro interno: {str(e)}'}
    finally:
        try:
            os.unlink(fpath)
        except Exception:
            pass

    if res.get('status') in ('success', 'no_risks'):
        cenario_label = 'Manutenção' if cenario == 'manutencao' else 'Criação'
        perfil_nome = res.get('perfil', 'desconhecido')
        n_riscos = len(res.get('data', {})) if res.get('status') == 'success' else 0
        registrar_atividade(
            'analise_sod',
            f'Análise SoD · Cenário {cenario_label} — Perfil {perfil_nome} — {n_riscos} conflitos encontrados'
        )

    return jsonify(res)


# --- ROTAS AJAX (API) ---

@app.route('/listar_historico')
@require_login
def listar_historico():
    conn = get_db_connection()
    historico = conn.execute('SELECT * FROM uploads_historico ORDER BY id DESC').fetchall()
    conn.close()

    return jsonify([
        {
            'id': h['id'],
            'nome_arquivo_original': h['nome_arquivo_original'],
            'timestamp_formatado': h['timestamp'],
            'status': h['status']
        }
        for h in historico
    ])


@app.route('/upload_var', methods=['POST'])
@require_login
def upload_var():
    file = request.files.get('file')
    if not (file and allowed_file(file.filename)):
        return jsonify({'success': False, 'message': 'Erro no upload. Verifique o arquivo.'}), 400

    ts_now = datetime.now()
    timestamp_str = ts_now.strftime('%d/%m/%Y %H:%M:%S')
    fname = f"{ts_now.strftime('%Y%m%d%H%M%S')}_{secure_filename(file.filename)}"
    file_content = file.read()

    conn = get_db_connection()
    conn.execute(
        'INSERT INTO uploads_historico (nome_arquivo_original, nome_arquivo_salvo, timestamp, status, conteudo) VALUES (?, ?, ?, ?, ?)',
        (file.filename, fname, timestamp_str, 'Válido', psycopg2.Binary(file_content))
    )
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Upload realizado com sucesso! Clique em ativar.'})


@app.route('/ativar_var/<int:upload_id>')
@require_login
def ativar_var(upload_id):
    conn = get_db_connection()
    row = conn.execute('SELECT * FROM uploads_historico WHERE id = ?', (upload_id,)).fetchone()

    if not row:
        conn.close()
        return jsonify({'success': False, 'message': 'Arquivo não encontrado.'}), 404

    fd, tmp_path = tempfile.mkstemp(suffix='.xlsx')
    os.close(fd)

    try:
        # Escreve conteúdo BYTEA em arquivo temporário para pandas processar
        with open(tmp_path, 'wb') as f:
            f.write(bytes(row['conteudo']))

        df = pd.read_excel(tmp_path)
        cols = {c.lower().strip(): c for c in df.columns}

        col_id = cols.get('id') or cols.get('codigo')
        if not col_id: raise ValueError("Coluna obrigatória 'id' (ou 'codigo') não encontrada.")

        col_func = cols.get('funcionalidade')
        if not col_func: raise ValueError("Coluna obrigatória 'funcionalidade' não encontrada.")

        col_mod = cols.get('modulo') or cols.get('módulo')
        if not col_mod: raise ValueError("Coluna obrigatória 'modulo' não encontrada.")

        col_id_mod = None
        for k in ['id modulo', 'modulo id', 'id módulo', 'módulo id', 'cod modulo']:
            if k in cols:
                col_id_mod = cols[k]
                break
        if not col_id_mod: raise ValueError("Coluna obrigatória 'id modulo' (ou 'modulo id') não encontrada.")

        df.rename(columns={
            col_id: 'ID Funcionalidade',
            col_func: 'Funcionalidade',
            col_mod: 'Módulo',
            col_id_mod: 'ID Módulo'
        }, inplace=True)

        df_final = df[['ID Funcionalidade', 'Funcionalidade', 'ID Módulo', 'Módulo']].copy()

        engine = get_engine()
        df_final.to_sql('dados_var', engine, if_exists='replace', index=False)

        conn.execute("UPDATE uploads_historico SET status = 'Arquivado' WHERE status = 'Ativo'")
        conn.execute("UPDATE uploads_historico SET status = 'Ativo' WHERE id = ?", (upload_id,))
        conn.commit()

        return jsonify({'success': True, 'message': f'Base ativada com {len(df_final)} registros!'})

    except Exception as e:
        return jsonify({'success': False, 'message': f'Erro: {str(e)}'}), 400

    finally:
        conn.close()
        try:
            os.unlink(tmp_path)
        except Exception:
            pass


@app.route('/excluir_var/<int:upload_id>')
@require_login
def excluir_var(upload_id):
    conn = get_db_connection()
    try:
        row = conn.execute('SELECT * FROM uploads_historico WHERE id = ?', (upload_id,)).fetchone()
        if not row:
            return jsonify({'success': False, 'message': 'Arquivo não encontrado.'}), 404
        if row['status'] == 'Ativo':
            return jsonify({'success': False, 'message': 'Ação Negada: Não é possível excluir a Base Ativa.'}), 400

        conn.execute('DELETE FROM uploads_historico WHERE id = ?', (upload_id,))
        conn.commit()
        return jsonify({'success': True, 'message': 'Arquivo excluído do histórico.'})
    except Exception as e:
        return jsonify({'success': False, 'message': str(e)}), 500
    finally:
        conn.close()


@app.route('/visualizar_base_ativa')
@require_login
def visualizar_base_ativa():
    try:
        conn = get_db_connection()
        try:
            count = conn.execute('SELECT COUNT(*) AS cnt FROM dados_var').fetchone()['cnt']
            if count == 0:
                conn.close()
                return jsonify({'erro': 'A base de dados está vazia.'}), 404
        except Exception:
            conn.close()
            return jsonify({'erro': 'Tabela de dados não encontrada.'}), 404

        dados = conn.execute('SELECT * FROM dados_var LIMIT 2000').fetchall()
        conn.close()
        return jsonify([dict(row) for row in dados])
    except Exception as e:
        return jsonify({'erro': str(e)}), 500


@app.route('/comparar', methods=['POST'])
@require_login
def comparar():
    conn = get_db_connection()
    try:
        try:
            if conn.execute('SELECT COUNT(*) AS cnt FROM dados_var').fetchone()['cnt'] == 0:
                return jsonify({'erro': 'Base VAR vazia. Ative uma base antes de comparar.'}), 400
        except Exception:
            return jsonify({'erro': 'Nenhuma Planilha VAR ativa no sistema.'}), 400
    except Exception as e:
        return jsonify({'erro': f'Erro de conexão com banco: {str(e)}'}), 500
    finally:
        conn.close()

    if 'arquivo_analise' not in request.files:
        return jsonify({'erro': 'Sem arquivo.'}), 400

    try:
        f = request.files['arquivo_analise']
        mod = request.form.get('modulo')
        fname = f.filename.lower()
        df_usr = pd.DataFrame()

        if fname.endswith('.pdf'):
            fd, tpath = tempfile.mkstemp(suffix='.pdf')
            os.close(fd)
            f.save(tpath)
            df_usr = extrair_funcionalidades_pdf(tpath)
            try:
                os.unlink(tpath)
            except Exception:
                pass
            if df_usr.empty:
                return jsonify({'erro': "Não foi possível extrair dados do PDF."}), 400
        elif fname.endswith(('.xlsx', '.xls')):
            df_usr = pd.read_excel(f, usecols=[0])
            if df_usr.empty:
                return jsonify({'erro': "Excel vazio."}), 400
        else:
            return jsonify({'erro': "Formato inválido."}), 400

        norm = lambda s: str(s).lower().strip()
        l_usr = [norm(x) for x in df_usr.iloc[:, 0].dropna()]
        if not l_usr:
            return jsonify({'erro': 'Nenhuma funcionalidade encontrada no arquivo enviado.'}), 400

        engine = get_engine()
        df_var = pd.read_sql_query(
            sa_text('SELECT "ID Funcionalidade", "Funcionalidade" FROM dados_var WHERE "Módulo" = :modulo'),
            engine, params={'modulo': mod}
        )

        if df_var.empty:
            return jsonify({'erro': f'A Base VAR não possui registros para o módulo "{mod}".'}), 400

        df_var['F_Norm'] = df_var['Funcionalidade'].apply(norm)
        v_map = {r['F_Norm']: {'id': r['ID Funcionalidade'], 'orig': r['Funcionalidade']} for _, r in df_var.iterrows()}
        v_list = df_var['F_Norm'].tolist()

        res = []
        for i, u_norm in enumerate(l_usr):
            display = df_usr.iloc[i, 0] if i < len(df_usr) else u_norm
            item = {'id': i, 'Funcionalidade Analisada': str(display), 'Status': 'Divergente', 'ID Encontrado': '', 'ID Sugerido': '', 'Sugestão Similar (VAR)': '', 'Similaridade (%)': 0.0}

            if u_norm in v_map:
                m = v_map[u_norm]
                item.update({'Status': 'Encontrado', 'ID Encontrado': str(m['id']), 'Sugestão Similar (VAR)': m['orig'], 'Similaridade (%)': 100.0})
            else:
                best = process.extractOne(u_norm, v_list, scorer=fuzz.WRatio, score_cutoff=80)
                if best:
                    sug, score, _ = best
                    m = v_map[sug]
                    item.update({'Status': 'Divergente com Sugestão', 'ID Sugerido': str(m['id']), 'Sugestão Similar (VAR)': m['orig'], 'Similaridade (%)': round(score, 2)})
            res.append(item)

        res.sort(key=lambda x: {'Divergente com Sugestão': 0, 'Divergente': 1, 'Encontrado': 2}.get(x['Status'], 99))
        divs = sum(1 for r in res if r['Status'] != 'Encontrado')

        msg = {"texto": "Sucesso! Tudo ok.", "tipo": "sucesso"}
        if divs > 0:
            msg = {"texto": f"Concluído com {divs} divergências.", "tipo": "ressalva"}
            if (divs / len(res)) >= LIMIAR_ALERTA:
                msg = {"texto": "Alerta: Muitas divergências.", "tipo": "alerta"}

        registrar_atividade(
            'validacao_var',
            f'Validação VAR · Módulo {mod} — {len(res)} funcionalidades comparadas, {divs} divergências'
        )

        return jsonify({'resultados': res, 'mensagem_status': msg})

    except Exception as e:
        logger.exception("Erro na rota /comparar")
        return jsonify({'erro': f'Erro interno: {str(e)}'}), 500


@app.route('/gerar_importacao', methods=['POST'])
@require_login
def gerar_importacao():
    data = request.get_json()
    res = data.get('resultados')
    pnome = data.get('perfil_nome')
    pid = data.get('perfil_id')

    if not res:
        return jsonify({"erro": "Sem dados"}), 400

    try:
        df = pd.DataFrame(res)
        get_id = lambda r: (str(r.get('ID Encontrado', '')).strip() or str(r.get('ID Sugerido', '')).strip()) or '❗ Não encontrado'
        get_fn = lambda r: str(r.get('Sugestão Similar (VAR)', '')).strip() if r.get('Status') in ['Encontrado', 'Divergente com Sugestão'] else str(r.get('Funcionalidade Analisada', ''))

        lista_ids = df.apply(get_id, axis=1).tolist()
        lista_funcs = df.apply(get_fn, axis=1).tolist()

        out_df = pd.DataFrame({
            'funcionalidade id': lista_ids,
            'funcionalidade': lista_funcs
        })
        out_df['id'] = pid
        out_df['perfil'] = pnome
        out_df = out_df[['id', 'perfil', 'funcionalidade id', 'funcionalidade']]

        out = io.BytesIO()
        with pd.ExcelWriter(out, engine='xlsxwriter') as writer:
            out_df.to_excel(writer, index=False, sheet_name='Importa VAR')
            worksheet = writer.sheets['Importa VAR']
            worksheet.set_column('A:A', 10)
            worksheet.set_column('B:B', 20)
            worksheet.set_column('C:C', 15)
            worksheet.set_column('D:D', 50)

        out.seek(0)
        return send_file(out, as_attachment=True, download_name=f"Importar_VAR_{secure_filename(pnome)}.xlsx",
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except Exception as e:
        logger.exception("Erro na rota /gerar_importacao")
        return jsonify({"erro": str(e)}), 500


if __name__ == '__main__':
    debug = os.environ.get('FLASK_DEBUG', 'false').lower() == 'true'
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=debug, port=port)
