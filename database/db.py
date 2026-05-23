import sqlite3
import os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.dirname(__file__)), "audit.db")

def get_connection():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn

def init_db():
    conn = get_connection()
    cur = conn.cursor()

    cur.executescript("""
        CREATE TABLE IF NOT EXISTS empresas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nome_fantasia TEXT,
            razao_social TEXT NOT NULL,
            cnpj TEXT UNIQUE,
            tipo_padrao TEXT DEFAULT 'Outro',
            observacoes TEXT,
            criado_em TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS importacoes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            arquivo_estoque TEXT,
            arquivo_forlions TEXT,
            total_estoque INTEGER DEFAULT 0,
            total_forlions INTEGER DEFAULT 0,
            total_conciliadas INTEGER DEFAULT 0,
            total_divergencias INTEGER DEFAULT 0,
            total_despesas INTEGER DEFAULT 0,
            total_devolucoes INTEGER DEFAULT 0,
            valor_total_notas REAL DEFAULT 0,
            valor_conciliado REAL DEFAULT 0,
            criado_em TEXT DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS divergencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            importacao_id INTEGER,
            numero_nota TEXT,
            cnpj TEXT,
            razao_social TEXT,
            data_emissao TEXT,
            valor_forlions REAL,
            valor_estoque REAL,
            status TEXT DEFAULT 'Pendente',
            classificacao TEXT,
            observacoes TEXT,
            risco TEXT DEFAULT 'Médio',
            aprovado_por TEXT,
            criado_em TEXT DEFAULT CURRENT_TIMESTAMP,
            FOREIGN KEY (importacao_id) REFERENCES importacoes(id)
        );

        CREATE TABLE IF NOT EXISTS logs (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            importacao_id INTEGER,
            nivel TEXT,
            mensagem TEXT,
            criado_em TEXT DEFAULT CURRENT_TIMESTAMP
        );
    """)
    conn.commit()
    conn.close()

def insert_empresa(nome_fantasia, razao_social, cnpj, tipo_padrao, observacoes=""):
    conn = get_connection()
    try:
        conn.execute("""
            INSERT OR REPLACE INTO empresas (nome_fantasia, razao_social, cnpj, tipo_padrao, observacoes)
            VALUES (?, ?, ?, ?, ?)
        """, (nome_fantasia, razao_social, cnpj, tipo_padrao, observacoes))
        conn.commit()
        return True
    except Exception as e:
        return str(e)
    finally:
        conn.close()

def get_empresas():
    conn = get_connection()
    rows = conn.execute("SELECT * FROM empresas ORDER BY razao_social").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def delete_empresa(empresa_id):
    conn = get_connection()
    conn.execute("DELETE FROM empresas WHERE id = ?", (empresa_id,))
    conn.commit()
    conn.close()

def get_tipo_by_cnpj(cnpj):
    conn = get_connection()
    row = conn.execute("SELECT tipo_padrao FROM empresas WHERE cnpj = ?", (cnpj,)).fetchone()
    conn.close()
    return row["tipo_padrao"] if row else None

def save_importacao(data: dict) -> int:
    conn = get_connection()
    cur = conn.execute("""
        INSERT INTO importacoes (arquivo_estoque, arquivo_forlions, total_estoque, total_forlions,
            total_conciliadas, total_divergencias, total_despesas, total_devolucoes,
            valor_total_notas, valor_conciliado)
        VALUES (:arquivo_estoque, :arquivo_forlions, :total_estoque, :total_forlions,
            :total_conciliadas, :total_divergencias, :total_despesas, :total_devolucoes,
            :valor_total_notas, :valor_conciliado)
    """, data)
    conn.commit()
    imp_id = cur.lastrowid
    conn.close()
    return imp_id

def save_divergencia(data: dict):
    conn = get_connection()
    conn.execute("""
        INSERT INTO divergencias (importacao_id, numero_nota, cnpj, razao_social,
            data_emissao, valor_forlions, valor_estoque, status, classificacao, observacoes, risco)
        VALUES (:importacao_id, :numero_nota, :cnpj, :razao_social,
            :data_emissao, :valor_forlions, :valor_estoque, :status, :classificacao, :observacoes, :risco)
    """, data)
    conn.commit()
    conn.close()

def get_divergencias(importacao_id=None):
    conn = get_connection()
    if importacao_id:
        rows = conn.execute("SELECT * FROM divergencias WHERE importacao_id = ? ORDER BY criado_em DESC", (importacao_id,)).fetchall()
    else:
        rows = conn.execute("SELECT * FROM divergencias ORDER BY criado_em DESC").fetchall()
    conn.close()
    return [dict(r) for r in rows]

def update_divergencia(div_id, classificacao, observacoes, status):
    conn = get_connection()
    conn.execute("""
        UPDATE divergencias SET classificacao=?, observacoes=?, status=? WHERE id=?
    """, (classificacao, observacoes, status, div_id))
    conn.commit()
    conn.close()

def add_log(importacao_id, nivel, mensagem):
    conn = get_connection()
    conn.execute("INSERT INTO logs (importacao_id, nivel, mensagem) VALUES (?, ?, ?)",
                 (importacao_id, nivel, mensagem))
    conn.commit()
    conn.close()

def get_historico():
    conn = get_connection()
    rows = conn.execute("SELECT * FROM importacoes ORDER BY criado_em DESC LIMIT 20").fetchall()
    conn.close()
    return [dict(r) for r in rows]
