import pandas as pd
import re
from datetime import datetime
from database.db import get_tipo_by_cnpj

TIPOS_AUTOMATICOS = {
    "despesa": ["energia", "agua", "luz", "telefone", "internet", "aluguel", "seguro",
                "limpeza", "contabilidade", "escritório", "manutencao", "servico"],
    "devolucao": ["devoluc", "retorno", "estorno", "cancelad"],
}

def normalize_text(text):
    if pd.isna(text) or text is None:
        return ""
    return re.sub(r'\s+', ' ', str(text).strip().upper())

def normalize_number(val):
    if pd.isna(val) or val is None:
        return ""
    s = str(val).strip()
    s = re.sub(r'[^\d]', '', s)
    return s.lstrip('0') or "0"

def normalize_value(val):
    if pd.isna(val) or val is None:
        return 0.0
    try:
        if isinstance(val, (int, float)):
            return float(val)
        s = str(val).replace("R$", "").replace(" ", "").replace(".", "").replace(",", ".")
        return float(s)
    except:
        return 0.0

def load_estoque(filepath, log_cb=None):
    def log(msg): log_cb and log_cb("INFO", msg)
    def err(msg): log_cb and log_cb("ERRO", msg)

    log(f"Lendo arquivo de estoque: {filepath}")
    try:
        df = pd.read_excel(filepath, dtype=str)
        log(f"Arquivo lido: {len(df)} linhas brutas")

        col_map = {}
        for col in df.columns:
            nc = normalize_text(col)
            if "DOCUMENTO" in nc: col_map["documento"] = col
            elif "CÓDIGO" in nc or "CODIGO" in nc: col_map["codigo"] = col
            elif "FORNECEDOR" in nc: col_map["fornecedor"] = col
            elif "EMISSÃO" in nc or "EMISSAO" in nc or "EMISS" in nc: col_map["data_emissao"] = col
            elif "TOTAL" in nc and "R$" in nc: col_map["valor"] = col
            elif "TOTAL" in nc: col_map["valor"] = col

        missing = [k for k in ["documento", "fornecedor", "valor"] if k not in col_map]
        if missing:
            err(f"Colunas não encontradas no estoque: {missing}")
            return None, f"Colunas ausentes: {missing}"

        df_out = pd.DataFrame()
        df_out["documento"] = df[col_map["documento"]].apply(normalize_number)
        df_out["fornecedor"] = df[col_map["fornecedor"]].apply(normalize_text)
        df_out["valor"] = df[col_map["valor"]].apply(normalize_value)
        if "data_emissao" in col_map:
            df_out["data_emissao"] = df[col_map["data_emissao"]].apply(normalize_text)
        else:
            df_out["data_emissao"] = ""
        if "codigo" in col_map:
            df_out["codigo"] = df[col_map["codigo"]].apply(normalize_text)
        else:
            df_out["codigo"] = ""

        df_out = df_out[df_out["documento"] != ""].dropna(subset=["documento"])
        log(f"Estoque processado: {len(df_out)} registros válidos")
        return df_out, None
    except Exception as e:
        err(f"Erro ao ler estoque: {e}")
        return None, str(e)

def load_forlions(filepath, log_cb=None):
    def log(msg): log_cb and log_cb("INFO", msg)
    def err(msg): log_cb and log_cb("ERRO", msg)

    log(f"Lendo arquivo ForLions: {filepath}")
    try:
        df = pd.read_excel(filepath, header=4, dtype=str)
        log(f"Arquivo lido: {len(df)} linhas (cabeçalho na linha 5)")

        col_map = {}
        for col in df.columns:
            nc = normalize_text(col)
            if "NÚMERO" in nc or "NUMERO" in nc or nc == "NÚMERO" or nc == "NUMERO": col_map["numero"] = col
            elif "CHAVE" in nc: col_map["chave"] = col
            elif "CNPJ" in nc: col_map["cnpj"] = col
            elif "RAZÃO SOCIAL" in nc or "RAZAO SOCIAL" in nc: col_map["razao_social"] = col
            elif "DT. EMISSÃO" in nc or "DT.EMISSÃO" in nc or ("DT" in nc and "EMISS" in nc): col_map["data_emissao"] = col
            elif "VALOR TOTAL" in nc and "PRODUTO" not in nc: col_map["valor"] = col
            elif nc == "VALOR": col_map["valor"] = col
            elif "CANCELADA" in nc: col_map["cancelada"] = col

        missing = [k for k in ["numero", "valor"] if k not in col_map]
        if missing:
            err(f"Colunas não encontradas no ForLions: {missing}")
            return None, f"Colunas ausentes no ForLions: {missing}"

        df_out = pd.DataFrame()
        df_out["numero"] = df[col_map["numero"]].apply(normalize_number)
        df_out["valor"] = df[col_map["valor"]].apply(normalize_value)
        df_out["cnpj"] = df[col_map["cnpj"]].apply(normalize_number) if "cnpj" in col_map else ""
        df_out["razao_social"] = df[col_map["razao_social"]].apply(normalize_text) if "razao_social" in col_map else ""
        df_out["data_emissao"] = df[col_map["data_emissao"]].apply(normalize_text) if "data_emissao" in col_map else ""
        if "cancelada" in col_map:
            df_out["cancelada"] = df[col_map["cancelada"]].apply(normalize_text)
        else:
            df_out["cancelada"] = ""
        if "chave" in col_map:
            df_out["chave"] = df[col_map["chave"]].apply(normalize_text)
        else:
            df_out["chave"] = ""

        df_out = df_out[df_out["numero"] != ""].dropna(subset=["numero"])
        # Remove canceladas
        df_out = df_out[~df_out["cancelada"].isin(["SIM", "S", "X", "1", "TRUE"])]
        log(f"ForLions processado: {len(df_out)} registros válidos")
        return df_out, None
    except Exception as e:
        err(f"Erro ao ler ForLions: {e}")
        return None, str(e)

def infer_risco(row):
    val = row.get("valor_forlions", 0) or 0
    if val > 50000:
        return "Alto"
    elif val > 10000:
        return "Médio"
    return "Baixo"

def infer_classificacao(razao_social, cnpj):
    rs = normalize_text(razao_social)
    tipo_db = get_tipo_by_cnpj(str(cnpj).strip()) if cnpj else None
    if tipo_db:
        return tipo_db
    for tipo, palavras in TIPOS_AUTOMATICOS.items():
        if any(p in rs.lower() for p in palavras):
            return tipo.capitalize()
    return "Pendente"

def compare(df_estoque, df_forlions, log_cb=None, progress_cb=None):
    def log(nivel, msg): log_cb and log_cb(nivel, msg)

    results = {
        "conciliadas": [],
        "divergencias": [],
        "despesas": [],
        "devolucoes": [],
        "sem_entrada": [],
    }

    estoque_map = {}
    for _, row in df_estoque.iterrows():
        doc = row["documento"]
        if doc not in estoque_map:
            estoque_map[doc] = []
        estoque_map[doc].append(row)

    total = len(df_forlions)
    for i, (_, fnota) in enumerate(df_forlions.iterrows()):
        if progress_cb:
            progress_cb(i + 1, total)

        num = fnota["numero"]
        razao = fnota.get("razao_social", "")
        cnpj = fnota.get("cnpj", "")
        valor_f = fnota["valor"]
        data_f = fnota.get("data_emissao", "")

        auto_class = infer_classificacao(razao, cnpj)

        base_div = {
            "numero_nota": num,
            "cnpj": cnpj,
            "razao_social": razao,
            "data_emissao": data_f,
            "valor_forlions": valor_f,
            "valor_estoque": None,
            "risco": infer_risco({"valor_forlions": valor_f}),
            "classificacao": auto_class,
            "observacoes": "",
            "status": "Pendente",
        }

        if num in estoque_map:
            matches = estoque_map[num]
            melhor = None
            for m in matches:
                dif = abs(m["valor"] - valor_f)
                if melhor is None or dif < abs(melhor["valor"] - valor_f):
                    melhor = m

            dif_val = abs(melhor["valor"] - valor_f)
            pct = (dif_val / valor_f * 100) if valor_f else 0

            if pct < 1:
                log("OK", f"CONCILIADA: NF {num} | R$ {valor_f:.2f}")
                results["conciliadas"].append({**base_div, "valor_estoque": melhor["valor"], "status": "Conciliada"})
            else:
                log("AVISO", f"DIVERGÊNCIA DE VALOR: NF {num} | ForLions R$ {valor_f:.2f} | Estoque R$ {melhor['valor']:.2f} | Dif {pct:.1f}%")
                results["divergencias"].append({
                    **base_div,
                    "valor_estoque": melhor["valor"],
                    "status": "Divergência de Valor",
                    "observacoes": f"Diferença de {pct:.1f}% (R$ {dif_val:.2f})"
                })
        else:
            if auto_class in ["Despesa"]:
                log("INFO", f"DESPESA AUTOMÁTICA: NF {num} | {razao}")
                results["despesas"].append({**base_div, "status": "Despesa"})
            elif auto_class in ["Devolucao", "Devolução"]:
                log("INFO", f"DEVOLUÇÃO AUTOMÁTICA: NF {num} | {razao}")
                results["devolucoes"].append({**base_div, "status": "Devolução"})
            else:
                log("ALERTA", f"SEM ENTRADA: NF {num} | {razao} | R$ {valor_f:.2f}")
                results["sem_entrada"].append({**base_div, "status": "Sem Entrada"})

    log("INFO", f"Conciliadas: {len(results['conciliadas'])}")
    log("INFO", f"Divergências: {len(results['divergencias'])}")
    log("INFO", f"Sem entrada: {len(results['sem_entrada'])}")
    log("INFO", f"Despesas: {len(results['despesas'])}")
    log("INFO", f"Devoluções: {len(results['devolucoes'])}")

    return results
