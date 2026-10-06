#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
consolidar.py — lê as planilhas de pagamentos na rede e grava `dados.js`
ao lado do `index.html`. O dashboard carrega esse arquivo sozinho ao abrir.

Uso:
    python consolidar.py                      # usa as PASTAS configuradas abaixo
    python consolidar.py \\\\servidor-exemplo\\financeiro\\Pagamentos
    python consolidar.py C:\\dados --saida C:\\dash --excel

Requisitos: pip install pandas openpyxl xlrd
"""
import argparse
import json
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import pandas as pd

# ------------------------------------------------------------------ configuração
# Pastas varridas quando nenhum caminho é passado na linha de comando.
# Subpastas são incluídas automaticamente (ex.: ...\Pagamentos\2025, ...\2026).
PASTAS = [
    r"dados_exemplo",
]
SAIDA = Path(__file__).resolve().parent          # onde gravar dados.js
EXTENSOES = {".xlsx", ".xlsm", ".xls", ".csv"}

SINONIMOS = {
    "data": ["data pagamento", "data de pagamento", "data do pagamento", "dt pagamento",
             "data baixa", "data", "dt", "vencimento", "pagamento em", "competencia"],
    "empresa": ["beneficiario", "empresa", "fornecedor", "razao social", "credor",
                "cliente", "nome fornecedor", "destinatario", "favorecido"],
    "favorecido": ["favorecido", "nome favorecido", "favorecido pagamento", "titular",
                   "conta favorecido"],
    "valor": ["valor pago", "valor do pagamento", "valor liquido", "montante",
              "valor total", "valor", "vlr", "total pago", "total"],
    "banco": ["banco", "instituicao financeira", "instituicao", "bco",
              "conta bancaria", "banco pagador"],
    "departamento": ["departamento solicitante", "departamento", "depto", "setor",
                     "area solicitante", "area", "solicitante", "requisitante",
                     "centro de custo", "cc"],
    "natureza": ["descricao das naturezas", "descricao da natureza", "naturezas",
                 "natureza da despesa", "natureza", "rubrica", "conta contabil",
                 "classificacao"],
    "descricao": ["descricao", "historico", "observacao", "obs", "documento",
                  "nota fiscal", "nf", "titulo", "referencia", "finalidade"],
}
OBRIGATORIOS = ("data", "empresa", "valor")


def norm(texto) -> str:
    s = unicodedata.normalize("NFD", str(texto if texto is not None else ""))
    s = "".join(c for c in s if unicodedata.category(c) != "Mn").lower()
    return re.sub(r"[^a-z0-9]+", " ", s).strip()


def pontua(linha) -> int:
    achados = 0
    for celula in linha:
        n = norm(celula)
        if not n:
            continue
        if any(n == k or k in n for chaves in SINONIMOS.values() for k in chaves):
            achados += 1
    return achados


def acha_cabecalho(bruto: pd.DataFrame) -> int:
    melhor, pontos = -1, 1
    for i in range(min(20, len(bruto))):
        p = pontua(bruto.iloc[i].tolist())
        if p > pontos:
            melhor, pontos = i, p
    return melhor


def mapeia(colunas) -> dict:
    usados, mapa = set(), {}
    for campo, chaves in SINONIMOS.items():
        escolha, melhor = None, -1
        for i, col in enumerate(colunas):
            if i in usados:
                continue
            n = norm(col)
            if not n:
                continue
            for ordem, k in enumerate(chaves):
                peso = -1
                if n == k:
                    peso = 1000 - ordem
                elif n.startswith(k) or n.endswith(k):
                    peso = 600 - ordem
                elif k in n:
                    peso = 300 - ordem
                if peso > melhor:
                    melhor, escolha = peso, i
        if escolha is not None and melhor > 0:
            mapa[campo] = escolha
            usados.add(escolha)
    return mapa


def para_numero(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return None if pd.isna(v) else float(v)
    s = str(v or "").strip()
    if not s:
        return None
    negativo = s.startswith("(") and s.endswith(")")
    s = re.sub(r"[^\d,.\-]", "", s)
    if not s:
        return None
    uv, up = s.rfind(","), s.rfind(".")
    if uv > -1 and up > -1:
        s = s.replace(".", "").replace(",", ".") if uv > up else s.replace(",", "")
    elif uv > -1:
        s = s.replace(",", ".")
    elif re.fullmatch(r"-?\d{1,3}(\.\d{3})+", s):
        s = s.replace(".", "")
    try:
        n = float(s)
    except ValueError:
        return None
    return -n if negativo and n > 0 else n


def para_data(v):
    if isinstance(v, (pd.Timestamp, datetime)):
        return None if pd.isna(v) else pd.Timestamp(v).to_pydatetime().date()
    d = pd.to_datetime(v, dayfirst=True, errors="coerce")
    return None if pd.isna(d) else d.date()


def le_planilha(caminho: Path):
    """Devolve (registros, avisos) de um arquivo."""
    registros, avisos = [], []
    try:
        if caminho.suffix.lower() == ".csv":
            abas = {"csv": pd.read_csv(caminho, header=None, dtype=object,
                                       sep=None, engine="python", encoding="utf-8-sig")}
        else:
            abas = pd.read_excel(caminho, sheet_name=None, header=None, dtype=object)
    except Exception as e:
        return [], [f"{caminho.name}: não foi possível abrir ({e})"]

    for nome_aba, bruto in abas.items():
        if bruto.empty:
            continue
        hr = acha_cabecalho(bruto)
        if hr < 0:
            if len(bruto) >= 4 and bruto.shape[1] >= 3:
                avisos.append(f"{caminho.name} › {nome_aba}: cabeçalho não reconhecido")
            continue
        colunas = bruto.iloc[hr].tolist()
        mapa = mapeia(colunas)
        faltando = [c for c in OBRIGATORIOS if c not in mapa]
        if faltando:
            avisos.append(f"{caminho.name} › {nome_aba}: sem coluna de {', '.join(faltando)}")
            continue

        dados = bruto.iloc[hr + 1:]
        for _, linha in dados.iterrows():
            empresa = str(linha.iloc[mapa["empresa"]] or "").strip()
            data = para_data(linha.iloc[mapa["data"]])
            valor = para_numero(linha.iloc[mapa["valor"]])
            if not empresa or data is None or valor is None:
                continue
            if re.match(r"^(total|subtotal|soma|geral|totais)\b", norm(empresa)):
                continue

            def col(campo, padrao=""):
                if campo not in mapa:
                    return padrao
                v = linha.iloc[mapa[campo]]
                return (str(v).strip() if v is not None and not pd.isna(v) else "") or padrao

            registros.append({
                "data": data.strftime("%d/%m/%Y"),
                "empresa": empresa,
                "valor": round(valor, 2),
                "favorecido": col("favorecido"),
                "banco": col("banco", "Não informado"),
                "natureza": col("natureza", "Não informada"),
                "departamento": col("departamento", "Não informado"),
                "descricao": col("descricao"),
                "origem": caminho.name,
            })
    return registros, avisos


def main():
    ap = argparse.ArgumentParser(description="Consolida planilhas de pagamentos em dados.js")
    ap.add_argument("pastas", nargs="*", default=None,
                    help="pastas a varrer (padrão: lista PASTAS no topo do script)")
    ap.add_argument("--saida", default=str(SAIDA), help="pasta onde gravar dados.js")
    ap.add_argument("--excel", action="store_true", help="gravar também pagamentos_consolidado.xlsx")
    args = ap.parse_args()

    raizes = [Path(p) for p in (args.pastas or PASTAS)]
    arquivos = []
    for raiz in raizes:
        if not raiz.exists():
            print(f"[aviso] pasta não encontrada: {raiz}")
            continue
        arquivos += [p for p in raiz.rglob("*")
                     if p.suffix.lower() in EXTENSOES and not p.name.startswith(("~$", "."))]
    if not arquivos:
        print("Nenhuma planilha encontrada. Ajuste a lista PASTAS ou passe o caminho por argumento.")
        sys.exit(1)

    todos, avisos = [], []
    for arq in sorted(arquivos):
        regs, avs = le_planilha(arq)
        todos += regs
        avisos += avs
        print(f"{arq}  →  {len(regs)} pagamentos")

    if not todos:
        print("Nenhum pagamento reconhecido. Confira os nomes das colunas.")
        sys.exit(1)

    todos.sort(key=lambda r: (r["data"][6:], r["data"][3:5], r["data"][:2], r["empresa"]))
    saida = Path(args.saida)
    saida.mkdir(parents=True, exist_ok=True)
    destino = saida / "dados.js"
    destino.write_text(
        "/* Gerado por consolidar.py em "
        + datetime.now().strftime("%d/%m/%Y %H:%M")
        + " — não editar à mão. */\n"
        + 'window.DADOS_ATUALIZADO = "' + datetime.now().isoformat(timespec="seconds") + '";\n'
        + "window.DADOS_PAGAMENTOS = "
        + json.dumps(todos, ensure_ascii=False, separators=(",", ":"))
        + ";\n",
        encoding="utf-8")

    if args.excel:
        pd.DataFrame(todos).to_excel(saida / "pagamentos_consolidado.xlsx", index=False)

    total = sum(r["valor"] for r in todos)
    print(f"\n{len(todos)} pagamentos · R$ {total:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    print(f"gravado em {destino}")
    for a in avisos:
        print(f"[aviso] {a}")


if __name__ == "__main__":
    main()
