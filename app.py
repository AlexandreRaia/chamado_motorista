import hashlib
import os
import secrets
import sqlite3
import threading
import time
from datetime import datetime, timedelta
from html.parser import HTMLParser

from flask import Flask, g, redirect, render_template, request, url_for
from openpyxl import load_workbook

BASE = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE, "chamada.db")
ANTECEDENCIA_PADRAO = 30  # minutos antes da saida para chamar o motorista

app = Flask(__name__)

SCHEMA = """
CREATE TABLE IF NOT EXISTS motoristas (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    nome TEXT UNIQUE NOT NULL
);
CREATE TABLE IF NOT EXISTS viagens (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    chave TEXT UNIQUE NOT NULL,
    motorista_id INTEGER NOT NULL REFERENCES motoristas(id),
    inicio TEXT NOT NULL,
    termino TEXT,
    placa TEXT,
    modelo TEXT,
    secretaria TEXT,
    solicitante TEXT,
    descricao TEXT,
    destino TEXT,
    passageiros INTEGER,
    status TEXT NOT NULL DEFAULT 'aguardando',
    chamado_em TEXT,
    ciente_em TEXT,
    concluido_em TEXT
);
CREATE TABLE IF NOT EXISTS config (
    chave TEXT PRIMARY KEY,
    valor TEXT NOT NULL
);
"""

STATUS_ROTULO = {
    "aguardando": "Aguardando",
    "chamado": "Chamado",
    "ciente": "Ciente",
    "concluido": "Concluído",
}


# ---------- banco ----------
def conectar():
    con = sqlite3.connect(DB_PATH)
    con.row_factory = sqlite3.Row
    return con


def db():
    if "db" not in g:
        g.db = conectar()
    return g.db


@app.teardown_appcontext
def fechar(_exc):
    con = g.pop("db", None)
    if con is not None:
        con.close()


def iniciar_db():
    con = conectar()
    con.executescript(SCHEMA)
    con.commit()
    con.close()


def get_antecedencia(con):
    row = con.execute("SELECT valor FROM config WHERE chave='antecedencia'").fetchone()
    return int(row["valor"]) if row else ANTECEDENCIA_PADRAO


def get_tv_chave(con):
    row = con.execute("SELECT valor FROM config WHERE chave='tv_chave'").fetchone()
    if row:
        return row["valor"]
    chave = secrets.token_urlsafe(9)
    con.execute("INSERT INTO config(chave, valor) VALUES ('tv_chave', ?)", (chave,))
    con.commit()
    return chave


def agora():
    return datetime.now().strftime("%Y-%m-%d %H:%M:%S")


# ---------- importacao ----------
def norm(texto):
    return " ".join(str(texto or "").split())


def parse_data(valor):
    if isinstance(valor, datetime):
        return valor
    texto = norm(valor)
    for fmt in ("%d/%m/%Y %H:%M:%S", "%d/%m/%Y %H:%M", "%Y-%m-%d %H:%M:%S"):
        try:
            return datetime.strptime(texto, fmt)
        except ValueError:
            pass
    return None


class _TabelaHTML(HTMLParser):
    """Extrai as linhas da maior tabela de um HTML (SisFrota exporta .xls assim)."""

    def __init__(self):
        super().__init__()
        self.tabelas, self._pilha = [], []

    def handle_starttag(self, tag, attrs):
        if tag == "table":
            self._pilha.append([])
        elif tag == "tr" and self._pilha:
            self._pilha[-1].append([])
        elif tag in ("td", "th") and self._pilha and self._pilha[-1]:
            self._pilha[-1][-1].append("")
        elif tag == "br" and self._pilha and self._pilha[-1] and self._pilha[-1][-1]:
            self._pilha[-1][-1][-1] += " "

    def handle_endtag(self, tag):
        if tag == "table" and self._pilha:
            self.tabelas.append(self._pilha.pop())

    def handle_data(self, data):
        if self._pilha and self._pilha[-1] and self._pilha[-1][-1]:
            self._pilha[-1][-1][-1] += data


def ler_html(caminho):
    bruto = open(caminho, "rb").read()
    try:
        texto = bruto.decode("utf-8")
    except UnicodeDecodeError:
        texto = bruto.decode("latin-1")
    p = _TabelaHTML()
    p.feed(texto)
    if not p.tabelas:
        raise ValueError("nenhuma tabela encontrada no arquivo")
    maior = max(p.tabelas, key=len)
    return iter([[norm(c) for c in linha] for linha in maior if linha])


def ler_linhas(caminho):
    """Devolve as linhas da planilha (xlsx, xls, html disfarçado de xls ou csv) como tuplas."""
    ext = os.path.splitext(caminho)[1].lower()
    with open(caminho, "rb") as f:
        inicio = f.read(512).lstrip().lower()
    if inicio.startswith((b"<", b"\xef\xbb\xbf<")):
        return ler_html(caminho)
    if ext in (".xlsx", ".xlsm"):
        ws = load_workbook(caminho, read_only=True, data_only=True).active
        return ws.iter_rows(values_only=True)
    if ext == ".xls":
        import xlrd
        ws = xlrd.open_workbook(caminho).sheet_by_index(0)
        wb_datemode = ws.book.datemode

        def valor(c):
            if c.ctype == xlrd.XL_CELL_DATE:
                return xlrd.xldate.xldate_as_datetime(c.value, wb_datemode)
            return c.value

        return ([valor(c) for c in ws.row(i)] for i in range(ws.nrows))
    import csv
    for enc in ("utf-8-sig", "latin-1"):
        try:
            with open(caminho, encoding=enc, newline="") as f:
                texto = f.read()
            break
        except UnicodeDecodeError:
            continue
    delim = ";" if texto.split("\n", 1)[0].count(";") >= texto.split("\n", 1)[0].count(",") else ","
    return iter(csv.reader(texto.splitlines(), delimiter=delim))


def importar_planilha(caminho):
    """Le a planilha do SisFrota. Retorna (novas, duplicadas, ignoradas)."""
    linhas = ler_linhas(caminho)
    cab = [norm(c) for c in next(linhas)]
    idx = {nome: i for i, nome in enumerate(cab)}

    def col(linha, nome):
        i = idx.get(nome)
        return linha[i] if i is not None and i < len(linha) else None

    con = db()
    novas = duplicadas = ignoradas = 0
    for linha in linhas:
        motorista = norm(col(linha, "Motorista")).upper()
        inicio = parse_data(col(linha, "DT_Inicio"))
        status = norm(col(linha, "DS_StatusSoli")).lower()
        if not motorista or not inicio or (status and not status.startswith("aprovad")):
            ignoradas += 1
            continue
        termino = parse_data(col(linha, "DT_Termino"))
        placa = norm(col(linha, "Placa")).upper()
        fim = f"{termino:%Y%m%d%H%M}" if termino else ""
        chave = hashlib.sha1(f"{motorista}|{placa}|{inicio:%Y%m%d%H%M}|{fim}".encode()).hexdigest()

        con.execute("INSERT OR IGNORE INTO motoristas(nome) VALUES (?)", (motorista,))
        mid = con.execute("SELECT id FROM motoristas WHERE nome=?", (motorista,)).fetchone()["id"]
        try:
            passageiros = int(col(linha, "Quant_Pass") or 0)
        except (TypeError, ValueError):
            passageiros = 0
        cur = con.execute(
            """INSERT OR IGNORE INTO viagens
               (chave, motorista_id, inicio, termino, placa, modelo, secretaria,
                solicitante, descricao, destino, passageiros)
               VALUES (?,?,?,?,?,?,?,?,?,?,?)""",
            (
                chave, mid, inicio.strftime("%Y-%m-%d %H:%M:%S"),
                termino.strftime("%Y-%m-%d %H:%M:%S") if termino else None,
                placa, norm(col(linha, "Modelo")), norm(col(linha, "Secretaria")),
                norm(col(linha, "NomeSolicitante")),
                norm(col(linha, "DS_SolicitacaoVeiculo")),
                norm(col(linha, "local")).rstrip("/ ").strip(),
                passageiros,
            ),
        )
        if cur.rowcount:
            novas += 1
        else:
            duplicadas += 1
    con.commit()
    return novas, duplicadas, ignoradas


# ---------- chamada automatica ----------
def chamar_pendentes():
    con = conectar()
    limite = (datetime.now() + timedelta(minutes=get_antecedencia(con))).strftime("%Y-%m-%d %H:%M:%S")
    con.execute(
        """UPDATE viagens SET status='chamado', chamado_em=?
           WHERE status='aguardando' AND inicio<=? AND substr(inicio,1,10)=?""",
        (agora(), limite, datetime.now().strftime("%Y-%m-%d")),
    )
    con.commit()
    con.close()


def agendador():
    while True:
        try:
            chamar_pendentes()
        except Exception as exc:  # o agendador nunca deve morrer
            print("erro no agendador:", exc)
        time.sleep(20)


# ---------- filtros de template ----------
@app.template_filter("hora")
def f_hora(valor):
    return valor[11:16] if valor else ""


@app.template_filter("dia")
def f_dia(valor):
    return datetime.strptime(valor[:10], "%Y-%m-%d").strftime("%d/%m/%Y") if valor else ""


app.jinja_env.globals["STATUS_ROTULO"] = STATUS_ROTULO


# ---------- despachante ----------
def viagens_do_dia(dia):
    return db().execute(
        """SELECT v.*, m.nome AS motorista FROM viagens v
           JOIN motoristas m ON m.id = v.motorista_id
           WHERE substr(v.inicio,1,10)=? ORDER BY v.inicio, m.nome""",
        (dia,),
    ).fetchall()


@app.route("/admin")
def admin():
    dia = request.args.get("dia") or datetime.now().strftime("%Y-%m-%d")
    return render_template(
        "admin.html", dia=dia, viagens=viagens_do_dia(dia),
        antecedencia=get_antecedencia(db()), msg=request.args.get("msg"),
        tv_chave=get_tv_chave(db()),
    )


@app.route("/admin/tabela")
def admin_tabela():
    dia = request.args.get("dia")
    return render_template("_tabela.html", viagens=viagens_do_dia(dia), dia=dia)


@app.post("/admin/importar")
def admin_importar():
    arq = request.files.get("planilha")
    ext = os.path.splitext(arq.filename)[1].lower() if arq else ""
    if ext not in (".xlsx", ".xlsm", ".xls", ".csv"):
        return redirect(url_for("admin", msg=f"Formato '{ext or 'desconhecido'}' não suportado. Envie .xlsx, .xls ou .csv."))
    caminho = os.path.join(BASE, "uploads", f"{datetime.now():%Y%m%d_%H%M%S}{ext}")
    arq.save(caminho)
    try:
        novas, dup, ign = importar_planilha(caminho)
    except Exception as exc:
        return redirect(url_for("admin", msg=f"Não consegui ler a planilha: {exc}"))
    primeira = db().execute("SELECT MIN(substr(inicio,1,10)) d FROM viagens").fetchone()["d"]
    return redirect(url_for(
        "admin", dia=primeira or "",
        msg=f"Importação concluída: {novas} novas, {dup} duplicadas, {ign} ignoradas.",
    ))


@app.post("/admin/config")
def admin_config():
    minutos = max(0, min(240, int(request.form.get("antecedencia", ANTECEDENCIA_PADRAO))))
    db().execute("INSERT OR REPLACE INTO config(chave, valor) VALUES ('antecedencia', ?)", (str(minutos),))
    db().commit()
    return redirect(url_for("admin", dia=request.form.get("dia"), msg="Antecedência atualizada."))


@app.post("/admin/chamar/<int:vid>")
def admin_chamar(vid):
    db().execute(
        "UPDATE viagens SET status='chamado', chamado_em=? WHERE id=? AND status='aguardando'",
        (agora(), vid),
    )
    db().commit()
    v = db().execute("SELECT substr(inicio,1,10) d FROM viagens WHERE id=?", (vid,)).fetchone()
    return admin_tabela_dia(v["d"])


@app.post("/admin/ciente/<int:vid>")
def admin_ciente(vid):
    """Despachante confirma o ciente de um motorista presente na garagem (sem celular)."""
    db().execute(
        "UPDATE viagens SET status='ciente', ciente_em=? WHERE id=? AND status='chamado'",
        (agora(), vid),
    )
    db().commit()
    v = db().execute("SELECT substr(inicio,1,10) d FROM viagens WHERE id=?", (vid,)).fetchone()
    return admin_tabela_dia(v["d"])


def admin_tabela_dia(dia):
    return render_template("_tabela.html", viagens=viagens_do_dia(dia), dia=dia)


@app.route("/admin/imprimir/<int:vid>")
def admin_imprimir(vid):
    v = db().execute(
        "SELECT v.*, m.nome AS motorista FROM viagens v JOIN motoristas m ON m.id=v.motorista_id WHERE v.id=?",
        (vid,),
    ).fetchone()
    return render_template("imprimir.html", v=v)


# ---------- TV (somente exibicao) ----------
def tv_autorizada():
    return secrets.compare_digest(request.args.get("chave", ""), get_tv_chave(db()))


def dados_tv():
    dia = datetime.now().strftime("%Y-%m-%d")
    todas = [v for v in viagens_do_dia(dia) if v["status"] != "concluido"]
    return [v for v in todas if v["status"] == "chamado"], [v for v in todas if v["status"] != "chamado"]


@app.route("/tv")
def tv():
    if not tv_autorizada():
        return "Acesso negado.", 403
    chamados, proximos = dados_tv()
    return render_template("tv.html", chamados=chamados, proximos=proximos, chave=request.args["chave"])


@app.route("/tv/painel")
def tv_painel():
    if not tv_autorizada():
        return "Acesso negado.", 403
    chamados, proximos = dados_tv()
    return render_template("_tv.html", chamados=chamados, proximos=proximos)


# ---------- motorista (celular) ----------
@app.route("/")
def inicio():
    motoristas = db().execute("SELECT * FROM motoristas ORDER BY nome").fetchall()
    return render_template("inicio.html", motoristas=motoristas)


def viagens_motorista(mid):
    hoje = datetime.now().strftime("%Y-%m-%d")
    return db().execute(
        """SELECT * FROM viagens WHERE motorista_id=? AND substr(inicio,1,10)=?
           ORDER BY inicio""",
        (mid, hoje),
    ).fetchall()


@app.route("/m/<int:mid>")
def motorista(mid):
    m = db().execute("SELECT * FROM motoristas WHERE id=?", (mid,)).fetchone()
    if not m:
        return redirect(url_for("inicio"))
    return render_template("motorista.html", m=m, viagens=viagens_motorista(mid))


@app.route("/m/<int:mid>/painel")
def motorista_painel(mid):
    return render_template("_painel.html", mid=mid, viagens=viagens_motorista(mid))


@app.post("/m/<int:mid>/viagem/<int:vid>/<acao>")
def motorista_acao(mid, vid, acao):
    if acao == "ciente":
        db().execute(
            "UPDATE viagens SET status='ciente', ciente_em=? WHERE id=? AND motorista_id=? AND status='chamado'",
            (agora(), vid, mid),
        )
    elif acao == "concluir":
        db().execute(
            "UPDATE viagens SET status='concluido', concluido_em=? WHERE id=? AND motorista_id=? AND status IN ('chamado','ciente')",
            (agora(), vid, mid),
        )
    db().commit()
    return render_template("_painel.html", mid=mid, viagens=viagens_motorista(mid))


iniciar_db()

if __name__ == "__main__":
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true" or not app.debug:
        threading.Thread(target=agendador, daemon=True).start()
    app.run(host="0.0.0.0", port=5000)
