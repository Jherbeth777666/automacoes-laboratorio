#!/usr/bin/env python3
"""
separar_laudos.py — Laboratório Municipal de Itapecuru-Mirim
Junta tudo num passo só: lê os LAUDOS (PDF do Worklab), os AGENDADOS e os ATENDIDOS,
descobre a UBS de cada paciente, gera 1 PDF por paciente na pasta da sua UBS e
uma planilha de conferência. Opcionalmente imprime (1 paciente = 1 trabalho).

COMO USAR
  1. Coloque na pasta "entrada":
       - o PDF com os laudos           (qualquer nome .pdf)
       - a exportação de agendados     (nome contendo "agend":  agendados.pdf / .xlsx / .csv)
       - a exportação de atendidos     (nome contendo "atend":  atendidos.pdf / .xlsx / .csv)  [opcional]
  2. Rode:  python separar_laudos.py            -> separa e gera a conferência
            python separar_laudos.py --imprimir -> separa e imprime
     (ou dê dois cliques em IMPRIMIR_LAUDOS.bat)

DE ONDE VEM A UBS (nesta ordem):
  1. Do próprio laudo, se ele tiver "Procedência / Posto de coleta / Unidade / Origem"
  2. Dos agendados, pelo nº de atendimento/pedido/prontuário
  3. Dos agendados, pelo nome (aceita pequena diferença de digitação)
  4. Dos atendidos, se a planilha tiver coluna de UBS

Requisitos: Python 3.9+  e  pip install pypdf pdfplumber reportlab openpyxl
Impressão no Windows: SumatraPDF (gratuito).
"""

import argparse
import csv
import datetime as dt
import difflib
import io
import re
import shutil
import subprocess
import sys
import unicodedata
from pathlib import Path

from pypdf import PdfReader, PdfWriter

# ======================== CONFIGURAÇÃO ========================
PASTA_ENTRADA = Path("entrada")
PASTA_SAIDA = Path(__file__).resolve().parent / "saida"

# Rótulos no laudo do Worklab (ajuste quando tiver um laudo real em mãos)
REGEX_NOME = re.compile(r"Paciente\s*[:\-]?\s*([A-ZÀ-Ú][A-ZÀ-Ú '\.]{3,})", re.IGNORECASE)
REGEX_CODIGO = re.compile(r"(?:Protocolo|Atendimento|Prontu[aá]rio|Pedido|Requisi[cç][aã]o)\s*[:N\.ºo\-]*\s*(\d{3,})", re.IGNORECASE)
# Só procurado no CABEÇALHO do laudo (antes de "Data Coleta"). "Unidade:" sozinho NÃO conta:
# no corpo do exame é a unidade de medida (mg/dL...).
REGEX_UBS_LAUDO = re.compile(r"(?:Proced[eê]ncia|Posto de coleta|Unidade de (?:coleta|sa[uú]de)|Unidade solicitante|Local de coleta)\s*[:\-]\s*([^\n]{3,60})", re.IGNORECASE)
# Worklab de Itapecuru-Mirim:  "Paciente  : NOME DO PACIENTE Protocolo : 1234567"
# (no PDF real o nome vem COLADO: "...DOS SANTOSProtocolo", por isso Protocolo sem \b)
CORTES_NOME = r"\s{2,}|\n|Protocolo|(?:Idade|Sexo|Data|CNS|CPF|Atendimento|Prontu\w*|Pedido|Proced\w*|Posto|Unidade|Origem)\s*:"

SIMILARIDADE_MINIMA = 0.92   # 92% parecido = considera o mesmo nome
SEM_UBS = "_SEM_UBS_CONFERIR"

SUMATRA = r"C:\Program Files\SumatraPDF\SumatraPDF.exe"
IMPRESSORA = None            # None = padrão; ou ex.: "Brother HL-L5102DW"
PRINT_SETTINGS = "simplex"   # ou "duplex"
# ==============================================================


def norm(txt) -> str:
    s = unicodedata.normalize("NFKD", str(txt or "")).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip(" .:-").upper()


def nome_arquivo(txt) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "_", norm(txt)).strip("_")[:60]


def mesmo_codigo(a, b) -> bool:
    return bool(a) and bool(b) and str(a).lstrip("0") == str(b).lstrip("0")


def parecido(a, b) -> float:
    return difflib.SequenceMatcher(None, a, b).ratio()


# ------------------------- LAUDOS (PDF) -------------------------
def ler_laudos(pdf: Path):
    reader = PdfReader(str(pdf))
    grupos, ordem, sem_id, atual = {}, [], [], None
    for i, pag in enumerate(reader.pages):
        texto = (pag.extract_text() or "").replace("\xa0", " ")
        cabecalho = re.split(r"Data Coleta", texto, flags=re.IGNORECASE)[0]
        m = REGEX_NOME.search(cabecalho)
        if m:
            nome = norm(re.split(CORTES_NOME, m.group(1), flags=re.IGNORECASE)[0])
            mc = REGEX_CODIGO.search(cabecalho)
            codigo = mc.group(1) if mc else ""
            mu = REGEX_UBS_LAUDO.search(cabecalho)
            ubs_laudo = re.split(r"\s{2,}", mu.group(1).strip())[0] if mu else ""
            chave = f"{codigo}|{nome}"
            if chave not in grupos:
                grupos[chave] = {"nome": nome, "codigo": codigo, "ubs_laudo": ubs_laudo, "paginas": []}
                ordem.append(chave)
            atual = chave
        elif atual is None:
            sem_id.append(i + 1)
            continue
        grupos[atual]["paginas"].append(i)   # sem cabeçalho = continuação
    return reader, [grupos[k] for k in ordem], sem_id


# ------------------- PLANILHAS (agendados/atendidos) -------------------
COLS = {
    "nome":   ("PACIENTE", "NOME"),
    "codigo": ("ATEND", "PEDIDO", "PRONT", "REQUISI", "CODIGO", "N "),
    "ubs":    ("UBS", "UNIDADE", "POSTO", "PROCEDENCIA", "ORIGEM", "LOCAL", "PSF", "ESF", "SOLICITANTE"),
    # comparecimento (quando o relatório trouxer): "Compareceu", "Faltou", "Presente"...
    "situacao": ("SITUA", "STATUS", "COMPAREC", "PRESEN", "FALT"),
}


def ler_tabela(caminho: Path):
    """Lê .pdf/.xlsx/.csv/.txt e devolve [{nome, codigo, ubs}]. Acha o cabeçalho sozinho."""
    if caminho is None:
        return []
    ext = caminho.suffix.lower()
    if ext == ".pdf":
        return ler_relatorio_pdf(caminho)
    if ext in (".xlsx", ".xlsm"):
        from openpyxl import load_workbook
        ws = load_workbook(caminho, read_only=True, data_only=True).active
        tabela = [["" if c is None else str(c).strip() for c in r] for r in ws.iter_rows(values_only=True)]
    elif ext in (".csv", ".txt"):
        txt = caminho.read_text(encoding="utf-8-sig", errors="ignore")
        if ext == ".txt" and not re.search(r"[;\t,]", txt.splitlines()[0] if txt else ""):
            return ler_txt_blocos(txt)
        try:
            dialeto = csv.Sniffer().sniff(txt[:3000], delimiters=";,\t")
        except csv.Error:
            dialeto = csv.excel
        tabela = [[c.strip() for c in r] for r in csv.reader(io.StringIO(txt), dialeto)]
    else:
        sys.exit(f"Formato não suportado: {caminho.name} (use .pdf, .xlsx, .csv ou .txt)")
    regs = tabela_para_registros(tabela)
    if regs is None:
        sys.exit(f"Não achei a coluna de PACIENTE/NOME em {caminho.name}.")
    return regs


def ler_relatorio_pdf(caminho: Path):
    """Relatório do Worklab em PDF (atendidos/agendados).
    1º tenta ler como tabela; se não houver tabela, lê linha a linha."""
    import pdfplumber
    tabela, linhas = [], []
    with pdfplumber.open(str(caminho)) as pdf:
        for pag in pdf.pages:
            for t in pag.extract_tables():
                tabela += [[(c or "").replace("\n", " ").strip() for c in r] for r in t]
            linhas += (pag.extract_text() or "").splitlines()
    regs = tabela_para_registros(tabela) if tabela else None
    if regs:
        return regs
    regs = linhas_para_registros(linhas)
    if not regs:
        sys.exit(f"Não consegui ler pacientes em {caminho.name}. Se for PDF escaneado (imagem), "
                 f"exporte em Excel/CSV ou me mande um exemplo para ajustar.")
    return regs


RX_DATA = re.compile(r"\b\d{1,2}/\d{1,2}(?:/\d{2,4})?\b|\b\d{1,2}:\d{2}\b")
RX_UBS_TXT = re.compile(r"\b(?:UBS|PSF|ESF|USF|POSTO|UNIDADE|CS)\b.*", re.IGNORECASE)
RX_LIXO = re.compile(r"RELAT|P[AÁ]GINA|EMITIDO|TOTAL|LABORAT|PREFEITURA|SECRETARIA|IMPRESS", re.IGNORECASE)


def linhas_para_registros(linhas):
    """Relatório em texto corrido. Entende linhas como
         10231  MARIA DA SILVA SOUZA   14/10/2026   UBS CENTRO
       e também blocos com a UBS como título e os pacientes embaixo."""
    regs, ubs_bloco = [], ""
    for bruta in linhas:
        l = bruta.strip()
        if not l or RX_LIXO.search(l):
            continue
        palavras_cab = re.findall(r"\b(NOME|PACIENTE|ATEND\w*|DATA|UNIDADE|IDADE|SEXO|PRONT\w*|PEDIDO|CNS|CPF)\b", norm(l))
        if len(set(palavras_cab)) >= 2:   # linha de cabeçalho da tabela
            continue
        sem_data = RX_DATA.sub(" ", l)
        mu = RX_UBS_TXT.search(sem_data)
        ubs = mu.group(0).strip() if mu else ""
        resto = sem_data[:mu.start()] if mu else sem_data
        # linha que é só o nome da UBS = título de bloco
        if ubs and not re.search(r"[A-Za-zÀ-ú]{2,}\s+[A-Za-zÀ-ú]{2,}", resto):
            ubs_bloco = ubs
            continue
        mc = re.search(r"\b\d{3,}\b", resto)
        codigo = mc.group(0) if mc else ""
        # nome = maior sequência de 2+ palavras (só letras)
        candidatos = re.findall(r"[A-Za-zÀ-ú'][A-Za-zÀ-ú'\.]*(?:\s+[A-Za-zÀ-ú'][A-Za-zÀ-ú'\.]*)+", resto)
        candidatos = [c for c in candidatos if not re.fullmatch(r"(?i)(nome|paciente|atendimento|data|unidade|idade)(\s.*)?", c.strip())]
        if not candidatos:
            continue
        nome = norm(max(candidatos, key=len))
        if len(nome.split()) < 2 or nome in ("NOME DO PACIENTE", "NOME PACIENTE"):
            continue
        regs.append({"nome": nome, "codigo": codigo, "ubs": ubs or ubs_bloco})
    return regs


def tabela_para_registros(tabela):
    # procura a linha de cabeçalho nas 15 primeiras linhas
    for h, linha in enumerate(tabela[:15]):
        cab = [norm(c) for c in linha]
        idx = {}
        for campo, chaves in COLS.items():
            for j, c in enumerate(cab):
                if j not in idx.values() and any(c.startswith(k) or f" {k}" in f" {c}" for k in chaves):
                    if campo == "nome" and "MAE" in c:      # ignora "Nome da mãe"
                        continue
                    idx[campo] = j
                    break
        if "nome" in idx:
            break
    else:
        return None

    regs, cab_norm = [], [norm(c) for c in tabela[h]]
    for r in tabela[h + 1:]:
        if [norm(c) for c in r] == cab_norm:   # cabeçalho repetido em cada página do PDF
            continue
        pega = lambda k: r[idx[k]] if k in idx and idx[k] < len(r) else ""
        if norm(pega("nome")):
            ubs = pega("ubs").strip()
            regs.append({"nome": norm(pega("nome")), "codigo": pega("codigo"),
                         "ubs": "" if ubs in ("-", "—", "–", "0", "N/A", "NA") else ubs,
                         "situacao": norm(pega("situacao"))})
    return regs


def ler_txt_blocos(txt):
    regs, ubs = [], ""
    for bruta in txt.splitlines():
        l = re.sub(r"^\s*\d+[\.\)\-]\s*|^\s*[-•*]\s*", "", bruta).strip()
        if not l:
            continue
        if re.match(r"^(UBS|PSF|ESF|USF|POSTO|UNIDADE|CS)\b", norm(l)) or l.endswith(":"):
            ubs = l.rstrip(":").strip()
        else:
            regs.append({"nome": norm(l), "codigo": "", "ubs": ubs})
    return regs


PARTICULAS = {"DA", "DE", "DO", "DAS", "DOS", "E", "D"}


def palavras(nome):
    return [t.strip(".") for t in nome.replace(".", ". ").split() if t.strip(".") and t.strip(".") not in PARTICULAS]


def palavra_ok(a, b):
    if len(a) == 1 or len(b) == 1:        # inicial: "C." casa com "COSTA"
        return a[0] == b[0]
    return parecido(a, b) >= 0.8


def casa_palavras(x, y):
    """Mesmo nome com sobrenome faltando ou abreviado:
       'ANA PAULA SOUZA' ~ 'ANA PAULA SOUZA LIMA', 'MARIA CLARA R. DE OLIVEIRA' ~ '... RIBEIRO DE OLIVEIRA'.
       Exige o 1º nome igual e TODAS as palavras do nome menor, na ordem."""
    A, B = palavras(x), palavras(y)
    curto, longo = sorted((A, B), key=len)
    if len(curto) < 2 or sum(len(t) > 1 for t in curto) < 2 or not palavra_ok(curto[0], longo[0]):
        return False
    j = 1
    for t in curto[1:]:
        k = next((k for k in range(j, len(longo)) if palavra_ok(t, longo[k])), None)
        if k is None:
            return False
        j = k + 1
    return True


def procurar(paciente, registros):
    """Acha o registro do paciente numa lista: código > nome exato > nome parecido."""
    for r in registros:
        if mesmo_codigo(r["codigo"], paciente["codigo"]):
            return r, "nº atendimento"
    exatos = [r for r in registros if r["nome"] == paciente["nome"]]
    if exatos:
        if len({norm(r["ubs"]) for r in exatos if r["ubs"]}) > 1:
            return None, "homônimo em UBS diferentes"
        return exatos[0], "nome"
    melhor = max(registros, key=lambda r: parecido(r["nome"], paciente["nome"]), default=None)
    if melhor and parecido(melhor["nome"], paciente["nome"]) >= SIMILARIDADE_MINIMA:
        return melhor, f"nome parecido: {melhor['nome']}"
    candidatos = [r for r in registros if casa_palavras(r["nome"], paciente["nome"])]
    if candidatos:
        if len({norm(r["ubs"]) for r in candidatos}) > 1 or len({r["nome"] for r in candidatos}) > 1:
            return None, "nome incompleto, mais de um candidato: " + " / ".join(sorted({r["nome"] for r in candidatos}))
        return candidatos[0], f"nome incompleto/abreviado: {candidatos[0]['nome']}"
    return None, ""


# ----------------------------- SAÍDA -----------------------------
def gerar_capa(nome, codigo, ubs, qtd):
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas
    buf = io.BytesIO()
    c = canvas.Canvas(buf, pagesize=A4)
    w, h = A4
    c.setFont("Helvetica-Bold", 16); c.drawCentredString(w / 2, h - 80, ubs.upper())
    c.setFont("Helvetica-Bold", 22); c.drawCentredString(w / 2, h / 2 + 20, nome[:40])
    c.setFont("Helvetica", 14)
    if codigo:
        c.drawCentredString(w / 2, h / 2 - 15, f"Protocolo: {codigo}")
    c.drawCentredString(w / 2, h / 2 - 40, f"{qtd} página(s) de laudo")
    c.setFont("Helvetica", 9)
    c.drawCentredString(w / 2, 50, "Documento com dados de saúde — entregar somente ao paciente ou responsável")
    c.showPage(); c.save(); buf.seek(0)
    return buf


def planilha_conferencia(destino: Path, separados, sem_laudo, avisos):
    from openpyxl import Workbook
    from openpyxl.styles import Font, PatternFill
    wb = Workbook()
    abas = [
        ("Separados", ["UBS", "Paciente", "Protocolo", "Págs no PDF", "Qtd págs", "Origem da UBS", "Atendido?", "Arquivo"], separados),
        ("Sem laudo ainda", ["UBS", "Paciente", "Protocolo", "Situação"], sem_laudo),
        ("CONFERIR", ["Problema", "Paciente", "Protocolo", "Detalhe"], avisos),
    ]
    for n, (titulo, cab, linhas) in enumerate(abas):
        ws = wb.active if n == 0 else wb.create_sheet()
        ws.title = titulo
        ws.append(cab)
        for c in ws[1]:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="C0392B" if titulo == "CONFERIR" else "2E6DA4")
        for l in linhas:
            ws.append(l)
        for col in ws.columns:
            ws.column_dimensions[col[0].column_letter].width = min(45, max(10, *(len(str(c.value or "")) + 2 for c in col)))
        ws.freeze_panes = "A2"
    wb.save(destino)


def imprimir(arquivo: Path):
    if sys.platform.startswith("win"):
        cmd = [SUMATRA] + (["-print-to", IMPRESSORA] if IMPRESSORA else ["-print-to-default"])
        cmd += ["-silent", "-print-settings", PRINT_SETTINGS, str(arquivo)]
    else:
        cmd = ["lp"] + (["-d", IMPRESSORA] if IMPRESSORA else []) + [str(arquivo)]
    subprocess.run(cmd, check=True)


# ----------------------------- FLUXO -----------------------------
EXT_LISTA = (".xlsx", ".xlsm", ".csv", ".txt")


def pasta_padrao() -> Path:
    """Usa a pasta 'entrada' se existir; senão, a própria pasta do script."""
    entrada = Path(__file__).resolve().parent / PASTA_ENTRADA
    return entrada if entrada.is_dir() else Path(__file__).resolve().parent


def eh_laudo(pdf: Path) -> bool:
    """PDF de laudo = tem o rótulo de paciente nas primeiras páginas."""
    try:
        r = PdfReader(str(pdf))
        return any(REGEX_NOME.search(p.extract_text() or "") for p in r.pages[:3])
    except Exception:
        return False


def amostra_texto(arq: Path) -> str:
    """Primeiras linhas do arquivo, só para descobrir se é agendados ou atendidos."""
    try:
        if arq.suffix.lower() == ".pdf":
            return norm(PdfReader(str(arq)).pages[0].extract_text() or "")[:1500]
        if arq.suffix.lower() in (".xlsx", ".xlsm"):
            from openpyxl import load_workbook
            ws = load_workbook(arq, read_only=True, data_only=True).active
            return norm(" ".join(str(c) for r in ws.iter_rows(max_row=8, values_only=True) for c in r if c))
        return norm(arq.read_text(encoding="utf-8-sig", errors="ignore")[:1500])
    except Exception:
        return ""


def achar_arquivos(pasta: Path):
    """Identifica cada arquivo pelo nome e, se o nome não disser nada, pelo conteúdo."""
    arqs = sorted(p for p in pasta.iterdir() if p.is_file() and not p.name.startswith(("~", "."))
                  and p.suffix.lower() in (".pdf",) + EXT_LISTA and p.name != Path(__file__).name)
    laudos, agend, atend, indefinidos = [], None, None, []
    for p in arqs:
        nome = p.name.lower()
        if "agend" in nome and not agend:
            agend = p
        elif "atend" in nome and not atend:
            atend = p
        elif p.suffix.lower() == ".pdf" and eh_laudo(p):
            laudos.append(p)
        elif p.suffix.lower() == ".txt":
            continue  # .txt só entra se tiver "agend"/"atend" no nome
        else:
            indefinidos.append(p)
    # Relatórios sem pista no nome: decide pelo título do relatório
    for p in indefinidos:
        txt = amostra_texto(p)
        if "AGEND" in txt and not agend:
            agend = p
        elif "ATENDID" in txt and not atend:
            atend = p
        elif not agend:
            agend = p      # a 1ª lista sem título vira a fonte das UBS
        elif not atend:
            atend = p
    return laudos, agend, atend


def mascarar(linha: str) -> str:
    """Esconde nomes e números para o diagnóstico poder ser compartilhado sem dados de paciente."""
    rotulos = {"PACIENTE", "NOME", "DATA", "IDADE", "SEXO", "ATENDIMENTO", "PEDIDO", "PRONTUARIO",
               "PROCEDENCIA", "POSTO", "UNIDADE", "ORIGEM", "MEDICO", "SOLICITANTE", "EXAME", "RESULTADO",
               "MATERIAL", "METODO", "VALORES", "REFERENCIA", "LABORATORIO", "NASCIMENTO", "COLETA",
               "LAUDO", "LIBERADO", "LIBERACAO", "CONVENIO", "REQUISICAO", "CADASTRO", "ENTRADA", "EMISSAO",
               "MUNICIPAL", "SECRETARIA", "SAUDE", "PREFEITURA", "PAGINA", "CODIGO", "ANOS", "MESES", "DIAS", "PROTOCOLO", "OBSERVACAO", "LOCAL", "HORARIO", "TIPO"}
    def troca(m):
        w = m.group(0)
        if w.isdigit():
            return "#" * len(w)
        if len(w) <= 3 or norm(w).rstrip(":") in rotulos:
            return w
        return w[0] + "*" * (len(w) - 1)
    return re.sub(r"[A-Za-zÀ-ú]+|\d+", troca, linha)


def diagnostico(pasta: Path):
    print(f"Pasta lida: {pasta}\n")
    laudos, agend, atend = achar_arquivos(pasta)
    for p in sorted(p for p in pasta.iterdir() if p.is_file() and p.suffix.lower() in (".pdf",) + EXT_LISTA):
        papel = ("LAUDOS" if p in laudos else "AGENDADOS" if p == agend else
                 "ATENDIDOS" if p == atend else "ignorado")
        print(f"  {p.name:<40} -> {papel}")
    for p in sorted(p for p in pasta.glob("*.pdf") if p not in laudos and p not in (agend, atend)):
        laudos.append(p)  # mostra também PDFs não reconhecidos, para ajustar o rótulo
    for pdf in laudos:
        r = PdfReader(str(pdf))
        achou = sum(1 for pg in r.pages if REGEX_NOME.search(pg.extract_text() or ""))
        print(f"\n== {pdf.name}: {len(r.pages)} páginas, rótulo de paciente achado em {achou}")
        texto = (r.pages[0].extract_text() or "").splitlines()
        if not texto:
            print("   (página sem texto: PDF escaneado/imagem — vai precisar de OCR)")
        print("   Início da 1ª página (nomes e números escondidos):")
        for l in texto[:30]:
            print("   | " + mascarar(l))
    for f in (agend, atend):
        if f and f.suffix.lower() != ".pdf":
            regs = ler_tabela(f)
            print(f"\n== {f.name}: {len(regs)} pacientes lidos; "
                  f"{sum(1 for r in regs if r['ubs'])} com UBS, {sum(1 for r in regs if r['codigo'])} com nº")
            print("   Primeiras linhas (mascaradas):")
            print("   | " + mascarar(amostra_texto(f))[:600])
        elif f:
            regs = ler_tabela(f)
            print(f"\n== {f.name}: {len(regs)} pacientes lidos; {sum(1 for r in regs if r['ubs'])} com UBS")
    print("\nPode copiar e colar este diagnóstico no chat: nomes e números estão escondidos.")


def main():
    ap = argparse.ArgumentParser(description="Separa laudos do Worklab por paciente e UBS.")
    ap.add_argument("--entrada", type=Path, default=None, help="Pasta dos arquivos (padrão: 'entrada' ou a pasta do script)")
    ap.add_argument("--diagnostico", action="store_true", help="Mostra como os arquivos foram lidos, sem dados de paciente")
    ap.add_argument("--imprimir", action="store_true", help="Envia cada paciente para a impressora")
    ap.add_argument("--perguntar", action="store_true", help="Separa, mostra o resumo e pergunta se imprime")
    ap.add_argument("--sem-capa", action="store_true", help="Não colocar folha separadora")
    ap.add_argument("--forcar", action="store_true", help="Imprime mesmo com pendências em CONFERIR")
    args = ap.parse_args()
    args.entrada = args.entrada or pasta_padrao()
    if not args.entrada.is_dir():
        sys.exit(f"Pasta não encontrada: {args.entrada}")
    if args.diagnostico:
        try:
            diagnostico(args.entrada)
        except SystemExit as e:
            print(f"\n!! {e}")
        return

    pdfs, f_agend, f_atend = achar_arquivos(args.entrada)
    if not pdfs:
        sys.exit(f"Não achei PDF de laudo em '{args.entrada}' (procurei o rótulo 'Paciente:').\n"
                 f"Rode:  python separar_laudos.py --diagnostico   e me mande o resultado.")
    if not f_agend and not f_atend:
        print("Aviso: nenhuma lista de agendados/atendidos encontrada — todos irão para _SEM_UBS_CONFERIR,\n"
              "a menos que o próprio laudo traga a UBS (Procedência/Posto).")
    agendados, atendidos = ler_tabela(f_agend), ler_tabela(f_atend)
    print(f"Laudos: {', '.join(p.name for p in pdfs)}")
    print(f"Agendados: {f_agend.name if f_agend else '—'} ({len(agendados)})  |  "
          f"Atendidos: {f_atend.name if f_atend else '—'} ({len(atendidos)})")

    # Grafia padrão de cada UBS (a primeira que aparece nos agendados), para que
    # "UBS PIQUIÁ", "Ubs Piquia" e "UBS Piquiá" virem uma pasta só.
    grafia = {}
    for r in agendados + atendidos:
        if r["ubs"]:
            grafia.setdefault(norm(r["ubs"]), r["ubs"])
    padrao = lambda u: u if u == SEM_UBS else grafia.setdefault(norm(u), u)

    raiz = PASTA_SAIDA / dt.datetime.now().strftime("%Y-%m-%d_%Hh%M")
    separados, avisos, gerados, todos_pacientes = [], [], [], []
    contador = {}

    for pdf in pdfs:
        reader, pacientes, sem_id = ler_laudos(pdf)
        if sem_id:
            avisos.append(["Página sem paciente", "", "", f"{pdf.name}: págs {sem_id} não foram separadas"])
        if not pacientes:
            avisos.append(["Nenhum paciente lido", "", "", f"{pdf.name}: ajustar REGEX_NOME no script"])

        for p in pacientes:
            # ---- descobre a UBS ----
            reg_ag, como_ag = procurar(p, agendados) if agendados else (None, "")
            reg_at, como_at = procurar(p, atendidos) if atendidos else (None, "")
            if p["ubs_laudo"]:
                ubs, origem = p["ubs_laudo"], "laudo"
                if reg_ag and reg_ag["ubs"] and norm(reg_ag["ubs"]) != norm(ubs):
                    avisos.append(["UBS diferente laudo × agendado", p["nome"], p["codigo"],
                                   f"laudo: {ubs} / agendado: {reg_ag['ubs']} (usado o do laudo)"])
            elif reg_ag and reg_ag["ubs"]:
                ubs, origem = reg_ag["ubs"], f"agendados ({como_ag})"
            elif reg_at and reg_at["ubs"]:
                ubs, origem = reg_at["ubs"], f"atendidos ({como_at})"
            else:
                ubs = SEM_UBS
                origem = ("está nos agendados, mas sem UBS preenchida ('-')" if reg_ag
                          else como_ag or "não encontrado nos agendados")
                avisos.append(["Sem UBS", p["nome"], p["codigo"], origem])
            ubs = padrao(ubs)
            if "parecido" in origem or "incompleto" in origem:
                avisos.append(["Nome com diferença", p["nome"], p["codigo"], origem])
            atendido = "sim" if reg_at else ("—" if not atendidos else "NÃO consta")
            if atendidos and not reg_at:
                avisos.append(["Laudo sem registro de atendimento", p["nome"], p["codigo"], ubs])

            # ---- gera o PDF do paciente ----
            pasta = raiz / nome_arquivo(ubs)
            pasta.mkdir(parents=True, exist_ok=True)
            n = contador[ubs] = contador.get(ubs, 0) + 1
            w = PdfWriter()
            if not args.sem_capa and ubs != SEM_UBS:
                try:
                    w.append(gerar_capa(p["nome"], p["codigo"], ubs, len(p["paginas"])))
                except ImportError:
                    print("Aviso: 'reportlab' não instalado — seguindo SEM folha de capa.\n"
                          "       Para ter a capa: pip install reportlab")
                    args.sem_capa = True
            for i in p["paginas"]:
                w.add_page(reader.pages[i])
            arq = pasta / f"{n:03d}_{nome_arquivo(p['nome'])}{'_' + p['codigo'] if p['codigo'] else ''}.pdf"
            if arq.exists():  # segurança extra: nunca sobrescrever o laudo de outro paciente
                arq = arq.with_stem(f"{arq.stem}_{len(gerados) + 1}")
            with open(arq, "wb") as f:
                w.write(f)
            gerados.append((ubs, arq))
            todos_pacientes.append(p)
            separados.append([ubs, p["nome"], p["codigo"], ",".join(str(i + 1) for i in p["paginas"]),
                              len(p["paginas"]), origem, atendido, str(arq.relative_to(raiz))])

    # ---- quem foi atendido/agendado e ainda não tem laudo ----
    sem_laudo, vistos = [], set()
    for fonte, regs in (("Atendido, laudo ainda não saiu", atendidos), ("Agendado, sem laudo", agendados)):
        for r in regs:
            chave = r["codigo"] or r["nome"]
            if chave in vistos:
                continue
            achou = any(mesmo_codigo(r["codigo"], p["codigo"]) or parecido(r["nome"], p["nome"]) >= SIMILARIDADE_MINIMA
                        or casa_palavras(r["nome"], p["nome"]) for p in todos_pacientes)
            if not achou:
                vistos.add(chave)
                sit = r.get("situacao", "")
                if "FALT" in sit or "AUSENT" in sit or sit in ("NAO", "N"):
                    fonte = "Faltou (não colheu)"
                elif "COMPAREC" in sit or "PRESENT" in sit or sit in ("SIM", "S"):
                    fonte = "Compareceu, laudo ainda não saiu"
                ubs = r["ubs"] or next((a["ubs"] for a in agendados if a["nome"] == r["nome"]), "")
                sem_laudo.append([padrao(ubs) if ubs else "", r["nome"], r["codigo"], fonte])

    separados.sort(key=lambda l: (l[0] == SEM_UBS, norm(l[0]), l[1]))
    sem_laudo.sort(key=lambda l: (norm(l[0]), l[1]))
    raiz.mkdir(parents=True, exist_ok=True)
    planilha_conferencia(raiz / "CONFERENCIA.xlsx", separados, sem_laudo, avisos)

    # ---- resumo na tela ----
    print(f"\nPDFs gerados em: {raiz}")
    for ubs in sorted(contador, key=lambda u: (u == SEM_UBS, norm(u))):
        print(f"  {ubs:<30} {contador[ubs]:>3} paciente(s)")
    print(f"\n  Atendidos/agendados ainda sem laudo: {len(sem_laudo)}")
    graves = [a for a in avisos if a[0] in ("Sem UBS", "Página sem paciente", "Nenhum paciente lido", "UBS diferente laudo × agendado")]
    print(f"  Itens para conferir: {len(avisos)}  (abra CONFERENCIA.xlsx, aba CONFERIR)")

    if args.perguntar and not args.imprimir:
        print(f"\n  Confira a pasta '{raiz}' e a planilha CONFERENCIA.xlsx (aba CONFERIR).")
        args.imprimir = input("  Digite S para IMPRIMIR agora (qualquer outra tecla sai): ").strip().upper() == "S"

    if args.imprimir:
        if graves and not args.forcar:
            print("\n  Impressão CANCELADA: há pacientes sem UBS ou páginas não identificadas. "
                  "Corrija ou use --forcar (os sem UBS nunca são impressos).")
            return
        n = 0
        for ubs, arq in sorted(gerados, key=lambda x: (norm(x[0]), x[1].name)):
            if ubs != SEM_UBS:
                imprimir(arq); n += 1
        print(f"\n  {n} pacientes impressos, separados por UBS.")
        proc = args.entrada / "processados" / raiz.name
        proc.mkdir(parents=True, exist_ok=True)
        for f in [*pdfs, f_agend, f_atend]:
            if f:
                shutil.move(str(f), proc / f.name)


if __name__ == "__main__":
    main()