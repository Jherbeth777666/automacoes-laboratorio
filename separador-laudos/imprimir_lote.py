#!/usr/bin/env python3
"""
imprimir_lote.py — Depois de separar os laudos, decide o que fazer com cada UBS:
  IMPRIMIR -> imprime os PDFs dos pacientes um por um (1 paciente = 1 trabalho na impressora)
  PDF      -> junta os pacientes daquela UBS num PDF único em "para_enviar", para mandar por WhatsApp/e-mail

O modo de cada UBS fica salvo em ubs_modo.txt (pergunta só na primeira vez que a UBS aparece).
Para mudar depois, é só editar ubs_modo.txt no Bloco de Notas.

Uso:
  python imprimir_lote.py                 -> usa a separação mais recente (pasta saida)
  python imprimir_lote.py "C:\\pasta"     -> imprime TODOS os PDFs dessa pasta, um por um
  python imprimir_lote.py --simular       -> mostra o que faria, sem imprimir nada
"""

import os
import re
import subprocess
import sys
import time
import unicodedata
from pathlib import Path

from pypdf import PdfWriter

AQUI = Path(__file__).resolve().parent
PASTA_SAIDA = AQUI / "saida"
ARQ_MODOS = AQUI / "ubs_modo.txt"
SEM_UBS = "SEM_UBS_CONFERIR"
PAUSA_ENTRE_PACIENTES = 1.5   # segundos: dá tempo da fila da impressora manter a ordem

# Mesma impressora configurada no separar_laudos.py (se existir)
try:
    sys.path.insert(0, str(AQUI))
    import separar_laudos as _cfg
    IMPRESSORA, PRINT_SETTINGS = _cfg.IMPRESSORA, _cfg.PRINT_SETTINGS
except Exception:
    IMPRESSORA, PRINT_SETTINGS = None, "simplex"

SIMULAR = "--simular" in sys.argv


def norm(t):
    s = unicodedata.normalize("NFKD", str(t)).encode("ascii", "ignore").decode()
    return re.sub(r"\s+", " ", s).strip().upper()


# ----------------------------- impressão -----------------------------
def achar_sumatra():
    candidatos = [
        Path(os.environ.get("ProgramFiles", r"C:\Program Files")) / "SumatraPDF" / "SumatraPDF.exe",
        Path(os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")) / "SumatraPDF" / "SumatraPDF.exe",
        Path(os.environ.get("LOCALAPPDATA", "")) / "SumatraPDF" / "SumatraPDF.exe",
        AQUI / "SumatraPDF.exe",   # versão portátil na mesma pasta
    ]
    return next((c for c in candidatos if c.is_file()), None)


SUMATRA = achar_sumatra() if sys.platform.startswith("win") else None


def imprimir(pdf: Path):
    if SIMULAR:
        print(f"      [simulado] imprimiria {pdf.name}")
        return
    if sys.platform.startswith("win"):
        if not SUMATRA:
            sys.exit("\nSumatraPDF não encontrado. Instale (gratuito): https://www.sumatrapdfreader.org\n"
                     "ou coloque o SumatraPDF.exe portátil nesta pasta.")
        cmd = [str(SUMATRA)] + (["-print-to", IMPRESSORA] if IMPRESSORA else ["-print-to-default"])
        cmd += ["-silent", "-print-settings", PRINT_SETTINGS, str(pdf)]
    else:
        cmd = ["lp"] + (["-d", IMPRESSORA] if IMPRESSORA else []) + [str(pdf)]
    subprocess.run(cmd, check=True)
    time.sleep(PAUSA_ENTRE_PACIENTES)


def imprimir_lista(pdfs, titulo):
    print(f"\n  Imprimindo {titulo}: {len(pdfs)} PDF(s)")
    for i, pdf in enumerate(pdfs, 1):
        imprimir(pdf)
        print(f"    {i:>3}/{len(pdfs)}  {pdf.name}")


# ------------------------------ modos ------------------------------
def ler_modos():
    modos = {}
    if ARQ_MODOS.exists():
        for linha in ARQ_MODOS.read_text(encoding="utf-8").splitlines():
            if "=" in linha and not linha.strip().startswith("#"):
                ubs, modo = linha.rsplit("=", 1)
                modos[norm(ubs)] = (ubs.strip(), "PDF" if "PDF" in modo.upper() else "IMPRIMIR")
    return modos


def salvar_modos(modos):
    linhas = ["# Modo de cada UBS:  IMPRIMIR  ou  PDF  (pode editar no Bloco de Notas)"]
    linhas += [f"{nome} = {modo}" for nome, modo in sorted(modos.values())]
    ARQ_MODOS.write_text("\n".join(linhas) + "\n", encoding="utf-8")


def nome_bonito(pasta: Path):
    return pasta.name.replace("_", " ")


# ------------------------------ fluxos ------------------------------
def pasta_avulsa(pasta: Path):
    """Imprime todos os PDFs de uma pasta qualquer, um por um."""
    pdfs = sorted(pasta.glob("*.pdf"))
    if not pdfs:
        sys.exit(f"Nenhum PDF em {pasta}")
    print(f"Pasta: {pasta}\n{len(pdfs)} PDF(s) serão impressos um por um.")
    if input("Digite S para imprimir: ").strip().upper() == "S":
        imprimir_lista(pdfs, pasta.name)


def separacao_recente():
    lotes = sorted((p for p in PASTA_SAIDA.glob("*") if p.is_dir()), key=lambda p: p.name)
    if not lotes:
        sys.exit("Nenhuma separação encontrada. Rode primeiro: python separar_laudos.py")
    lote = lotes[-1]
    ubs_pastas = sorted(p for p in lote.iterdir() if p.is_dir() and SEM_UBS not in p.name
                        and p.name != "para_enviar")
    sem_ubs = lote / SEM_UBS
    print(f"Separação: {lote.name}")

    modos = ler_modos()
    novos = False
    for p in ubs_pastas:
        chave = norm(nome_bonito(p))
        if chave not in modos:
            r = input(f"  UBS nova: {nome_bonito(p)} -> (I)mprimir ou (P)DF para enviar? ").strip().upper()
            modos[chave] = (nome_bonito(p), "PDF" if r.startswith("P") else "IMPRIMIR")
            novos = True
    if novos:
        salvar_modos(modos)
        print(f"  (escolhas salvas em {ARQ_MODOS.name})")

    plano = [(p, modos[norm(nome_bonito(p))][1], sorted(p.glob("*.pdf"))) for p in ubs_pastas]
    print("\n  #  UBS                                   Modo       Pacientes  Situação")
    for i, (p, modo, pdfs) in enumerate(plano, 1):
        feito = "já feito" if (p / ".concluido").exists() else ""
        print(f"  {i:<2} {nome_bonito(p)[:36]:<37} {modo:<10} {len(pdfs):>5}      {feito}")
    if sem_ubs.exists():
        print(f"\n  !! {len(list(sem_ubs.glob('*.pdf')))} paciente(s) em {SEM_UBS} — não entram aqui; confira à mão.")

    esc = input("\nQuais processar? (T = todas pendentes, ou números: 1,3,4 | Enter sai): ").strip().upper()
    if not esc:
        return
    if esc == "T":
        escolhidos = [x for x in plano if not (x[0] / ".concluido").exists()]
    else:
        nums = {int(n) for n in re.findall(r"\d+", esc)}
        escolhidos = [x for i, x in enumerate(plano, 1) if i in nums]
    ja = [x for x in escolhidos if (x[0] / ".concluido").exists()]
    if ja and input(f"  {len(ja)} UBS já foram feitas antes. Refazer mesmo assim? (S/N) ").strip().upper() != "S":
        escolhidos = [x for x in escolhidos if x not in ja]

    enviar = lote / "para_enviar"
    for p, modo, pdfs in escolhidos:
        if not pdfs:
            continue
        if modo == "PDF":
            enviar.mkdir(exist_ok=True)
            destino = enviar / f"{p.name}_{lote.name[:10]}.pdf"
            w = PdfWriter()
            for pdf in pdfs:
                w.append(str(pdf))
            with open(destino, "wb") as f:
                w.write(f)
            print(f"\n  PDF para enviar: {destino.name}  ({len(pdfs)} pacientes)")
        else:
            imprimir_lista(pdfs, nome_bonito(p))
        if not SIMULAR:
            (p / ".concluido").write_text(time.strftime("%d/%m/%Y %H:%M"), encoding="utf-8")

    if enviar.exists() and any(m == "PDF" for _, m, _ in escolhidos):
        print(f"\n  Arquivos para enviar estão em: {enviar}")
        if sys.platform.startswith("win") and not SIMULAR:
            os.startfile(enviar)   # abre a pasta no Explorer
    print("\nPronto.")


if __name__ == "__main__":
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if args:
        pasta_avulsa(Path(args[0]))
    else:
        separacao_recente()
