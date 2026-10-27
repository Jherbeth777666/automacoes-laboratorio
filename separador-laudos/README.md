# Separador e impressão de laudos

Pega o PDF com todos os laudos do dia, exportado do Worklab, e entrega um PDF por paciente, organizado na pasta da sua unidade de saúde (UBS), com uma planilha de conferência e impressão em lote.

## Problema

O Worklab exporta os laudos do dia num único PDF. Separar paciente por paciente, descobrir a UBS de cada um e montar os lotes de entrega era trabalho manual, demorado e sujeito a laudo na pasta errada.

## Como funciona

1. **Leitura dos laudos** (`pypdf`, com `pdfplumber` como apoio): extrai paciente e protocolo do cabeçalho de cada página.
2. **Descoberta da UBS**, nesta ordem: o próprio laudo; a planilha de agendados pelo número de atendimento; a planilha de agendados pelo nome, aceitando pequenas diferenças de digitação (similaridade com `difflib` e regra para sobrenome abreviado); a planilha de atendidos.
3. **Separação:** um PDF por paciente, dentro da pasta da UBS. Quem não for encontrado vai para "SEM UBS", para conferência.
4. **Conferência** (`openpyxl`): planilha com cada paciente, a UBS atribuída e de onde veio a informação.
5. **Entrega** (`imprimir_lote.py`): para cada UBS, imprime os laudos um a um (SumatraPDF) ou junta tudo num PDF único para envio. A escolha de cada UBS fica salva e só é perguntada na primeira vez.

Agendados e atendidos são aceitos em PDF, Excel ou CSV.

## Executar

```bash
pip install pypdf pdfplumber reportlab openpyxl
python separar_laudos.py                # separa e gera a conferência
python separar_laudos.py --imprimir     # separa e imprime
python imprimir_lote.py --simular       # mostra o que seria impresso, sem imprimir
```

Coloque os arquivos na pasta `entrada`. A impressão no Windows usa o SumatraPDF (gratuito).

## Dados

Roda só no computador do laboratório. Nenhum laudo, planilha ou nome de paciente faz parte deste repositório.
