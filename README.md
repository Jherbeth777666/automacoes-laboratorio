# Automações de laboratório (RPA com Python)

Robôs que desenvolvi e mantenho no laboratório municipal onde sou Líder Administrativo. Cada um nasceu de uma tarefa manual e repetitiva da rotina: entendi o processo com a equipe, levantei as regras e automatizei o que era digitação ou conferência.

| Robô | Problema que resolve | Tecnologias |
|---|---|---|
| [Cadastro no Worklab](robo-cadastro-worklab/) | Cadastro manual de 70 a 90 pacientes por dia em um sistema legado sem integração | Python, PyAutoGUI, Google Sheets (gspread), IA multimodal (Gemini) |
| [Separador de laudos](separador-laudos/) | Separar os laudos em PDF por paciente e por unidade de saúde para entrega e impressão | Python, pypdf, pdfplumber, openpyxl |

## Cuidados com dados de saúde (LGPD)

- Nenhum dado de paciente está neste repositório: só o código.
- Os robôs rodam localmente, no computador do laboratório.
- Chaves de acesso (como a conta de serviço do Google) ficam fora do código e do repositório.
- Toda automação tem conferência humana antes de gravar ou imprimir.
