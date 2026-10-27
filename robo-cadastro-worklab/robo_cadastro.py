import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pyautogui
import time
from collections import OrderedDict

CIDADE_PADRAO = "SUA CIDADE"  # cidade preenchida no endereço de todos os cadastros

# ===============================
# 1. CONEXÃO COM GOOGLE SHEETS
# ===============================
scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
# A chave da conta de serviço fica fora do repositório (ver README).
creds = ServiceAccountCredentials.from_json_keyfile_name("chave.json", scope)
client = gspread.authorize(creds)

try:
    planilha = client.open("Cadastro_Worklab").sheet1
    dados_brutos = planilha.get_all_records()
    print(f">>> Conexão OK! {len(dados_brutos)} linhas lidas.")
except Exception as e:
    print(f"Erro ao abrir planilha: {e}")
    exit()

# ===============================
# 2. ORGANIZAÇÃO DOS DADOS (Agrupamento)
# ===============================
pacientes = OrderedDict()
for linha in dados_brutos:
    nome = str(linha.get('Nome') or linha.get('NOME', '')).strip().upper()
    if not nome: continue
    
    exame_raw = str(linha.get('exame', '') or linha.get('Exames', '')).strip()
    exames_da_linha = [e.strip().upper() for e in exame_raw.split(',') if e.strip()]
    
    if nome not in pacientes:
        pacientes[nome] = {
            'Sexo':       str(linha.get('Sexo') or linha.get('SEXO', 'F')).upper(),
            'Nascimento': str(linha.get('Nascimento', '')).replace("/", ""),
            'CPF':        str(linha.get('CPF', '')).replace(".", "").replace("-", ""),
            'Endereco':   str(linha.get('Endereco') or linha.get('Endereço', '')).upper(),
            'crm':        str(linha.get('CRM') or linha.get('crm', '')),
            'exames':     exames_da_linha
        }
    else:
        for ex in exames_da_linha:
            if ex and ex not in pacientes[nome]['exames']:
                pacientes[nome]['exames'].append(ex)

# ===============================
# 3. AUTOMAÇÃO TURBO
# ===============================
# Diminuí o PAUSE global para 0.3 para acelerar a navegação geral
pyautogui.PAUSE = 0.3 
pyautogui.FAILSAFE = True

for nome, p in pacientes.items():
    print(f"\n🚀 Paciente: {nome}")
    input("👉 Posicione o cursor no NOME e aperte ENTER...")
    time.sleep(1.5)

    # --- DADOS PESSOAIS ---
    pyautogui.write(nome, interval=0.01) # Digitação rápida
    pyautogui.press('tab')
    pyautogui.write(p['Sexo'])
    pyautogui.press('tab')
    pyautogui.write(p['Nascimento'])
    pyautogui.press('tab')
    pyautogui.press('tab', presses=4) 
    pyautogui.write(p['CPF'])
    pyautogui.press('tab', presses=2) 

    # --- ENDEREÇO ---
    pyautogui.press('tab') 
    pyautogui.write(p['Endereco'], interval=0.01)
    pyautogui.press('tab')
    pyautogui.write(CIDADE_PADRAO, interval=0.01)
    pyautogui.press('tab', presses=4) 

    # --- MÉDICO E CONVÊNIO ---
    pyautogui.write(p['crm'], interval=0.01)
    pyautogui.press('tab') 
    time.sleep(1.0) # Espera sistema validar médico
    
    pyautogui.press('tab') 
    pyautogui.press('tab', presses=3) # Pula Convênio
    
    # --- TRANSIÇÃO PARA EXAMES (5 TABS CONFORME VALIDADO) ---
    pyautogui.press('tab', presses=5) 

    # --- LANÇAMENTO DOS EXAMES (MODO VELOCIDADE) ---
    if p['exames']:
        for exame in p['exames']:
            print(f"   + Inserindo: {exame}")
            # Digitação quase instantânea
            pyautogui.write(exame, interval=0.01) 
            pyautogui.press('enter')
            # Tempo mínimo para o Worklab aceitar na grade
            # Se começar a pular exames, aumente para 0.7 ou 0.8
            time.sleep(0.6) 
    
    print(f"✅ Concluído: {nome}")
    print("-" * 40)

print("\n🏁 Processo finalizado no modo Turbo!")