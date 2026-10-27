# Robô de cadastro no Worklab (IA + RPA)

Automação do cadastro de pacientes e exames no Worklab, o sistema de informação do laboratório. O Worklab não tem API nem integração, então o robô trabalha pela interface, como um operador faria. Tema do meu TCC em Análise e Desenvolvimento de Sistemas (UEMA).

## Problema

- 70 atendimentos por dia, com picos de 80 a 90 pacientes.
- Cerca de 2 minutos de digitação por cadastro, mais de 2 horas por dia só nessa etapa.
- Erros comuns: datas trocadas, siglas de exame erradas, guias manuscritas mal lidas.

## Como funciona

```
Foto da guia → IA extrai os dados → equipe confere na planilha → Python padroniza → PyAutoGUI cadastra no Worklab
```

1. **Extração:** as guias e documentos são fotografados e um modelo de IA multimodal (Gemini) devolve os dados em tabela. O prompt guarda as regras do laboratório: siglas do Worklab, exames excluídos, um paciente por linha. CPF e endereço não são extraídos da foto.
2. **Conferência humana:** a equipe revisa os dados na planilha do Google Sheets antes de qualquer cadastro.
3. **Padronização (Python):** lê a planilha com `gspread`, agrupa as linhas por paciente, junta os exames sem repetição e normaliza datas e documentos.
4. **Cadastro (PyAutoGUI):** preenche os campos pelo teclado e lança os exames.

## Decisões de projeto

- **Navegação por Tab**, sem depender da posição dos campos na tela.
- **Confirmação do operador:** cada cadastro só começa depois do Enter de quem está no computador.
- **Pausas calibradas** para o tempo de resposta do Worklab (0,3 s entre comandos, 1 s após o CRM, 0,6 s por exame).
- **Parada de emergência:** levar o mouse ao canto da tela interrompe o robô (fail-safe do PyAutoGUI).

## Resultado

Estimativa da equipe: de cerca de 2 min para 20–30 s por cadastro, só na etapa de digitação. O tempo mínimo calculado a partir das pausas do script fica entre 15 s (5 exames) e 27 s (15 exames).

## Limitações conhecidas e próximos passos

- Tempos estimados, ainda sem cronometragem por paciente; taxa de erros não medida.
- Depende da ordem dos campos do Worklab: uma mudança na tela exige ajustar a sequência de Tabs.
- Próximos passos: registrar log de cada execução, tratar exceções por paciente e migrar `oauth2client` para `google-auth`.

## Executar

```bash
pip install gspread oauth2client pyautogui
python robo_cadastro.py
```

Coloque o arquivo `chave.json` da conta de serviço do Google na mesma pasta (ele é ignorado pelo git) e ajuste `CIDADE_PADRAO` no início do script.
