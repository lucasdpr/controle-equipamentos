# Requisitos: controle de vida dos equipamentos

## Objetivo

Controlar o desgaste de cada equipamento (molde, oscilador e componentes) até a meta de troca, a partir de:
- **Movimentos** de entrada e saída, digitados manualmente.
- **Lançamentos** de produção (toneladas ou corridas), que podem vir automaticamente por e-mail.

## Regras confirmadas

- **Meta**: limite de vida do item, até a troca.
- **Acumulado**: quanto o item já rodou.
- **Falta** = meta − acumulado. **%** = acumulado ÷ meta.
- **Unidade por item**: oscilador em dias; componentes em toneladas (t); molde em corridas.
- **Dias rodados** = saída − entrada. A entrada conta, a saída não. Período em aberto (sem saída) conta até hoje.
- **Lançamento só conta se a data cair dentro de um período em máquina** (entrada ≤ data < saída). Fora disso, fica marcado como "Não válido" e não entra no acumulado.
- **Sem previsão** de data de troca.

## Exemplos de referência

- Entrada 10/03, saída 15/03: 5 dias.
- Bender, meta 1.100.000 t, lançado 2.000 t: falta 1.098.000 t.
- Molde MCC4: meta 1100 corridas, acumulado inicial 1080, falta 20.

## Automação de e-mails (Outlook)

- **Ferramenta:** Outlook clássico (desktop) no Windows, lido por um script `.vbs` via COM.
- **Agendamento:** Agendador de Tarefas do Windows, uma vez por dia útil, com "executar assim que possível após início perdido" marcado.
- **Processamento:** lê todos os e-mails não processados da caixa de entrada, não só os do dia.
- **Data do lançamento:** vem do e-mail (data de recebimento), não da data de execução.
- **Sem duplicidade:** após lançado, o e-mail é movido para a pasta "Processados". Só lê a caixa de entrada.
- **Revisão:** e-mail que não bate com o padrão (equipamento não encontrado, valor não numérico) vai para a pasta "Revisar", sem lançar.
- **Remetentes:** várias pessoas enviam e-mails. Lista ainda não definida.

## Pendências (bloqueiam a automação)

1. **Exemplo de e-mail real** (com nomes e números trocados) e o padrão do corpo.
2. **Lista de remetentes** autorizados.
3. **Confirmação de Outlook clássico** na máquina que vai rodar o script.
4. **Política da empresa** sobre scripts `.vbs` (se bloqueado, precisa liberação de TI).
5. **Local da planilha** (pasta fixa, rede ou OneDrive/SharePoint, se mais de uma pessoa usar).

## Estado atual

- Planilha `controle_equipamentos.xlsx` gerada por `scripts/gerar_planilha.py` (abas: Equipamentos, Movimentos, Lancamentos, Regras).
- Script `.vbs`: não iniciado, aguarda o exemplo de e-mail.
- Entrada e saída: manuais, pela planilha.
