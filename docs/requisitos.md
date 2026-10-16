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

## Automação de e-mails (decidido)

- **Fonte:** e-mail diário "Relatórios GLA" no Outlook clássico, vários remetentes. Só dois anexos: `VIDA MOLDE.xlsx` e `CORTESEQ.xls`.
- **Robô:** `scripts/atualizar_controle.vbs`, no PC com Outlook e Excel, agendado uma vez por dia útil.
- **Não depende de lido/não lido:** guarda o ID de cada e-mail processado em `processados.txt`. Não move nem marca e-mails.
- **Sem soma de diferenças:** molde usa o acumulado do anexo; um registro por data e veio; reenvio substitui.
- **Troca de molde:** detectada quando o número do molde do veio muda. Vai para a aba Trocas; o molde novo começa em 0 com outra numeração. Meta fixa por veio.
- **Componentes:** descontam a tonelada real do MCC4 só entre entrada e saída. Oscilador conta dias.
- **Alertas no Excel:** acumulado diminuiu sem troca; veio não andou enquanto a MCC produziu; layout dos anexos mudou; troca com observação divergente.
- **Paradas e entrada/saída:** manuais.

## Pendências

1. Testar o `.vbs` no PC dele em DRY_RUN (não foi testado: sem Windows aqui).
2. Confirmar se a produção usada no alerta "veio não andou" deve ser a do dia anterior ao relatório (assumido) ou a do próprio dia.
3. Confirmar se os e-mails ficam na Caixa de Entrada padrão (o robô lê só ela).
4. Acumulado inicial dos componentes e do oscilador (hoje em branco = 0).
5. Política da empresa sobre scripts `.vbs`.
