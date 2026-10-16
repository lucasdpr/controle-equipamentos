# Controle de vida dos equipamentos

Planilha `controle_equipamentos.xlsx` (Excel) atualizada por um robô `.vbs` que lê o e-mail diário "Relatórios GLA" no Outlook clássico.

## Fontes (só estes dois anexos do e-mail)

- `VIDA MOLDE.xlsx`: acumulado de corridas por veio (C a H) e número do molde.
- `CORTESEQ.xls`, aba DADOS DA PRODUÇÃO: toneladas reais por MCC por dia.

## Abas da planilha

- **Resumo**: meta, acumulado, falta, %, situação. Itens: Molde G, Molde H, Oscilador, Bender, Bow, Straightener R1/R2, Horizontal.
- **Movimentos** (manual): entrada e saída de oscilador e componentes.
- **Paradas** (manual): início, fim, motivo.
- **Historico** (robô): um registro por data e veio (todos os veios C a H).
- **Producao** (robô): toneladas reais por MCC por dia.
- **Trocas** (robô): trocas de molde, com campo "Conferir" quando a observação não bate.
- **Alertas** (robô): acumulado diminuiu, veio não andou, layout mudou. Marque "Sim" ao resolver.
- **Regras**: limiar de atenção e explicação.

## Regras

- Molde: corridas, acumulado do último dia importado. Meta fixa por veio (G e H = 1100).
- Oscilador: dias (saída − entrada; entrada conta, saída não; em aberto conta até hoje). Meta 730.
- Componentes: toneladas reais do MCC4 nos dias entre entrada e saída.
- Falta = meta − acumulado. % = acumulado ÷ meta. Sem previsão.
- Reenvio do e-mail do mesmo dia substitui o registro (não soma). Dia perdido se corrige no próximo e-mail.

## Instalação no PC (Windows, Outlook clássico, Excel desktop)

1. Copie `controle_equipamentos.xlsx` e `scripts/atualizar_controle.vbs` para `C:\Controle\`.
2. Se usar outra pasta, ajuste `PASTA` e `PLANILHA` no topo do `.vbs`.
3. Feche a planilha. Rode o `.vbs` com `DRY_RUN = True` (padrão): ele não grava nada e escreve `C:\Controle\robo.log`.
4. Confira o log. Se estiver certo, troque para `DRY_RUN = False` e rode de novo.
5. Agende no Agendador de Tarefas do Windows (uma vez por dia útil, marcando "executar assim que possível após início perdido").

## Não testado

O `.vbs` foi escrito sem acesso a Windows, Outlook ou Excel desktop. A lógica de leitura segue as posições reais dos arquivos de exemplo (VIDA MOLDE: B..H linhas 5 a 10, data em I2; CORTESEQ: DADOS DA PRODUÇÃO colunas A, H, I). Rode primeiro em DRY_RUN e numa cópia da planilha.

## Regenerar a planilha

```
python3 scripts/gerar_planilha.py
```
