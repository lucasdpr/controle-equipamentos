# Controle de vida dos equipamentos

Planilha `controle_equipamentos.xlsx` (Excel) atualizada por um robô `.vbs` que lê o e-mail diário "Relatórios GLA" no Outlook clássico.

## Fontes (só estes dois anexos do e-mail)

- `VIDA MOLDE.xlsx`: acumulado de corridas por veio (C a H) e número do molde.
- `CORTESEQ.xls`, aba DADOS DA PRODUÇÃO: toneladas reais por MCC por dia.

## Abas da planilha

- **Início**: como usar, com links para cada aba.
- **Painel** (abre nele): indicadores (equipamentos / além da meta / atenção / OK / alertas), mais desgastados, resumo por tipo, todos os equipamentos e 3 gráficos.
- **Equipamentos**: uma linha por equipamento/tag (21 já listados: 2 moldes G e H, oscilador, Bender, 5 Bow, R1, R2 e 10 Horizontais; cabem até 40). Cada tag tem a sua meta, MCC, "Desde" e acumulado inicial (células amarelas). Metas e "Desde = 01/08/2026" são provisórios.
- **Moldes**: todos os veios (C a H) com molde atual, acumulado, %, quantidade de trocas, corridas médias na saída, última troca e dias parados.
- **TrocasEquip** (manual): registro da saída de oscilador e componentes (quando trocar um, registre aqui e digite a nova data em "Desde" no Resumo).
- **Paradas** (manual): início, fim, motivo.
- **Historico** (robô): um registro por data e veio (todos os veios C a H).
- **Producao** (robô): toneladas reais por MCC por dia.
- **Trocas** (robô): trocas de molde, com campo "Conferir" quando a observação não bate.
- **Alertas** (robô): acumulado diminuiu, veio não andou, layout mudou. Marque "Sim" ao resolver.
- **Regras**: limiar de atenção e explicação.

## Regras

- Molde: corridas, acumulado do último dia importado. Meta fixa por veio (G e H = 1100).
- Oscilador: dias corridos desde a data em "Desde" (Resumo) + Acumulado inicial. Meta 730.
- Componentes: toneladas reais do MCC4 desde a data em "Desde" (Resumo) + Acumulado inicial. Tudo atualiza sozinho quando o robô importa a produção do dia.
- Falta = meta − acumulado. % = acumulado ÷ meta. Sem previsão.
- Reenvio do e-mail do mesmo dia substitui o registro (não soma). Dia perdido se corrige no próximo e-mail.

## Instalação no PC (Windows, Outlook clássico, Excel desktop)

1. Copie `controle_equipamentos.xlsx` e `scripts/atualizar_controle.vbs` para a mesma pasta (qualquer lugar, ex.: Área de Trabalho\Controle). O robô usa a pasta onde o `.vbs` está.
2. Feche a planilha. Rode o `.vbs` com `DRY_RUN = True` (padrão): ele não grava nada e escreve `robo.log` na mesma pasta.
3. Confira o log. Se estiver certo, troque para `DRY_RUN = False` e rode de novo.
4. Agende no Agendador de Tarefas do Windows (uma vez por dia útil, marcando "executar assim que possível após início perdido").

## Não testado

O `.vbs` foi escrito sem acesso a Windows, Outlook ou Excel desktop. A lógica de leitura segue as posições reais dos arquivos de exemplo (VIDA MOLDE: B..H linhas 5 a 10, data em I2; CORTESEQ: DADOS DA PRODUÇÃO colunas A, H, I). Rode primeiro em DRY_RUN e numa cópia da planilha.

## Regenerar a planilha

```
python3 scripts/gerar_planilha.py
```

---
Autor: Lucas Gabriel de Paula Rafael | Matrícula cbk3574 | https://github.com/lucasdpr/controle-equipamentos
