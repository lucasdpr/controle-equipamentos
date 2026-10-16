# Controle de vida dos equipamentos

Planilha `controle_equipamentos.xlsx` para controlar o desgaste de moldes, osciladores e componentes.

## Abas

- **Equipamentos**: cadastro (meta, acumulado inicial, unidade) e resumo automático (acumulado, falta, %, situação).
- **Movimentos**: entrada e saída de cada equipamento. Dias rodados são calculados.
- **Lancamentos**: produção lançada (t ou corridas) por equipamento e data. A coluna *Válido* indica se o lançamento caiu dentro de um período em máquina.
- **Regras**: limiar de atenção e explicação do cálculo.

## Regras

- Falta = meta − acumulado. % = acumulado ÷ meta.
- Dias rodados = saída − entrada (entrada conta, saída não). Sem saída, conta até hoje.
- Oscilador: unidade em dias.
- Demais componentes: toneladas. Molde MCC4: corridas.
- Lançamento fora de período em máquina não entra no acumulado.
- Sem previsão de troca.

## Regenerar a planilha

```
python3 scripts/gerar_planilha.py
```
