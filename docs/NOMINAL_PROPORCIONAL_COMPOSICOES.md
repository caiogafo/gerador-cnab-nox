# Sugestão de nominal em composições — 14/09/2026

Atualização solicitada: composições estrela com aquisição consolidada e nominal
consolidado positivos recebem sugestão proporcional de valor de face.

Para cada componente: nominal consolidado × (valor presente do componente ÷
aquisição consolidada). Cada resultado é arredondado em centavos com Decimal e
ROUND_HALF_UP. O resíduo da soma é aplicado integralmente ao principal detectado,
mesmo quando ele não aparece primeiro na PFMI. A soma sugerida fecha exatamente
com o nominal consolidado. Valores ausentes, não finitos, totais não positivos ou
ajuste que produziria nominal negativo não recebem sugestão automática.

`VL_NOMINAL_SUGERIDO` preserva a sugestão como controle imutável.
`VL_NOMINAL` recebe inicialmente esse valor e continua editável. O registro
`PREENCHIMENTOS_AUTOMATICOS` identifica a regra
`PRO_RATA_VALOR_PRESENTE_COMPOSICAO` e a exigência de aprovação humana. Somente a
representação numérica desse JSON usa float conforme contrato; a aritmética e a
conferência financeira usam Decimal.

Nada aprova nem seleciona os títulos automaticamente. Todas as linhas precisam
de `COMPOSICAO_APROVADA=SIM`, além das outras validações existentes. Valores manuais
não proporcionais são aceitos se a soma nominal fechar exatamente; qualquer
diferença de um centavo bloqueia. Apagar um nominal volta a gerar pendência.

O esquema passa a `CNAB-NOX-V1-7` para incluir o novo controle protegido. Prepare
um novo intermediário: versões anteriores não são migradas silenciosamente.
Composições sem crédito consolidado resolvido continuam no fluxo manual existente.
Nenhuma regra de identidade, comissão ou emissão foi alterada.
