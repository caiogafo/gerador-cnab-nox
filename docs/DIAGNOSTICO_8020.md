# Diagnóstico offline antes da revisão do cruzamento

O comando abaixo analisa as fontes sem alterá-las, sem preparar intermediário
e sem gerar CNAB. Reutiliza os leitores e as regras atuais do produto.
Gera apenas JSON e Markdown em um destino novo. Os relatórios são locais;
não contêm nomes, documentos ou dados bancários, mas contêm valores e referências
de linhas que permitem ao operador consultar as fontes.

```powershell
python scripts/diagnose_reconciliation.py --pfmi "PFMI.xlsx" --analytic "ProspectAnalitico.xls" --base "Base.xlsx" --date 2026-09-10 --modality NORMAL --output "outputs/diagnostico_novo"
```

Executar em ambiente com o pacote `gerador_cnab_nox` e suas dependências disponíveis.
Data e modalidade são obrigatórias para tornar explícitas as hipóteses do diagnóstico.
Comissão pode ser informada por `--commission`; o padrão é 15 e não afeta Normal.
A primeira sequência permanece ausente: não existe validação completa de geração.

## Resultado do lote de exemplo de 10/09/2026

Com modalidade Normal como hipótese, liquidação em 10/09 e Base fornecida:

- 17 pagamentos reconhecidos, zero linhas desconhecidas, dez grupos físicos.
- Seis grupos pré-selecionados pelas regras atuais.
- Dois grupos adicionais têm sugestões com nome e valor correspondentes,
  mas a equivalência de falência não está confirmada nas regras atuais.
- Dois grupos permanecem sem sugestão individual de valor exato.
- Cinco grupos têm pendência de sacado na Base, incluindo três pré-selecionados.
- A comparação com os dois Analíticos fornecidos produziu as mesmas contagens.
  Isso não significa que os arquivos ou todos os seus dados sejam iguais.

Os pacotes estão em `outputs/diagnostico_8020_20260910/analitico_enviado`
e `outputs/diagnostico_8020_20260910/analitico_pasta`. Cada JSON registra hashes
das fontes, comparados novamente após a leitura para conferir preservação.

Seis de dez grupos pré-selecionados correspondem a 60% neste lote. As duas
sugestões adicionais não elevam a cobertura aprovada para 80%. Pré-seleção de
crédito também não comprova que documentos, vencimentos e demais campos estejam
prontos para geração.

## Como interpretar sugestões

O diagnóstico compara a aquisição esperada segundo modalidade/comissão com o
total do grupo em centavos. Não procura combinações de créditos nem junta blocos.
Para nomes, usa normalização existente, uma variante sem prefixo de espólio e
conteúdo entre parênteses como pistas exploratórias. Um índice textual de pelo
menos 0,8 filtra sugestões; não é uma probabilidade nem critério de aprovação.
Todas as alternativas que passam pelo filtro permanecem visíveis, inclusive
empates e campanhas diferentes. A referência é a linha do Analítico específico.

Nenhuma sugestão modifica o matching de produção, cria equivalência de falência,
troca nomes/documentos ou libera o CNAB. A seleção atual permanece a do produto.

## Decisões necessárias para a próxima alteração funcional

1. Confirmar a equivalência sugerida entre a falência da PFMI e a campanha dos
   dois grupos adicionais, e conferir sua correspondência na Base de Vencimentos.
2. Resolver os sacados ausentes/conflitantes usando a origem correta para os campos.
3. Definir a representação dos instrumentos separados de credora/advogado:
   há coincidência financeira agregada observada na análise da reunião, mas
   divergência de sobrenome e dois blocos físicos. Não juntar nem duplicar crédito
   automaticamente para fazer os valores fecharem.
4. Definir a revisão humana de candidatos aproximados no intermediário antes
   de ampliar o contrato de seleção. A reunião é contexto de requisitos; o resumo
   não autoriza sozinho geração parcial nem aprendizado automático de aliases.

As regras atuais do DRD/Specs e o esquema do intermediário permanecem vigentes.
O avanço entregue nesta etapa é o diagnóstico reproduzível e testado.
