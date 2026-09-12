# Specs 1.1 — Gerador de CNAB NOX

**Versão:** 1.1  
**Data da aprovação funcional:** 2026-09-04  
**Autoridade de negócio:** [DRD](drd.md)  
**Execução e evidências:** [ExecPlan](execplan.md)  
**Estado:** contrato revisado aprovado e implementado; candidato Linux M1–M6 comprovado; M7/M8 posteriores.

Este é o contrato de M1–M6; os resultados executados estão no ExecPlan.
M0 foi somente documental. O pacote Windows e a homologação operacional
continuam em M7/M8.

## 1. Arquitetura e invariantes

Manter Python 3.12+, openpyxl (XLSX), xlrd 2.x (XLS), CSV/TXT pela biblioteca padrão,
Tkinter/ttk e PyInstaller `onedir` no Windows. Sem Excel/COM/VBA no processamento,
sem banco, rede, IA ou integração ao Frontis. Não introduzir outro framework.

Responsabilidades existentes, a evoluir no próprio módulo:

| Módulo | Responsabilidade na revisão |
|---|---|
| `models` | Entradas por PFMI, ordens, proveniência, identidade de crédito, cálculo e resultado |
| `normalize` | Decimal/documentos/datas; comparação de nomes separada da saída |
| `readers` | Formatos conhecidos, blocos, auxiliares, espelhos e linhas não reconhecidas |
| `matching` | Candidatos determinísticos, seleção inicial, Base e comissão por crédito |
| `workbook` | Cinco abas, esquema novo, referências imutáveis e finais editáveis |
| `validation` | Releitura independente, integridade, campos finais e reconciliação |
| `cnab` | Renderizador posicional puro, sem fontes, matching ou comissão |
| `service` | Orquestração, gravação atômica, nomes exclusivos e log sanitizado |
| `gui` | Três etapas em português, lista de PFMIs/parâmetros, resultados e execução serial |

Invariantes:

- Pagamento PFMI, grupo físico, candidato e título são unidades diferentes.
- Um crédito selecionado do Analítico produz um título, independentemente do cedente final.
- Originais, sugestões calculadas e valores finais são camadas distintas.
- Pendência operacional não impede Excel; impede TXT quando ainda bloqueante.
- O segundo passo usa exclusivamente o intermediário, inclusive os parâmetros capturados.
- Nenhuma seleção de pessoa por similaridade, documento de beneficiário Normal ou fechamento financeiro isolado.

### Interface Linux — melhoria aprovada em 2026-09-05

- Uma etapa por vez: configurar lote, preparar/revisar Excel e gerar TXT.
  Preservar parâmetros entre telas e permitir entrada direta com Excel revisado.
- Tema azul baseado em `clam`, fontes disponíveis no Tk, navegação/ações fixas e
  rolagem central com foco de teclado visível. Nenhuma dependência nova.
  Refinamento de composição: cartões e controles arredondados, etapas numeradas,
  métricas separadas e tabela sem grade pesada. Manter widgets ttk para teclado e
  estados; recursos decorativos locais não podem acoplar a GUI ao processamento.
- A GUI captura entradas antes de executar um único trabalhador. Somente a thread
  principal acessa Tk; `after` entrega resultados/erros e mantém progresso
  indeterminado. Não há cancelamento de processamento; fechamento aguarda conclusão.
- Exibir caminho salvo e ações locais, contagens e valores em português brasileiro.
  Detalhes de validação continuam completos; traduções pertencem apenas à GUI.
  A contagem agregada da preparação não é quantidade de erros bloqueantes.
- Preservar assinaturas/resultados de `prepare_workbook` e `generate_cnab`, exceções,
  esquema `CNAB-NOX-V1-2`, leiaute e logs. Helpers de tema, widgets e mensagens são
  exclusivamente de apresentação. Abertura local usa argumentos sem shell no Linux
  e associação do sistema no Windows; falha preserva o arquivo e orienta abertura manual.

## 2. Contratos de entrada e preparação

### 2.1. Requisição do lote

Contrato lógico (nomes internos podem ser adaptados mantendo a semântica):

```text
PreparationRequest
  pfmis: lista ordenada, não vazia, de PfmiInput
    path
    ordem_arquivo
    modalidade: NORMAL | CESSAO_DA_CESSAO
    comissao_texto: entrada original da interface
  analitico_path: um arquivo
  vencimentos_path: um arquivo
  data_liquidacao: valor informado, inclusive ausente/inválido para revisão
  primeira_sequencia: valor informado, inclusive ausente/inválido para revisão
```

Capturar SHA-256 dos arquivos e a ordem da seleção. Um mesmo path ou conteúdo
operacional repetido na lista não é removido silenciosamente: sinalizar revisão.

Na interface, modalidade inicial Normal; ao configurar Cessão da Cessão pela primeira
vez, preencher 15%. Depois preservar explicitamente o valor digitado, inclusive zero
ou vazio; mudanças de modalidade não podem substituir uma taxa já informada
silenciosamente. Em Normal a taxa é inativa e seu conteúdo não gera comissão nem
pendência de taxa. Reativar Cessão mantém o valor próprio daquele item.

A data de liquidação e a primeira sequência podem ser corrigidas no `RESUMO`.
Modalidade, ordem de origem e taxa são parâmetros da preparação: para mudá-los,
alterar a interface e preparar um novo Excel.

### 2.2. Percentual e aritmética

- Entrada decimal textual simples com vírgula ou ponto, sem separador de milhar.
- Remover espaços nas extremidades; `15`, `15,0`, `15.0`, `0`, `0,5` são válidos.
- Campo vazio é pendência, nunca fallback 15. Não interpretar 15 como 0,15%.
- Rejeitar negativos, texto, NaN, infinito, múltiplos separadores e notação não suportada.
- Limite técnico inicial: até 28 algarismos significativos, 28 casas decimais e
  28 posições inteiras; usar contexto Decimal de pelo menos 80 dígitos nos cálculos.
  Este é limite de representação, não teto comercial. 100% e 101% são valores
  de parâmetro válidos; o resultado ainda precisa caber nos campos monetários.
- Excesso de capacidade deve gerar `COMISSAO_PARAMETRO_INVALIDO`, nunca exceção
  não tratada, arredondamento silencioso da taxa ou queda da preparação.
- Preservar texto informado e Decimal normalizado no snapshot. Para Normal, registrar
  taxa efetiva zero e modalidade; não aplicar a taxa visual inativa.
- Quantizar valores financeiros com `Decimal("0.01")` e `ROUND_HALF_UP`. Ao converter
  números de planilha, usar representação decimal, não operações monetárias em float.
- Validar a capacidade antes de quantizar/renderizar números extremos.

Por crédito:

```text
A = quantizar_centavos(Valor Aquisicao original)
N = quantizar_centavos(Valor Receber original)
B = N - A

Normal:
  C = 0,00
  E = A

Cessão da Cessão:
  C = quantizar_centavos(B * p / 100)
  E = A + C
```

Para Cessão a 0%, C = 0,00 e E = A. Base negativa com p positivo gera
`COMISSAO_BASE_NEGATIVA`, impede sugestão automática válida e exige revisão,
sem zerar a base nem aplicar comissão negativa silenciosamente.
Dados de crédito ausentes/inválidos permanecem visíveis; não abortam a preparação.

A soma do grupo é `sum(E por crédito)`, após arredondamento individual. Nunca
calcular comissão sobre a soma para redistribuir, buscar combinação ou tolerar
centavos. Comissão não alimenta `TAXA_INDEXADOR`.

Taxa inválida é pendência imutável da preparação e exige novo Excel com parâmetro
corrigido. Já um campo final de crédito pode ser corrigido no Excel: a validação
não recalcula E a partir de finais editados nem impõe E como valor final obrigatório.
Preservar A, N, B, p, C e E originais/calculados para comparação.

### 2.3. PFMI XLSX

Cabeçalhos comparados sem acento, caixa, pontuação e espaços repetidos:

| Campo interno | Aliases mínimos |
|---|---|
| `valor_operacao` | VALOR DA OPERAÇÃO, VALOR OPERACAO |
| `falencia` | FALÊNCIA, FALENCIA, CAMPANHA |
| `cedente` | CREDOR / CEDENTE, CREDOR, CEDENTE |
| `beneficiario` | BENEFICIÁRIOS / FAVORECIDOS, BENEFICIARIO, FAVORECIDO |
| `documento_beneficiario` | CPF/ CNPJ, CPF/CNPJ, DOCUMENTO |
| `valor_beneficiario` | VALOR |

Localizar tabelas operacionais por assinatura dos cabeçalhos em todas as abas
relevantes. Não selecionar/excluir aba pelo nome ou por estar oculta.

Regras de classificação:

1. Cabeçalho inicia tabela; linha vazia encerra o trecho físico de pagamentos.
   Cabeçalho repetido de continuação não é pagamento nem razão suficiente para
   duplicar operações. Uma nova tabela/bloco fisicamente separado mantém nova identidade.
2. Credor/falência podem ser herdados apenas dentro do trecho corrente, em continuação
   que possua beneficiário e valor; nunca de outro bloco ou arquivo.
3. Agrupar por arquivo + aba efetiva + bloco + cedente normalizado + falência canônica.
   Preservar todas as linhas, inclusive iguais, e a primeira ocorrência como ordem do grupo.
4. Comparar VALOR DA OPERAÇÃO e VALOR por pagamento em centavos. Ausência/invalidez/
   diferença vira pendência de origem, sem excluir a linha. A base financeira do grupo
   é a soma de VALOR dos beneficiários (coluna J na PFMI atual); se incompleta,
   marcar total inválido, não tratar como zero. VALOR DA OPERAÇÃO é conferência,
   nunca substitui VALOR no total.
5. Reconhecer títulos de fundo, cabeçalhos repetidos, resumo
   DATA / QUANTIDADE DE TEDS / VALOR TOTAL A PAGAR e os demais formatos auxiliares
   demonstrados no corpus. Tabela de gerador legado é auxiliar, não segunda fonte de títulos.
6. Guardar resumo declarado e resumo calculado separadamente. Resumo vazio,
   desatualizado ou divergente produz alerta; não bloqueia por si só.
7. Linha preenchida fora dessas assinaturas aparece em `PENDENCIAS_PFMI` com
   arquivo/aba/bloco/linha e orientação `REVISAR`. Não ignorar por posição, cor ou palavra isolada.
8. Aba espelho: comparar sequência integral de pagamentos, conteúdo operacional e
   marcadores de blocos. Somente igualdade integral permite uma aba efetiva; registrar
   aba principal, espelho, critério e contagem não duplicada. Igualdade de totais ou
   subconjunto de linhas é insuficiente.
9. Entre arquivos, sinalizar conteúdo operacional duplicado sem removê-lo. Nunca
   deduplicar créditos por documento/nome/valor. A prevenção global de reutilização
   continua obrigatória, inclusive depois da escolha manual.

Linhas de origem desconhecidas/inconsistentes exigem corrigir PFMI e preparar
novamente. Resumos e espelhos reconhecidos são avisos, não pendências bloqueantes
disfarçadas. Não copiar dados bancários para campos CNAB.

### 2.4. ProspectAnalítico XLS/XLSX/CSV

| Campo | Cabeçalhos |
|---|---|
| Documento original | Cpf, CPF/CNPJ, Documento |
| Nome original | Contato, Nome Cedente, Cedente |
| Falência/campanha | Campanha, Falência, Falencia |
| Nominal | Valor Receber, Valor Nominal |
| Aquisição | Valor Aquisicao, Valor Aquisição |
| Assinatura original | Data Assinatura, Data de Assinatura |
| Referência | Primeiro disponível entre Id, ID Prospect, Prospect, Código; fallback linha física |

Localizar cabeçalho nas primeiras 30 linhas. No formato conhecido com dois
`Contato`, usar o contato principal associado ao CPF (primeira ocorrência no
export verificado), sem sobrescrever pelo segundo contato. Fixar esse formato em
fixture sintética. Um formato novo não pode ser interpretado silenciosamente como o conhecido.

CSV: UTF-8 com/sem BOM ou Windows-1252; delimitadores vírgula, ponto e vírgula,
tab ou barra vertical. Respeitar registros lógicos e aspas. Registrar aba/linha
ou registro lógico de origem, sem confundir referência comercial com identidade.

`ID_CREDITO` é composto por hash do snapshot + aba/registro de origem + linha +
referência. Referências Prospect repetidas mantêm créditos distintos.
Linhas iguais em posições diferentes não são deduplicadas.

Valores e datas operacionais inválidos, inclusive extremos, viram pendências/
referências inválidas no Excel, não falha de leitura geral. Cabeçalhos essenciais
ausentes ou conteúdo ilegível são falhas estruturais. Não inventar assinatura,
documento ou registro quando faltarem.

### 2.5. Base de Vencimentos XLSX

Exigir NOME_SACADO, DOC_SACADO e DATA_VENCIMENTO. Aplicar a mesma função de identidade
de falência da PFMI e Analítico, incluindo o de-para MOGIANO confirmado.
Duplicatas de mesmo registro canônico/documento/data não criam conflito.
Duas combinações materiais de nome, documento ou vencimento para a mesma identidade
geram `CONFLITO_BASE_VENCIMENTOS`; não escolher por ordem de linha.

Sacado ausente ou campos incompletos geram `PREENCHER_MANUALMENTE`.
Preservar a evidência da Base e deixar campos ambíguos pendentes.
O operador do fundo pode resolvê-los nos finais do Excel; o segundo passo não exige atualizar
a Base externa nem considera o status original um bloqueio permanente.

## 3. Matching e autoridade dos campos

O preparo aceita um snapshot opcional `failure_aliases`, validado e independente
do estado global. A interface carrega as equivalências locais confirmadas antes
de iniciar o trabalho. O resolvedor aplica normalização e cadeias sem ciclos
às três fontes, preservando nomes originais e conflitos materiais da Base.
O snapshot fica no manifesto e é exibido no RESUMO. A geração não consulta o
cadastro local. Ver [cadastro de equivalências](docs/EQUIVALENCIAS_FALENCIAS.md).

### 3.1. Comparação, equivalências e candidatos

`normalizar_comparacao` remove acentos, caixa, pontuação e espaços redundantes
e canoniza apenas LIMITADA→LTDA, COMPANHIA→CIA, SOCIEDADE ANONIMA→SA.
Para falência, aplicar adicionalmente o de-para explícito
MOGIANO ↔ MOGIANO TRANSP GERAIS. Não usar substring/fuzzy matching como equivalência.

Algoritmo em duas passagens, para detectar concorrência antes da seleção inicial:

1. Formar candidatos principais por nome normalizado + falência canônica.
2. Calcular a aquisição esperada dos candidatos usando modalidade/taxa do seu grupo.
3. Se não existir solução principal automática, oferecer alternativas individuais
   da mesma falência com aquisição esperada exatamente igual ao total do grupo.
   Marcar motivo `ALTERNATIVA_VALOR_EXATO_ESCOLHA_MANUAL`, nunca escolha automática.
4. Indexar participação dos créditos em todos os grupos; crédito concorrente não
   recebe pré-seleção automática em nenhum deles.
5. Pré-selecionar o conjunto principal integral apenas se todos tiverem identidade
   consistente, mesmo documento original válido, cálculo utilizável, soma exata
   e ausência de concorrência. Não escolher um subconjunto para fechar.
6. Demais candidatos começam `INCLUIR_CNAB=NAO`, com motivo claro de revisão.
   Um grupo sem selecionados bloqueia TXT, mas não Excel.
7. Sem candidato útil, inserir placeholder `CREDITO_NAO_LOCALIZADO`, sem
   `ID_CREDITO` válido. Orientar nova preparação após corrigir PFMI/Analítico.

Não oferecer todo o Analítico indiscriminadamente, nem inferir pessoa pela taxa
ou por documento de favorecido Normal. CPF/CNPJ é a chave de identidade após
localização; a origem do crédito é obrigatória também para seleção manual.
A escolha de alternativa diferente exige conferência de identidade/nome, além
de reconciliação, e não converte `APROVADO` em aprovação financeira.

Ordenação determinística de saída: ordinal do arquivo → ordinal do grupo físico
→ ordinal da linha no Analítico. Ordenar a grade no editor de Excel não altera
a ordem protegida da geração.

### 3.2. Campos finais sugeridos

Manter os 21 campos já existentes:

```text
DOC_CEDENTE, NOME_CEDENTE, SEU_NUMERO, NU_DOCUMENTO, DT_VENCIMENTO,
VL_NOMINAL, DOC_SACADO, NOME_SACADO, VL_PRESENTE, TIPO_PESSOA_SACADO,
ENDERECO, CEP, TP_TITULO, DT_EMISSAO_TITULO, COOBRIGACAO,
TIPO_PESSOA_CEDENTE, NFE, VALOR_PAGO_TITULO, INDEXADOR, TAXA_INDEXADOR, MOVIMENTO
```

| Campo | Sugestão e regra |
|---|---|
| DOC_CEDENTE / NOME_CEDENTE Normal | Documento e nome do crédito do Analítico; PFMI é referência de comparação |
| DOC_CEDENTE / NOME_CEDENTE Cessão da Cessão | CNPJ confirmado da Conexcred / somente CONEXCRED INTERMEDIACAO |
| VL_NOMINAL | N original normalizado |
| VL_PRESENTE | E calculado segundo modalidade |
| DT_EMISSAO_TITULO Normal | Data Assinatura original |
| DT_EMISSAO_TITULO Cessão da Cessão | Vazio; operador do fundo informa assinatura da nova cessão |
| DOC_SACADO / NOME_SACADO / DT_VENCIMENTO | Base, com origem e conflitos visíveis |
| SEU_NUMERO / NU_DOCUMENTO | Mesmo inteiro contíguo a partir de PRIMEIRA_SEQUENCIA |
| Tipos PF/PJ | Derivados dos documentos finais, validando coerência |
| ENDERECO / CEP / NFE | Vazios inicialmente |
| TP_TITULO / COOBRIGACAO / MOVIMENTO | 24 / 2 / 1 |
| VALOR_PAGO_TITULO | Zero |
| INDEXADOR / TAXA_INDEXADOR | Vazios |

Cessão da Cessão: procurar especificamente o beneficiário Conexcred. Somente um
CNPJ único, válido e consistente pode sugerir documento da principal, com origem
visível. Documento ausente ou múltiplos CNPJs incompatíveis deixam pendência para
correção final; não escolher o primeiro e não hardcodificar CNPJ real. O nome final
não inclui o credor original. O crédito e cedente originais continuam nas referências.

Falência Gama (caso de exemplo com advogados envolvidos): identificar o caso para
revisão, sem deduzir identidade dos parênteses. Nome final somente do advogado
(Fulano de Tal ou Beltrano conforme crédito/instrumento), documento confiável da
fonte ou correção do operador responsável. Não completar
CPF incompleto. A aprovação de divergência de nome é a mesma coluna APROVADO, sem
novo fluxo de justificativa. Não automatizar substituição por qualquer beneficiário.

### 3.3. Data e nomes finais

Emissão válida, não posterior à liquidação. A assinatura original da primeira cessão
não preenche emissão na Cessão da Cessão. Nunca derivar liquidação − 1 dia.

Comparar nome final com referência esperada para a modalidade: indicação PFMI no
Normal, principal Conexcred no modo correspondente, tratamento explícito da exceção
dos advogados. Correção válida de nome/documento fica rastreada e alertada.
Divergência material de nome exige correção ou APROVADO=SIM; equivalências explícitas
não bloqueiam. APROVADO não autoriza mudar a regra de cedente final da modalidade.

`formatar_nome_cnab` é diferente de `normalizar_comparacao`:

- preservar nome completo no Excel;
- saída em maiúsculas, sem acentos, espaços especiais normalizados, trim e colapso;
- preservar pontuação representável compatível com o leiaute;
- não converter LIMITADA em LTDA, por exemplo, só por equivalência de comparação;
- gerar campo técnico de até 40 caracteres (40 bytes após CP1252), alertando
  `NOME_TRUNCADO_CNAB` quando houver corte;
- exigir conteúdo alfanumérico e rejeitar caracteres de controle/incompatíveis
  ainda presentes; não esconder erro de codificação com substituição genérica.

## 4. Contrato do Excel intermediário

### 4.1. Esquema e abas

Novo identificador: **`CNAB-NOX-V1-2`**, referente às Specs 1.1.
O esquema anterior `CNAB-NOX-V1-1` é incompatível: recusar geração com mensagem
“Este Excel foi preparado por uma versão anterior. Prepare novamente as fontes
na versão revisada.” Não migrar, sobrescrever ou aceitar parcialmente.

Preservar exatamente as cinco abas:

| Aba | Conteúdo/edição |
|---|---|
| RESUMO | DATA_LIQUIDACAO e PRIMEIRA_SEQUENCIA editáveis; resumo próprio, lista/ordem/modalidade/taxa das PFMIs e instruções |
| CREDITOS | Uma linha por candidato ou placeholder; referências visíveis, 21 finais, seleção, nome aprovado, diferenças e orientação |
| PENDENCIAS_PFMI | Linhas não reconhecidas/inconsistentes com origem; sem exclusão manual para liberar TXT |
| ORIGINAIS_CONTROLE | Oculta; snapshot original, sugestões, fonte da Base, cálculos e pendências de origem |
| MANIFESTO | Oculta; versão, IDs/conjuntos, parâmetros, ordem, contagens, evidências de espelhos e digests de integridade |

Resumos/espelhos conhecidos devem ter registros de alerta/decisão no RESUMO e
controles, sem serem classificados como bloqueio só por ocuparem linhas auxiliares.
Incluir legenda de cores, filtros e formatos legíveis; referências não podem exigir
abrir abas ocultas para o double check cotidiano.

Campos adicionais mínimos, agrupados visualmente:

| Grupo | Informações |
|---|---|
| PFMI | ID/ordem do arquivo, aba/bloco/linhas, modalidade, falência indicada/canônica, nome indicado, pagamentos e total do grupo |
| Crédito | ID_CREDITO, hash do Analítico, aba/linha, referência, nome/documento originais, motivo principal/alternativa |
| Cálculo | A original, N original, assinatura original, p usado, B, C e E; sem substituir os originais |
| Finais | Campos editáveis, nomes completos/técnicos, diferença VL_PRESENTE−E e VL_NOMINAL−N |
| Orientação | STATUS, PENDENCIAS, ALERTAS, ação necessária e se exige nova preparação |

No RESUMO: quantidades de pagamentos, grupos, candidatos e títulos selecionados
separadas; totais PFMI válidos/incompletos, nominal e aquisição selecionados; faixa
sequencial; resumo declarado versus calculado; quantidades por PFMI e do lote.
Pagamentos/grupos/candidatos não podem aparecer com rótulo ambíguo “títulos”.

### 4.2. Edição e integridade

Editáveis: finais, INCLUIR_CNAB=SIM/NAO, APROVADO=SIM/NAO ou vazio, data de liquidação
e primeira sequência no RESUMO. Nas alternativas, NAO é valor inicial, não silêncio
interpretado como consentimento. Campo INCLUIR apagado vira pendência de seleção.

Originais, parâmetros, IDs, ordens, total PFMI, participação em grupos e ocorrências
de fonte são imutáveis. Bloquear TXT em inserção/exclusão/duplicação de linhas,
alteração desses controles, remoção de pendências ou inconsistência do manifesto.

Proteção de células/abas deve prevenir acidentes sem impedir a edição autorizada
nem criar senha, desbloqueio especial ou justificativa formal. Digests verificam
integridade acidental; não alegar assinatura digital ou proteção contra adulteração
intencional por quem pode editar todo o arquivo.

Snapshot deve ser suficiente para reler, validar e gerar sem fontes e sem estado
da GUI: incluir os dados originais relevantes, sugestões, controles de escolha,
pendências de origem, identificação das fontes e parâmetros usados.
Preservar originais sem sobrescrever A por E. Fórmulas em campos lidos não são
aceitas como valores finais; não depender de Excel recalcular fórmulas.

A verificação compara finais aos originais/sugestões por ID_LINHA e refaz
validações. STATUS/PENDENCIAS/ALERTAS editados visualmente não liberam o TXT.
Dados de candidatos não selecionados não entram em totais/títulos; pendências de
origem do grupo/lote não somem ao desmarcar candidatos.

## 5. Segundo botão e validação

Ordem lógica:

1. Ler somente o intermediário; validar esquema, estrutura, IDs e integridade.
2. Recuperar parâmetros capturados e finais editados; validar seleção em todos os grupos.
3. Bloquear placeholder selecionado e ID_CREDITO reutilizado, inclusive entre arquivos.
4. Validar finais selecionados, classificar mudanças e nomes, formatar nomes técnicos.
5. Reconciliar aquisição final por grupo e lote com PFMI original.
6. Ordenar pelas origens protegidas, calcular sequência e renderizar em memória.
7. Validar estrutura/bytes completos, salvar atomicamente e emitir resultado/log.

Regras finais:

- Pelo menos um título e pelo menos um crédito selecionado por grupo.
- Documentos completos: máscaras removidas, zeros reais preservados, dígitos
  verificadores válidos; PF/PJ coerente. Não reconstruir dígitos faltantes.
- Datas reais; emissão não posterior à liquidação; vencimento obrigatório.
- Valores monetários obrigatórios não negativos, finitos, normalizados em centavos
  e dentro das larguras. Limites monetários de nominal/presente são 13 dígitos em
  centavos (R$ 99.999.999.999,99); testar limite e um centavo acima.
- Primeira sequência inteira positiva, NU_DOCUMENTO de até 10 posições e SEU_NUMERO
  de até 25. Incrementar por título na ordem protegida. Sequência física/trailer
  de seis posições limita registros totais a 999999.
- SEU_NUMERO e NU_DOCUMENTO são recalculados; edição individual é ignorada com
  alerta. Para mudar a faixa, editar PRIMEIRA_SEQUENCIA.
- Campos técnicos editados continuam sujeitos às constantes 24/2/1, pago zero,
  indexador/taxa vazios, larguras e conteúdo do leiaute.
- Cada grupo: soma de VL_PRESENTE final selecionado exatamente igual ao total
  PFMI protegido. Lote: mesmo fechamento somando grupos. Diferenças compensadas
  entre grupos não liberam TXT.
- Não impor novamente A/E sobre correção final válida. Diferença entre final e
  sugestão é alerta; diferença residual entre final e PFMI é bloqueio.
- Parametrização de taxa inválida e pendências de origem PFMI continuam bloqueantes
  até nova preparação. Base ausente/conflitante pode ser resolvida nos finais,
  preservando alerta/evidência original.
- APROVADO=SIM resolve apenas DIVERGENCIA_NOME; nunca financeira/documento/origem/
  concorrência/configuração/leiaute.
- No modo Conexcred, exigir nome final da principal conforme contrato; não aceitar
  texto composto legado por mera aprovação genérica de nome.

Catálogo mínimo e correção:

| Código / condição | Excel | Como liberar TXT |
|---|---|---|
| CREDITO_NAO_LOCALIZADO | Placeholder REVISAR | Corrigir fonte e preparar novamente |
| CREDITO_CANDIDATO_EM_MULTIPLOS_GRUPOS | Sem pré-seleção | Seleção humana sem reutilizar crédito e com todos os grupos reconciliados |
| COMISSAO_PARAMETRO_INVALIDO | Pendência de preparação | Corrigir taxa na PFMI da interface e preparar novamente |
| COMISSAO_BASE_NEGATIVA | Revisar cálculo/referências | Corrigir fonte e preparar novamente ou preencher finais válidos reconciliados; sem cálculo automático negativo |
| DIVERGENCIA_VALOR | Valores/diferença visíveis | Corrigir finais ou fontes; fechar cada grupo/lote em centavos |
| DIVERGENCIA_NOME | Revisão | Corrigir nome ou APROVADO=SIM, respeitando modalidade |
| PREENCHER_MANUALMENTE / CONFLITO_BASE_VENCIMENTOS | Gerar com campos pendentes | Corrigir campos finais do sacado |
| Linha desconhecida / pagamento inconsistente | PENDENCIAS_PFMI | Corrigir PFMI e preparar novamente |
| Resumo ausente/divergente, espelho integral | Alerta/registro da decisão | Sem bloqueio por si só |
| Arquivo operacional duplicado | REVISAR, sem descarte | Conferir lista; remover cópia e preparar novamente quando duplicação indevida |
| NOME_TRUNCADO_CNAB / final válido alterado | Alerta | Permitido, com conferência |
| Esquema antigo / controle alterado | Recusa de geração | Nova preparação, sem migração silenciosa |

## 6. Leiaute CNAB 444

Todas as posições abaixo são inclusivas e baseadas em 1.

### Header

| Pos. | Conteúdo |
|---|---|
| 1 | `0` |
| 2 | `1` |
| 3–9 | `REMESSA` |
| 10–11 | `01` |
| 12–26 | `COBRANCA` + 7 espaços |
| 27–46 | originador `00000000000000000125` |
| 47–76 | zeros |
| 77–79 | `001` |
| 80–94 | zeros |
| 95–100 | liquidação `ddmmaa` |
| 101–108 | espaços |
| 109–110 | `MX` |
| 111–117 | `0000001` |
| 118–438 | espaços |
| 439–444 | `000001` |

### Detalhe

| Pos. | Conteúdo |
|---|---|
| 1 | `1` |
| 2–7 | espaços |
| 8 | indexador ou espaço |
| 9–10 | espaços |
| 11–20 | taxa com 7 decimais ou espaços |
| 21–22 | coobrigação, 2 dígitos |
| 23–37 | zeros |
| 38–62 | `SEU_NUMERO`, direita, espaços |
| 63–65 | `001` |
| 66–70 | zeros |
| 71–81 | zeros |
| 82 | `1` |
| 83–92 | valor pago em centavos |
| 93 | `1` |
| 94 | `N` |
| 95–100 | liquidação `ddmmaa` |
| 101–104 | espaços |
| 105 | espaço |
| 106 | `1` |
| 107–108 | espaços |
| 109–110 | movimento |
| 111–120 | `NU_DOCUMENTO`, direita, espaços |
| 121–126 | vencimento `ddmmaa` |
| 127–139 | nominal em centavos |
| 140–142 | zeros |
| 143–147 | zeros |
| 148–149 | tipo de título |
| 150 | espaço |
| 151–156 | emissão `ddmmaa` |
| 157–159 | zeros |
| 160–161 | tipo PF/PJ cedente (`01`/`02`) |
| 162–192 | zeros |
| 193–205 | presente em centavos |
| 206–218 | zeros |
| 219–220 | tipo PF/PJ sacado (`01`/`02`) |
| 221–234 | sacado: CPF `000` + 11 dígitos; CNPJ 14 dígitos |
| 235–274 | nome técnico sacado, 40, esquerda |
| 275–314 | endereço, 40, esquerda |
| 315–326 | espaços |
| 327–334 | CEP, esquerda, zeros à direita |
| 335–380 | nome técnico cedente, máximo 40 + 6 espaços |
| 381–394 | cedente: CPF **3 espaços** + 11 dígitos; CNPJ 14 dígitos |
| 395–438 | NFE, zeros à esquerda |
| 439–444 | sequência física |

### Trailer e arquivo

Trailer: `9`, 437 espaços e quantidade total de registros em 6 dígitos.
Cada registro codifica exatamente 444 bytes em Windows-1252. O arquivo usa CRLF,
possui CRLF final e não possui BOM.

## 7. Segurança, persistência e resultados

Nunca alterar fontes. Nomes exclusivos de saída e gravação atômica; nenhum arquivo
parcial deve parecer final. Cancelamento de diálogo não é falha operacional e
não cria saída. Falha de leitura/gravação/permissão deve informar a ação possível,
sem stack trace ou dado sensível na interface.

Neutralizar textos de entrada iniciados por =, +, - ou @ ao escrever Excel,
preservando a referência textual e impedindo execução de fórmulas.
Dados numéricos legítimos não são convertidos arbitrariamente em texto.

Log local JSONL somente com data/hora, extensão e hash abreviado das entradas,
contagens, totais, faixa sequencial, tipo da saída e resultado. Não registrar
nome/caminho de entrada, documentos, pessoas, células ou dados bancários.
Referências detalhadas são privadas no intermediário, não no log compartilhável.

Se o log falhar após a saída atômica concluída, retornar
`LOG_AUDITORIA_NAO_GRAVADO` como alerta não fatal e indicar a saída já criada.
Não repetir geração silenciosamente. Resultado de preparação informa caminho,
contagens e pendências; resultado de geração informa títulos, nominal, aquisição,
faixa e alertas. Sem telemetria.

## 8. Matriz obrigatória de testes

Fixtures versionadas devem ser sintéticas e expectativas independentes da função
testada. Não calcular o esperado chamando o mesmo cálculo/renderizador sob teste.

| ID | Casos e resultado esperado | Marco |
|---|---|---|
| T01 | Uma/várias PFMIs; adicionar/remover/ordenar; modalidades e taxas independentes; origens preservadas | M1, M6 |
| T02 | 0%, 15%, fração, vírgula/ponto; vazio não vira 15; Normal ignora taxa; 101% não é rejeitado por teto comercial | M1, M3 |
| T03 | Negativo, texto, NaN, infinito, capacidade extrema: Excel com pendência, TXT bloqueado até nova preparação | M1, M4 |
| T04 | Meio centavo ROUND_HALF_UP; arredondamento por crédito vs total; base negativa com taxa positiva; sem rateio | M3 |
| T05 | Cabeçalhos por conteúdo, abas renomeadas, continuação, blocos distintos e linhas idênticas legítimas | M2 |
| T06 | Resumo correto/vazio/divergente; tabelas auxiliares legadas; linha desconhecida continua REVISAR | M2, M4 |
| T07 | Espelho integral contado uma vez com prova; semelhança parcial/total igual não deduplica; arquivo duplicado não some | M2 |
| T08 | Quatro beneficiários → um crédito/título; dois créditos distintos → dois títulos; nenhuma busca de subconjuntos | M3 |
| T09 | Homônimos, valores iguais, referências repetidas, documentos diferentes, concorrência entre grupos e reutilização manual | M3, M4 |
| T10 | Dois Contato; principal associado ao CPF; XLS/XLSX/CSV; campos operacionais inválidos não derrubam preparação | M1, M3 |
| T11 | MOGIANO nas três fontes; falência só parecida não casa; Base ausente/conflitante gera Excel corrigível | M3, M4 |
| T12 | Alternativa mesma falência + valor individual exato começa NAO; placeholder não vira título; similaridade não escolhe | M3, M4 |
| T13 | Conexcred final sem composição; CNPJ válido único vs ausente/conflitante; beneficiário Normal não preenche cedente | M3 |
| T14 | Advogados em Papéis Independência; documento parcial não completado; nome original entre parênteses não vira identidade | M3, M5 |
| T15 | Emissão Normal sugerida; Cessão vazia e original preservada; preencher/corrigir data; emissão após liquidação bloqueia | M3, M4 |
| T16 | Finais válidos alterados em valor/nome/documento/emissão geram alerta e prevalecem; A/N/p/C/E originais intactos | M4 |
| T17 | R$ 0,01 residual por grupo ou lote bloqueia; diferenças opostas entre grupos não compensam; APROVADO não libera | M4 |
| T18 | Cinco abas; alterações de IDs/ordem/contagens/originais/parâmetros; inserção/exclusão; fórmula; esquema antigo recusado | M4 |
| T19 | Preparar, corrigir, tornar fontes inacessíveis, mudar parâmetros da GUI e gerar só pelo Excel: saída não muda | M4, M6 |
| T20 | Nomes 40/41 caracteres, acentos, pontuação, formas jurídicas equivalentes mas saída textual preservada | M5 |
| T21 | Posições numéricas no limite/acima; CP1252, BOM, CRLF final, 444 bytes, CPF/CNPJ, sequência, header/trailer | M5 |
| T22 | Gravação atômica, nomes exclusivos, diretório sem permissão, cancelamento, log indisponível após saída salva | M5, M6 |
| T23 | Comandos e widgets reais Tk: lista, modalidade/taxa, dois botões, pendências, correção e geração com resumo | M6 |
| T24 | Três provas separadas: renderizador, E2E sintético revisado e histórico conforme disponibilidade temporal | M5, M6 |

Importar a GUI com CNAB_NOX_SMOKE=1 não comprova widgets nem fluxo. Testes
de comandos devem acionar os handlers/widgets, isolar diálogos sem pular a lógica
e complementar com exercício visual em sessão gráfica ou Xvfb documentado.
Preservar testes unitários/integração, lint, privacidade e cobertura por módulo/
ramos críticos. Não fixar um percentual isolado como substituto de T01–T24.

### 8.1. Regressão da comissão de 28/08

Valores financeiros dos cinco créditos já confrontados; linhas abaixo são rótulos
sanitizados, não fixtures de identidade. Somente a fórmula não comprova o match.

| Crédito | N | A | Comissão 15% por crédito | E |
|---|---:|---:|---:|---:|
| C1 | 181.800,00 | 69.993,00 | 16.771,05 | 86.764,05 |
| C2 | 25.594,84 | 10.807,55 | 2.218,09 | 13.025,64 |
| C3 | 23.191,21 | 11.767,38 | 1.713,57 | 13.480,95 |
| C4 | 181.800,00 | 95.000,00 | 13.020,00 | 108.020,00 |
| C5 | 16.588,17 | 4.459,95 | 1.819,23 | 6.279,18 |
| Total | 428.974,22 | 192.027,88 | 35.541,94 | 227.569,82 |

O cálculo sobre a base total resultaria 35.541,95 e deve falhar como expectativa
do algoritmo correto. Também deve existir teste sintético independente com valores
pequenos demonstrando meio centavo, sem copiar pessoas/documentos reais.

### 8.2. Corpus histórico e limites de prova

| Pasta em tests/Testes Reais | Data de 2026 | Títulos | Aquisição esperada |
|---|---|---:|---:|
| Teste 01 | 28/08 | 17 | 330.683,32 |
| Teste 02 | 04/09 | 10 | 140.472,89 |
| Teste 03 | 05/08 | 9 | 245.895,55 |
| Teste 04 | 07/08 | 10 | 62.441,84 |
| Teste 05 | 12/08 | 8 | 97.891,51 |
| Teste 06 | 14/08 | 2 | 36.844,60 |
| Teste 07 | 19/08 | 11 | 276.052,73 |
| Teste 08 | 26/08 | 19 | 349.556,90 |
| Total | Oito lotes | 86 | 1.539.839,34 |

28/08 usa duas PFMIs: 12 títulos Normais e cinco de Cessão da Cessão.
Não contar cópia operacional adicional como novo insumo. 19/08 tem quatro pagamentos
de 19.394,89 formando um título de 77.579,56; seu resumo declara 19 TEDs, mas são
18 pagamentos. 26/08 exige reconhecer o espelho integral entre abas sem perder
linhas iguais legítimas.

Separar:

1. **Renderizador:** dados finais conhecidos → bytes esperados; testar todas as posições,
   padding, documentos e pontuação. Referência principal SEM_ZERO...V2.
2. **Fluxo revisado sintético:** fontes completas independentes → preparação → correções
   explícitas → TXT. Não extrair do TXT o próprio input de um teste chamado “matching”.
3. **Histórico real:** comprovar apenas o que PFMI/Analítico/Base disponíveis permitem.
   Snapshot atual do Analítico não prova preenchimento automático na data passada.
   XLSB de gerador não é substituto do Analítico.

Nomes compostos legados da Conexcred e dos advogados serão alterados intencionalmente.
Cada exceção deve ter intervalo, valor antigo e esperado novo definidos localmente,
com relatório compartilhável sanitizado. Não ignorar genericamente campos de nome
nem aceitar qualquer diferença nesses intervalos.

Registrar correções manuais de emissão, identidade ou outros finais como correções,
sem anunciar preenchimento automático. Comparação byte a byte do fluxo completo só
é exigível com entradas correspondentes e expectativas revisadas explícitas.

## 9. Definition of Done e gates

M0: documentação reconciliada e conferida, sem código alterado.

M1–M6: todos os contratos acima implementados, T01–T24 comprovados, suíte/lint/
privacidade verdes, cobertura e lacunas explicadas, GUI exercitada, corpus avaliado
sem perda silenciosa, correções/retestes registrados no ExecPlan e roteiro de
reprodução entregue. Não encerrar por terem passado apenas os 48 testes antigos.

M7: pacote Windows onedir executado e validado na máquina do responsável técnico,
sem depender de Python, Excel ou administrador no ambiente de uso.

M8: importação real/reconciliação no Frontis, aceite operacional do operador do
fundo e aceite final do responsável pela aprovação final. O Goal Linux termina em
M6; a V1 produto somente após M7/M8.
