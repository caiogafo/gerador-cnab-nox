# DRD — Gerador de CNAB NOX

**Versão funcional:** 1.1 — revisão aprovada em 2026-09-04  
**Estado:** revisão funcional implementada; candidato Linux M1–M6 comprovado. M7/M8 posteriores.  
**Responsável pelo produto:** responsável pela aprovação final  
**Uso operacional e aceite:** operador do fundo  
**Alvo:** Windows 10/11 x64, local e offline; desenvolvimento inicial no Linux.

Este documento é a autoridade de negócio. [Specs 1.1](specs.md) detalha contratos e
testes; [ExecPlan](execplan.md) registra execução e evidências. M0 foi documental;
a implementação revisada é o candidato Linux de M1–M6. Isso não representa
executável Windows homologado nem aceite operacional/final.

## 1. Objetivo e limites

Substituir o preenchimento manual do gerador por um fluxo simples:

**Selecionar entradas → preparar Excel → revisar e corrigir → gerar TXT.**

O operador do fundo mantém o double check e a importação manual no Frontis. O programa não escolhe
pessoas por aproximação nem transforma pagamentos bancários em títulos.

Dentro da V1:

- fundo NOX, novas cessões, movimento `1`, título `24` e coobrigação `2` (sem coobrigação);
- PFMIs Normais e de Cessão da Cessão, juntas ou processadas separadamente;
- vários créditos por lote, processamento determinístico e Excel intermediário obrigatório;
- somente os campos e o leiaute CNAB 444 já utilizados, com as correções aprovadas aqui.

Fora da V1: outros fundos, movimentos, XML, IA, leitura automática dos instrumentos,
integração/envio ao Frontis, pagamentos, liquidação parcial, recompra, baixa,
prorrogação `REC`, atualização automática da Base, serviços externos e banco de dados.

## 2. Fluxo e interface mínima

### Configurar lote e preparar/revisar Excel

A interface Linux apresenta uma etapa por vez: **1. Configurar lote →
2. Preparar e revisar Excel → 3. Gerar TXT**, com navegação que preserva os campos.
O atalho **Já tenho um Excel revisado** abre diretamente a terceira etapa.

O operador do fundo revisa as indicações e os pagamentos das PFMIs antes de selecioná-las e confere
os instrumentos quando necessário. Não serão criadas as colunas
`NOME_CEDENTE_CONFIRMADO` e `CPF_CNPJ_CEDENTE_CONFIRMADO`, nem uma grade de edição
dentro do programa.

A configuração mantém:

1. Lista ordenável de uma ou várias PFMIs, com adicionar, remover e mover.
2. Modalidade explícita por arquivo: **Normal** ou **Cessão da Cessão**.
3. **Comissão (%)** por PFMI de Cessão da Cessão, inicialmente preenchida com `15`.
4. Um ProspectAnalítico e uma Base de Vencimentos.
5. Data de liquidação e primeira sequência.
6. Ação **Continuar** para preparar e revisar o Excel na etapa seguinte.

Na segunda etapa, **Preparar Excel** solicita o destino. Após salvar, apresenta
contagens, caminho copiável, **Abrir Excel para revisar**, **Abrir pasta** e roteiro
das abas. A revisão continua no editor de planilhas associado pelo sistema; sem
associação, orientar abertura manual. Não criar grade de edição na aplicação nem
marcar automaticamente o lote como revisado/validado.

Nome do arquivo e presença da Conexcred podem ajudar a sugerir a modalidade, mas
não substituem a escolha explícita. Alterar uma PFMI ou sua taxa não altera as demais.

O Excel será gerado mesmo com pendências de dados ou parâmetros. Somente ausência
de PFMI, Analítico ou Base, formato não suportado, arquivo corrompido/ilegível,
ausência dos cabeçalhos mínimos ou impossibilidade de gravar a saída impedem a
preparação. Data de liquidação, primeira sequência e comissão ausentes/inválidas
permanecem pendências no Excel. Data e sequência podem ser corrigidas no RESUMO;
comissão exige corrigir o parâmetro e preparar novamente. O programa explicará
qual arquivo ou ação precisa de correção.

### Gerar TXT

O operador do fundo abre o Excel em um editor compatível, confere e corrige os campos finais,
seleciona os créditos e salva. **Validar e gerar TXT** recebe **somente esse Excel**:
não reabre fontes, não consulta parâmetros atuais da interface e não refaz o
cruzamento. Recalcula as validações e gera TXT apenas sem pendências bloqueantes.

Pendências mostram orientação e localização disponível, com todos os apontamentos
originais acessíveis em **Detalhes técnicos**. Revisões e avisos da preparação não
são apresentados como uma contagem de bloqueios. Alterar a configuração depois
de preparar exige novo Excel para aplicar essas mudanças; o anterior permanece
independente. Durante processamento, mostrar atividade sem porcentagem fictícia
e impedir edição, geração duplicada e fechamento até a conclusão.

A ausência de Excel instalado não impede o processamento pelo programa. A revisão
humana do intermediário exige um editor compatível; não exige macros.

## 3. Função e autoridade das fontes

### PFMI

É referência da operação/cessão e dos pagamentos, não cadastro definitivo de pessoas.
`CREDOR/CEDENTE` é a indicação revisada pelo operador do fundo para localizar o crédito.
O nome e o documento finais devem estar corretos no intermediário e no CNAB.

O CPF/CNPJ da PFMI pode ser do beneficiário bancário. Em **Normal**, ele não pode
escolher nem preencher automaticamente o documento do cedente. Quando beneficiário
e cedente forem a mesma pessoa, serve apenas como conferência adicional.

`VALOR DA OPERAÇÃO` é o valor financeiro de cada pagamento e deve coincidir com
`VALOR` do beneficiário. O total usa exclusivamente `VALOR` dos beneficiários
(coluna J na PFMI atual); `VALOR DA OPERAÇÃO` fica como conferência.
Somar esses pagamentos dentro do grupo serve para
reconciliar os créditos, não para criar um título por beneficiário ou inferir sozinho
a aquisição de um crédito.

### ProspectAnalítico

Fonte do registro de crédito original, nome/documento original, `Valor Receber`,
`Valor Aquisicao` e `Data Assinatura`. Aceitar XLS, XLSX e CSV por cabeçalhos.
Nos formatos com dois `Contato`, usar o contato principal associado ao CPF,
conforme o formato verificado; não o contato operacional secundário.

Cada registro selecionado do Analítico produz **um título**. Uma pessoa pode ter
dois créditos na mesma falência e ambos continuam separados. `Prospect` sozinho
não é único: a identidade inclui snapshot, linha de origem e referência.

### Base de Vencimentos

Fonte do nome final do sacado, documento e vencimento, com
`NOME_SACADO`, `DOC_SACADO` e `DATA_VENCIMENTO`.

- Duplicatas integralmente equivalentes podem ser reduzidas.
- Sacado ausente: `PREENCHER_MANUALMENTE`.
- Conflito material de nome, documento ou vencimento: `CONFLITO_BASE_VENCIMENTOS`.
- Ambos permitem preparar Excel e bloqueiam apenas TXT até a correção final.
- A Base nunca será atualizada automaticamente.

## 4. Pagamentos, blocos e preservação

As tabelas operacionais serão reconhecidas pelos cabeçalhos, sem depender de nomes
como `Planilha1` ou `Planilha3`. Preservar arquivo, aba, bloco e linha de origem.

Reunir pagamentos apenas dentro do mesmo bloco físico, cedente e falência.
Linhas iguais legítimas e blocos distintos devem permanecer. Não consolidar
globalmente por nome, CPF/CNPJ ou falência.

Reconhecer também os resumos e formatos auxiliares encontrados nos históricos,
incluindo tabelas do gerador legado. O resumo usado pelo operador do fundo no e-mail é auxiliar:
ausência, desatualização ou divergência gera alerta, não bloqueio por si só.
O programa calcula suas próprias quantidades e totais.

Linhas vazias, títulos de fundo e cabeçalhos repetidos reconhecidos não são
pagamentos. Qualquer outra linha preenchida não reconhecida aparece como
`REVISAR`; não pode desaparecer silenciosamente e bloqueia apenas TXT.

Abas espelho podem ser contadas uma única vez **somente** se repetirem integralmente
conteúdo operacional, ordem e estrutura dos blocos de outra aba. Registrar a decisão
no Excel. Algumas linhas iguais ou o mesmo total não comprovam um espelho.

Arquivos separados com conteúdo operacional duplicado são sinalizados para revisão,
sem descarte automático. O operador do fundo pode remover uma cópia da lista e preparar novamente.
Em qualquer caso, o mesmo crédito do Analítico não pode ser utilizado em dois grupos.

Beneficiários não multiplicam títulos. Espólio, herdeiro, advogado ou terceiro segue
a identidade do crédito/instrumento, não apenas a condição de favorecido bancário.
A exceção de casos com advogados configurados está explicitada na seção 7.

## 5. Identificação determinística

Atualização autorizada: o aplicativo permite cadastrar equivalências de falência
confirmadas pela operadora. São persistidas localmente e aplicadas à PFMI, ao
Analítico e à Base na preparação. Nomes originais são preservados; regras usadas
ficam no Excel. Mudanças exigem nova preparação. Não inferir equivalências nem
cadastrar identidades de pessoas automaticamente. Consulte
[operação do cadastro](docs/EQUIVALENCIAS_FALENCIAS.md).

1. Localizar candidatos por nome normalizado do cedente e falência equivalente.
2. Aplicar o de-para confirmado `MOGIANO ↔ MOGIANO TRANSP GERAIS` na PFMI,
   no Analítico e na Base. Não estender por semelhança a outras falências.
3. Pré-selecionar o conjunto integral de candidatos somente se a identidade for
   consistente, os documentos originais forem válidos, a soma esperada fechar
   exatamente e nenhum crédito concorrer a outro grupo.
4. Sem solução automática, mostrar candidatos para escolha por
   `INCLUIR_CNAB=SIM/NAO`. Alternativas com nome diferente só podem ser oferecidas
   pela mesma falência e valor esperado individual exatamente igual ao grupo;
   começam não selecionadas, com indicação de escolha humana.
5. Não procurar subconjuntos/combinações de créditos automaticamente para fechar
   valores. A seleção manual ainda precisa ter origem e fechar financeiramente.
6. Sem candidato útil, mostrar a falta e orientar corrigir a indicação na PFMI
   ou atualizar o Analítico e preparar novamente. Preencher valores em uma linha
   sem origem no Analítico não a transforma em título.

Diferenças apenas de acento, caixa, pontuação, espaços ou abreviações comuns
`LTDA/LIMITADA`, `CIA/COMPANHIA`, `SA/SOCIEDADE ANONIMA` não bloqueiam.
Diferença material gera `DIVERGENCIA_NOME`: corrigir o nome final ou usar
`APROVADO=SIM`. Similaridade classifica a divergência; nunca seleciona uma pessoa.

A identidade original do crédito permanece mesmo quando o cedente final for a
Conexcred ou um advogado. A substituição aprovada da Conexcred não é, por si só,
uma divergência entre pessoas a exigir aprovação de nome.

## 6. Comissão e reconciliação em centavos

A comissão pertence a cada PFMI de **Cessão da Cessão**. `15` significa 15%;
`0` permanece zero. Aceitar inteiros e frações com vírgula ou ponto decimal.
Apagar o campo não restaura 15%. Em Normal, comissão não tem efeito.

Percentual negativo, não numérico, não finito ou fora da capacidade técnica
suportada gera pendência. Não há teto comercial arbitrário de 100%.
Essa pendência permite Excel, mas exige corrigir o parâmetro e **preparar novamente**
antes do TXT; não é um campo final editável para contornar no intermediário.

Por crédito, usando Decimal e arredondamento `ROUND_HALF_UP`:

```text
A = Valor Aquisicao do Analítico, normalizado para duas casas
N = Valor Receber do Analítico, normalizado para duas casas
p = percentual da PFMI

Comissão = arredondar((N − A) × p / 100, 2 casas)
Aquisição esperada = A + Comissão
```

Normal: aquisição esperada = A. Cessão da Cessão a 0%: aquisição esperada = A.
Base negativa para comissão positiva fica para revisão; não converter em zero.
Arredondar por crédito e somente depois somar. Não usar float, ratear nem distribuir
centavos. A fórmula não determina identidade.

No caso de 28/08, os cinco créditos identificados devem reproduzir:

| Aquisição original | Comissão a 15% | Aquisição esperada |
|---:|---:|---:|
| R$ 192.027,88 | R$ 35.541,94 | R$ 227.569,82 |

Calcular a comissão sobre o total produziria R$ 35.541,95, um centavo incorreto
para esta regra. Esse é um caso obrigatório de regressão.

Após normalizar para duas casas, a comparação é exata em centavos. Diferença de
R$ 0,01 gera `DIVERGENCIA_VALOR` no Excel e bloqueia TXT por grupo e no lote.
Uma correção válida nos valores finais prevalece sobre a sugestão calculada,
desde que a reconciliação com a PFMI feche. Não reaplicar a fórmula para desfazer
a correção. Se a origem PFMI estiver errada, corrigir a fonte e preparar novamente.

## 7. Campos finais e casos especiais

| Informação | Normal | Cessão da Cessão |
|---|---|---|
| Crédito original | Registro correspondente do Analítico | Mesmo registro original do Analítico |
| Nominal sugerido | `Valor Receber` | `Valor Receber` |
| Aquisição sugerida | `Valor Aquisicao` | Aquisição original + comissão por crédito |
| Cedente final sugerido | Cedente do crédito, sujeito a conferência | Somente `CONEXCRED INTERMEDIACAO` |
| Documento do cedente | Analítico; validar e permitir correção | CNPJ confirmado da Conexcred |
| Sacado e vencimento | Base de Vencimentos | Base de Vencimentos |
| Emissão sugerida | `Data Assinatura` | Em branco para informar a assinatura da nova cessão |

Somente no modo Cessão da Cessão, um CNPJ válido, único e consistente do beneficiário
Conexcred pode sugerir o documento da principal, com origem explícita para conferência.
Ausência ou conflito deixa pendência. Essa exceção não autoriza usar qualquer
documento de beneficiário em Normal nem hardcode de dados reais.

Na exceção de casos com advogados configurados (`lawyer_rules`, ex.: **Falência Gama**),
o CNAB contém somente o advogado como cedente, sem o credor original entre
parênteses. Se as fontes não trouxerem documento confiável, o operador do fundo
corrige nome e documento no Excel. Não extrair identidade dos
parênteses nem completar CPF incompleto. Manter crédito original e motivo da
substituição para conferência, sem exigir justificativa formal.

Emissão é a assinatura do instrumento/termo da cessão correspondente. Na Cessão
da Cessão, preservar a assinatura original apenas como referência e pedir a data
da nova cessão no Excel. Não usar automaticamente a assinatura original nem
liquidação menos um dia. A emissão final deve ser válida e não posterior à liquidação.

### Campos do gerador

| Campo final | Regra |
|---|---|
| `DOC_CEDENTE`, `NOME_CEDENTE` | Conforme modalidade/exceção acima; conferência e correção final |
| `SEU_NUMERO`, `NU_DOCUMENTO` | Iguais, sequência contígua calculada a partir do primeiro número |
| `DT_VENCIMENTO` | Sugestão da Base |
| `VL_NOMINAL` | Sugestão de Valor Receber |
| `DOC_SACADO`, `NOME_SACADO` | Sugestão da Base |
| `VL_PRESENTE` | Sugestão de aquisição segundo a modalidade |
| `TIPO_PESSOA_SACADO`, `TIPO_PESSOA_CEDENTE` | Derivados dos documentos válidos |
| `ENDERECO`, `CEP`, `NFE` | Inicialmente vazios; preenchimento técnico conforme leiaute |
| `TP_TITULO`, `COOBRIGACAO`, `MOVIMENTO` | Fixos 24, 2 e 1 |
| `DT_EMISSAO_TITULO` | Regra da assinatura acima |
| `VALOR_PAGO_TITULO` | Zero |
| `INDEXADOR`, `TAXA_INDEXADOR` | Vazios; não recebem a comissão |

O operador do fundo pode editar os campos finais, inclusive valores, documentos e emissão. Isso
não amplia os movimentos da V1 nem autoriza valores técnicos incompatíveis.
A sequência é calculada: para mudá-la, editar `PRIMEIRA_SEQUENCIA` no resumo,
não os números individuais. Tipos PF/PJ devem ser coerentes com os documentos.

A primeira sequência é informada pelo operador do fundo com base no estoque do fundo, coluna
número de documento/seu número. Seguir: ordem dos arquivos na interface → ordem
dos grupos físicos → ordem dos créditos no Analítico. Não inferir pelo intervalo
de um gerador antigo. A liquidação é única por lote.

## 8. Intermediário e rastreabilidade

Preservar as cinco abas: `RESUMO`, `CREDITOS`, `PENDENCIAS_PFMI`,
`ORIGINAIS_CONTROLE` e `MANIFESTO`, com novo esquema incompatível com o anterior.

Mostrar origem/modalidade, crédito original, cedente final, motivo da correspondência,
aquisição original, nominal, percentual, comissão, aquisição calculada, valores
finais, diferenças, alertas e instrução de correção. Separar claramente:

- quantidade de pagamentos bancários;
- quantidade de grupos físicos de operação;
- quantidade de títulos selecionados.

Originais, parâmetros usados e identificadores ficam protegidos contra alteração
acidental. Os valores originais do Analítico não são substituídos pelos calculados.
As referências necessárias à conferência ficam visíveis; controles detalhados
podem permanecer nas duas abas ocultas existentes.

`INCLUIR_CNAB` permite escolher créditos. `APROVADO=SIM` confirma **somente**
divergência material de nome, sem justificar por texto. Não libera documento
inválido, diferença financeira, origem ausente, reutilização de crédito ou
parametrização inválida.

Alterações válidas em campos finais geram alertas. Os textos de status não são
autoridade: o segundo botão revalida os valores efetivos e os controles.
Intermediários antigos são recusados com orientação de nova preparação, sem
migração silenciosa.

## 9. O que bloqueia cada etapa

| Ocorrência | Preparar Excel | Gerar TXT / resolução |
|---|---|---|
| Arquivo obrigatório ausente, ilegível ou sem cabeçalhos mínimos | Não é possível preparar | Corrigir seleção/arquivo |
| Falta de candidato ou ambiguidade | Gerar com REVISAR | Escolher candidato válido ou corrigir fontes e preparar novamente |
| Sacado ausente/conflitante | Gerar com pendência | Corrigir campos finais da Base no Excel |
| Parâmetro de comissão inválido ou pagamento PFMI inconsistente | Gerar com pendência de origem | Corrigir parâmetro/fonte e preparar novamente |
| Resumo conhecido vazio/divergente | Gerar com alerta | Não bloqueia por si só |
| Linha preenchida desconhecida | Gerar com REVISAR | Corrigir origem e preparar novamente |
| Documento/data/valor final inválido | Gerar com pendência | Corrigir no Excel |
| Diferença financeira de um centavo | Gerar com DIVERGENCIA_VALOR | Reconciliar exatamente |
| Mudança final válida | Gerar/mostrar alerta | Permitida se demais validações passarem |
| Nome longo | Preservar completo e alertar | Nome técnico limitado a 40, não bloqueante por comprimento |
| Controles alterados ou esquema antigo | Não se aplica à preparação nova | Recusar TXT e orientar nova preparação |

Além disso, cada grupo deve ter seleção válida, cada crédito só pode aparecer uma vez
no lote, e todos os campos fixos, documentos, datas, totais e limites do leiaute
devem passar antes de gravar o TXT.

## 10. Nomes e TXT

Separar comparação de nomes da formatação de saída. O Excel preserva o nome completo.
O nome técnico usa maiúsculas, sem acentos, espaços normalizados e pontuação compatível.
Limitar a 40 caracteres com alerta não bloqueante; não substituir automaticamente
formas jurídicas só porque são equivalentes no cruzamento.

Preservar o padrão homologado `SEM_ZERO_cnab_liquidacao_NOX_28082026V2.txt`:

- 444 bytes por registro, Windows-1252, CRLF inclusive final, sem BOM;
- header 0, detalhes 1 e trailer 9, sequência física de seis posições;
- trailer inclui header e trailer na contagem;
- CPF do cedente: três espaços + 11 dígitos; CPF do sacado: 000 + 11 dígitos;
- CNPJ: 14 dígitos, preservando zeros reais; nada de completar documentos incompletos;
- código do originador preservado: `00000000000000000125`.

O mapa completo está nas Specs. A referência `COM_ZERO` não é padrão correto
para CPF de cedente. Textos compostos antigos da Conexcred e dos advogados mudam
intencionalmente; não se exige repetir esses nomes legados.

## 11. Segurança e operação

Nunca alterar entradas reais. Gravar novas saídas de forma atômica com nomes
exclusivos. Neutralizar fórmulas injetáveis em textos exportados ao Excel.
Não usar APIs, telemetria ou nuvem.

Logs locais e evidências compartilháveis contêm apenas contagens, totais, hashes
e resultados sanitizados; não nomes de pessoas, documentos, conteúdo bruto ou
caminhos de entrada que revelem dados. Dados reais permanecem locais; fixtures
e goldens versionados são sintéticos. Falha no log após uma saída já salva gera
alerta, não afirmação falsa de que a saída falhou.

## 12. Aceite e etapas posteriores

M0 fecha apenas estes documentos. M1–M6 entregarão candidato comprovado no Linux:
entradas múltiplas, leitura correta, matching/cálculos, intermediário, renderizador,
interface realmente exercitada e testes revisados. Os 48 testes e 81% de cobertura
confirmados anteriormente pertencem ao comportamento anterior; não encerram a revisão.

A prova técnica separa renderizador byte a byte, fluxo sintético completo e
históricos reais conforme as fontes disponíveis. O corpus de oito lotes tem
86 títulos e R$ 1.539.839,34 de aquisição. Não prometer reprodução automática
integral sem snapshots correspondentes do Analítico/Base. Registrar toda correção
manual empregada nos testes, sem chamá-la de preenchimento automático.

M7: gerar e executar o pacote Windows `onedir` na máquina do responsável técnico,
sem dependência de Python/Excel/administrador no computador de uso; testar
diálogos, caminhos, permissões, codificação, antivírus e SmartScreen.

M8: o operador do fundo processa lote real no Frontis, confere aceitação, quantidade,
nominal, aquisição, sequência, cedentes, sacados e vencimentos. O operador do fundo
dá aceite operacional; o responsável pela aprovação final dá aceite final do produto.
Candidato Linux não equivale a V1 produto concluída.

Referências privadas de consulta: geradores XLSB legados, PFMIs, Analítico, Base,
TXT homologado e pastas `tests/Testes Reais/Teste 01` a `Teste 08`.
XLSB legado é fonte de auditoria, não novo formato obrigatório de entrada da aplicação.
