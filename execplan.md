# ExecPlan vivo — V1 revisada do Gerador CNAB NOX

**Início do projeto:** 2026-09-03  
**Revisão do plano:** 2026-09-05  
**Estado inicial:** implementação anterior disponível; revisão funcional aprovada; implementação da revisão pendente.  
**Marco atual:** M1–M6 concluídos e comprovados no Linux; M7/M8 pendentes externos.  
**Contrato:** [DRD](drd.md) e [Specs 1.1](specs.md).  
**Execução atual:** melhoria da interface azul concluída e comprovada no Linux em 2026-09-05. Tk/ttk, cartões e controles arredondados, Noto Sans com Xft no Tk autorizado do Ubuntu. Suíte final: 262 testes aprovados, sem skips, 93% de cobertura; Ruff, importação e privacidade aprovados. Quatro cenários ampliados de layout/teclado e dois fluxos sintéticos finais até TXT concluídos. M1–M6 e a auditoria real anterior permanecem encerrados; seus 20 cenários ficaram bloqueados para TXT. Windows/Frontis permanecem posteriores.

## 1. Objetivo e limites desta execução

Implementar e comprovar M1–M6 conforme autorização do responsável pela aprovação final nesta tarefa.
O Goal entrega candidato funcional e comprovado no Linux, pelo
fluxo entradas → Excel → revisão humana → TXT.

O fechamento local não representa executável Windows testado, aceitação no
Frontis ou aceite do produto. Esses resultados pertencem a M7 e M8.

Não ampliar para banco de dados, rede, outros fundos/movimentos, IA, XML, automação
do Frontis ou pagamentos. Não reabrir decisões de negócio já aprovadas por simples
dificuldade técnica. Corrigir falhas técnicas e retestar dentro do marco; não pedir
nova autorização a cada milestone.

## 2. Baseline e revalidação antes do Goal

### Estado conhecido em 2026-09-04

- A implementação anterior está em `src/gerador_cnab_nox`, com testes e scripts.
- 48 testes, 81% de cobertura, lint e privacidade foram confirmados na etapa anterior.
  São evidências do comportamento anterior, não da revisão 1.1.
- O smoke CNAB_NOX_SMOKE=1 importa a interface; não comprova uso dos dois botões.
- A implementação anterior usa uma PFMI e esquema CNAB-NOX-V1-1.
- A revisão requer múltiplas PFMIs, modalidade/taxa por arquivo, novos cruzamentos,
  leitura de auxiliares/espelhos e esquema CNAB-NOX-V1-2.
- O checkout estava em `main`, sem commit inicial/HEAD, com arquivos não rastreados.
  Preservar esse trabalho. Não usar git diff vazio como prova de ausência de mudanças.
- Referências reais estão locais, incluindo `tests/Testes Reais/Teste 01` a `Teste 08`.
  Não modificar, versionar ou publicar esses insumos.
- Executável Windows, Frontis e aceites não foram comprovados.

Antes de M1, ler instruções vigentes e estes documentos, verificar branch/HEAD,
status, worktrees/escritor ativo, ambiente Python/Tk, scripts e arquivos disponíveis.
Reproduzir o baseline e registrar divergências. Se houver alterações preexistentes,
preservá-las e trabalhar sobre evidência atual; não restaurar arquivos pela memória.

Comandos de entrada, no diretório do projeto:

```bash
git status --short
git branch --show-current
git rev-parse --verify HEAD
uv sync --group dev
./scripts/check.sh
```

A ausência de HEAD é esperada no baseline acima, não um teste de produto vermelho.
Se já existir HEAD na retomada, registrar o novo estado, sem presumir que o histórico
aqui continua atual. `uv sync` deve usar o lock existente; resolver dependências
locais necessárias sem atualização indiscriminada de bibliotecas.

## 3. Decisões fechadas

1. Uma lista ordenável de PFMIs; cada item Normal ou Cessão da Cessão; um Analítico
   e uma Base por lote. Um arquivo também continua sendo fluxo válido.
2. Taxa por PFMI de Cessão: inicialmente 15%, editável, inclusive 0 ou fração,
   vírgula/ponto. Vazio não volta para 15; Normal ignora taxa; sem teto comercial de 100%.
3. Comissão por crédito: ROUND_HALF_UP((N−A)×p/100, 2), com A/N já em centavos;
   aquisição esperada A+comissão. Não distribuir centavos nem escolher identidade pelo valor.
4. Preservar blocos e linhas; auxiliares conhecidos são alertas; espelho apenas integral;
   duplicação entre arquivos sinalizada, não descartada; crédito original nunca reutilizado.
5. Mesmo nome/falência não une créditos. Candidatos principais determinísticos;
   conjunto integral exato pode ser pré-selecionado; sem busca de subconjuntos.
   Alternativas da mesma falência/valor individual exato começam não selecionadas.
6. Equivalência confirmada MOGIANO ↔ MOGIANO TRANSP GERAIS nas três fontes.
7. Conexcred como cedente final da Cessão da Cessão, sem nome composto; documento
   confiável sugerido do beneficiário Conex ou preenchido no Excel. Advogados de
   Papéis Independência sem nome do credor entre parênteses e sem documento inventado.
8. Emissão Normal vem da assinatura original; na Cessão da Cessão, fica vazia
   para a nova assinatura, preservando a anterior apenas como referência.
9. Finais corrigíveis no Excel, originais/parâmetros protegidos, nenhuma justificativa
   formal. APROVADO resolve apenas nome. Segundo botão usa somente o intermediário.
10. Esquema novo recusa antigo; nome técnico até 40 com alerta; preservar pontuação
    compatível e leiaute 444/CP1252/CRLF/sem BOM, inclusive CPF homologado.
11. Sequência segue arquivos → grupos físicos → linhas do Analítico. O operador
    do fundo informa início pelo estoque do fundo; não inferir pelo último arquivo de exemplo.
12. O contrato detalhado e limites técnicos de representação estão nas Specs 1.1.
    Não confundir capacidade numérica com regra comercial.

## 4. Milestones, dependências e aceite

| Marco | Dependência | Estado | Entrega |
|---|---|---|---|
| M0 — Revisão documental | Aprovação funcional | Concluído | DRD, Specs 1.1, ExecPlan, uso/homologação e prompt reconciliados |
| M1 — Entradas e parâmetros | M0 | Comprovado | Modelos/API para PFMIs ordenadas, modalidade, taxa e proveniência |
| M2 — Leitura das PFMIs | M1 | Comprovado | Blocos, auxiliares, resumos, espelhos e duplicidades |
| M3 — Cruzamento e cálculos | M1–M2 | Comprovado | Identidade, candidatos, Base, comissão e cedentes finais |
| M4 — Intermediário e validação | M1–M3 | Comprovado | Novo esquema, edição rastreável e geração independente das fontes |
| M5 — TXT e regressões | M3–M4 | Comprovado | Renderização e três tipos de prova |
| M6 — Interface e candidato Linux | M1–M5 | Comprovado | Dois botões exercitados, suíte, revisão e evidências locais |
| M7 — Windows | M6 + ambiente Windows do responsável técnico | Pendente externo | Pacote onedir e fluxo completo sem Python/Excel/administrador no uso |
| M8 — Homologação operacional | M7 + operador do fundo/Frontis/responsável pela aprovação final | Pendente externo | Lote real reconciliado e aceites |

### M0 — Revisão documental

Superfície autorizada: drd.md, specs.md, execplan.md, README.md,
docs/HOMOLOGACAO_WINDOWS.md e docs/PROMPT_GOAL_V1_1.md.

- Substituir regras antigas conflitantes, sem apagar a distinção entre histórico e alvo.
- Preservar mapa posicional já documentado.
- Registrar cálculos, exceções, limites, matriz T01–T24 e corpus.
- Deixar roteiro do Goal sem iniciá-lo.
- Verificar links, estados, consistência, soma dos históricos e integridade dos
  arquivos de código/testes/scripts/insumos não alterados.
- Fazer revisão independente documental; corrigir achados dentro de M0.

Aceite: nenhum conflito normativo conhecido; documentação pronta para execução;
nenhuma alteração de código/dados nem alegação de recurso revisado já implementado.

### M1 — Entradas e parâmetros

Superfície futura: models, normalize, service e testes correspondentes.
Dependência: M0.

Implementar lista de entradas com modalidade/ordem/taxa individual, proveniência,
snapshot/hash e representação de pendências de parâmetros sem quebrar a preparação.
Definir claramente o contrato de cálculo e o estado “taxa inválida”. Preservar
assinatura original e identidade composta de crédito. Não montar ainda uma segunda GUI.

Testes T01–T03 e contrato de T10: uma/várias PFMIs, taxas diferentes, 0/15/fração,
vírgula/ponto, vazio, negativos, NaN/infinito/capacidade, Normal inativo e 101%.
Planejar `tests/test_parameters.py` (novo em M1) e ampliar testes existentes.

Verificação após implementação:

```bash
uv run pytest tests/test_parameters.py tests/test_normalize.py
uv run ruff check src tests
```

Aceite: contrato permite Excel com taxa inválida identificada e preserva parâmetros
independentes; nenhum zero perdido nem percentual global escondido.

### M2 — Leitura das PFMIs

Superfície futura: readers, modelos necessários e tests/test_readers.py.
Dependência: M1.

Implementar classificação por cabeçalhos, proveniência completa, continuação física,
resumos/auxiliares conhecidos, espelho integral e duplicação entre arquivos.
Preservar linhas iguais legítimas e desconhecidas. Não preencher uma falha de total
como zero e não transformar tabela do gerador em títulos extras.

Testes T05–T07: formatos sintéticos derivados das estruturas, nomes arbitrários de
abas, resumo correto/vazio/divergente, parcial vs espelho integral, arquivo duplicado,
blocos distintos e divergência entre valores do pagamento. Conferir contagens do
corpus disponível localmente, sem usar datas/nomes para hardcode.

```bash
uv run pytest tests/test_readers.py
```

Aceite: toda linha operacional contabilizada ou explicitamente pendente; decisões
de espelho auditáveis; nenhum resumo auxiliar bloqueante por si só.
Registrar por pasta arquivos efetivamente usados, pagamentos/grupos, espelhos,
duplicações e linhas pendentes em relatório local sanitizado.

### M3 — Cruzamento e cálculos

Superfície futura: matching, normalize, modelos e testes.
Dependências: M1–M2.

Aplicar equivalência única de falência em todas as fontes; corrigir Contato principal;
identidade por snapshot/linha/referência. Implementar candidatos principais,
alternativas não selecionadas, detecção global de concorrência e proibição de
subconjuntos automáticos. Implementar comissão individual e autoridade dos campos,
Conexcred/advogados, documento confiável e emissão da nova cessão vazia.

Testes T04, T08–T15. Acrescentar `tests/test_commission.py` (novo em M3).
Obrigatórios: homônimo/mesmo valor, Prospect repetido, crédito concorrente, quatro
pagamentos/um título, dois créditos/dois títulos, ausência de candidato, Base
conflitante/ausente, Conex com CNPJ único/conflitante, Normal sem identidade do favorecido.

```bash
uv run pytest tests/test_matching.py tests/test_commission.py tests/test_normalize.py
```

Aceite: cinco créditos de referência de 28/08 totalizam A=192.027,88,
C=35.541,94 e E=227.569,82; cálculo sobre total de comissão 35.541,95 não é aceito
como algoritmo. Testes de identidade são independentes dos valores financeiros.
Nenhum caso ambíguo recebe pessoa automaticamente.

### M4 — Intermediário e validação

Superfície futura: workbook, validation, service e testes.
Dependências: M1–M3.

Implementar CNAB-NOX-V1-2 mantendo cinco abas, referências/cálculos visíveis e
controles protegidos. Distinguir originais, esperado e finais; matriz de pendências
corrigíveis no Excel versus origem/configuração que exige preparar novamente.
Rejeitar esquema antigo, alteração estrutural, placeholder e reutilização.

Testes T16–T19 e validações anteriores: editar nome/documento/valor/emissão;
correção válida prevalece sobre a sugestão; APROVADO só nome; R$ 0,01 residual;
diferenças que se compensam no lote mas não no grupo; controles/taxa/ordem alterados;
fontes removidas e GUI modificada depois da preparação sem afetar segundo passo.

```bash
uv run pytest tests/test_workbook.py tests/test_validation.py tests/test_e2e.py
```

Aceite: Excel corrigido gera pela sua própria informação, sem consulta às fontes;
pendências reais bloqueiam; referências não são sobrescritas nem falsamente liberadas.
Inspecionar legibilidade do Excel sintético e registrar correções empregadas.

### M5 — TXT e regressões

Superfície futura: cnab, normalização de saída, testes/goldens sintéticos e
verificação histórica local necessária.
Dependências: M3–M4.

Separar normalização de comparação e saída. Manter todos os campos posicionais.
Criar expectativas de bytes independentes; nome 40/41, pontuação/forma jurídica,
CP1252 e documentos, limites monetários/sequenciais, trailer e gravação atômica.
A comissão nunca vai para indexador.

```bash
uv run pytest tests/test_cnab.py tests/test_e2e.py
```

Testes T20–T22/T24 e corpus da seção 5. Se necessário, criar neste marco um pequeno
verificador local reproduzível de históricos; registrar o comando real e o
manifesto de entradas/expectativas no diário, sem publicar entradas/saídas privadas.
Não usar o renderizador sob teste para calcular seus próprios bytes esperados.

Aceite: prova do renderizador e E2E revisado verdes; 86 títulos de referência
avaliados com limites de disponibilidade explicitados; diferenças de nomes
intencionais exatamente delimitadas; nenhuma máscara genérica de campo.

Ausência de snapshot histórico contemporâneo não impede terminar M5 se as outras
provas e a análise do material disponível estiverem completas. Não inventar fonte,
identidade ou preencher automaticamente o que foi corrigido manualmente.

### M6 — Interface e candidato Linux

Superfície futura: gui, service, scripts/testes necessários, README/runbook e
evidências do ExecPlan. Dependências: M1–M5.

Implementar lista ordenável, modalidade/taxa por item e resumos separados de
pagamentos/grupos/títulos. Manter os dois botões. Cancelamentos não geram saída;
erros orientam o operador do fundo; alertas pós-gravação não afirmam falha inexistente.

Criar `tests/test_gui.py` neste marco: exercitar comandos e widgets reais, não
somente importar o módulo. Usar diálogos controlados e sessão Tk gráfica; em
ambiente headless, Xvfb quando disponível. Não declarar experiência visual validada
apenas pela simulação de handlers. Registrar limitações e resolvê-las antes de
marcar o marco como comprovado.

```bash
uv run pytest tests/test_gui.py
./scripts/check.sh
uv run python -m gerador_cnab_nox
```

Testes T01/T19/T22/T23: adicionar/remover/reordenar, alternar modalidade/taxa,
cancelar seleções/gravação, preparar pendências, corrigir Excel e gerar TXT.
Exercitar ao menos lote Normal, Cessão e misto com taxas independentes.

Revisar diff/arquivos, testar regressões, conferir logs e privacidade. Medir cobertura
atual com ramos críticos e explicar lacunas; os 81% antigos não são meta suficiente.
Atualizar documentação para o comportamento realmente entregue, sem adiantar M7/M8.

Aceite: suíte/lint/privacidade verdes, matriz T01–T24 coberta, fluxo visual e comandos
comprovados, evidências reproduzíveis e riscos residuais explícitos.
Somente aqui o Goal Linux poderá ser marcado concluído.

### M7 — Windows

Estado: pendente externo, não executar por analogia no Linux.
Dependência: M6 e ambiente Windows do responsável técnico.

Seguir [roteiro Windows](docs/HOMOLOGACAO_WINDOWS.md). Gerar pacote onedir completo,
testar em Windows 10/11 x64 com ambiente de uso sem Python/Excel/administrador.
Verificar diálogos, caminhos acentuados/espaços, gravação/permissão, codificação,
antivírus e SmartScreen. Registrar ambiente, hash do pacote e resultados.
Não desativar proteções para “aprovar” o teste.

Aceite: fluxo completo no pacote demonstrado, não apenas build concluído.
Somente depois encaminhar à máquina do operador do fundo.

### M8 — Homologação operacional

Estado: pendente externo. Dependência: M7 e operador do fundo/Frontis/responsável
pela aprovação final.

O operador do fundo prepara lote real, faz double check, gera e importa
manualmente no Frontis. Conferir quantidade de títulos, nominal, aquisição,
sequência, cedentes, sacados, emissão/vencimentos e resultado da importação.
Registrar aceite operacional do operador do fundo e aceite final do responsável
pela aprovação final separadamente.

Aceite: ambos explícitos. Não presumir aceitação porque TXT tem 444 bytes ou porque
um arquivo do gerador antigo já entrou no Frontis.

## 5. Plano de prova e corpus

A matriz detalhada T01–T24 está em [Specs, seção 8](specs.md#8-matriz-obrigatória-de-testes).
Cada ID deverá apontar para teste/caso executado e evidência; nenhum ID será marcado
coberto só por existir uma função com nome parecido.

Três provas separadas:

| Prova | Fonte do esperado | Limite |
|---|---|---|
| Renderizador | Campos finais conhecidos, mapa posicional e golden independente | Não prova matching |
| E2E revisado sintético | Entradas sintéticas completas e correções planejadas | Não prova automaticamente histórico real |
| Histórico real | PFMIs/TXTs e Analítico/Base disponíveis por lote | Só comprova os campos e automatismos sustentados pelas fontes |

| Teste / data de 2026 | Títulos esperados | Aquisição |
|---|---:|---:|
| 01 — 28/08 | 17 | 330.683,32 |
| 02 — 04/09 | 10 | 140.472,89 |
| 03 — 05/08 | 9 | 245.895,55 |
| 04 — 07/08 | 10 | 62.441,84 |
| 05 — 12/08 | 8 | 97.891,51 |
| 06 — 14/08 | 2 | 36.844,60 |
| 07 — 19/08 | 11 | 276.052,73 |
| 08 — 26/08 | 19 | 349.556,90 |
| Total | 86 | 1.539.839,34 |

Cuidados obrigatórios:

- Teste 01: duas PFMIs efetivas (12 Normal + 5 Cessão), não incluir a cópia operacional.
- Teste 02: XLSB legado não é Analítico; não apresentar valores extraídos do gerador
  como match automático do pipeline.
- Teste 07: 18 pagamentos/11 títulos; quatro pagamentos de 19.394,89 formam
  77.579,56; resumo com 19 TEDs gera alerta.
- Teste 08: espelho integral de abas contado uma vez, sem apagar linhas iguais legítimas.
- Nomes Conexcred/advogados diferem intencionalmente dos textos compostos antigos.
  Especificar cada expectativa; não ignorar todo o campo de nome.
- Analítico atual pode ter datas/valores distintos da época. Registrar lacuna,
  campo manual e efeito; não exigir histórico impossível nem maquiar automatismo.
- Corpus bruto e resultados nominais ficam locais/ignorados. Evidências compartilháveis
  usam contagens, totais, hashes e resultados sanitizados.

## 6. Definition of Done do Goal Linux

- [x] M1–M6 concluídos com evidência, não apenas código escrito.
- [x] T01–T24 associados a testes/casos reais executados.
- [x] Preparação funciona com pendências operacionais, sem perda silenciosa.
- [x] Taxas por PFMI e cálculo por crédito corretos, inclusive 0% e arredondamento.
- [x] Identidade/créditos não escolhidos por aproximação nem reutilizados.
- [x] Excel novo rastreável e corrigível; esquema antigo recusado.
- [x] Geração independente das fontes e da GUI; centavos exatos por grupo/lote.
- [x] Renderizador/golden/limites e diferenças intencionais verificados.
- [x] Interface exercitada pelos comandos/widgets e conferência visual.
- [x] Suíte, lint, privacidade e análise de cobertura atuais registradas.
- [x] Código revisado, falhas corrigidas e retestadas, dados/alterações preservados.
- [x] Documentos reconciliados com o resultado e roteiro M7/M8 reproduzível.
- [x] Relatório final distingue candidato Linux e V1 produto; M7/M8 permanecem pendentes.

## 7. Como manter este plano vivo

Um escritor por superfície compartilhada. Especialistas podem revisar/levar tarefas
independentes com ownership explícito; coordenação não autoriza alteração de negócio.
Registrar progresso após cada marco e depois de correções materiais.

Modelo de registro:

```text
Data/hora e marco:
Objetivo / estado inicial:
Arquivos e comportamento alterados:
Comandos executados e resultado:
Testes Txx atendidos / contagens / cobertura:
Evidência sanitizada (arquivo, hash, relatório ou saída resumida):
Problema encontrado → correção → reteste:
Limitações reais / gates externos:
Estado final do marco:
Próximo passo:
```

Distinguir “planejado”, “implementado”, “testado”, “comprovado” e “pendente externo”.
Falha técnica exige diagnóstico/correção, não novo pedido de autorização.
Mudança material de negócio, risco ou autoridade externa exige o responsável pela aprovação final.
Ausência de fonte histórica contemporânea é limite da prova histórica, não licença
para inventar dados nem motivo para ignorar testes sintéticos possíveis.

## 8. Riscos e mitigação

| Risco | Resposta |
|---|---|
| Taxa global/zero perdido | Parâmetro por arquivo, texto original, testes de transição/vazio |
| Identidade por valor/fórmula | Candidatos determinísticos, alternativa NAO e origem obrigatória |
| Espelho confundido com repetição legítima | Assinatura integral ordenada com estrutura de blocos |
| Corretor financeiro escondido | Centavos exatos, sem subset search, tolerância ou distribuição |
| Final reescrito pelo cálculo | Snapshot separado e teste de correção válida independente |
| Alteração de controles no Excel | Esquema, IDs e digests; não alegar segurança criptográfica |
| Teste circular/golden complacente | Esperado independente e diferenças intencionais restritas |
| Cobertura aparente da GUI | Comandos/widgets + fluxo visual, import smoke insuficiente |
| Dados reais em evidências | Insumos locais, fixtures sintéticas, logs e relatórios sanitizados |
| Linux declarado equivalente a Windows | M7 separado, execução real do pacote |
| Produto declarado pronto antes do operador do fundo | M8 com dois aceites explícitos |

## 9. Histórico preservado — implementação anterior

A numeração abaixo pertence ao plano antigo e não conclui marcos da revisão 1.1.

- 2026-09-03: fontes e macro inspecionadas; padding de CPF confirmado; contratos
  iniciais e implementação anterior desenvolvidos (antigos M1–M5).
- 2026-09-03: scripts/check.sh passou com 48 testes, lint, 81% de cobertura total
  e privacidade; GUI importada sem exercício do fluxo gráfico.
- 2026-09-03: smoke local sanitizado leu 6 pagamentos, 4.414 créditos do Analítico
  e 44 registros da Base; gerou duas candidatas, uma OK e uma REVISAR, e bloqueou
  TXT com pendências. Esses números são daquele insumo, não do corpus revisado.
- 2026-09-03: cinco abas sintéticas inspecionadas visualmente, controles/manifesto
  coerentes com o esquema anterior.
- 2026-09-04: auditorias dos novos históricos e respostas do operador do fundo e do
  responsável pela aprovação final revelaram requisitos não cobertos pela
  implementação anterior; baseline 48/81% reconfirmado.
- A regra antiga que tornava resumo conhecido da PFMI um bloqueio foi substituída
  por alerta na revisão 1.1. O bloqueio antigo não é comportamento-alvo agora.
- Empacotamento Windows, Frontis e aceites permaneceram sem comprovação.

## 10. Registro da revisão atual

- 2026-09-04 — M0 concluído: plano aprovado aplicado somente aos seis documentos
  indicados na superfície de M0. DRD, Specs 1.1, ExecPlan, README, roteiro Windows/
  Frontis e prompt da próxima conversa reconciliados.
- Verificações de M0: 19 links/âncoras internos resolvidos; blocos de código Markdown
  balanceados; matriz T01–T24 sem omissão; seção posicional 6 das Specs preservada
  literalmente em relação ao documento anterior.
- Valores documentados reconferidos com aritmética inteira em centavos: oito lotes,
  86 títulos e 153.983.934 centavos; cinco cálculos individuais da comissão corretos,
  total 3.554.194 centavos, contra 3.554.195 no cálculo incorreto sobre o agregado.
- SHA-256 antes/depois confirmou 51 arquivos de código, testes, scripts e insumos
  locais integralmente preservados. Nenhum código, fixture ou entrada foi editado.
  Como não há HEAD inicial, a prova de preservação é por hashes, não por git diff.
- Revisão independente documental concluída sem achados materiais remanescentes.
  A única ambiguidade encontrada foi corrigida: faltar PFMI/Analítico/Base impede
  preparação, mas data/sequência/comissão pendentes não impedem o Excel.
- A suíte da aplicação não foi reexecutada em M0: não houve mudança de implementação.
  Os 48 testes/81% permanecem baseline confirmado na etapa anterior; não são evidência
  da revisão. M1 deve revalidar o estado antes de começar.
- Próximo passo: usar o prompt na nova conversa e iniciar o Goal M1–M6. Nenhum Goal
  ou tarefa de implementação foi iniciado durante M0.
- M1–M6: não iniciados nesta revisão.
- M7/M8: pendentes nos ambientes e com os responsáveis correspondentes.

## 11. Próxima conversa

Usar [prompt do Goal V1.1](docs/PROMPT_GOAL_V1_1.md).
Ele autoriza criar o Goal **na próxima conversa**, revalidar o estado e executar
M1–M6 até candidato Linux comprovado. Este documento não iniciou Goal nem tarefa
paralela de implementação.

Resultado esperado do próximo fechamento: o que funciona, evidências/tests,
correções e limitações históricas, caminhos de reprodução e passos de M7/M8.
A V1 como produto só termina com Windows validado e aceites do operador do fundo
e do responsável pela aprovação final.

### Revalidação do Goal — 2026-09-04

- Instruções fornecidas pelo responsável pela aprovação final, DRD 1.1, Specs 1.1
  e este plano lidos. Um único
  worktree em main, sem commit/HEAD; todos os arquivos anteriores preservados.
- Manifesto local pré-edição: `outputs/v1_1_baseline/manifest.json`, 71 arquivos,
  SHA-256 `0a956a79a549eb4a1a69286a59d9b83770bdfe910bcaea80e896b72a3e97ee87`.
  Cópia dos arquivos editáveis no mesmo diretório; fontes reais apenas lidas/hashadas.
- `UV_CACHE_DIR=/tmp/cnab-nox-uv uv sync --locked --offline --group dev` e
  `UV_CACHE_DIR=/tmp/cnab-nox-uv ./scripts/check.sh`: 48 testes/81%, Ruff, import
  smoke e privacidade verdes. São o baseline antigo, não aprovação da revisão.
- Cache padrão do uv é somente leitura no sandbox; cache temporário resolveu sem
  atualizar dependências. Runtime da venv: Python 3.13.14/Tk 9.0. Python de sistema
  sem Tk; display :0 não acessível no sandbox e Xvfb ausente. Prova gráfica será
  resolvida em M6, ainda não comprovada.
- Especialistas somente leitura: inventário histórico e revisão de invariantes.
  Próximo passo: contratos/modelos e parâmetros de M1; depois leitores de M2.

### M1 comprovado — entradas e parâmetros (2026-09-04)

- `models/normalize/service`: `PreparationRequest` e PFMIs ordenadas com taxa textual
  própria, modalidade, IDs/hash e proveniência; API aceita lista e preserva o caso
  de uma PFMI. Percentual com 28 dígitos/casas e cálculo em contexto Decimal 80.
- `uv run pytest tests/test_parameters.py tests/test_normalize.py`: 37 testes verdes
  no cache local temporário. T01–T03: zero/fração/vírgula/ponto/101%, Normal inativo,
  vazio e extremos; taxa inválida permite Excel e não é liberada por correção final.
- Correção/reteste: neutralização de `-1` no Excel exige apóstrofo visível; texto
  original continua no snapshot JSON. Teste passou verificando ambas as camadas.
- Ruff em `src tests` verde. Limite: prova dos widgets pertence a M6. Próximo M2.

### M2 comprovado — leitura e preservação (2026-09-04)

- `readers`: aliases por conteúdo (inclusive FALÊNCIA / SACADO), blocos/continuações,
  valores inválidos preservados, resumos e tabelas auxiliares demonstradas no corpus,
  projeções legadas sem cabeçalho reconhecidas por forma completa e origem ordenada.
  Conteúdo bruto/formulado das linhas preservado no snapshot; fórmulas não executadas.
- Espelho exige os seis campos operacionais na mesma ordem e partição integral dos
  grupos. Mudança de posição das colunas não cria novo pagamento. Igual total,
  subconjunto, ordem diferente ou fusão de blocos iguais não comprova espelho.
- `uv run pytest tests/test_readers.py`: 12 testes verdes (T05–T07/T10). Continuação,
  resumo correto/vazio/divergente, legado, projeções, linha desconhecida, arquivos
  duplicados e XLS/XLSX/CSV com dois Contato e registros lógicos cobertos.
- Leitura privada das 12 PFMIs históricas: pagamentos/grupos por arquivo efetivo:
  T01 12/12 + 5/5 (cópia 12/12 identificada separadamente); T02 2/2 + 8/8;
  T03 9/9; T04 10/10; T05 8/8; T06 2/2; T07 18/11; T08 17/15 + 4/4.
  Zero linhas desconhecidas remanescentes nesses formatos; os dois espelhos T08
  mantêm uma contagem cada. Nenhuma fonte foi alterada.
- Problemas resolvidos por evidência: assinatura bruta falhava por colunas bancárias
  diferentes; adotada assinatura semântica com partições. Cabeçalhos anotados de
  gerador, resumo com dois campos, rodapé externo e rótulos PFMI foram reconhecidos
  por assinaturas específicas. Releitura de todo corpus e testes sintéticos verdes.
- Limite: isso prova classificação/contagens, não matching nem preenchimento histórico.
  Próximo M3; relatório reproduzível do corpus consolidado em M5.

### M3 implementado e testado — revisão integrada ainda em andamento (2026-09-04)

- Identidade snapshot/aba/linha/referência, seleção integral em duas passagens,
  alternativas individuais NAO e concorrência global; sem subset search. MOGIANO
  aplicado às três fontes. Comissão por crédito e campos originais separados.
- Conexcred final sem composição e emissão nova vazia; advogado fora de parênteses,
  sem reconstruir documentos nem usar beneficiário como cedente Normal.
- `uv run pytest tests/test_parameters.py tests/test_readers.py tests/test_matching.py
  tests/test_commission.py tests/test_normalize.py`: 71 testes verdes; Ruff verde.
  T04/T08–T15 exercitados com expectativas sintéticas independentes e cinco cálculos
  literais de 28/08. Revisão final do conjunto segue em M6.
- Próximo: completar testes de edição/integridade/reconciliação de M4, prova posicional
  e relatório histórico M5, depois comandos/widgets e visual M6.


### M3/M4 — revisão, correções e retestes integrados (2026-09-04)

- Esquema CNAB-NOX-V1-2 com cinco abas, IDs e originais protegidos por manifestos,
  fontes/parametrização preservadas, CNPJ Conexcred com estado e aba/linha de origem
  visível, totais por PFMI e lote. Campos finais continuam editáveis; emissão nova
  exige correção explícita. Gerar lê exclusivamente o Excel salvo.
- Validação bloqueia controles/linhas/fórmulas alterados, esquema antigo, fonte sem
  crédito, reutilização, documento inválido e centavos divergentes por grupo/lote.
  APROVADO confirma só nome; correções financeiras válidas prevalecem sem recalcular E.
- Revisão independente de Turing corrigida: Cessão exige CNPJ e nome técnico exato
  da principal; agregados PFMI não usam teto de título; montantes acima da precisão
  de 15 dígitos do Excel ficam como texto decimal exato, inclusive centavos.
  Conflito de CNPJ distingue-se de ausência e referências ficam visíveis por arquivo.
- Achado inicial sobre rejeitar todo sobrenome novo de advogado foi retirado após
  cotejar DRD/Specs: o operador do fundo pode corrigir nome completo com APROVADO. A regra preserva
  a identidade de advogado indicada na fonte e rejeita composição comprovada com
  credor; não cria catálogo de nomes nem infere pessoa dos parênteses.
- `uv run pytest tests/test_review_regressions.py tests/test_matching.py
  tests/test_revision_e2e.py -q`: verde após corrigir duas expectativas da nova fixture
  (a PFMI sintética tem dois pagamentos, não um). Depois de sugerir a sequência no
  Excel, `uv run pytest tests/test_commission.py tests/test_revision_e2e.py
  tests/test_review_regressions.py -q` também verde. Ruff em src/tests/scripts verde.
- Provas adicionadas: grupo de 120 bilhões distribuído em dois títulos válidos,
  round trip de 100000000000000.01 preservado, CNPJ de outro grupo/isolamento entre
  arquivos, fórmula inclusive linha NAO e abas protegidas, assinatura Normal,
  totais por PFMI e soma dos cinco cálculos (35.541,94 de comissão).
- Primeiro check integrado: 200 passaram, dois erros de fixture na integração dos
  testes GUI adicionais. Corrigido registro explícito da fixture; reteste final em
  andamento. Um lint de assinatura longa foi corrigido antes de repetir a suíte.
- Limites externos não mudaram. Próximo: consolidar M5/M6 com artefatos pós-correções,
  inspeção visual, suíte completa sem skips gráficos e conferência das fontes.


### M5 comprovado — renderização e histórico (2026-09-04)

- Leiaute 444 preservado; nomes técnicos separados da comparação de identidade,
  CPF/CNPJ/padding, faixas máximas e um centavo acima, header/trailer, sequência,
  CP1252/CRLF final/sem BOM. Golden novo independente em `tests/golden_expected.py`
  e `tests/fixtures/golden_cnab_v1_2.txt`; referência antiga preservada.
- `tests/test_cnab.py`, `tests/test_e2e.py`, `tests/test_revision_e2e.py` e regressões
  de revisão passaram na suíte final abaixo. Gravação atômica e não sobrescrita
  verificadas, inclusive falha injetada, pasta sem permissão e log indisponível.
- Comando: `UV_CACHE_DIR=/tmp/cnab-nox-uv uv run python scripts/verify_historical.py
  --output outputs/v1_1_historico_final` — exit 0. Relatório sanitizado:
  [relatorio_sanitizado.json](outputs/v1_1_historico_final/relatorio_sanitizado.json).
- Renderizador histórico: oito lotes, 86 títulos, aquisição total 153.983.934 centavos;
  zero bytes divergentes da expectativa revisada. Há 35 substituições exatas de
  intervalos de nome; valores antigos/novos e hashes ficam somente no arquivo
  privado `outputs/v1_1_historico_final/expectativas_privadas.json`. Não há máscara
  genérica de nomes. As expectativas de advogados usam referência limpa do mesmo
  documento no lote 05 apenas como prova do renderizador, nunca como fonte de matching.
- Leitura das 12 PFMIs: contagens de M2 confirmadas, zero linhas desconhecidas,
  resumo divergente do lote 07 sinalizado e dois espelhos integrais no lote 08.
- Cruzamento real do lote 01, com 12 Normal + cinco Cessão/15% conforme contrato:
  snapshot da pasta tem 4.341 créditos, 18 candidatos e cinco grupos pré-selecionados
  de 17. Snapshot atual da raiz também gera 18 candidatos, mas só dois grupos
  pré-selecionados. Pendências de identidade, nomes, valores, Base e emissão ficam
  visíveis; nenhuma correção manual nem TXT de importação real foi produzido.
- Hash da cópia operacional efetiva e da omitida registrado; igualdade operacional
  integral verificada antes de excluir a cópia adicional apenas deste roteiro.
  A aplicação continua preservando arquivos duplicados quando selecionados.
- Modalidades dos demais lotes permanecem indeterminadas na prova de leitura;
  nome do arquivo não decide modalidade. Não há prova de snapshots contemporâneos
  nem alegação de E2E histórico integral. XLSB não substitui o Analítico.
- Problemas/correções: expectativas antigas divergiam em nomes; intervalos exatos
  individuais e formato técnico aprovado corrigiram a comparação, sem ignorar
  diferenças financeiras ou outros bytes. Script foi reexecutado após a revisão do
  core. Limite histórico permanece explícito. Próximo M6.

### M6 comprovado — interface, candidato e verificações finais (2026-09-04)

- GUI Tk/ttk: lista ordenada, adicionar/remover/subir/descer, modalidade e taxa
  por item, seletores e dois botões. Normal mantém taxa inativa; alternar modalidade
  não perde zero ou vazio. Resultado com rolagem mostra erros e resumo completo.
- Seis testes GUI executados na sessão gráfica Linux real, incluindo entrypoint
  `python -m gerador_cnab_nox` via `runpy`, construção real da aplicação e event loop;
  widgets/handlers/serviços reais, diálogos com respostas sintéticas controladas.
  Normal, Cessão isolada e lote misto, cancelamentos, emissão bloqueada/corrigida,
  fontes inacessíveis e parâmetros da interface alterados cobertos.
- O teste de entrypoint inicialmente manteve o loop Tk porque outra raiz da fixture
  seguia viva. O processo de teste próprio foi encerrado; teste corrigido com
  `quit` explícito e destruição posterior. Os seis testes gráficos passaram isolados,
  depois a suíte completa. Não houve alteração operacional para contornar a falha.
- Comando final em sessão gráfica disponível:
  `UV_CACHE_DIR=/tmp/cnab-nox-uv ./scripts/check.sh` — **exit 0, 204 testes aprovados,
  nenhum skip, 92% de cobertura (branch coverage ligada), Ruff/import smoke/privacidade OK**.
  Duração da suíte: 51,36 s. [Log final](outputs/v1_1_final/check.log).
- A rodada anterior tinha 203 testes aprovados; o teste POSIX de permissão foi separado
  da gravação atômica portátil para o roteiro M7. Smoke PowerShell agora exige Tk e
  interrompe em erro de comando. Sua execução Windows continua não comprovada.
- Cobertura por superfície: matching 98%, workbook 94%, GUI 92%, service 92%,
  validation 90%, normalize 89%, CNAB 87%, readers 86%. Lacunas são combinações
  defensivas, erros de abertura/parser, variantes de auxiliares e exceções de GUI;
  corpus histórico exercita caminhos adicionais fora do pytest. T01–T24 foram
  reconciliados com casos executados; percentual não substitui a matriz.
- Prova visual pós-core: `uv run python scripts/visual_smoke.py --output
  outputs/v1_1_visual_final --capture-python <python-com-Pillow>` — exit 0.
  O Python de captura usado foi um interpretador local com Pillow instalado
  (caminho específico da máquina de homologação, omitido nesta cópia).
  Captura X11 restrita à janela do teste, sem captura do desktop inteiro.
- [Manifesto sintético](outputs/v1_1_visual_final/evidencia_sintetica.json): três PFMIs,
  três pagamentos/grupos/candidatos/títulos; nominal 53.000 centavos, aquisição
  27.050 centavos, sequência 500–502. Duas datas novas foram preenchidas explicitamente
  no Excel do cenário. Fontes sintéticas foram removidas e parâmetros visuais mudados;
  resultado continuou exclusivamente determinado pelo Excel.
- Quatro PNGs gerados. Inspeção visual direta de lista, bloqueio e sucesso confirmou
  os dois botões visíveis, modalidades/taxas legíveis e mensagens navegáveis.
  [Lista](outputs/v1_1_visual_final/01_lista_mista.png),
  [bloqueio](outputs/v1_1_visual_final/03_emissao_bloqueada.png),
  [sucesso](outputs/v1_1_visual_final/04_txt_gerado.png).
- Problemas visuais corrigidos e reexercitados: caracteres de três bytes mal renderizados
  pelo Tk local, status cortado e cabeçalhos/legenda do Excel. Tentativas iniciais de
  ImageGrab do root XWayland falharam; após diagnóstico anti-loop, captura da própria
  janela por XGetImage funcionou. Artefatos finais substituem a autoridade dos ensaios.
- QA das cinco abas por importação/renderização Artifact: seis recortes PNG em
  `outputs/v1_1_excel_qa_final`, zero erros de fórmula. Originais/manifesto ocultos,
  filtros/colunas, referências e finais inspecionados. O preview Artifact mostra
  CNPJs em notação científica apesar de strings integrais no próprio importador e
  no XLSX com formato texto; essa limitação do preview foi registrada em
  [verificacao_tipos.json](outputs/v1_1_excel_qa_final/verificacao_tipos.json).
  Nenhuma célula foi reexportada por esse renderizador. Editor nativo é prova de M7.
- Preservação: 37 arquivos de referência/entradas do manifesto inicial conferidos
  byte a byte por SHA-256; zero alterados/ausentes. [Evidência](outputs/v1_1_final/preservacao_fontes.json).
  `main` continua sem commit/HEAD e com um worktree; arquivos preexistentes não
  rastreados preservados. Sem commit/push, release, deploy, importação ou publicação.
- Revisões: Turing verificou invariantes e correções; Galileo reconciliou cobertura
  da matriz em leitura; Gauss audita o fechamento. Nenhuma regra de produto nova.
- Próximo: somente M7/M8 conforme roteiro; Goal Linux não depende de presumir essas provas.

### Matriz executada T01–T24

Todos os arquivos abaixo estão em `tests/`, salvo indicação de script. Os testes
foram executados na suíte final de 204 casos; IDs repetidos indicam provas complementares.

| ID | Prova executada | Resultado |
|---|---|---|
| T01 | `test_parameters.py`; `test_gui.py::test_t01_real_widgets_keep_order_modalities_and_individual_rates` | Lista/ordem/parametrização própria preservadas |
| T02 | `test_parameters.py`; `test_commission.py`; widgets Normal/0/15/0,5 | Vazio não restaura 15; 101 aceito; taxa Normal ignorada |
| T03 | `test_parameters.py`; `test_readers.py::test_extreme_operational_values_become_pending_instead_of_aborting` | Excel com pendência; TXT bloqueado até nova preparação |
| T04 | `test_commission.py` (cinco cálculos, total 35.541,94, meio centavo, base negativa) | ROUND_HALF_UP por crédito; sem rateio |
| T05 | `test_readers.py` (blocos, continuação, aliases, partições) | Ordem/blocos e linhas legítimas preservados |
| T06 | `test_readers.py` (resumos/auxiliares/projeções/desconhecidas); histórico | Auxiliares com alerta; desconhecida bloqueante |
| T07 | `test_readers.py` (espelho semântico integral e duplicação entre arquivos) | Parcial/total igual/ordem distinta não descartam créditos |
| T08 | `test_matching.py::test_t08_four_payments_one_credit_and_no_subset_search` | Pagamentos não multiplicam títulos; conjunto integral |
| T09 | `test_matching.py` (homônimos/ref repetidas); `test_revision_e2e.py` (reuso manual) | Crédito distinto preservado e não reutilizado |
| T10 | `test_readers.py` (dois Contato, XLS/XLSX/CSV multiline, erros operacionais) | Identidade principal/registro lógico preservados |
| T11 | `test_matching.py` MOGIANO/Base; `test_revision_e2e.py` conflito corrigível | Sem fuzzy; finais Base corrigíveis |
| T12 | `test_matching.py` alternativa; `test_revision_e2e.py` placeholder | Alternativa NAO; ausência de origem não vira título |
| T13 | `test_matching.py`, `test_commission.py`, `test_review_regressions.py` | Conex CNPJ/nome exato/origem por PFMI; Normal sem substituição indevida |
| T14 | `test_matching.py` advogado/CPF parcial; `test_review_regressions.py` correção/identidade | Sem identidade dos parênteses; composição/troca bloqueadas |
| T15 | `test_review_regressions.py` Normal; `test_revision_e2e.py` Cessão/data; GUI extra | Assinatura Normal; emissão nova manual; data futura bloqueia |
| T16 | `test_revision_e2e.py` finais prevalecem/originais intactos | Correções válidas geram alerta, sem recomputar E |
| T17 | `test_revision_e2e.py` um centavo/opostos; `test_validation.py` | Reconciliação por grupo/lote sem compensação |
| T18 | `test_revision_e2e.py`, `test_review_regressions.py`, `test_workbook.py` | IDs/digests/fórmulas/esquema/linhas/centavo grande verificados |
| T19 | `test_revision_e2e.py`, `test_gui.py`, `test_gui_extra.py`, visual final | Fontes removidas e GUI mudada não afetam geração |
| T20 | `test_cnab.py`, `test_revision_e2e.py`, `test_normalize.py` | Nomes 40/41, acentos, pontuação e equivalências separados |
| T21 | `test_cnab.py`, `test_review_regressions.py`, `test_revision_e2e.py` | Golden independente, limites, 444/CRLF/CP1252, documentos e sequência |
| T22 | `test_cnab.py`, `test_revision_e2e.py`, `test_gui.py`, `test_gui_extra.py` | Atômico/não sobrescreve; permissão/cancelamento/log tratados |
| T23 | Seis testes em `test_gui.py`/`test_gui_extra.py`; visual final | Comandos/widgets reais e fluxo visual comprovados |
| T24 | Golden/`verify_historical.py`; E2E sintético; dois snapshots T01 | Três provas separadas e limites históricos registrados |

### Reprodução e entrega posterior

1. No Linux, com Tk/display, executar `uv sync --locked --group dev` e
   `./scripts/check.sh`. Cache alternativo `UV_CACHE_DIR=/tmp/cnab-nox-uv` é opção
   deste sandbox. Sem display, Tk obrigatório falha; usar sessão gráfica válida.
2. Abrir o candidato com `uv run python -m gerador_cnab_nox`.
3. Repetir o visual sintético usando novo diretório inexistente em `--output`; o
   script só remove as próprias fontes sintéticas. Captura PNG é opcional e requer
   Python com Pillow/X11 no Linux. Não passar diretório de dados reais.
4. Repetir histórico com `scripts/verify_historical.py --output outputs/<novo-diretorio>`;
   mantém expectativas nominais privadas e relatório sanitizado separados.
5. Identidade final de código/documentos/testes/scripts e evidências locais em
   `outputs/v1_1_final/manifesto_candidato.json`; sem Git HEAD, hashes são a prova
   de integridade, não um suposto diff/commit.
6. M7: [roteiro Windows](docs/HOMOLOGACAO_WINDOWS.md), ambiente/build onedir na
   máquina do responsável técnico, fluxo no executável sem dependências de
   desenvolvimento, editor nativo do Excel, caminhos/permissões/hard links/antivírus.
   Não comprovado no Linux.
7. M8: confirmar parâmetros reais, revisar/reconciliar, importar manualmente no
   Frontis sob autoridade do responsável pela aprovação final, registrar aceite
   operacional do operador do fundo e aceite final do responsável pela aprovação
   final separadamente. Nenhuma dessas ações foi executada neste Goal.


### Encerramento do Goal local (2026-09-04)

Gauss reconciliou objetivo, DoD, matriz T01–T24, correções de Turing, documentos e
SHA-256: **M1–M6 aceitos como candidato Linux comprovado; nenhum bloqueio material**.
A auditoria confirmou os 47 arquivos do candidato e 22 evidências então existentes.
Seu veredito sanitizado foi acrescentado em `outputs/v1_1_final/auditoria_gauss.json`,
e o manifesto foi atualizado apenas para esse registro e este encerramento documental.

Estado final: 204 testes sem skips, 92% de cobertura, Ruff/import/privacidade verdes;
provas visual, sintética e histórica separadas; 37 referências preservadas.
O código técnico permaneceu idêntico ao auditado. Goal local concluído; M7/M8 e
V1 como produto continuam pendentes dos ambientes e aceites externos descritos.

### Testes reais locais — execução autorizada em 2026-09-04

O responsável pela aprovação final aprovou executar 20 preparações dos lotes
01, 06, 07 e 08; comparar dois Analíticos, isolar o efeito das Bases no lote 01,
testar modalidades desconhecidas como hipóteses explícitas, revisar cópias com
intervenções rastreadas e gerar somente quando válido. Bugs sistêmicos podem ser
corrigidos e retestados; mudança de regra exige decisão do responsável pela
aprovação final. Fontes reais permanecem intactas e privadas; sem Frontis.

- Baseline revalidado: 47 arquivos do candidato e 23 evidências idênticos ao manifesto.
- Pacote privado: `outputs/testes_locais_20260904_232828`; manifesto inicial cobre
  40 fontes/referências/evidências, sem modificar os originais.
- DoD atendida: 20 cenários classificados, correções rastreadas, diferenças por título/campo/
  bytes explicadas ou explicitamente inconclusivas, simulações e GUI verificadas,
  fontes preservadas, relatório sanitizado e procedimento reproduzível.

#### Resultado e evidência de encerramento (2026-09-05)

- Rodada de referência: `outputs/testes_locais_20260904_232828/rodada_final`.
  Quatro cenários T01 (Normal/Cessão 15% × dois Analíticos × duas Bases), quatro
  T06, quatro T07 (Normal/Cessão × dois Analíticos) e oito T08 (quatro combinações
  de modalidade × dois Analíticos). Modalidades fora de T01 permanecem hipóteses;
  datas dos snapshots não foram presumidas.
- 20 preparações reais via `PreparationRequest`/`PfmiInput` → `prepare_workbook`,
  com cinco abas verificadas. Preservados 20 Excel originais e 20 cópias revisadas.
  `generate_cnab` recebeu somente cada Excel e foi exercitado antes e depois da
  revisão: **40 tentativas bloqueadas, zero TXT real**. Isso não estabelece
  igualdade ou divergência de bytes E2E para nenhum lote.
- 48 vínculos de título aceitos ao longo dos cenários (não títulos únicos),
  1.008 campos comparados, 198 diferenças iniciais e 112 alterações de célula
  rastreadas. Após revisão, restam 96 diferenças de sequência prévia, derivada
  somente na geração, e zero diferenças nos demais campos vinculados.
  Nomes completos e técnicos possuem registros separados; o trecho além dos
  40 caracteres não pode ser reconstruído a partir do TXT histórico.
- Cópia das fontes conferida independentemente, com zero divergências;
  122 cálculos individuais conferidos, sem divergência. Origem, seleção e
  vínculo histórico são contagens distintas. Documentos finais de principal/
  advogado não identificam por si só o crédito original. Títulos sem vínculo
  são inconclusivos, e não automaticamente ausentes/extras.
- Bloqueios: T01 tem oito grupos sem candidato de origem; T06 tem um, com
  documento/nominal correspondentes porém campanha diferente no Analítico;
  T07 tem nove ou dez; T08 tem nove a onze. Há diferenças concretas de Base,
  emissão, nomes e valores (inclusive um centavo em T08). A causa temporal
  permanece indeterminada quando não demonstrada por comparação controlada.
- Comparador independente em `scripts/local_audit_compare.py`: bytes integrais,
  registros, campos e intervalos, 444/CP1252/CRLF final/sem BOM, nomes sem máscara
  genérica; identidade de títulos exige pares explícitos e únicos. Os oito
  históricos também foram comparados aos oráculos revisados por intervalos
  exatos, como prova separada da geração E2E.
- `sinteticos_01`: 11 cenários com expectativas independentes; nove bloqueios
  esperados e dois TXT válidos. Cobertos centavo, compensação entre grupos,
  comissão por crédito, alternativas, reutilização, espelho parcial, documento,
  emissão ausente/futura, controles, pontuação e nomes longos. Nenhuma fonte real
  foi corrigida artificialmente para forçar um resultado.
- `gui_real_final`: dois botões Tk reais em T01 com dois arquivos; preparação e
  duas tentativas bloqueadas. Diálogos de arquivo/mensagem controlados.
  `visual_sintetico`: quatro capturas locais, incluindo bloqueio de emissão e
  TXT válido; estados de bloqueio e geração inspecionados visualmente.
- `check_encerramento.log`: **229 testes aprovados, sem skips, cobertura 92%,
  Ruff/import/privacidade OK**. Inclui divergências deliberadas do comparador
  (byte, centavo, nome, ordem, registro ausente/extra) e regressões dos vínculos
  e da distinção entre vínculo aceito e seleção atual.
- Tentativas anteriores preservadas: `rodada_01` interrompida por conversão de
  datas do auditor; `rodada_02` superada após restringir vínculos de Cessão;
  `rodada_03` antecede ajustes de evidência. Rascunhos de relatório também foram
  preservados; usar somente `relatorio_encerramento/RELATORIO.md` como conclusão.
- Revisão independente confirmou o escopo e os limites. Foram corrigidos erros
  do auditor; **nenhum bug do produto confirmado, nenhuma regra alterada**.
  Fontes reais, 23 evidências anteriores e código `src` preservados por SHA-256.
  README/ExecPlan e os novos scripts/testes são a superfície alterada nesta rodada.

Relatório sanitizado: [RELATORIO.md](outputs/testes_locais_20260904_232828/relatorio_encerramento/RELATORIO.md).
Preservação e identidade final: `outputs/testes_locais_20260904_232828/encerramento_sanitizado.json`
e `manifesto_encerramento_privado.json`. Detalhes identificáveis permanecem somente
nos arquivos privados do pacote, protegido localmente e ignorado pelo Git.
Reprodução: seção “Reproduzir a auditoria local” do README, usando destinos novos.

Conclusão desta rodada não exige forçar TXT. A reprodução integral depende de
reconciliar origens/campanhas/identidades pendentes e confirmar os parâmetros
históricos. M7/Windows, M8/Frontis e aceites operacionais continuam posteriores.

### Interface azul em três etapas — 2026-09-05

O responsável pela aprovação final aprovou implementar o planejamento: Tk/ttk sem novas dependências, uma etapa
por vez (configurar lote → preparar/revisar Excel → gerar TXT), orientação em
português, progresso responsivo e abertura local de arquivo/pasta. Não alterar
serviços, cálculos, validações, esquema, rastreabilidade ou leiaute; somente dados
sintéticos na validação. Um escritor: Steve.

- Baseline revalidado: main sem HEAD, arquivos preexistentes não rastreados;
  49 arquivos técnicos copiados/hashados em `outputs/interface_azul_20260905/baseline`.
- Ambiente gráfico acessível fora do sandbox: Python 3.13.14, Tk 9.0.3, tema clam
  disponível. Baseline atual: 229 testes aprovados, 92% de cobertura, sem skips
  (`outputs/interface_azul_20260905/baseline_check.log`).
- Entregas: estrutura/tema; orientação e trabalhador único; testes gráficos,
  inspeção visual e documentação diretamente afetada.
- DoD: fluxo sintético completo e entrada direta com Excel; parâmetros individuais
  e ordem preservados; bytes independentes das fontes/configuração atual; estados
  de erro/aviso/cancelamento/ocupado verificados; teclado, escalas e tamanhos
  exercitados; suíte/lint/privacidade; núcleo preservado por hashes.
- Estado: funcionalidade implementada, refinamento visual em andamento. Nenhum processamento real nem pacote Windows
  integra esta execução.

#### Validação funcional do primeiro candidato e refinamento

- [x] Três painéis exclusivos, cabeçalho e navegação fixos, tema azul `clam`,
  fontes de 11 pontos, rolagem central por teclado/mouse e caminhos copiáveis.
  PFMIs mostram nome, ordem, modalidade e comissão individual.
- [x] Preparação e geração em trabalhador único, com captura de entradas antes
  de iniciar e atualização de widgets exclusivamente na thread principal.
  Progresso indeterminado, bloqueio de dupla operação/edição/fechamento e
  recuperação dos controles depois de sucesso ou erro.
- [x] Roteiro de revisão, abertura local de Excel/pasta e resultado por etapa.
  Todos os apontamentos permanecem acessíveis, com orientações e detalhes
  técnicos originais. Alterar a configuração não altera o arquivo anterior.
- [x] Fluxo sintético real pelas três etapas: três PFMIs com Normal, Cessão 0%
  e Cessão 0,5%; bloqueio por emissão, correção/salvamento do Excel e geração
  após remover as fontes sintéticas e modificar parâmetros da tela.
  Resultado: três títulos, nominal R$ 530,00, aquisição R$ 270,50, sequência 500–502.
- [x] Suíte final: **262 testes aprovados, zero skips, 92% de cobertura**
  (`check_encerramento.log`). São 33 casos adicionais; 20 casos da suíte exercitam
  Tk real. Cobertos entrada direta com Excel, mesmos bytes sem fontes/configuração,
  taxas individuais, ordem, cancelamentos, erros, auditoria após saída salva,
  todos os apontamentos, responsividade e recuperação de estados.
- [x] Áreas de tela 1366×768 e 1920×1080 simuladas em janelas Tk reais, escalas
  100%/150%, mínimo 900×600, 30 caminhos longos/arquivos homônimos, Tab e roda.
  Display físico desta máquina: 1680×1050; não se trata de dois monitores físicos
  homologados. As capturas do primeiro candidato estão em `visual_final_100` e
  `visual_final_150`; foram superadas visualmente após o feedback do responsável pela aprovação final.
- [x] Ruff, importação/smoke e privacidade aprovados. 13 arquivos protegidos
  (núcleo original fora da GUI e contratos de dependências) idênticos ao baseline
  por SHA-256. Nenhuma instalação necessária.
- [x] Revisão independente de Turing concluída sem achados remanescentes;
  corrigida a orientação de comissão para indicar retorno à configuração e
  novo Excel. README, DRD e Specs reconciliados somente no fluxo de interface.

O Tk desta máquina não expõe Noto Sans/DejaVu Sans entre suas famílias; foi usado
o fallback previsto de TkDefaultFont (Nimbus Sans L). Nenhuma fonte foi instalada.
O caminho de abertura Linux e suas falhas foram testados com processo controlado;
a associação específica do editor depende do sistema. O ramo Windows teve teste
unitário, sem homologação Windows.

Relatório: [RELATORIO.md](outputs/interface_azul_20260905/RELATORIO.md).
Preservação e manifesto: `outputs/interface_azul_20260905/preservacao_nucleo.json`
e `manifesto_composicao.json`. O pacote local contém exclusivamente evidências
sintéticas desta melhoria e é ignorado pelo Git. Nenhum commit foi criado: o
checkout preexistente permanece em `main`, sem HEAD e com arquivos não rastreados.
Lote real completo até TXT, distribuição Windows e aceites operacionais continuam
sendo trabalhos separados.

#### Correção de direção visual — 2026-09-05 (histórico anterior à retomada)

O responsável pela aprovação final considerou a primeira composição antiquada,
quadrada e ainda distante da qualidade visual esperada. O encerramento visual foi
reaberto; a suíte verde não substitui esse critério de produto.

- Implementados cartões arredondados, indicador de etapas com círculos numerados,
  tabela sem grade pesada/rolagem horizontal, seleção azul suave, botões e campos
  arredondados, uma ação principal por etapa e métricas em blocos separados.
- Controles continuam sendo ttk, com semântica nativa de teclado e estado.
  As imagens decorativas são locais; fundos são gerados com biblioteca padrão.
  Cartões usam quatro cantos pequenos e preenchimentos sólidos, evitando o custo
  de repetir imagens por toda a área. Nenhuma biblioteca de interface acrescentada.
- A revisão de design especializada confirmou melhora concreta de composição e apontou fonte,
  separação das métricas e repintura das capturas como verificações finais.
- Dois fluxos sintéticos em `refinamento_100`/`refinamento_150` concluíram até TXT;
  os 39 casos de interface passaram em `refinamento_visual/gui.log`. Métricas e
  captura receberam depois ajustes pequenos, sujeitos à conferência final.
- Confirmado no runtime: Tk 9.0.3, `tk::pkgconfig get fontsystem = x11`, sem Xft;
  Noto Sans não é exposta pelo Tk mesmo estando disponível no sistema. Isso deixa
  a fonte serrilhada. O fallback previsto não resolveu o critério visual do responsável pela aprovação final.
- Solicitada autorização específica para instalar `python3-tk` padrão do Ubuntu
  e usar Python/Tk do sistema com suavização, pois o plano aprovado proibia novas
  dependências. **Autorização pendente; pacote não instalado, ambiente preservado.**
- Captura anterior `visual_encerramento_100` não é evidência de sucesso: o ensaio
  apresentou seis PFMIs em vez das três esperadas e bloqueou TXT. O roteiro agora
  protege suas janelas contra cliques/teclas externos e verifica a quantidade.
  As execuções posteriores mantiveram três PFMIs e geraram o TXT esperado.

Estado: composição refinada e funcional; conclusão visual depende da tipografia.
A validação atual está registrada abaixo. Não declarar a melhoria encerrada
enquanto esse acabamento permanecer pendente.

As capturas intermediárias de composição são `composicao_100` e `composicao_150`, com
sete imagens e fluxo sintético completo em cada escala. Cabeçalhos conferidos
por pixels em `evidencia_capturas.json`; métricas agora têm blocos próprios e o
contorno da etapa aparece somente quando o controle recebe foco. Verificações
de Ruff, importação e privacidade aprovadas; 13 arquivos protegidos preservados.

O candidato passou a manter **Abrir Excel para revisar** e **Abrir pasta** como
ações principais fixas nos respectivos estados concluídos. O teste ampliado
detectou rodapé sem espaço em 900×600/150%: o Canvas solicitava 397 pixels primeiro.
Corrigida a prioridade do layout para reservar as ações antes do conteúdo rolável.
As capturas seguintes são `candidato_100`/`candidato_150`, com fluxo completo,
cabeçalho íntegro e ações dentro da janela. O TXT é byte a byte idêntico nas duas
escalas e ao ensaio funcional anterior. Testes agora aguardam o estado efetivo
de mapeamento/foco e verificam explicitamente essas ações.
O runbook Windows inclui a coleta dos ícones locais pelo PyInstaller; não houve
build, instalação de dependências Windows ou homologação Windows.

As capturas dessa composição ficaram em `layout_final_100` (1040×668/100%) e
`layout_minimo_150` (900×600/150%), ambas após a correção do rodapé. Cada ensaio
concluiu o fluxo sintético com sete capturas. `evidencia_capturas.json` confirma
cabeçalhos íntegros, ações principais dentro dos limites nos dois eixos e o mesmo
TXT nas cinco execuções comparadas. Núcleo e dependências: 13 arquivos preservados;
10 arquivos técnicos existentes alterados e 16 novos, incluindo 12 ícones locais.

A suíte ampliada identificou interferência de foco no teste de layout, que criava
dois intérpretes Tk simultâneos. Turing revisou essa causa e o teste passou a
manter apenas um Tk por cenário, conservando as verificações de geometria,
travessia por teclado e rolagem. A execução fica em `check_isolamento_final.log`.

Validação desta composição: **262 testes aprovados em 239,48 segundos, sem skips,
93% de cobertura**, incluindo os quatro cenários gráficos corrigidos e os demais
casos de interface. Ruff, smoke de importação e privacidade aprovados. O Goal
está bloqueado aguardando a autorização para resolver a tipografia e validar o
runtime final; nenhuma instalação foi executada. A mesma pendência persistiu por
três turnos consecutivos e foi reconfirmada antes de atualizar o estado do Goal.

Na continuação, foi descartada por prova a alternativa de usar o runtime já
disponível nas ferramentas locais: Python 3.12.14/Tk 9.0.4 também usa `x11`
sem as famílias Noto Sans/DejaVu Sans. O pacote Ubuntu continua ausente. Evidência:
`runtime_existente_verificado.json`. A mesma autorização permanece pendente.

A auditoria final de Gauss confirmou cobertura dos requisitos de interface/fluxo,
65 hashes atuais e preservação dos 13 arquivos protegidos. A lacuna de guardar
as saídas dos checks auxiliares foi resolvida: `ruff_final.log`, `import_final.log`
e `privacy_final.log` registram sucesso. Resta somente o acabamento tipográfico.

#### Retomada autorizada e ambiente final — 2026-09-05

A autorização posterior do responsável pela aprovação final resolveu o bloqueio histórico descrito acima.
Instalados os pacotes Ubuntu `python3-tk`, `python3.14-tk`, `libtk8.6` e `libxss1`,
sem atualizar/remover outros pacotes. A `.venv` anterior foi preservada no pacote
de evidências; a nova usa Python 3.14.4, Tk 8.6.17/Xft e Noto Sans. As versões
Python do lock permaneceram iguais; não foi adicionada biblioteca de interface.
O helper temporário de autenticação gráfica foi removido.

A inspeção real confirmou fontes suaves. O ensaio ampliado configura Fontconfig
em 144 DPI apenas no processo de QA e registra altura de linha de 31 pixels,
contra 21 em 100%. Quatro cenários de layout, foco e rolagem passaram em
`layout_final_fontes_150.log`, incluindo 900×600 e 30 PFMIs com caminhos longos.
A requisição de tamanho dos controles foi preservada ao ampliar os trechos
decorativos repetidos pelo ttk, reduzindo o custo de pintura. Revisões técnicas
e visuais independentes não apontaram problema material remanescente.

A validação completa usa a configuração normal do desktop; o ensaio de 150%
permanece separado. A execução da suíte inteira com fontes forçadas em 150%
atingiu o limite do comando sem reportar falhas e não conta como check concluído.


#### Conclusão da melhoria de interface — 2026-09-05

- [x] Suíte final no ambiente preparado: **262 testes aprovados em 415,80 segundos,
  zero skips, 93% de cobertura**, Ruff, importação e privacidade aprovados
  (`check_final_desktop.log`). Inclui 39 casos de GUI, 20 com widgets Tk reais.
- [x] Quatro cenários ampliados de tamanho, teclado e rolagem aprovados
  (`layout_final_fontes_150.log`); fontes a 144 DPI limitadas ao processo de QA.
- [x] Dois fluxos sintéticos finais completos em `interface_concluida_100` e
  `interface_concluida_150`: bloqueio por emissão, correção e salvamento do Excel,
  remoção das fontes sintéticas, mudança da configuração e TXT gerado somente
  pelo Excel. Três títulos, R$ 530,00 nominal e R$ 270,50 aquisição; mesmos bytes.
- [x] Capturas finais conferidas nas três etapas, com tipografia suave,
  ações fixas, rolagem e orientação da revisão preservadas.
- [x] Núcleo e contratos de dependências preservados por SHA-256; documentação
  diretamente afetada reconciliada. Não houve processamento de dados reais.

Na conferência de 150%, a confirmação final ficava abaixo da dobra. Corrigida
a rolagem automática para revelar o título ao concluir sucesso ou bloqueio, sem
mudar o foco. Após a suíte completa acima, esse delta passou em três testes
focados, Ruff, importação, privacidade e nos dois fluxos sintéticos finais.
O título visível passou a ser uma asserção permanente do roteiro de capturas.

A melhoria está concluída. Evidências e comandos reproduzíveis permanecem no
[relatório local](outputs/interface_azul_20260905/RELATORIO.md). O fechamento desta
interface não homologa Windows, lote real completo ou aceites operacionais.
