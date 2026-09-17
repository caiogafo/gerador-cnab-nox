# Gerador de CNAB NOX

Novo: [cadastro local de equivalências de falência](docs/EQUIVALENCIAS_FALENCIAS.md)
na configuração do lote. Regras confirmadas são reutilizadas automaticamente
na preparação e preservadas no Excel revisado.

Diagnóstico offline do cruzamento 80/20: [procedimento e resultados do lote de
10/09/2026](docs/DIAGNOSTICO_8020.md). Essa análise separa sugestões exploratórias
da seleção atual e não gera CNAB nem altera as fontes.

Aplicativo local: entradas operacionais → Excel intermediário para revisão →
TXT CNAB de 444 bytes por registro.

## Estado atual

**Candidato revisado 1.1 comprovado no Linux: M1–M6 concluídos.**
Interface azul em três etapas concluída no Linux em 2026-09-05: cartões e
controles arredondados, fontes suaves e próxima ação sempre acessível. Mantém
Tk/ttk, sem nova biblioteca de interface; o Tk padrão do Ubuntu foi instalado
com autorização. Última suíte: 262 testes aprovados, sem skips, 93% de cobertura,
Ruff e privacidade verdes; 20 casos com widgets Tk reais. O ajuste final de
rolagem foi retestado em três casos focados e nos fluxos sintéticos de 100%/150%.
O programa recebe PFMIs ordenadas, modalidade/comissão por arquivo e gera o novo
intermediário CNAB-NOX-V1-2. A etapa **Gerar TXT** usa somente esse Excel revisado.
Os 48 testes/81% e 229 testes/92% são baselines anteriores; a matriz T01–T24 e o
histórico estão no [ExecPlan](execplan.md). Evidências e capturas da melhoria estão
no [relatório local da interface](outputs/interface_azul_20260905/RELATORIO.md).

A rodada real local dos lotes 01, 06, 07 e 08 executou 20 cenários: todos
prepararam Excel e permaneceram bloqueados para TXT, inclusive após a revisão
assistida. Nenhum lote foi declarado igual ao histórico em bytes E2E. Os detalhes
por cenário, causas e limites estão no
[relatório sanitizado local](outputs/testes_locais_20260904_232828/relatorio_encerramento/RELATORIO.md).
Esse pacote privado é ignorado pelo Git e só existe nesta máquina.

Windows (M7), importação no Frontis e aceites do operador do fundo e do responsável
pela aprovação final (M8) continuam posteriores.
Nenhum executável Windows homologado é declarado por esta entrega.

Documentos de referência:

- [DRD — regras de negócio](drd.md).
- [Specs 1.1 — contratos e testes](specs.md).
- [ExecPlan — marcos e evidências](execplan.md).
- [Prompt histórico que iniciou este Goal](docs/PROMPT_GOAL_V1_1.md).
- [Roteiro Windows e homologação](docs/HOMOLOGACAO_WINDOWS.md).

## Desenvolvimento e execução local

Python 3.12+ e ambiente Tk são necessários no desenvolvimento; o processamento
não exige Excel. A revisão humana do intermediário exige um editor compatível.

No Ubuntu, use o Python/Tk do sistema para renderizar as fontes com suavização
(Xft). O Python gerenciado pelo uv disponível nesta máquina inclui Tk sem Xft.
Prepare o ambiente explicitamente com o interpretador do sistema:

```bash
sudo apt-get install --no-install-recommends python3-tk
uv sync --locked --group dev --python /usr/bin/python3 --no-managed-python
uv run --locked --no-sync python -m gerador_cnab_nox
```

Na máquina validada, esse ambiente usa Python 3.14.4, Tk 8.6.17 e Noto Sans.
As dependências do aplicativo continuam nas mesmas versões de `uv.lock`.
O comando de preparação cria a `.venv`; execução e checks reutilizam esse ambiente.
Para distribuição Windows, siga o roteiro próprio, sem os comandos Ubuntu.

Verificações existentes:

```bash
./scripts/check.sh
```

O check exige sessão gráfica Tk disponível e inclui testes dos comandos/widgets
reais. Sem display ele falha; não tratar skips gráficos como comprovação de M6.
A inspeção visual e a regressão histórica são comandos separados no ExecPlan. Dados operacionais, saídas e logs
permanecem locais/ignorados pelo Git; fixtures versionadas devem ser sintéticas.

## Uso do candidato revisado

1. Em **Configurar lote**, adicione uma ou várias PFMIs na ordem desejada. Escolha **Normal** ou
   **Cessão da Cessão** em cada arquivo. Para Cessão, confira a **Comissão (%)**:
   começa em 15 e pode ser 0 ou fração, com vírgula/ponto. A taxa é individual;
   campo apagado é pendência, não volta sozinho para 15.
2. Selecione um Analítico e uma Base de Vencimentos. Informe liquidação e primeira
   sequência, conferida pelo operador responsável no estoque do fundo.
3. Clique **Continuar** e, na etapa **Preparar e revisar Excel**, clique **Preparar Excel**.
   Escolha um destino novo. Pendências operacionais não impedem Excel. Confira
   separadamente pagamentos, grupos e títulos; resumo auxiliar vazio/divergente
   gera alerta, enquanto linha desconhecida permanece REVISAR.
4. Use **Abrir Excel para revisar**. Revise as linhas na aba CREDITOS.
   Escolha candidatos com INCLUIR_CNAB=SIM/NAO
   e corrija campos finais, inclusive valores, documentos e emissão.
   Na Cessão da Cessão, informe a assinatura da nova cessão: a original é referência.
   APROVADO=SIM serve somente para divergência de nome.
5. Não altere originais, parâmetros, IDs ou quantidade de linhas. Taxa/modalidade,
   erros na origem PFMI ou ausência de crédito exigem corrigir a fonte/parâmetro
   e preparar novamente; dados finais do sacado podem ser corrigidos no Excel.
6. Salve o Excel e use **Continuar para gerar TXT**. Em **Gerar TXT**, escolha somente
   o Excel revisado e clique **Validar e gerar TXT**. O programa revalida sem
   reabrir fontes nem ler parâmetros atuais da interface. Aquisição deve fechar
   em centavos por grupo e lote; crédito não pode ser reutilizado.

**Já tenho um Excel revisado** permite entrar diretamente na terceira etapa.
Se salvou uma cópia com outro nome, selecione essa cópia. Alterar ordem, modalidade,
comissão ou configuração depois de preparar não muda o Excel existente: prepare
outro arquivo para aplicar essas mudanças. A tela mostra esse aviso.

O caminho salvo fica selecionável/copiável. **Abrir pasta** abre o destino pelo
sistema; o Excel abre no editor associado. Se não houver associação, use o caminho
para abrir manualmente em um editor compatível; o programa não instala um editor.
Arquivos de saída existentes não são sobrescritos.

Resultados aparecem na própria etapa, com rolagem automática para mostrar a
confirmação ou o bloqueio ao concluir. **Detalhes técnicos** conserva todos os
apontamentos de validação; a contagem conjunta de revisões/avisos da preparação não
é uma quantidade de erros bloqueantes. Se o TXT foi salvo e apenas o registro de
auditoria falhou, a tela informa ambos sem afirmar que a geração falhou.
Após preparar, **Abrir Excel para revisar** fica em destaque no rodapé; depois de
gerar o TXT, **Abrir pasta** ocupa esse lugar. As ações continuam acessíveis mesmo
quando o conteúdo exige rolagem.

Durante uma operação, a janela permanece responsiva e mostra progresso sem
porcentagem. Alterações e novas operações ficam bloqueadas; aguarde a conclusão
para fechar. Navegue com Tab; o conteúdo central rola para manter o foco visível.

Uma alteração final válida prevalece sobre a sugestão calculada e gera alerta.
Nome completo fica no Excel; saída técnica limitada a 40 caracteres com alerta.
Intermediários antigos são recusados com orientação para preparar novamente.

### Reproduzir a verificação visual sintética

Com Tk/display e dependências de desenvolvimento já disponíveis, use destinos novos:

```bash
uv run --locked --no-sync python scripts/visual_smoke.py --output outputs/visual_azul_100 --window 1040x668
FONTCONFIG_FILE="$PWD/scripts/fontconfig_144.conf" \
  uv run --locked --no-sync python scripts/visual_smoke.py --output outputs/visual_azul_150 --window 900x600 --scale 1.5
```

O roteiro usa serviços e widgets reais com fontes sintéticas; cobre configuração,
revisão, bloqueio e sucesso sem fontes acessíveis. Capturas opcionais requerem
`--capture-python /caminho/python-com-Pillow` já disponível e são restritas à janela
de teste no Linux. Isso não adiciona Pillow ao aplicativo distribuído.
No ensaio Linux com Xft, a configuração de 144 DPI amplia também as fontes em
150%, somente nesse processo. `--scale` sozinho controla a escala lógica do Tk;
os metadados salvos registram a fonte e sua altura efetiva. As preferências de
fontes do desktop permanecem sob controle do usuário.

## Reproduzir a auditoria sintética de cenários bloqueantes

Esta cópia higienizada inclui apenas a trilha 100% sintética de
`scripts/verify_local_scenarios.py`: ela prepara e tenta gerar CNAB a partir de
fixtures fabricadas para comprovar que cada bloqueio esperado (centavo
divergente, documento inválido, emissão ausente/futura, controle alterado,
reuso de crédito, etc.) realmente impede a saída. Os auxiliares que auditavam
lotes reais (`verify_local_e2e.py`, `report_local_audit.py` e o modo
`--real-revised` desta trilha) são exclusivos do ambiente operacional com dados
reais e não fazem parte deste repositório.

Exemplo a partir da raiz, com um destino novo:

```bash
uv run python scripts/verify_local_scenarios.py --output outputs/auditoria_sintetica
```

O comando usa apenas serviços reais do aplicativo sobre fixtures sintéticas,
recusa sobrescrever um destino existente e grava `relatorio_sintetico.json` com
o resultado de cada cenário.

## Windows e operação real

Na etapa posterior M7, na máquina Windows do responsável técnico, seguir o
[roteiro M7/M8](docs/HOMOLOGACAO_WINDOWS.md). Distribuir a pasta onedir completa,
não apenas o executável isolado. Validar o pacote sem Python/Excel/administrador
no ambiente de uso antes de entregar ao operador do fundo.

Importação no Frontis é manual. O operador do fundo deve reconciliar o lote e dar
aceite operacional; o responsável pela aprovação final dá o aceite final da V1.


O resumo contém contagens e totais da preparação por PFMI e por lote. Após editar
finais, salve o Excel; Validar e gerar TXT revalida e mostra o resumo final. Em CREDITOS,
amarelo indica campo editável e azul indica referência protegida. Os detalhes de
proveniência incluem a origem do CNPJ Conexcred. As abas ORIGINAIS_CONTROLE e
MANIFESTO são ocultas/protegidas; os campos necessários à revisão estão visíveis.

A proteção previne alterações acidentais e verifica integridade; não é assinatura
criptográfica nem mecanismo contra adulteração deliberada de todos os controles.
# Tolerância na reconciliação pré-TXT

`MAX_TOLERANCE_DIFF_REAIS` define o limite **absoluto por operação reconciliada**,
em reais, com padrão `100.00`. Informe decimal com ponto, não negativo e com até
duas casas. Configuração inválida bloqueia a operação. Exemplo no PowerShell,
antes de abrir a aplicação:

```powershell
$env:MAX_TOLERANCE_DIFF_REAIS = '100.00'
```

- Diferença zero: `PERFECT_MATCH`.
- Diferença absoluta até o limite, inclusive: `TOLERATED_WARNING`. O TXT usa
  o valor da PFMI, com aviso na interface e registro no log JSONL.
- Acima do limite: `CRITICAL_ERROR`, sem gerar TXT, mesmo que os campos antigos
  de aprovação de divergência estejam preenchidos. `0` configura fechamento exato.

A tolerância não resolve identidade, documentos inválidos, ambiguidades ou
aprovação de composição. Também não altera o fechamento dos valores nominais.
Em composições, compara a soma dos componentes com o crédito compartilhado uma
única vez e conserva o valor PFMI de cada componente. Não distribui diferenças
entre vários créditos independentes que não tenham um valor PFMI individual.

O resultado de geração expõe `reconciliation` (`BatchValidationSummary`), também
presente no `ValidationError` quando a reconciliação foi realizada e bloqueou.
O log `logs/gerador_cnab_nox.jsonl` registra o resumo, com valores inteiros em
centavos, nos eventos de geração e bloqueio. `diffCents` e `totalDiffCents` usam
PFMI menos Analítico; diferenças opostas não anulam bloqueios individuais.
`clearedRecords` contém correspondências perfeitas; `warnings`, as toleradas.
Os totais consideram os registros identificados e selecionados para a remessa.
Outras pendências mantêm `canGenerateTxt=false`, mesmo sem divergência financeira.
O retorno em memória contém documento normalizado e nome; os logs mantêm a
proteção existente, com documentos mascarados e nomes omitidos, vinculados por ID.
## STP: preenchimento automático com auditoria

O match por CPF/CNPJ válido e idêntico nas fontes pode aprovar a divergência de
nome com `APROVADO=SIM_SISTEMA` e `OK_COM_ALERTA_NOME`. Não substitui documentos
por similaridade e continua respeitando a falência e a identidade do crédito.

Para composições, o motor compara o total PFMI com o Analítico usando a tolerância
configurada. Exige um único candidato de valor **antes** de conferir CPF/CNPJ ou
nome. Dois candidatos, mesmo que só um pareça ter o nome certo, abortam a
auto-aprovação com `COMPOSICAO_BLOQUEADA_SELECAO_MANUAL`. Uma correspondência
única com documento ou nome forte preenche aba, linha e aprovações `SIM_SISTEMA`.
Nomes abreviados exigem pelo menos dois tokens completos iniciais e metade dos
tokens do nome maior; pequenas diferenças de grafia exigem primeiro token igual
e similaridade mínima de 92%. Valor isolado não confirma identidade.

Composições usam referências explícitas entre parênteses. Para o caso de advogado
sem referência, `JOSE VALDIR` (ou identidade nas regras locais de advogado da
falência) só se associa quando há um único titular no mesmo arquivo, aba, bloco
físico e falência. O motor não testa combinações arbitrárias entre operações.

O nominal é preenchido pela distribuição proporcional já existente, com soma
exata e resíduo de centavos no titular. A composição automaticamente aprovada
permanece `INCLUIR_CNAB=NAO` até a decisão do operador. Basta auditar e selecionar
SIM/NAO para o grupo inteiro; inclusão parcial continua proibida. Operações
inteiramente excluídas não bloqueiam as demais operações selecionadas.

As regras, fontes e preenchimentos ficam em `ALERTAS` e no histórico protegido.
Digitar `SIM_SISTEMA` não fabrica uma aprovação: a geração confere o snapshot
original e invalida aprovações automáticas quando os dados aprovados são alterados.
O TXT usa o intermediário sem reabrir o Analítico para seleções automáticas.

Fallbacks de sacados são lidos de [EQUIVALENCIAS_FALENCIAS.md](docs/EQUIVALENCIAS_FALENCIAS.md).
Herdam a data de liquidação do lote; os CNPJs de desenvolvimento cadastrados
precisam ser substituídos por documentos reais válidos antes de uma remessa.
### Intervenção humana: SIM prevalece sobre SIM_SISTEMA

Zero digitação obrigatória não significa impedir edição. Para substituir um
match, preencha `COMPOSICAO_SELECAO_MANUAL_ABA` e
`COMPOSICAO_SELECAO_MANUAL_LINHA` com a localização correta do Analítico e marque
`COMPOSICAO_APROVADA=SIM` em todas as linhas da composição. Para uma operação
individual, use os mesmos campos de aba/linha e `APROVADO=SIM`.
`SELECAO_MANUAL_APROVADA=SIM` também é uma assinatura humana válida da escolha.

A seleção humana pode substituir um crédito automático anterior. Os campos
finais digitados pelo operador são preservados. Valores automáticos ainda
intocados são atualizados a partir do crédito escolhido (incluindo o rateio
nominal exato); a geração registra `MANUAL_OVERRIDE=SIM` nos avisos.
Para corrigir somente o nome, basta editar o nome e marcar `APROVADO=SIM`.
Alterações manuais de valores de uma composição exigem `COMPOSICAO_APROVADA=SIM`.

Uma nova escolha de aba/linha requer o Analítico disponível na configuração do
lote. Seleções automáticas inalteradas usam o snapshot protegido, sem reabrir
fontes. CPF/CNPJ, limite financeiro, soma dos nominais, unicidade do crédito e
inclusão integral da composição continuam sendo validados. A origem protegida
permanece preservada para comparar a decisão humana com a sugestão anterior.
