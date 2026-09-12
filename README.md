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
