# Equivalências de falência no aplicativo

Na etapa **Configurar lote**, abra **Equivalências de falência**.
Informe um nome alternativo e um nome de referência, confirme que representam
a mesma falência e clique **Salvar equivalência**. O cadastro fica no computador
e é aplicado automaticamente ao preparar os próximos Excel, sem API.

Exemplo fictício: `FALÊNCIA ALFA` e `MASSA FALIDA ALFA` podem ser associados
somente depois de conferidos pela operadora. Nenhuma equivalência nova dos lotes
reais foi cadastrada automaticamente nesta implementação.

O botão **Remover selecionada** exclui a regra para preparações futuras.
Para corrigir o destino, selecione uma regra, altere o nome de referência e
confirme novamente. Use o mesmo nome de referência para todas as variações
daquela falência. Nomes de pessoas não devem ser cadastrados nesta tela.

As regras são aplicadas à PFMI, à campanha do Analítico e à busca na Base de
Vencimentos. Os nomes de origem permanecem intactos; o nome final do sacado
continua vindo da Base. Registros da Base com diferenças materiais continuam
gerando conflito, mesmo depois de confirmada a equivalência.

O Excel registra as equivalências usadas no RESUMO e no manifesto protegido.
Alterar o cadastro não muda um Excel já preparado. Prepare outro arquivo para
aplicar mudanças. A geração de TXT usa somente o Excel revisado e não lê o
cadastro local nem as fontes originais.

No Windows, o cadastro fica em
`%LOCALAPPDATA%/GeradorCNABNOX/equivalencias_falencias.json`.
Em outros sistemas, o padrão é `~/.local/share/GeradorCNABNOX/`.
Gravações usam substituição atômica; formato inválido, ciclos e destinos
conflitantes são recusados. Um cadastro ilegível bloqueia a preparação com
mensagem de erro em vez de ignorar regras silenciosamente.

Para chamadas Python, `prepare_workbook(..., failure_aliases={...})` aceita um
snapshot explícito. A interface lê o cadastro e captura esse snapshot antes de
iniciar o processamento. Chamadas diretas sem esse parâmetro continuam usando
somente as equivalências já existentes no produto.

Esta entrega implementa o cadastro e seu uso nos cruzamentos. Não altera a
comissão CONEX, não resolve a exceção de instrumentos separados e não aprova
pessoas por similaridade. Candidatos de valor exato e mesma falência continuam
disponíveis para seleção manual no intermediário conforme as regras existentes.
