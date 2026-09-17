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

O cadastro de equivalências é independente das heurísticas STP. O Smart Match
usa candidato único por valor e evidência de documento/nome, com `SIM_SISTEMA`
rastreável. O operador pode substituir essa escolha informando aba/linha e
aprovação humana `SIM`; a comissão CONEX e as validações financeiras permanecem.
# Fallback de sacados para STP

A preparação lê exclusivamente o bloco JSON delimitado abaixo. CNPJs são
normalizados para 14 dígitos. O vencimento de cada fallback recebe a **data de
liquidação do lote**, com formato de célula Excel `DD/MM/YYYY`. A regra só atua
quando a Base de Vencimentos não resolve o sacado de forma consistente.
Cada injeção aparece em `ALERTAS` e no histórico protegido do intermediário.

**Dados de desenvolvimento fornecidos pelo operador:** os dois CNPJs abaixo
têm 14 dígitos, mas não passam no checksum. São preenchidos para testar o fluxo;
a validação de geração permanece bloqueada até cadastrar CNPJs válidos confirmados.
Não representam CNPJs reais verificados dessas massas falidas.

<!-- CNAB_NOX_FALLBACKS_BEGIN -->
```json
{
  "COSTEIRA": {
    "cnpj": "48060297000107",
    "nome": "MASSA FALIDA COSTEIRA TRANSPORTES"
  },
  "LP DISPLAYS": {
    "cnpj": "04119093000129",
    "nome": "MASSA FALIDA LP DISPLAYS BRASIL"
  }
}
```
<!-- CNAB_NOX_FALLBACKS_END -->

Para uma instalação em outra pasta, `CNAB_NOX_EQUIVALENCIAS_PATH` aponta para
este documento. Chamadas Python podem fornecer `fallback_document_path`.
JSON inválido, chaves duplicadas normalizadas ou CNPJs com tamanho incorreto
bloqueiam a preparação; o aplicativo não ignora silenciosamente um cadastro inválido.
