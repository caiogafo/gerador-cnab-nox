# Homologação Windows e Frontis — V1 revisada

**Contrato:** [Specs 1.1](../specs.md) e [DRD](../drd.md).  
**Estado em 2026-09-04:** roteiro futuro; M7 e M8 pendentes.
O candidato Linux e suas provas não comprovam build nem execução do pacote Windows.

## 1. Pré-condições

M1–M6 concluídos e comprovados no Linux, com registros no
[ExecPlan](../execplan.md). Transferir código/lock e fixtures sintéticas ao Windows
do responsável técnico preservando alterações existentes. Dados reais permanecem
privados; não incluí-los no pacote nem no repositório.

Build exige ferramentas de desenvolvimento no Windows do responsável técnico. O computador
de uso do pacote não deve exigir Python, Excel, conexão de rede ou administrador.
A revisão humana do Excel pode ser feita em editor compatível; a aplicação não
deve depender dele para ler planilhas ou gerar TXT.

Não migrar intermediário antigo: preparar um novo no esquema CNAB-NOX-V1-2.

## 2. M7 — Build na máquina Windows do responsável técnico

No diretório do projeto, com Python/uv e Tk disponíveis para desenvolvimento:

```powershell
uv sync --locked --group dev --group windows
.\scripts\smoke_windows.ps1
uv run pyinstaller --noconfirm --clean --windowed --onedir --paths src --collect-data gerador_cnab_nox --name GeradorCNABNOX scripts/launch_gui.py
```

O artefato é a pasta completa `dist\GeradorCNABNOX`, com
`GeradorCNABNOX.exe` e suas dependências. Não distribuir o exe isolado.
`--collect-data gerador_cnab_nox` inclui os ícones locais das etapas. Conferir
esses recursos e as três telas no pacote final; a melhoria Linux não homologa
automaticamente o build Windows. A opção é descrita na
[documentação do PyInstaller](https://www.pyinstaller.org/en/stable/usage.html#what-to-bundle-where-to-search).
O script executa a suíte revisada, incluindo widgets Tk, e um smoke de importação.
Exigir sessão gráfica; nenhum teste GUI pode ser pulado. O teste de chmod POSIX
será pulado no Windows e não substitui testar ACL/permissão real no pacote. Os testes em Python não
substituem o roteiro abaixo executado no pacote final.

Registrar ambiente Windows/arquitetura, versão das ferramentas, versão de código
(ou hashes quando não houver commit), hash do pacote e resultados sanitizados.
Build concluído não equivale a M7 aceito.

## 3. M7 — Teste do pacote, não do Python de desenvolvimento

1. Executar em Windows 10/11 x64 como usuário comum, com aplicação sem depender
   de Python/Excel instalados. Validar as versões de Windows efetivamente disponíveis;
   registrar o que foi testado, sem presumir a outra versão.
2. Extrair a pasta onedir em caminho comum e caminho com espaços/acentos.
   Abrir o exe diretamente, testar seleção e gravação nesses caminhos.
3. Registrar avisos de antivírus/SmartScreen e eventuais bloqueios. Não desativar
   proteções, instalar certificados nem exigir administrador para fabricar aprovação.
4. Cancelar cada diálogo de entrada/saída: nenhuma saída indevida ou mensagem de sucesso falsa.
5. Adicionar/remover/reordenar PFMIs e confirmar a ordem de geração.
6. Exercitar Normal, Cessão da Cessão e lote misto. Usar taxas próprias, 15%, 0%,
   percentual fracionário com vírgula/ponto e campo apagado.
7. Preparar fontes sintéticas válidas e fontes com pendências. Excel deve sair
   para sacado ausente/conflitante, dado inválido e linha desconhecida; erro de
   arquivo ilegível/corrompido/cabeçalho ausente deve orientar a correção.
8. Conferir cinco abas, visibilidade de originais/cálculos, campos finais editáveis,
   pagamentos/grupos/títulos separados e diferenças financeiras.
9. Verificar resumo auxiliar vazio/divergente como alerta. Linha desconhecida ou
   pagamento inconsistente exige corrigir PFMI e preparar novamente.
10. Corrigir finais no Excel: nome, documento, valor e emissão. Para Cessão da Cessão,
    preencher assinatura nova; preservar original como referência.
11. Confirmar que correção final válida pode diferir da sugestão calculada, desde que
    feche com PFMI. R$ 0,01 divergente bloqueia; APROVADO=SIM não libera diferença,
    documento inválido ou origem inexistente.
12. Taxa inválida deve exigir nova preparação, não ser contornada por edição de
    status ou mudança posterior da interface. Esquema antigo/controles alterados recusados.
13. Após preparar e corrigir, tornar cópias de fontes indisponíveis e mudar os
    parâmetros visuais: gerar apenas pelo intermediário e conferir resultado estável.
14. Conferir Conexcred sem nome composto e advogados sem nome do credor em parênteses.
    Confirmar nomes técnicos longos com alerta e nome completo mantido no Excel.
15. Validar TXT com verificador independente: CP1252, 444 bytes por registro,
    CRLF final, sem BOM, documentos/padding, posições, sequência e trailer.
16. Validar publicação atômica por hard link no sistema de arquivos de destino
    (inclusive o destino efetivamente usado pelo operador do fundo). Repetir gravação em pasta
    sem permissão e nome de saída já existente:
    erro compreensível ou nova saída exclusiva, sem sobrescrever dados existentes.
17. Testar falha do log após gravação usando caso controlado: saída salva continua
    identificada como sucesso com alerta, sem gerar duplicata silenciosa.

Todos os testes controlados usam cópias/fixtures sintéticas, nunca adulteram fontes
reais. Para comprovar independência das fontes, mover apenas cópias de teste.

## 4. Regressão histórica no Windows

Separar renderizador, fluxo completo sintético e histórico real, conforme Specs.
Quando as entradas contemporâneas correspondentes estiverem disponíveis, comparar
o resultado com o esperado byte a byte.

Os textos compostos legados de Conexcred e advogados mudam intencionalmente:
delimitar a expectativa exata de cada campo alterado, não ignorar nomes em geral.
Sem Analítico/Base correspondentes à época, documentar a limitação e correções
manuais usadas. Não apresentar dado extraído do TXT/gerador como preenchimento
automático do cruzamento.

Registrar contagens, totais e hashes; não anexar planilhas, documentos ou nomes
reais ao repositório. Os oito lotes somam 86 títulos/R$ 1.539.839,34, mas essa
referência não substitui a prova de cada fluxo e de cada diferença intencional.

## 5. Evidência e saída de M7

Registrar no ExecPlan:

- ambiente efetivamente testado, identidade/hash do pacote e comando do build;
- itens do roteiro aprovados/reprovados, problemas, correções e retestes;
- prova de execução no pacote sem dependências de desenvolvimento;
- riscos residuais de Windows/antivírus e instruções reproduzíveis.

Se faltar prova de ambiente ou comportamento, M7 permanece pendente/parcial,
não concluído por suposição. Somente depois do aceite técnico do pacote seguir
para a máquina do operador do fundo.

## 6. M8 — Lote real com o operador do fundo / Frontis

1. O operador do fundo seleciona uma ou várias PFMIs reais e confirma
   modalidade/comissão de cada arquivo, Analítico, Base, liquidação e primeiro
   número pelo estoque do fundo.
2. Prepara Excel, revisa créditos e corrige finais. Resolve erros de origem com
   nova preparação, sem apagar controles. Resumos auxiliares inconsistentes não
   devem bloquear por si só.
3. Antes da importação, confere quantidade de títulos, total nominal, aquisição,
   faixa sequencial, cedentes, sacados, emissão e vencimentos entre Excel e TXT.
   Beneficiários/pagamentos não são contados como títulos.
4. O operador do fundo importa manualmente no Frontis, verifica aceitação/rejeição
   e reconcilia quantidade, valores e cadastros efetivamente registrados pelo documento.
5. Havendo rejeição, registrar motivo sanitizado, corrigir no escopo e retestar
   com controle humano para evitar cargas duplicadas.
6. Registrar **aceite operacional do operador do fundo**: carga entrou e foi reconciliada.
7. Registrar separadamente **aceite final do responsável pela aprovação final**: V1 aceita como produto.

Não presumir aceite nem automatizar envio ao Frontis. Sem ambos os aceites, M8
e V1 produto permanecem pendentes. Depois do uso real, refinamentos novos são
priorizados pelo responsável pela aprovação final, não incorporados silenciosamente
ao Goal encerrado.
