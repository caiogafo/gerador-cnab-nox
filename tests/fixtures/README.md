# Fixtures sintéticas

Todos os arquivos desta pasta são dados fabricados exclusivamente para teste
automatizado. Nenhum conteúdo aqui vem de lotes, clientes, PFMIs ou operações
reais do fundo.

- `golden_cnab_v1_2.txt` / `golden_cnab_v1_2.sha256`: registro CNAB de exemplo
  (header + 1 detalhe + trailer) com nomes, documentos e valores fictícios,
  usado para travar o layout de 444 bytes por posição.
- `golden_cnab.sha256`: hash de uma variação sintética anterior do mesmo cenário.

Os documentos e nomes usados (`ACME LIMITADA`, `FALENCIA ALFA SOCIEDADE ANONIMA`,
CPF/CNPJ canônicos de teste como `52998224725` e `11222333000181`) são valores
de teste amplamente conhecidos, sem qualquer correspondência com pessoas,
empresas ou documentos reais.
