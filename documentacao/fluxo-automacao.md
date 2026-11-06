# Fluxo de Automação

## Visão geral

A automação foi desenvolvida em VBA com o objetivo de simplificar e padronizar o processo de atualização e geração de relatórios gerenciais que são feitos mensalmente (Excel e Power Point).

A solução integra recursos do Excel, Power Query, Power Pivot e PowerPoint, reduzindo a necessidade de intervenções manuais durante a preparação dos relatórios e permitindo o armazenamento de versões históricas.

A solução utiliza-se de 5 bases:

- Base Mês Anterior
- Base Mês Atual
- Base Mês Anterior - OI
- Base Mês Atual  - OI
- Referências

O Power Pivot une as bases do mês anterior e mês atual através das "Referências" possibilitando a criação da tabela dinâmica comparando os dois meses.

A automação é dividida em duas etapas - Atualização e Geração de Arquivos - cada uma operada por um botão.

Poderia ser feita tudo no mesmo botão mas optei por dividir em duas etapas para que primeiro atualize o modelo de dados e planilhas e antes de gerar os arquivos fazer a conferência se está tudo OK. 

Estando correto dispara a segunda etapa da automação.