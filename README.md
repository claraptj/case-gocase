O arquivo db_preparation.py faz a limpeza e a transformação dos dados brutos localizados na planilha .XLSX de base corporativa mais recente presente na pasta.
Isso significa que o script está preparado para atualizações indefinidas da base de dados, contanto que:
  - Ela siga a mesma nomenclatura de arquivo (iniciada por "BASE CORPORATIVA" e com extensão .XLSX)
  - Ela contenha uma aba chamada "Página1", e nesta aba esteja localizada a base principal dos dados

Vale ressaltar para que as colunas presentes na versão atual da planilha não sejam renomeadas em versões futuras.
Caso contrário, o dashboard em 'Analise Comercial.PBIX' encontrará erro nas colunas da base, e precisará ser corrigido manualmente.
