# 📋 Geração Automática de Lista Mestra de Verificações

Este script automatiza a criação de uma planilha mestra de verificações de arquivos para projetos armazenados no Google Drive. Ele é acionado sempre que um formulário é preenchido e executa uma série de etapas para organizar, registrar e gerar um relatório atualizado com os arquivos encontrados nas pastas do projeto.


## 🚀 Funcionamento

- A função é acionada automaticamente por um gatilho `onEdit` (sempre que a planilha for editada, normalmente através de um formulário).
- Lê a planilha principal e identifica se há novas linhas sem link de planilha gerada.
- Para cada nova linha:
   - Registra a data/hora da atualização.
   - Verifica a pasta do projeto no Drive.
   - Exclui a planilha anterior (se houver).
   - Cria uma nova planilha e a move para a pasta do projeto.
   - Percorre as subpastas (disciplinas > formatos > arquivos).
   - Gera o conteúdo da nova planilha (Nome do arquivo, disciplina, formato e link direto para o Google Drive)
   - Preenche a nova planilha com o conteúdo gerado, a formata e adiciona checkboxes para marcação de verificação por parte do usuário.
   - Atualiza a planilha principal com o link da nova planilha criada e a data/hora da criação.


## 📅 Última Atualização

28/06/2024


## 👨‍💻 Autor

Polyana Ramos Araújo
