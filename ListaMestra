//FUNÇÃO "ListaMestra" RODANDO TODA VEZ QUE A PLANILHA FOR MODIFICADA >> ACIONADOR

function ListaMestra() {

  //Pegando a data e horário atual
  var d = new Date()
  var now = Utilities.formatDate(d, 'GMT-03','dd/MM/yyyy HH:mm');

  //Pegando a aba "Atualizações" da planilha atual e coletando os dados da última linha (da primeira até a quinta coluna)
  var planilha1 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Atualizações')
  var ultima_linha1 = planilha1.getLastRow()
  var infos1 = planilha1.getRange(ultima_linha1,1,1,10).getValues()
  var infos2 = planilha1.getRange(1,1,ultima_linha1,10).getValues()

  //Estrutura de repetição que vai rodar o código apenas para as linhas da aba "Atualizações" onde a última coluna está vazia
  for (linha=0; linha<ultima_linha1; linha++){

    if (infos2[linha][9] == ""){

      //Adicionando data e horário de atualização na aba "Atualizações"
      planilha1.getRange(linha+1,10).setValue(now)

      //Achando a pasta de projetos atualizados no drive e excluindo a planilha mestra antiga, caso exista
      var nome_projeto = infos1[0][0]
      var id_pasta = infos1[0][5]

      Logger.log(nome_projeto+': '+id_pasta)

      var pasta_projeto = DriveApp.getFolderById(id_pasta)

      Logger.log('Pasta: '+pasta_projeto)

      var arquivos_existentes = pasta_projeto.getFilesByName('Lista mestra verificações - '+ nome_projeto);

      Logger.log('Arquivos existentes: '+arquivos_existentes)

      if (arquivos_existentes.hasNext()) {
        var planilha_existente = arquivos_existentes.next();
        Logger.log('Planilha existente: '+planilha_existente)
        planilha_existente.setTrashed(true)
      }

      //Criando a planilha mestra na pasta do projeto
      var planilha_lista_mestra = SpreadsheetApp.create('Lista mestra verificações - '+ nome_projeto)
      DriveApp.getFileById(planilha_lista_mestra.getId()).moveTo(pasta_projeto)
      var aba_lista_mestra = planilha_lista_mestra.getSheetByName('Página1')
      aba_lista_mestra.setName('Verificações');

      //Pegando o link da planilha final
      var link_planilha = 'https://docs.google.com/spreadsheets/d/' + (planilha_lista_mestra.getId())

      //Configurando planilha
      aba_lista_mestra.setFrozenRows(1)
      aba_lista_mestra.setColumnWidth(1, 150)
      aba_lista_mestra.setColumnWidth(3, 300)
      aba_lista_mestra.setColumnWidth(4, 550)
      aba_lista_mestra.getRange('A1:Z1000').setBorder(true, true, true, true, true, true, "#ffffff", SpreadsheetApp.BorderStyle.SOLID)

      //Criando o cabeçalho e adicionando ao array final
      var cabecalho = [['Disciplina','Formato','Nome do arquivo','Link do arquivo','Verificado']]
      range_cabecalho = aba_lista_mestra.getRange('A1:E1')
      range_cabecalho.setValues(cabecalho)
      range_cabecalho.setBackground('#ffdd00')
      range_cabecalho.setFontFamily('Montserrat')
      range_cabecalho.setFontWeight("bold")
      range_cabecalho.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      //Adicionando o link da planilha na aba "Atualizações"
      planilha1.getRange(linha+1,9).setValue(link_planilha)

      //Verificando as disciplinas presentes no projeto e adicionando nomes e IDs das pastas  no array x
      var disciplinas = DriveApp.getFolderById(id_pasta).getFolders()
      var disciplinas_array = []

      while (disciplinas.hasNext()) {
        var disciplina = disciplinas.next()
        var name_disciplina = disciplina.getName().substring(disciplina.getName().length - 3);
        var id_disciplina = disciplina.getId()
        var x = [name_disciplina,id_disciplina]
        disciplinas_array.push(x)
      }
      
      var final_data = []

      //Estrutura de repetição para cada disciplina presente no projeto que verifica os formatos de arquivos presentes e adiciona o nome e ID das pastas no array y
      for (i = 0; i < disciplinas_array.length; i++) {
        var formatos = DriveApp.getFolderById(disciplinas_array[i][1]).getFolders()
        var formatos_array = []

        while (formatos.hasNext()) {
          var formato = formatos.next()
          var name_formato = formato.getName().substring(formato.getName().length - 3);
          var id_formato = formato.getId()
          var y = [name_formato,id_formato]
          formatos_array.push(y)
        }

        //Estrutura de repetição para cada formato de arquivo que verifica quais os arquivos presentes e adiciona o nome e ID dos arquivos no array z
        for (j = 0; j < formatos_array.length; j++) {
          var arquivos = DriveApp.getFolderById(formatos_array[j][1]).getFiles()
          var arquivos_array = []

          while (arquivos.hasNext()) {
            var arquivo = arquivos.next()
            var name_arquivo = arquivo.getName()
            var id_arquivo = arquivo.getId()
            var z = [name_arquivo,id_arquivo]
            arquivos_array.push(z)

            //Pegando informações finais sobre os arquivos em questão e adicionando no array dados_linha
            var final_disciplina=disciplinas_array[i][0]
            var final_formato=formatos_array[j][0]
            var final_arquivo_nome=name_arquivo
            var final_arquivo_link='https://drive.google.com/file/d/'+id_arquivo
            var dados_linha = [final_disciplina,final_formato,final_arquivo_nome,final_arquivo_link,'']
            final_data.push(dados_linha)
          }
        }
      }

      //Adicionando os dados finais na planilha do projeto e configurando as bordas das células
      var final_range = aba_lista_mestra.getRange(2,1,final_data.length,5)
      final_range.setValues(final_data)
      final_range.setFontFamily('Montserrat')
      final_range.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

      //Adicionando checks na última coluna da planilha
      var ultima_coluna = aba_lista_mestra.getRange(2,5,final_data.length,1)
      ultima_coluna.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())

      //Verificando qual linha da aba "Projeto" é correspondente ao projeto em questão e adicionando o link da planilha atualizada + data de atualização nas colunas finais
      var planilha2 = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Projetos')
      var ultima_linha2 = planilha2.getLastRow()

      for (k = 0; k < ultima_linha2; k++) {
        var nome_projeto2 = planilha2.getRange(k+1,1).getValue()    

        if (nome_projeto2 == nome_projeto) {
          var infos_finais = planilha2.getRange(k+1,9,1,2)
          var dados_finais = [[link_planilha,now]]
          infos_finais.setValues(dados_finais)
        }
      }
    }
  }
}
