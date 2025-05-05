const abaListaMestra = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Lista Mestra - Forms')

function gerarListaMestra(){

  //Pegando os dados da planilha preenchida pelo forms
  let dadosListaMestra = abaListaMestra.getRange(1,1,abaListaMestra.getLastRow(),abaListaMestra.getLastColumn()).getValues()
  let novosDadosListaMestra = dadosListaMestra.filter(function(linha){return linha[6]=="" || linha[6]==null})

  //Para cada projeto
  for(let i=0; i<novosDadosListaMestra.length; i++){
    let projetoNome = novosDadosListaMestra[i][1]
    let clienteNome = novosDadosListaMestra[i][2]
    let pastaAtualizadosId = PlanilhaBaseDadoseInovacao.getIdFromUrl(novosDadosListaMestra[i][3])
    let discordId = novosDadosListaMestra[i][4]
    let index = dadosListaMestra.indexOf(novosDadosListaMestra[i])

    Logger.log(`${i+1}. Projeto "${projetoNome}" do cliente "${clienteNome}"`)

    //Pegando os arquivos presentes na pasta do projeto
    let pastaAtualizados = DriveApp.getFolderById(pastaAtualizadosId)
    let arquivosPastaAtualizados = pastaAtualizados.getFilesByName(`Lista mestra verificações - ${projetoNome}`)

    //Excluindo a antiga planilha de lista mestra caso exista
    if(arquivosPastaAtualizados.hasNext()){
      arquivosPastaAtualizados.next().setTrashed(true)
      Logger.log('Lista mestra já existente foi excluída')
    }

    //Criando a planilha de lista mestra
    let planilhaListaMestra = SpreadsheetApp.create(`Lista mestra verificações - ${projetoNome}`)

    //Movendo para a pasta do projeto
    DriveApp.getFileById(planilhaListaMestra.getId()).moveTo(pastaAtualizados)

    //Pegando o link da planilha final
    let planilhaLink = planilhaListaMestra.getUrl()
    Logger.log(`Lista mestra criada: ${planilhaLink}`)

    //Adicionando o link da planilha
    abaListaMestra.getRange(index+1,7).setValue(planilhaLink)

    //Renomeando a aba da planilha
    let abaVerificacoes = planilhaListaMestra.getSheetByName('Página1')
    abaVerificacoes.setName('Verificações')

    //Configurando planilha
    abaVerificacoes.setFrozenRows(1)
    abaVerificacoes.setColumnWidth(1,150)
    abaVerificacoes.setColumnWidth(3,300)
    abaVerificacoes.setColumnWidth(4,550)
    abaVerificacoes.getRange('A1:Z1000').setBorder(true,true,true,true,true,true,"#ffffff",SpreadsheetApp.BorderStyle.SOLID)

    //Criando o cabeçalho e adicionando ao array final
    let cabecalhoPlanilha = abaVerificacoes.getRange('A1:E1')
    cabecalhoPlanilha.setValues([['Disciplina','Formato','Nome do arquivo','Link do arquivo','Verificado']])
    cabecalhoPlanilha.setBackground('#ffdd00')
    cabecalhoPlanilha.setFontFamily('Montserrat')
    cabecalhoPlanilha.setFontWeight("bold")
    cabecalhoPlanilha.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //Verificando as disciplinas existentes no projeto
    let disciplinas = DriveApp.getFolderById(pastaAtualizadosId).getFolders()
    var disciplinasExistentes = []

    //Pegando nome e ID das disciplinas
    while(disciplinas.hasNext()){
      let disciplina = disciplinas.next()
      let disciplinaNome = disciplina.getName().substring(disciplina.getName().length-3)
      var disciplinaId = disciplina.getId()
      disciplinasExistentes.push([disciplinaNome,disciplinaId])
    }

    let dadosFinais = []

    //Para cada disciplina existente
    for(let j=0; j<disciplinasExistentes.length; j++){
      Logger.log(`${i+1}.${j+1}. Disciplina: ${disciplinasExistentes[j][0]}`)
      let formatos = DriveApp.getFolderById(disciplinasExistentes[j][1]).getFolders()
      let formatosExistentes = []

      //Pegando nome e ID dos formatos
      while (formatos.hasNext()){
        let formato = formatos.next()
        let formatoNomeCompleto = formato.getName()
        if(formatoNomeCompleto.substring(formatoNomeCompleto.length-6)=="OUTROS"){
          var formatoNome = "OUTROS"
        }
        else{
          var formatoNome = formatoNomeCompleto.substring(formatoNomeCompleto.length-3)
        }
        let formatoId = formato.getId()
        formatosExistentes.push([formatoNome,formatoId])
      }

      //Para cada formato existente
      for(let k=0; k<formatosExistentes.length; k++){
        Logger.log(`${i+1}.${j+1}.${k+1}. Formato: ${formatosExistentes[k][0]}`)
        var arquivos = DriveApp.getFolderById(formatosExistentes[k][1]).getFiles()

        //Pegando nome e ID dos arquivos
        while (arquivos.hasNext()){
          let arquivo = arquivos.next()
          let arquivoNome = arquivo.getName()
          let arquivoId = arquivo.getId()
          Logger.log(`Arquivo ${arquivoNome}`)

          //Pegando informações finais sobre os arquivos
          // let disciplinaArquivo = disciplinasExistentes[i][0]
          // let formatoArquivo = formatosExistentes[j][0]
          let disciplinaArquivo = disciplinasExistentes[j][0]
          let formatoArquivo = formatosExistentes[k][0]
          dadosFinais.push([disciplinaArquivo,formatoArquivo,arquivoNome,`https://drive.google.com/file/d/${arquivoId}`,''])
        }
      }
    }

    Logger.log(dadosFinais)

    //Adicionando os dados finais na planilha do projeto e configurando as bordas das células
    var corpoPlanilha = abaVerificacoes.getRange(2,1,dadosFinais.length,5)
    corpoPlanilha.setValues(dadosFinais)
    corpoPlanilha.setFontFamily('Montserrat')
    corpoPlanilha.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

    //Adicionando checks na última coluna da planilha
    abaVerificacoes.getRange(2,5,dadosFinais.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())

    let mensagem = `\n **-------------------------------------------------------------------------------------------------------------------------**\n:loudspeaker:   **NOTIFICAÇÃO DO PROJETO: ${projetoNome} **\n \n:arrow_right:  A lista de arquivos presentes na pasta de Projetos Atualizados foi gerada com sucesso. Confira no seguinte link: ${planilhaLink}`

    try{PlanilhaBaseDadoseInovacao.notificacaoDiscord(discordId,mensagem)}catch(e){Logger.log(`Erro ao enviar a notificação no Discord: ${e}`)}
  }
}
