const masterListSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Master List - Forms')

function generateMasterList(){

  //Fetching data from the spreadsheet filled out by the form
  let masterListData = masterListSheet.getRange(1,1,masterListSheet.getLastRow(),masterListSheet.getLastColumn()).getValues()
  let masterListNewData = masterListData.filter(function(linha){return linha[6]=="" || linha[6]==null})

  //For each project
  for(let i=0; i<masterListNewData.length; i++){
    let projectName = masterListNewData[i][1]
    let clientName = masterListNewData[i][2]
    let updatedFolderId = PlanilhaBaseDadoseInovacao.getIdFromUrl(masterListNewData[i][3])
    let discordId = masterListNewData[i][4]
    let index = masterListData.indexOf(masterListNewData[i])

    Logger.log(`${i+1}. Project "${projectName}" of client "${clientName}"`)

    //Fetching the present files in the project folder
    let updatedFolder = DriveApp.getFolderById(updatedFolderId)
    let updatedFolderFiles = updatedFolder.getFilesByName(`Master list - ${projectName}`)

    //Deleting the old master list spreadsheet if it exists
    if(updatedFolderFiles.hasNext()){
      updatedFolderFiles.next().setTrashed(true)
      Logger.log('Existing master list has been deleted')
    }

    //Creating the master list spreadsheet
    let masterListSpreadsheet = SpreadsheetApp.create(`Master list - ${projectName}`)

    //Moving to the project folder
    DriveApp.getFileById(masterListSpreadsheet.getId()).moveTo(updatedFolder)

    //Fetching the final spreadsheet link
    let spreadsheetLink = masterListSpreadsheet.getUrl()
    Logger.log(`Master list created: ${spreadsheetLink}`)

    //Adding the spreadsheet link
    masterListSheet.getRange(index+1,7).setValue(spreadsheetLink)

    //Renaming the spreadsheet sheet
    let checkSheet = masterListSpreadsheet.getSheetByName('Page1')
    checkSheet.setName('Checks')

    //Setting up the spreadsheet
    checkSheet.setFrozenRows(1)
    checkSheet.setColumnWidth(1,150)
    checkSheet.setColumnWidth(3,300)
    checkSheet.setColumnWidth(4,550)
    checkSheet.getRange('A1:Z1000').setBorder(true,true,true,true,true,true,"#ffffff",SpreadsheetApp.BorderStyle.SOLID)

    //Creating the header and adding it to the final array
    let spreadsheetHeader = checkSheet.getRange('A1:E1')
    spreadsheetHeader.setValues([['Discipline','Format','File name','File link','Checked']])
    spreadsheetHeader.setBackground('#ffdd00')
    spreadsheetHeader.setFontFamily('Montserrat')
    spreadsheetHeader.setFontWeight("bold")
    spreadsheetHeader.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

    //Checking the existing disciplines in the project
    let disciplines = DriveApp.getFolderById(updatedFolderId).getFolders()
    var existingDisciplines = []

    //Fetching the disciplines' name and ID
    while(disciplines.hasNext()){
      let discipline = disciplines.next()
      let disciplineName = discipline.getName().substring(discipline.getName().length-3)
      var disciplineId = discipline.getId()
      existingDisciplines.push([disciplineName,disciplineId])
    }

    let finalData = []

    //For each existing discipline
    for(let j=0; j<existingDisciplines.length; j++){
      Logger.log(`${i+1}.${j+1}. Discipline: ${existingDisciplines[j][0]}`)
      let formats = DriveApp.getFolderById(existingDisciplines[j][1]).getFolders()
      let existingFormats = []

      //Fetching the formats' name and ID
      while (formats.hasNext()){
        let format = formats.next()
        let fullNameFormat = format.getName()
        if(fullNameFormat.substring(fullNameFormat.length-6)=="OTHERS"){
          var formatName = "OTHERS"
        }
        else{
          var formatName = fullNameFormat.substring(fullNameFormat.length-3)
        }
        let formatId = format.getId()
        existingFormats.push([formatName,formatId])
      }

      //For each existing format
      for(let k=0; k<existingFormats.length; k++){
        Logger.log(`${i+1}.${j+1}.${k+1}. Format: ${existingFormats[k][0]}`)
        var files = DriveApp.getFolderById(existingFormats[k][1]).getFiles()

        //Fetching the files' name and ID
        while (files.hasNext()){
          let file = files.next()
          let fileName = file.getName()
          let fileId = file.getId()
          Logger.log(`File ${fileName}`)

          //Fetching the files' final information
          let fileDiscipline = existingDisciplines[j][0]
          let fileFormat = existingFormats[k][0]
          finalData.push([fileDiscipline,fileFormat,fileName,`https://drive.google.com/file/d/${fileId}`,''])
        }
      }
    }

    Logger.log(finalData)

    //Adding the final data to the project spreadsheet and setting up cell borders
    var spreadsheetBody = checkSheet.getRange(2,1,finalData.length,5)
    spreadsheetBody.setValues(finalData)
    spreadsheetBody.setFontFamily('Montserrat')
    spreadsheetBody.setBorder(true, true, true, true, true, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM)

    //Adding checkboxes in the last column of the spreadsheet
    checkSheet.getRange(2,5,finalData.length,1).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build())

    let message = `\n **-------------------------------------------------------------------------------------------------------------------------**\n:loudspeaker:   **PROJECT NOTIFICATION: ${projectName} **\n \n:arrow_right: The list of files in the Updated Projects folder has been successfully generated. Check it at the following link: ${spreadsheetLink}`

    try{DataAndInnovationSpreadsheet.discordNotification(discordId,message)}catch(e){Logger.log(`Error sending notification on Discord: ${e}`)}
  }
}
