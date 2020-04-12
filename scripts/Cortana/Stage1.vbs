dim accessApp
set accessApp = createObject("Access.Application")
accessApp.OpenCurrentDataBase("C:\Transcription\Database\AQCProduction.accdb")
accessApp.Run "Stage1BatchFile"
accessApp.Quit
set accessApp = nothing