dim accessApp
set accessApp = createObject("Access.Application")
accessApp.OpenCurrentDataBase("C:\Transcription\Database\AQCProduction.accdb")
accessApp.Run "IncompleteTasks"
accessApp.Quit
set accessApp = nothing