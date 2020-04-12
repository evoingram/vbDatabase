dim accessApp
set accessApp = createObject("Access.Application")
accessApp.OpenCurrentDataBase "C:\Transcription\Database\AQCProduction.accdb"
accessApp.visible = true
accessApp.Run "ArchiveEmails"
'call DailyTaskAddFunction()
'accessApp.Quit
set accessApp = nothing