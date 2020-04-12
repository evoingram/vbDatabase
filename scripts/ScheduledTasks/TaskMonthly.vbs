dim accessApp
dim cAdminF as new cmAdminFunctions
set accessApp = createObject("Access.Application")
accessApp.OpenCurrentDataBase("C:\Transcription\Database\AQCProduction.accdb")
accessApp.visible = true
accessApp.Run "pfMonthlyTaskAddFunction"
'CALL pfMonthlyTaskAddFunction()
accessApp.Quit
set accessApp = nothing