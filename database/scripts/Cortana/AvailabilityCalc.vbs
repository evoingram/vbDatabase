dim accessApp
set accessApp = createObject("Access.Application")
accessApp.OpenCurrentDataBase("C:\Transcription\Database\AQCProduction.accdb")
accessapp.run "AvailabilitySchedule"
accessApp.Quit
set accessApp = nothing