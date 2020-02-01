dim accessApp
dim cInvoice as new cmInvoice
set accessApp = createObject("Access.Application")
accessApp.OpenCurrentDataBase("C:\Transcription\Database\AQCProduction.accdb")
accessApp.Run "pfAutoCalculateFactorInterest"
'cInvoice.pfAutoCalculateFactorInterest()
accessApp.Quit
set accessApp = nothing