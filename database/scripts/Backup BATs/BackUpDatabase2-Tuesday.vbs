Const DestinationFile = "T:\Database\Backups\AQCProduction-backend-Tuesday.accdb"
Const SourceFile = "T:\Software\AQCProduction-backend.accdb"

Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile SourceFile, DestinationFile, True
    Set fso = Nothing