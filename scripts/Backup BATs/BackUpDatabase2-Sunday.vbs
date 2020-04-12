Const DestinationFile = "T:\Database\Backups\AQCProduction-backend-Sunday.accdb"
Const SourceFile = "T:\software\AQCProduction-backend.accdb"

Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile SourceFile, DestinationFile, True
    Set fso = Nothing