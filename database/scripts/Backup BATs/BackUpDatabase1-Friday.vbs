Const DestinationFile = "T:\Database\Backups\AQCProduction-Friday.accdb"
Const SourceFile = "C:\Transcription\Database\AQCProduction.accdb"

Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CopyFile SourceFile, DestinationFile, True
    Set fso = Nothing