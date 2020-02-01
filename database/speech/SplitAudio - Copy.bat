echo sox infile.wav output.wav trim 0 30 : newfile : restart
for /D %%D in ("S:\UnprocessedAudio\Prepared\*") do (
            if exist "%%D\NewAudio\*.*" (
                echo Processing %%D ...
                if not exist echo No audio to process...
                for %%I in ("%%D\NewAudio\*.*") do (
		    sox %%D\NewAudio\%%I %%D\FinalAudio\%%I-%1n.wav trim 0 1800 : newfile : restart
                )
            )
        )




echo @off
for /D %%D in ("S:\UnprocessedAudio\Prepared\111\NewAudio\*") do (
            if exist "%%D\1.wav" (
                echo Processing %%D ...
            if not exist echo No audio to process...
                for %%I in ("%%D\*.*") do (
		    "C:\Program Files (x86)\sox-14-4-1\sox.exe" %%I "%%D\%%I-%1.wav" trim 0 1800 : newfile : restart
                )
            )
        )


"C:\Program Files (x86)\sox-14-4-1\sox.exe" S:\UnprocessedAudio\Prepared\111\NewAudio\1.wav S:\UnprocessedAudio\Prepared\111\FinalAudio\1-%1.wav trim 0 1800 : newfile : restart