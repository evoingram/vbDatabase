
for /D %%D in ("S:\UnprocessedAudio\Prepared\*") do (
            if exist "%%D\NewAudio\*.wav" (
                echo Processing %%D ...
                if not exist echo No audio to process...
                for %%I in ("%%D\NewAudio\*.*") do (
		    "C:\Program Files (x86)\sox-14-4-1\sox.exe" %%I %%D\FinalAudio\%1n.wav trim 0 1800 : newfile : restart
                )
            )
        )
