for /D %%D in ("S:\UnprocessedAudio\Prepared\*") do (
            if exist "%%D\Audio\*.*" (
                echo Processing %%D ...
                if not exist echo No audio to process...
                for %%I in ("%%D\NewAudio\*.*") do (
                    start *.*
                    sox -d -r 16000 *.wav
                )
            )
        )