@echo off
SETLOCAL ENABLEDELAYEDEXPANSION
set "Apath=S:\UnprocessedAudio\Prepared\*.wav"
SET "CountFileNumber=1"
for /f "delims=" %%a in ('dir /s /b "%Apath%"') do (
	CALL ECHO ----------------------------------------------------------------
	CALL ECHO Loop Number !CountFileNumber!
	SET "p_dir=%%~dpa"
	set "wavfilename=%%a"
	set wavfilename1=!wavfilename:~0,-14!
	set "wavfilename2=%%~na"
for %%a in (%p_dir:~0,-1%) do (
	SET "p2_dir=%%~dpa"
	SET "p3_dir=!p2_dir!SREngine\"
	CALL ECHO p3dir is !p3_dir!
	CALL ECHO wavfilename is !wavfilename!
	CALL ECHO wavfilename1a is !wavfilename1!
	CALL ECHO wavfilename2 is !wavfilename2!
	SET "navariable2=WorkingFiles\base.fileids"
	SET "navariable3=en-us/mdef"
	SET "navariable4=WorkingFiles\base.transcription"
	SET "navariable5=WorkingFiles\base.dic"
	SET "navariable6=WorkingFiles\base.lm"
	SET "navariable7=WorkingFiles\base.hyp"
	SET "navariable2=NewAudio\"
	SET "ext=.wav"
	CALL Echo CountFileNumber in second loop is !CountFileNumber!
	Call ECHO "S:\pocketsphinx\bin\Release\Win32\pocketsphinx_continuous.exe -infile !wavfilename! -hmm !p3_dir!en-us-adapt -lm !p2_dir!!navariable6! -dict !p2_dir!!navariable5! >> !wavfilename1!!wavfilename2!-full-output.txt"
	CALL ECHO Output at !wavfilename1!!wavfilename2!-full-output.txt
	SET /A "CountFileNumber=CountFileNumber+1"
		)
)
ENDLOCAL