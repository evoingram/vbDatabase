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
for %%a in (%p_dir:~0,-1%) do (
	SET "p2_dir=%%~dpa"
	SET "p3_dir=!p2_dir!SREngine\"
	CALL ECHO p3dir is !p3_dir!
	CALL ECHO wavfilename is !wavfilename!
	CALL ECHO wavfilename1a is !wavfilename1!
	SET "navariable2=WorkingFiles\base.fileids"
	SET "navariable3=en-us/mdef"
	SET "navariable4=WorkingFiles\base.transcription"
	SET "navariable5=WorkingFiles\base.dic"
	SET "navariable6=WorkingFiles\base.lm"
	SET "navariable7=WorkingFiles\base.hyp"
	CALL "S:\sphinxtrain\bin\Release\Win32\sphinx_fe.exe -argfile !p3_dir!/en-us/feat.params -samprate 16000 -c !wavfilename1!!navariable2! -di . -do . -ei wav -eo mfc -mswav yes -nfft 2048"
	CALL "S:\sphinxtrain\bin\Release\Win32\pocketsphinx_mdef_convert.exe" -text !p3_dir!!navariable3! !p3_dir!!navariable3!.txt"
	CALL "S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir !p3_dir!en-us -moddeffn !p3_dir!!navariable3!.txt -ts2cbfn .ptm. -svspec 0-12/13-25/26-38 -feat 1s_c_d_dd -cmn current -agc none -dictfn !wavfilename1!!navariable5! -ctlfn !wavfilename1!!navariable2! -lsnfn !wavfilename1!!navariable4! -accumdir ."
	CALL "S:\sphinxtrain\bin\Release\Win32\mllr_solve.exe -meanfn !p3_dir!en-us/means -varfn !p3_dir!en-us/variances -outmllrfn !p3_dir!mllr_matrix -accumdir ."
	CALL "S:\sphinxtrain\bin\Release\Win32\map_adapt.exe -moddeffn !p3_dir!en-us/mdef.txt -ts2cbfn .ptm. -meanfn !p3_dir!en-us/means -varfn !p3_dir!en-us/variances -mixwfn !p3_dir!en-us/mixture_weights -tmatfn !p3_dir!sen-us/transition_matrices -accumdir . -mapmeanfn !p3_dir!en-us-adapt/means -mapvarfn !p3_dir!en-us-adapt/variances -mapmixwfn !p3_dir!en-us-adapt/mixture_weights -maptmatfn !p3_dir!en-us-adapt/transition_matrices"
	CALL "S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl !wavfilename1!!navariable2! -lm !wavfilename1!!navariable6! -dict !wavfilename1!!navariable5! -hmm !p3_dir!en-us-adapt -hyp !wavfilename1!!navariable7!"
	CALL "S:\sphinxtrain\scripts\decode\word_align.pl" !wavfilename1!!navariable4! !wavfilename1!!navariable7! >> word_align_output.txt"
	CALL ECHO fileids path is !wavfilename1!!navariable2!
	CALL ECHO transcription path is !wavfilename1!!navariable4!
	CALL ECHO DIC path is !wavfilename1!!navariable5!
	CALL ECHO LM path is !wavfilename1!!navariable6!
	CALL ECHO HYP path is !wavfilename1!!navariable7!
	CALL ECHO SREngine path is !p3_dir!
	SET /A "CountFileNumber=CountFileNumber+1"
	CALL ECHO ----------------------------------------------------------------
		)
)
ENDLOCAL