---------------------------------------------------------

HOW TO PREPARE FILES FOR TRAINING

You may use other numbers in \\HUBCLOUD\evoingram\speech\UnprocessedAudio\ for samples of other working documents you may want to look at.

PREPARING AUDIO:

	Find a completed job from either \\HUBCLOUD\evoingram\Comp Needs paperwork\ or \Completed with Payment\.
		I have done 453 and 3 dates from 446, 2/8/18, 2/12/18, & 3/9/18.  Everything else up for grabs.
	ADAM: If you want to skip the rest of this section, please just copy the direct audio + transcript into a new folder in
		\UnprocessedAudio\ and go to "PREPARING ALL OTHER FILES"
	Copy its corresponding final transcript docx into \UnprocessedAudio\ OR this folder you just made within \UnprocessedAudio\.
	Open audio player (FTR, liberty, javs, wmp, etc) & load corresponding audio from folder you selected.
	Open Audacity.
	Make sure project rate is 16000 ONLY and record moderately loudly, not quietly.
	Record corresponding audio into audacity file (hit record in audacity and then hit play to play the file).
	When complete, zoom out so you can see entire audio file in half-hour marks across the ruler on one screen.
	Select first half-hour.
	Go to file --> export-selected audio -->save as short name with "-##of##" at end 16-bit pcm something.
		See other folders for examples.
	Save in \UnprocessedAudio\ directory or next number folder in that directory.
	Repeat for every half-hour until you are all the way through the audio.

PREPARING ALL OTHER FILES:
	Create a folder that is one higher than the highest number in \UnprocessedAudio\
	OR rename one you just put audio+transcript in to one number higher than the highest number in \UnprocessedAudio\
	Move your audio and transcript from \UnprocessedAudio\ into that folder.
	Create the following 'new text documents': (make * the same as your wav name without the ##of## part)
		*.fileids 
		*.transcription
		*.dic
		*.lm 
		*.docx 
		*-corpus.txt 
	Open final transcript in number folder in \\HUBCLOUD\evoingram\speech\UnprocessedAudio\
	Copy everything from the SECOND line of the transcript body, so everything below the "CITY, STATE, DATE, TIME"
		line, down to and including one line BEFORE (Hearing/Proceedings concluded at time.) at the end.
	Paste as text only into wavfilename.docx.
	Close final transcript, don't save.  You're done with this file.
	Save wavfilename.docx, but do not close it.  This is the one we are mainly working with.
	Find/replace so that every sentence begins with "<s> " and ends with "</s> (wavfilename)"
		ALL SENTENCES CAPITALIZED but lowercase <s>, </s> and (wavfilename).
		One sentence per line
	I suggest replacements in the following order:
		++ Turn on Wildcards ++
		" {8,8}(*{1,})\: {2,2}" to "<s> "
		"^t{2,2}(*{1,})\: {2,2}" to "<s> "
		-- Turn off Wildcads --
		"^p {8,8}" to nothing
		" -- ^p" to " </s> (wavfilename)^p"
		" --^p" to " </s> (wavfilename)^p"
		"?  ^p" to " </s> (wavfilename)^p"
		"? ^p" to " </s> (wavfilename)^p"
		"?p" to " </s> (wavfilename)^p"
		".  ^p" to " </s> (wavfilename)^p"
		". ^p" to " </s> (wavfilename)^p"
		".^p" to " </s> (wavfilename)^p"
		"?  " to " </s> (wavfilename)^p<s> "
		".  " to " </s> (wavfilename)^p<s> "
		" -- " to " "
		"." "," "'" "?" to nothing
	Delete all lines with:
		Headings
		SHORT PHRASES IN ALL CAPS LIKE 'COURT'S RULING' 'ARGUMENT BY THE...' 'DIRECT EXAMINATION' etc
		Lines with parentheses'd comments
		Any extra blank lines or anything that doesn't fit the <s> text here </s> (wavfilename) format
	Select all, lowercase all, and copy.  Do not close.
	Open wavfilename-corpus.txt file and paste the lowercase stuff in there.
	Delete "<s> " and "</s> (wavfilename)" in corpus file (find/replace with nothing)
	Save, close, and go back to previous word document you were making all the replacements on.
	Select all, capitalize all.
	Find/replace, match case, <S>, </S> and (WAVFILENAME) to make lowercase <s>, </s> and (wavfilename).
	Select all and copy.
	Open wavfilename.transcription file and paste in there.  Save.
	Copy again and open Visual studio or other programming software that shows you lines 
		Visual basic component in office also does this.
	Paste your transcription file in there.
	Note number of lines.
	Save and close.
	Open wavfilename.fileids from your folder.
	Put wavfilename on each line for the EXACT SAME number of lines in your wavfilename.transcription.
		i copy it enough times to where i think it's close and then verify line count in programming software.
	Save and close.
	Go to http://www.speech.cs.cmu.edu/tools/lmtool-new.html
	Upload wavfilename-corpus.txt file.  When it completes its task:
	"save link as" *.dic as wavfilename.dic
	"save link as" *.lm as wavfilename.lm
	Now, open up the word document, wavfilename.docx, again and load up each wav file in a player.
	Note first and last line of words of the wav file.
	Search for them in the wavfilename.docx so that you select everything that's on the audio you just played.
	In each selected line, ensure the 'wavfilename' matches each of the wav files it belongs in.
		You want to make sure the wavfilename includes the "-##of##" part.
		So each selection, you will find/replace "(wavfilename)" with your new file name, "(wavfilename-##of##)"
		Repeat matching audio to lines and find/replace within selection for each half-hour audio file you have.
	When complete, save, select all, and copy.
	Then open wavfilename.transcription.
	Select all and paste.  Make sure there are no blank lines at the bottom.
	Save and close wavfilename.transcription.
	Go back to wavfilename.docx one more time.
	Find and replace, check "use wildcards", find "\)(*{1,})\(", and replace with ")^p("
	That should leave only the wavfilename a bunch of times in a list.
	Select all and copy.
	Open wavfilename.fileids.
	Select all and paste.
	Save and close wavfilename.fileids.
	Without saving, close wavfilename.docx.

---------------------------------------------------------

HOW TO TRAIN AUDIO
NOTE: "S:\training BATs\G-runWAlignforAccuracyReport.bat"

Change "wavfilename" to your wav file name.
Run the following commands in a command window with administrator:

1. cd /D S:\UnprocessedAudio\3

2. S:\sphinxtrain\bin\Release\Win32\sphinx_fe.exe -argfile en-us/feat.params -samprate 16000 -c wavfilename.fileids -di . -do . -ei wav -eo mfc -mswav yes -nfft 2048

3. S:\sphinxtrain\bin\Release\Win32\pocketsphinx_mdef_convert.exe -text en-us/mdef en-us/mdef.txt

4. S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -svspec 0-12/13-25/26-38 -feat 1s_c_d_dd -cmn current -agc none -dictfn wavfilename.dic -ctlfn wavfilename.fileids -lsnfn wavfilename.transcription -accumdir .

#### S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -svspec 0-12/13-25/26-38 -feat 1s_c_d_dd -cmn current -agc none -dictfn cmudict-en-us.dict -ctlfn wavfilename.fileids -lsnfn wavfilename.transcription -accumdir .

5. S:\sphinxtrain\bin\Release\Win32\mllr_solve.exe -meanfn en-us/means -varfn en-us/variances -outmllrfn mllr_matrix -accumdir .

6. S:\sphinxtrain\bin\Release\Win32\map_adapt.exe -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -meanfn en-us/means -varfn en-us/variances -mixwfn en-us/mixture_weights -tmatfn en-us/transition_matrices -accumdir . -mapmeanfn en-us-adapt/means -mapvarfn en-us-adapt/variances -mapmixwfn en-us-adapt/mixture_weights -maptmatfn en-us-adapt/transition_matrices

7. S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl ryan-weed-sample.fileids -lm wavfilename.lm -dict wavfilename.dic -hmm en-us-adapt -hyp wavfilename.hyp

### S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl wavfilename.fileids -lm wavfilename.lm -dict wavfilename.dic -hmm en-us-adapt -hyp wavfilename.hyp

8. S:\sphinxtrain\scripts\decode\word_align.pl wavfilename.transcription wavfilename.hyp >> word_align_output.txt


---------------------------------------------------------
HOW TO TRANSCRIBE FILES AFTER TRAINING

Run the following commands in a command window with administrator:

1. cd /D S:\UnprocessedAudio\3 (or number you're using)

2. S:\pocketsphinx\bin\Release\Win32\pocketsphinx_continuous.exe -infile wavfilename.wav -hmm en-us-adapt -lm wavfilename.lm -dict wavfilename.dic >> full-output.txt

3. Check output in full-output.txt in S:\UnprocessedAudio\##

#### S:\pocketsphinx\bin\Release\Win32\pocketsphinx_continuous.exe -infile sample-5m-wav.wav -hmm en-us-adapt -lm en-us.lm.bin -dict cmudict-en-us.dict >> full-output.txt

#### S:\pocketsphinx\bin\Release\Win32\pocketsphinx_continuous.exe -infile sample-5m-wav.wav -hmm en-us-adapt -lm wavfilename.lm -dict wavfilename.dic >> full-output.txt

---------------------------------------------------------

1. Get your files together
	note wav quality of audio/type
	append transcripts together for all files in one batch in format below
	copy folder Template from S:\UnprocessedAudio\## (latest number)
	rename folder copy new number (1, 2, 3, 4, 5, etc)
	audio wavs only
	copy wavs in S:\UnprocessedAudio\#\wav (new number instead of 'template')
	copy wavs in S:\UnprocessedAudio\#\ (new number instead of 'template')
	place wavname.fileids in S:\UnprocessedAudio\#\
		AudioFileNameNoExtensionOnePerLine
	place wavname.transcription in S:\UnprocessedAudio\#\
		format per line : <s> the sentence goes here without any punctuation </s> (AudioFileNameinParenthesisNoExtension)



2. Ensure the following files are in UnprocessedAudio/#/:
	*.fileids (make same as your wav name)
	*.transcription (make same as your wav name)
	cmudict-en-us.dict
	en-us folder
	en-us-adapt folder
	wav folder
	en-us.lm.bin
	en-us-phone.lm.bin
	init_mixw.exe
	map_adapt.exe
	mk_mllr_class.exe
	mk_s2sendump.exe
	mk_ts2cb.exe
	mllr_solve.exe
	mllr_transform.exe
	pocketsphinx.dll
	pocketsphinx_mdef_convert.exe
	sphinx_fe.exe
	sphinxbase.dll

3. Change the samprate in "S:\training BATs\A-ConvertWAVtoMFC.bat" to sample rate of audio wav.

4. edit all bats in the following folder; change 1 to your number:
	S:\training BATs\

5. run the following bat
	S:\training BATs\AtoZ-allbatchfiles.bat
		which runs all bats you just edited in order
		
-----------------------------------------------------
HOW TO RUN MULTIPLE AUDIO FILES LIST


When running multiple audio files, PRIOR TO ABOVE STEPS:

	Append transcripts of desired audio together into one file in order you are processing audio.  

	Use the following format for transcripts:

		<s> the sentence goes here without any punctuation </s> (AudioFileNameinParenthesisNoExtension)

	Within *.transcription, place audio file name only at last sentence of that file's transcription

	*.FileIDs should be in the following format

		arctic_001
		arctic_002

	Follow rest of steps above starting at #1.
	

