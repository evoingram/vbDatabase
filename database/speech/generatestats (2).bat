`

S:\UnprocessedAudio\1\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -feat 1s_c_d_dd -svspec 0-12/13-25/26-38 -cmn current -agc none -dictfn cmudict-en-us.dict -ctlfn sample-5m-wav.fileids -lsnfn sample-5m-wav.transcription -accumdir accumdir

S:\UnprocessedAudio\1\mllr_solve.exe -meanfn en-us/means -varfn en-us/variances -outmllrfn 'mllr_matrix' -accumdir

S:\sphinxtrain\bin\Release\Win32\mllr_solve.exe -meanfn en-us/means -varfn en-us/variances -outmllrfn 'mllr_matrix' -accumdir

S:\UnprocessedAudio\1\map_adapt.exe -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -meanfn en-us/means -varfn en-us/variances -mixwfn en-us/mixture_weights -tmatfn en-us/transition_matrices -accumdir -mapmeanfn en-us-adapt/means -mapvarfn en-us-adapt/variances -mapmixwfn en-us-adapt/mixture_weights -maptmatfn en-us-adapt/transition_matrices

--------------------------------------------

S:\sphinxtrain\bin\Release\Win32\sphinx_fe.exe -argfile en-us/feat.params -samprate 44000 -c sample-5m-wav.fileids -di . -do . -ei wav -eo mfc -mswav yes -nfft 2048

pocketsphinx_mdef_convert -text en-us/mdef en-us/mdef.txt

S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -feat 1s_c_d_dd -cmn current -agc none -dictfn cmudict-en-us.dict -ctlfn sample-5m-wav.fileids -lsnfn sample-5m-wav.transcription -accumdir

S:\sphinxtrain\bin\Release\Win32\bw.exe -hmmdir en-us -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -svspec 0-12/13-25/26-38 -feat 1s_c_d_dd -cmn current -agc none -dictfn cmudict-en-us.dict -ctlfn sample-5m-wav.fileids -lsnfn sample-5m-wav.transcription -accumdir .

S:\sphinxtrain\bin\Release\Win32\mllr_solve.exe -meanfn en-us/means -varfn en-us/variances -outmllrfn 'mllr_matrix' -accumdir .

S:\sphinxtrain\bin\Release\Win32\map_adapt.exe -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -meanfn en-us/means -varfn en-us/variances -mixwfn en-us/mixture_weights -tmatfn en-us/transition_matrices -accumdir -mapmeanfn en-us-adapt/means -mapvarfn en-us-adapt/variances -mapmixwfn en-us-adapt/mixture_weights -maptmatfn en-us-adapt/transition_matrices

S:\sphinxtrain\bin\Release\Win32\map_adapt.exe -moddeffn en-us/mdef.txt -ts2cbfn .ptm. -meanfn en-us/means -varfn en-us/variances -mixwfn en-us/mixture_weights -tmatfn en-us/transition_matrices -accumdir . -mapmeanfn en-us-adapt/means -mapvarfn en-us-adapt/variances -mapmixwfn en-us-adapt/mixture_weights -maptmatfn en-us-adapt/transition_matrices


S:\sphinxtrain\bin\Release\Win32\

-mllr mllr_matrix 

WARN: "accum.c", line 628: Over 500 senones never occur in the input data. This is normal for context-dependent untied senone training or for adaptation, but could indicate a serious problem otherwise.
ERROR: "s3io.c", line 277: Unable to open accumdir/mixw_counts for writing: No such file or directory
ERROR: "accum.c", line 733: Couldn't revert to backup of accumdir/mixw_counts

S:\sphinxtrain\scripts\decode\word_align.pl sample-5m-wav.transcription test.hyp >> word_align_output.txt


bin/Release/pocketsphinx_continuous.exe -inmic yes -lm 8521.lm -dict 8521.dic -hmm model/en-us/en-us

cd S:\UnprocessedAudio\1
S:\pocketsphinx\bin\Release\Win32\pocketsphinx_continuous.exe -infile sample-5m-wav.wav -hmm en-us-adapt -lm sample-5m-wav.lm -dict sample-5m-wav.dic >> full-output.txt


pocketsphinx_continuous -infile <your_file.wav> -keyphrase <your keyphrase> -kws_threshold <your_threshold> -time yes

pocketsphinx_continuous -hmm `<your_new_model_folder>` -lm `<your_lm>` -dict `<your_dict>` -infile test.wav

S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -hmm en-us-adapt -lm en-us.lm.bin -dict cmudict-en-us.dict -infile test.wav

S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl sample-5m-wav.fileids -lm en-us.lm.bin -dict cmudict-en-us.dict -hmm en-us-adapt -hyp test.hyp -samprate 8000

S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepdir wav -cepext .wav -ctl sample-5m-wav.fileids -lm en-us.lm.bin -dict cmudict-en-us.dict -hmm en-us-adapt -hyp test.hyp

S:\sphinxtrain\bin\Release\Win32\pocketsphinx_batch.exe -adcin yes -cepext .wav -ctl sample-5m-wav.fileids -lm en-us.lm.bin -dict cmudict-en-us.dict -hmm en-us-adapt -hyp test.hyp

S:\sphinxtrain\scripts\decode\word_align.pl sample-5m-wav.transcription test.hyp >> word_align_output.txt


S:\sphinxtrain\scripts\decode\


cd /D S:\UnprocessedAudio\1

-----------------------------------------------------------------------
RANDOM NOTES

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" -t "\\HUBCLOUD\evoingram\In Progress\speech recognition\CMUSphinx\sa1" setup

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" "\\HUBCLOUD\evoingram\In Progress\speech recognition\CMUSphinx\sphinxtrain\scripts\sphinxtrain" -t "\\HUBCLOUD\evoingram\speech\sa1" setup

python ../sphinxtrain/scripts/sphinxtrain -t an4 setup

python ../sphinxtrain/scripts/sphinxtrain run

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" "\\HUBCLOUD\evoingram\speech\sphinxtrain\scripts\sphinxtrain.py" -t sa1 setup

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" "\\HUBCLOUD\evoingram\speech\sphinxtrain\scripts\sphinxtrain.py" run

"//HUBCLOUD/evoingram/speech/CMUSphinx/sphinxtrain/python/dist"

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" "\\HUBCLOUD\evoingram\speech\sphinxtrain\scripts\sphinxtrain.py" run

"C:/Program Files (x86)/Microsoft Visual Studio\Shared\Python36_64\python.exe" "//HUBCLOUD/evoingram/speech//sphinxtrain/scripts/sphinxtrain" -t "//HUBCLOUD/evoingram/speech/sa1" setup

---------------------------
cd /D S:\sa1

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" \\HUBCLOUD\evoingram\speech\sphinxtrain\scripts\sphinxtrain.py -t sa1 setup

"C:\Program Files (x86)\Microsoft Visual Studio\Shared\Python36_64\python.exe" "\\HUBCLOUD\evoingram\speech\sphinxtrain\scripts\sphinxtrain.py" run

export SPHINXBASE="\\HUBCLOUD\evoingram\In Progress\speech recognition\CMUSphinx\sphinxbase" # change this to your sphinxbase  source tree
export PYTHONPATH=$SPHINXBASE/swig/python/build/lib.*
export LD_LIBRARY_PATH=$SPHINXBASE/src/libsphinxbase/.libs

"//HUBCLOUD/evoingram/speech/CMUSphinx/sphinxtrain/scripts/sphinxtrain" -t sa1 setup

_test.fileid

$CFG_LIST_DIR



import os

os.chdir(path)


S:\pocketsphinx\bin\Release\Win32\sphinx_fe.exe -argfile S:\pocketsphinx\model\en-us\en-us\feat.params \
        -samprate 16000 -c sample-5m-wav.fileids \
       -di . -do . -ei wav -eo mfc -mswav yes

S:\pocketsphinx\bin\Release\Win32\sphinx_fe.exe -argfile //HUBCLOUD/evoingram/speech/pocketsphinx/model/en-us/en-us/feat.params -samprate 16000 -c sample-5m-wav.fileids -di . -do . -ei wav -eo mfc -mswav yes -nfft 2048











./bw \
 -hmmdir en-us \
 -moddeffn en-us/mdef.txt \
 -ts2cbfn .ptm. \
 -feat 1s_c_d_dd \
 -svspec 0-12/13-25/26-38 \
 -cmn current \
 -agc none \
 -dictfn cmudict-en-us.dict \
 -ctlfn sample-5m-wav.fileids \
 -lsnfn sample-5m-wav.transcription \
 -accumdir .













