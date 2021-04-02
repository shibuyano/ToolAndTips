@echo off
echo "Change audio file type to mp3"
set exe="C:\usr\ffmpeg\bin\ffmpeg.exe"

set tmpinfile="C:\usr\tmp_input%~x1"
set tmpoutfile="C:\usr\tmp_output.mp3"

REM フルパス
REM set outputfile="%~d1%~p1%~n1_encrypted%~x1"
set outputfile="%~d1%~p1%~n1_toEnc.mp3"


copy %1 %tmpinfile%

REM エンコード
%exe% -y -i %tmpinfile% %tmpoutfile%

copy %tmpoutfile% %outputfile%
del %tmpinfile% >>NUL
del %tmpoutfile% >>NUL

echo %1
move /Y %1 finished

echo "End of process"
