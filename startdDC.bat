@ECHO OFF
setlocal

SET CURRENT=%~pd0

SET DATESTAMP=%DATE:/=%
SET TMPV=%TIME: =0%
SET TMPV=%TMPV::=%
SET TIMESTAMP=%DATESTAMP%_%TMPV:.=%

SET OUTFILE=%CURRENT%tmp\out_%TIMESTAMP%.txt

ECHO START>"%OUTFILE%"
START "" "%OUTFILE%"
ECHO .|CLIP


Cscript %CURRENT%script\dumpClipbord.vbs "%OUTFILE%"

EXIT /B



