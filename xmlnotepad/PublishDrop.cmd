set TargetDir=%1
PUSHD "%~dp0"

if not exist  "drop" mkdir  "drop"
if not exist  "drop\samples" mkdir  "drop\samples"

xcopy /d /y "Updates.*" "%TargetDir%"
xcopy /d /y "Updates.*" "drop"
xcopy /d /y "Readme.*" "drop"
xcopy /d /y "Help\Help.chm" "drop"
xcopy /d /y "Help\Images\xmlicon.*" "drop"
xcopy /d /y "Samples\*.*" "drop\samples"
xcopy /d /y "XML NotePad 2007 EULA.rtf" "drop"
xcopy /d /y "%TargetDir%*.*" "drop"