#!/bin/sh
TargetDir="$1"
cd ../../..
mkdir -p drop/samples
cp -Rv Updates.* "$TargetDir"
cp -Rv Updates.* "drop"
cp -Rv Readme.* "drop"
cp -Rv Help/Help.chm "drop"
cp -Rv Help/Images/xmlicon.* "drop"
# cp -Rv Samples/*.* "drop/samples"
# cp -Rv "XML NotePad 2007 EULA.rtf" "drop"
cp -Rv "$TargetDir"*.* "drop"
exit 0
