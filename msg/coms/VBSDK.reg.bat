
regsvr32 /u VBSDK.dll
del %windir%\system\VBSDK.dll


copy VBSDK.dll %windir%\system\
cd %windir%\system\
regsvr32 VBSDK.dll


echo.
echo Register VBSDK OK!!!!!
echo Bat File by Peace
cmd



