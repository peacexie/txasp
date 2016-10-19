

regsvr32 /u ZipCOM.dll
del %windir%\system\ZipCOM.dll


copy ZipCOM.dll %windir%\system\
cd %windir%\system\
regsvr32 ZipCOM.dll


echo.
echo Register ZipCOM OK!!!!!
echo Bat File by Peace
cmd


