if "%programfiles(x86)%XXX"=="XXX" goto 32BIT
echo 64-bit Windows installed
copy mapisvc.inf "%programfiles(x86)%\Common Files\System\MSMAPI\1033" /y
copy Debug\sabp32.dll "%programfiles(x86)%\Microsoft Office\Office12\" /y
copy Debug\sabp32.pdb "%programfiles(x86)%\Microsoft Office\Office12\" /y
goto END

:32BIT
echo 32-bit Windows installed
copy mapisvc.inf "%programfiles%\Common Files\System\MSMAPI\1033" /y
copy Debug\sabp32.dll "%programfiles%\Microsoft Office\Office12\" /y
copy Debug\sabp32.pdb "%programfiles%\Microsoft Office\Office12\" /y

:END
