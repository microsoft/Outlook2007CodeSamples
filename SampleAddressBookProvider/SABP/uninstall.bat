if "%programfiles(x86)%XXX"=="XXX" goto 32BIT
echo 64-bit Windows installed
del "%programfiles(x86)%\Microsoft Office\Office12\sabp32.dll"
del "%programfiles(x86)%\Microsoft Office\Office12\sabp32.pdb"
del "%programfiles(x86)%\Common Files\System\MSMAPI\1033\mapisvc.inf"
goto END

:32BIT
echo 32-bit Windows installed
del "%programfiles%\Microsoft Office\Office12\sabp32.dll"
del "%programfiles%\Microsoft Office\Office12\sabp32.pdb"
del "%programfiles%\Common Files\System\MSMAPI\1033\mapisvc.inf"

:END
