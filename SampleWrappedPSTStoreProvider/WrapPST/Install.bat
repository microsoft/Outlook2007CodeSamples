if "%programfiles(x86)%XXX"=="XXX" goto 32BIT
echo 64-bit Windows installed
copy Debug\wrppst32.dll %windir%\syswow64
copy Debug\wrppst32.pdb %windir%\syswow64
copy wrappst.ecf "%ProgramFiles(x86)%\Microsoft Office\OFFICE11\ADDINS"
copy wrappst.ecf "%ProgramFiles(x86)%\Microsoft Office\OFFICE12\ADDINS"
regedit /s ReloadExtensions.reg
rundll32 %windir%\syswow64\wrppst32.dll MergeWithMAPISVC
goto END

:32BIT
echo 32-bit Windows installed
copy Debug\wrppst32.dll %windir%\system32
copy Debug\wrppst32.pdb %windir%\system32
copy wrappst.ecf "%ProgramFiles%\Microsoft Office\OFFICE11\ADDINS"
copy wrappst.ecf "%ProgramFiles%\Microsoft Office\OFFICE12\ADDINS"
regedit /s ReloadExtensions.reg
rundll32 %windir%\system32\wrppst32.dll MergeWithMAPISVC

:END

