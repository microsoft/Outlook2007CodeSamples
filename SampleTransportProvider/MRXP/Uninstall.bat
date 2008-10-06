if "%programfiles(x86)%XXX"=="XXX" goto 32BIT
echo 64-bit Windows installed
%windir%\syswow64\rundll32.exe %windir%\syswow64\mrxp32.dll RemoveFromMAPISVC
del %windir%\sysWow64\mrxp32.dll 
goto END

:32BIT
echo 32-bit Windows installed
%windir%\system32\rundll32.exe %windir%\system32\mrxp32.dll RemoveFromMAPISVC
del %windir%\system32\mrxp32.dll 

:END

