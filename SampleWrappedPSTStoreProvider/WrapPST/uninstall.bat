rundll32 %windir%\system32\wrppst32.dll RemoveFromMAPISVC
del  %windir%\system32\wrppst32.dll
del  %windir%\system32\wrppst32.pdb
del "%ProgramFiles%\Microsoft Office\OFFICE11\ADDINS"\wrappst.ecf 
del "%ProgramFiles%\Microsoft Office\OFFICE12\ADDINS"\wrappst.ecf 
del "%ProgramFiles(x86)%\Microsoft Office\OFFICE12\ADDINS"\wrappst.ecf 
regedit /s ReloadExtensions.reg
