copy Debug\wrppst32.dll %windir%\system32
copy debug\wrppst32.pdb %windir%\system32
copy Debug\wrppst32.dll %windir%\system
copy debug\wrppst32.pdb %windir%\system
copy wrappst.ecf "%ProgramFiles%\Microsoft Office\OFFICE11\ADDINS"
copy wrappst.ecf "%ProgramFiles%\Microsoft Office\OFFICE12\ADDINS"
copy wrappst.ecf "%ProgramFiles(x86)%\Microsoft Office\OFFICE12\ADDINS"
regedit /s ReloadExtensions.reg
rundll32 %windir%\system32\wrppst32.dll MergeWithMAPISVC