copy .\unicode_debug\mrxp32.dll %windir%\system32
copy .\unicode_debug\mrxp32.dll %windir%\sysWow64
%windir%\system32\rundll32.exe %windir%\system32\mrxp32.dll MergeWithMAPISVC