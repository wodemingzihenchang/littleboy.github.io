set FMWK="v4.0.30319"

CD "%~dp0"

IF EXIST "%Windir%\Microsoft.NET\Framework64\%FMWK%\regasm.exe" "%Windir%\Microsoft.NET\Framework64\%FMWK%\regasm" /codebase "./bin/x64/Release/SWVizAPISample.dll" /unregister /tlb

pause