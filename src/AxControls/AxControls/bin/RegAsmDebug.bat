@echo off
setlocal
%SystemRoot%\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe "C:\Users\MareninDL\Documents\Visual Studio 2015\Projects\AxControls\AxControls\bin\Debug\AxControls.dll" /codebase /tlb

%SystemRoot%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe "C:\Users\MareninDL\Documents\Visual Studio 2015\Projects\AxControls\AxControls\bin\Debug\AxControls.dll" /codebase /tlb


@pause