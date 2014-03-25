@echo off
set REGASM=C:\Windows\Microsoft.NET\Framework\v2.0.50727\RegAsm.exe
REM カレントディレクトリをバッチファイルのディレクトリにする
cd /d %~dp0
REM %REGASM% bin\Release\ClassLibraryForVBA.dll /unregister /verbose /tlb:bin\release\ClassLibraryForVBA.tlb
REM %REGASM% bin\release\ClassLibraryForVBA.dll /verbose /tlb:bin\release\ClassLibraryForVBA.tlb
%REGASM% /codebase bin\release\ClassLibraryForVBA.dll
pause
