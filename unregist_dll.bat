@echo off
set REGASM=C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\regasm.exe
REM カレントディレクトリをバッチファイルのディレクトリにする
cd /d %~dp0
%REGASM% bin\Release\ClassLibraryForVBA.dll /unregister /verbose /tlb:bin\release\ClassLibraryForVBA.tlb
REM %REGASM% bin\release\ClassLibraryForVBA.dll /verbose /tlb:bin\release\ClassLibraryForVBA.tlb
REM %REGASM% /codebase bin\release\ClassLibraryForVBA.dll
pause
