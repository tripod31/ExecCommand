@echo off
set REGASM=C:\Windows\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe

REM カレントディレクトリをバッチファイルのディレクトリにする。
cd /d %~dp0
%REGASM% /codebase bin\release\ExecCommand.dll
pause
