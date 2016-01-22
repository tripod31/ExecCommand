@echo off
set REGASM=C:\WINDOWS\Microsoft.NET\Framework\v4.0.30319\regasm.exe
REM カレントディレクトリをバッチファイルのディレクトリにする。
cd /d %~dp0
%REGASM% bin\Release\ExecCommand.dll /unregister /verbose
pause
