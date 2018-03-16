@echo off
ECHO please select test program:
ECHO 1.test_exec.pl
ECHO 2.test_exec.py
ECHO 3.test_exec.rb
choice /c 123

if %errorlevel% == 1 goto 1:
if %errorlevel% == 2 goto 2:
if %errorlevel% == 3 goto 3:
goto exit:

:1
perl test_exec.pl
goto exit:
:2
python test_exec.py
goto exit:
:3
ruby test_exec.rb

:exit
pause
