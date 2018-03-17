@echo off
:loop
ECHO please select test program:
ECHO 1.test_exec.pl
ECHO 2.test_exec.py
ECHO 3.test_exec.rb
ECHO 4.exit 
choice /c 1234

if %errorlevel% == 1 goto 1:
if %errorlevel% == 2 goto 2:
if %errorlevel% == 3 goto 3:
if %errorlevel% == 4 goto exit:
goto loop:

:1
perl test_exec.pl
goto loop:
:2
python test_exec.py
goto loop:
:3
ruby test_exec.rb
goto loop:

:exit
pause
