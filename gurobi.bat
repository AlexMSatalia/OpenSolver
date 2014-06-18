@ECHO OFF

if "%GUROBI_HOME%"=="" (
  echo.
  echo Gurobi installer changes have not taken effect yet.
  echo Please restart your machine before continuing.
  echo.
  set /p JUNK= [Hit ENTER to exit]
  exit
)

set PYTHONSTARTUP=%GUROBI_HOME%\lib\gurobi.py

"%GUROBI_HOME%\python27\bin\python" %*
