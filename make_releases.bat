:: Makes all releases for OpenSolver
:: Call using: make_releases.bat <version number>
:: Needs 7-zip installed and the 7-zip folder on the system path
@echo off

:: Empty the Release folder
del Release\*

:: Get version number for release from the first parameter
set version=%1

:: Common elements of the 7-zip command
set start=7z a Release\
set end=. -xr!.git* -xr!OpenSolver.xlam.src -xr!Release -x!make_releases.*

:: All files to exclude in the Windows and Mac Releases
set windows=-xr!Solvers\osx
set osx=-xr!Solvers\win32 -xr!Solvers\win64

:: All files to be excluded from the linear release
set linear=-xr!*bonmin* -xr!*couenne* -xr!*libipoptfort* -xr!*NOMAD* -xr!*Nomad*

:: MAKE COMMANDS

:: Windows Linear
%start%OpenSolver%version%_LinearWin.zip %end% %windows% %linear%

:: Windows Advanced
%start%OpenSolver%version%_AdvancedWin.zip %end% %windows%

:: Mac Linear
%start%OpenSolver%version%_LinearMac.zip %end% %osx% %linear%

:: Mac Linear
%start%OpenSolver%version%_AdvancedMac.zip %end% %osx%

