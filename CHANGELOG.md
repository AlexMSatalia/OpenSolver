# OpenSolver Release Notes

## v2.9.0 - 2017-11-10
### Added
- Add support for Satalia SolveEngine as a new solver.

## v2.8.6 - 2017-03-06
### Changed
- Validate bad inputs for solver options.
- Show iteration limit in status bar while solving.

### Fixed
- Use HTTPS for NEOS connections since HTTP is no longer permitted.
- Fix errors when using Gurobi 7.
- Set Application.Calculation to Automatic to avoid altering user settings.
- Better error messages when no sheet is available.
- Various bug fixes (overflow, error handling, etc.).

## v2.8.5 - 2016-11-03
### Added
- Add support for user-supplied callback macro when using NOMAD
### Fixed
- Workaround for Excel 2016 update 73xx which broke z-order on forms
- Don't show "OpenSolver.xlam is unsaved" message on release copies

## v2.8.4 - 2016-10-12
### Fixed
- Fix crash on 64-bit Office on Windows 7
- Fix error with piping on new versions of 64-bit Office
- Use case-insensitive matching when dealing with solver short names

## v2.8.3 - 2016-10-05
### Added
- Console can now be resized
- Added support for using current solution on sheet as starting point in Gurobi and CBC on NEOS
- Add initial support for Mac Excel 2016 (model building and linear solvers)

### Changed
- Full system information is included in "About OpenSolver" box for easier support
- Starting solution is checked for small infeasibilities and adjusted if possible before passing to solver
- Use current Solver algorithm choice to pick solver if no solver has been chosen
- Error message form is now contextual and shows help links that are relevant to the error

### Fixed
- Fix in non-linear parser for multiple arg SUMPRODUCT
- Speedups for non-linear parsing
- Fix sensitivity analysis when constraints are presolved out
- Persist model dialog options across multiple showings of the form
- Fix error when deleting last constraint in model form
- Use tiled horizontal shapes for wide models in visualizer
- Fix overflow errors in option dialog
- More locale-robust storage of model options
- Fix showing model when parts of sheet are hidden
- Only check update setting once per Excel session
- Better checking for error values in linear model building
- Revert to linearity offset of zero for binary variables
- Warmstarts only loaded if feasible
- Fix appearance of forms on Excel 2016
- Ignore integrality restrictions in feasibility check when solving relaxation
- Fix bugs with 64-bit Office

## v2.8.2 - 2016-02-14
### Fixed
- Another fix for "Automation Error"

## v2.8.1 - 2016-02-11
### Fixed
- Fix "Automation Error"
- Fix error on "Set QuickSolve Parameters" cancel press

## v2.8.0 - 2016-02-01
### Added
- Can use named ranges to define the model
- Support for PRODUCT function in nonlinear solvers
- New console now shows optimisation progress including streaming output from solver and easy cancel
- New menu items: "View error log file" and "View all OpenSolver temp files"
- "About OpenSolver" form includes a check for whether OpenSolver is installed correctly

### Changed
- Ranges don't have to be on the model sheet
- Changes to API
  - Deleted `GetObjectiveFunctionCellWithValidation`
  - `GetObjectiveFunctionCell` validate parameter defaults to `True`
  - Addition of `RefersTo` methods for working with named ranges (and optional `RefersTo` return parameters in Get methods)
  - API methods now validate the inputs and outputs
- Temp files live in a subfolder and are cleaned up when Excel exits
- Revamped menu layout, generated programmatically for Windows and Mac
- New default for linearity check, and better detection of non-linearity, including before solve

### Fixed
- Fix duplicate key error when loading model form
- NL writing to file is buffered for increased speed
- Only load an NL solution if the solver returns solution
- Fix for Mac form bug on Yosemite
- Improvements to Gurobi solution reader (c/o Greg Glockner at Gurobi) for better cloud support
- Improved reading of .sol file for NL solvers
- Fix errors when setting QuickSolve parameters range
- New backend for running external commands, Improvements for piping and termination
- Fix headings on sensitivity report
- Add better suggestion to restart Excel when log file is locked
- Make ConvertLocale a no-op if already in US locale for speed
- Only use numeric RHS values as bounds
- Fix overflow error in CBC solution reader
- Fix variable count in status bar while solving
- Fix localisation error in update checker
- Better escape handling in all solvers
- Fix garbage in error reports on Mac
- Fix ordering of variables in LP file to Make sensitivity table correct
- Fix focus issue with error dialog
- Raise error on non-numeric constants in parser
- Refactor COpenSolver to remove redundant code and better separate difference model into CModelDiff
- Raise error when sheet has been Deleted in reference
- Fix /amplin error on NEOS
- Increase time between NEOS checks to support changes made by NEOS server-side
- Model form doesn't delete missing references on load
- Fix RefEdit focus bugs on model form
- backend changes to model form to handle named ranges as model parameters
- Tidyup of names storage backend

## v2.7.1 - 2015-06-28
### Fixed
- Various bugfixes

## v2.7.0 - 2015-06-16
### Added
- Update Checker
- OpenSolver API VBA interface
- New error reporting with email reports
- All solvers support extra parameters
- Experimental support for NOMAD on Mac

### Changed
- Update solvers to CBC 2.9.4, Bonmin 1.8.1, Couenne 0.5.3, NOMAD 3.7.2

### Fixed
Various bugfixes (including lots of locale issues)

## v2.6.1 - 2015-02-16
### Added
- Extra parameter support for Gurobi (see http://opensolver.org/using-opensolver/#gurobi-params)
- Abort NEOS solve if Cancel button is clicked
- Bonmin, Couenne and Gurobi support Solver Options (e.g. time limit, tolerance)
- Make the model dialogue box resizable.
- Support for more formulae in non-linear solvers (e.g. SUMIF)

### Changed
- Add better error message for when non-linear solvers create no solution file
- Misc improvements in non-linear parsing.
- Update solvers to CBC 2.9.0, Bonmin 1.8.0 and Couenne 0.5.1
- Better logging and 'Show Optimisation Progress' support. Now works on Mac
- Expand check for `solver_rlx` to all new Excel versions.
- Better error message if NOMAD is missing

### Fixed
- Fix issues with sensitivity analysis and missing variables
- Support for running from multiple disks and network drives in OS X
- Fix for misc OS X Yosemite issues
- Bugfixes for locale-based issues
- Implement proper sheet name escaping for the formula parser.


## v2.6.0 - 2014-10-08
### Added
- Add support for Office 2011 on Mac - nearly all features supported
- Add support for local COIN-OR non-linear solvers
- NEOS solvers write AMPL files to disk before sending to NEOS. These can be used to run the model locally

### Changed
- Upgrade NOMAD to v3.6.2

### Fixed
- Resolve bugs introduced by system-locale settings
- Bugfixes for non-linear NOMAD solver
- Bugfixes for NEOS solvers

## v2.5.4-alpha - 2014-07-03
### Fixed
- NOMAD bug fixes for when errors are encountered

## v2.5.3-alpha - 2014-07-02
### Added
- Add support for NOMAD in 64-bit Office.

## v2.5.2-alpha - 2014-06-27
### Fixed
- Fix memory bug causing Excel 2013 to crash when using NOMAD
- Non-linear NEOS bug fixes

## v2.5.1-alpha - 2014-06-25
### Added
- Inclusion of 64-bit CBC with release - appropriate version is selected automatically
- Re-add NEOS non-linear solvers to release with lots of bug fixes.

### Fixed
- Stability fixes for NOMAD non-linear solver
- Bug fixes for sensitivity analysis methods.

## v2.5.0-alpha - 2014-06-20
### Added
- Support for using the Gurobi LP/IP solver if a user has this installed on their machine
- Support for cloud-based NEOS server for CBC solver
- Support for solving non-linear models using both NOMAD and the cloud-based NEOS servers (assuming non-negativity currently doesn't work correctly for non-linear NEOS, all variables are assumed positive, not just unconstrained ones)
- Reporting of dual variables and sensitivity analysis

### Changed
- Updated CBC.exe to version 2.8.8

### Fixed
- Many small bux fixes and feature enhancements

## v2.4.0-beta - 2013-02-19
### Added
- The model window has a 'Clear Model' button that clears all the information in the model. Unfinished models can also be saved now.
- Changed the model dialogue so you can now save uncompleted models
- OpenSolver can now handle models with no objective
- Has the option of solving with Gurobi rather then CBC. Individual read CBC or gurobi solved models
- Can write sensitivity analysis either on the same sheet (with shadow price/reduced cost, possible increase, possible decrease) or on a new sheet like solver does. Choice can be made in the model dialogue. Can also choose to overwrite old sensitivity sheets or to make new ones
- Uses gurobi.bat from python and can get the sensitivity information for gurobi. Writes it to the same tables that are used for CBC
- The model window can now show the name of a range if it has been defined in Defined Names and there is the option to turn this on and off ( thanks Andres Sommerhoff for your help adding in this functionality)
- Added beta functionality for a non-linear blackbox solver (NOMAD) which uses the models saved in solver and OpenSolver
- Extra options in the menu: to AutoModel and Solve in one, to change solver, to view NOMAD log file, view gurobi solution file
- Model Dialogue: New option to change Solver, New Options under sensitivity analysis
- QuickSolve Example worksheet

### Changed
- The AutoModel now groups adjacent constraints of the same sense rather then an individual constraint for every single one
- Constraints can now refer to a column LHS and a row RHS (or vice versa)
- Can put multiple bounds on variables and it writes them all in the cell rather then using multiple boxes on top of each other. For example 7≤,5≥.
- OpenSolver now uses the environment variable "OpenSolverTempPath" as the path to save files to if this has been defined by the user (thanks again Andres Sommerhoff)

### Fixed
- Fixed bug for drawing the model with large rectangles that were over the excel limit
- Trap bad numeric number errors

## v2.3-beta - 2012-10-29
### Added
Added support for 64 bit versions of the COIN-OR CBC solver, allowing bigger models to be solved without CBC failing

### Fixed
- Fixed bugs for sheet names with spaces (a few) and with a single quote (which caused OpenSolver to fail with an error msg); thanks to Fenny for this bug report.

## v2.2-beta - 2012-09-24
### Changed
- We now only look in the same folder as OpenSolver.xlam for the CBC file, but now look first for the 64-bit cbc64.exe if it exists and the systems is 64 bit

### Fixed
- Fixed a minor bug leaving Excel in manual calculation mode if cbc.exe was missing

## v2.1-beta - 2012-09-05
### Fixed
- Added better handling of non-US systems, in particular decision variables with multiple areas in their range. (Thanks to Brenhard Aeschbacher for pointing out this bug on his German system.)

## v2.0-beta - 2012-02-24
### Added
- Added an option to let the user turn off linearity checking

### Changed
- Updated CBC.exe to version 2.7.6
- Changed from displaying inequalities as < and > to ≤ and ≥ on screen, and <= and >= in message dialogs

### Fixed
- Improved the linearity checking to work around coefficients of many different magnitudes in which case numerical rounding can cause problems

## v1.9-beta - 2011-12-05
### Fixed
- Fixed some display issues in the model dialog to give more compact displays, and force formulae RHS to be in absolute terms
- Fixed a bug that saw models being wrongly built and then reported as non-linear (incorrectly) for some complex models if Excel calculation mode was set to manual.
- Fixed passing of multiple parameters to CBC if the user defines a parameter table on the sheet; previously only the last parameter was passed.
- Deleted the DLL's downloaded in the .zip file which don't seem to be needed now that we statically link everything in CBC.

## v1.8-beta - 30/11/2011
### Added
- Updated OpenSolver to properly handle models with an objective and constraints on multiple sheets

### Fixed
- Fixed a bug stopping OpenSolver loading on 64 bit systems

## v1.7-beta - 2011-11-11
### Added
- Added controls in the About box to allow easy installation and uninstallation
- Added code to interact nicely with the forthcoming OpenSolver Studio
- Improved OpenSolver for use from VBA:
  - Build and Solve operations now throw errors (instead of popping up dialogs), allowing dialog-free usage from VBA
  - Return codes are better handled (and Solver compatible)
  - A new optional parameter has been added to RunOpenSolver to avoid dialogs even if infeasible/unbounded solutions are generated

## v1.6-beta - 2011-09-29
### Added
- Modified the Open Last Model in CBC functionality so that it passes any Solver options and any CBC solve parameters to CBC if they are available in any current worksheet
- Added "Show optimisation progress while solving"  (being Solver's "Show Iteration Results") to the OpenSolver options dialog
- Added output of dual prices onto the sheet; this is set using the Model dialog

### Changed
- Rearranged Model dialog to better fit new Duals option, and better use space around constraint listing

### Fixed
- Fixed display and editting of an objective target value in the Model dialog.
- Fixed a minor issue in Model dialog where a RHS could be entered for a new constraint if the user had previously had a Bin or Int constraint selected
- Fixed a redim bug in the quick non-linearity checker for models with no constraints (which can happen if there is only a target objective value)
- Improved operation of Options dialog, including proper sycnronisation of values when opened from the Model dialog
- Better handling of the Excel 2010 "Simplex engine" option as used in parallel with "Assume linear model"
- Fixed an error in the full non-linearity checker
- Better handling of the Excel solver options - OpenSolver now sets all these to sensible defaults
- Better handling of users entering formulae in the Model dialog for a constraint RHS in terms of non-English localisation issues, but this still needs work
- Fixed a size limitation in Quick Solve, and converted Quick Solve to sparse matrix handling for better memory usage.

## v1.5-beta - 2011-08-09
### Changed
- Recompiled CBC with static linking to libraries so it works on machines without Visual Studio 2010
- We now pass the solve options (such as tolerance) to CBC both when solving the problem, and when opening the last model in CBC. This is useful for checking the CBC arguments.

### Fixed
- Fixed an issue where an objective or decision cells formatted as "currency" or "date" caused errors; we now use .Value2 (not .Value) to get cell values.
Properly pass Tolerance and RatioGap to CBC (in English) on internationalised systems

## v1.4-beta - 2011-07-31
### Added
- Added partial locale support. Entering number like 180,2 will work, but will display as 180.2
- Added an Options window, available under Model in menu, and from a button on the Model form.
- Added different locale support to Options window.
- Added error catching in Solve for int/bin constraints on non-decision variable cells.
- Added error catching in Visualiser too.
- Added warning message to Model tool.
- Added custom icons for toolbar.

### Changed
- Updated CBC to version 2.7
- Debug.Prints commented out

### Fixed
- Fixed 2003 menus
- Fixed the issue with .HorizontalAlignment in 2003: http://www.officekb.com/Uwe/Forum.aspx/excel-prog/159706/Shape-TextEffect-HorizontalAlignment-throws-error
- Quick Auto Model with no spreadsheet open doesn't crash.
- AutoModel window doesn't show if AutoModel works.
- no text colouring in AutoModel.
- If a Model is showing, and the user does a QuickAutomodel, but does not choose to show the Model, then the current Model display is hidden.
- Clear the status bar after the AutoModel tool changes it.
- Fixed an edge case with AutoModel relating to double-tracking cells.
- Fixed a ref-edit related focus bug in Model tool that was causing some strange behaviour.
    Uses spreadsheet internal to the OpenSolver add-in and Range.FormulaLocal to do a conversion.



## v1.3-beta - 2011-07-07
### Added
- Now completely independent of Solver - GUI for building models created.
- Menus for Excel 2003 added
- Excel 64-bit support

### Fixed
- Various bug fixes

## v1.2-beta - 2011-03-08
### Changed
- OpenSolver now treats empty cells as containing the value zero, which mimics Solver's approach

## v1.1-beta - 2011-03-04
### Added
- Support for much larger problemns
  - Sparse A matrix handling
  - Much faster data transfer between VBA and Excel
  - Using Ranges in the VBA instead of arrays for efficiency
  - New model display routines to handle large problems
  - Handling of Excel calculations that didn't complete; we observed these on large models (70,000 variables)
  - We have a problem submitted by a user with over 70,000 variables and 70,000 constraints that we can now solve (although it takes hours!)
- Correct Handling of Assume Non Negative
  - Models no longer require "Assume Non-Negative" to be turned on
  - If it is, then we only add 0 lower bounds to variables that do not have an explicit lower bound set in Solver
  - Note: Excel 2010 and 2007 seem to handle this differently when a single range includes decision variables and other cells; we follow the 2007 approach
- Support for Excel 2010
  - We recognise choosing the Simplex engine as being equivalent to Assume Linear
  - We present version-specific dialogs to easily turn on one of these options
- Reporting of Infeasible solutions
  - If CBC reports the solution as infeasible, we load it in anyway to show the user
- Auto Model (by Engineering Science student Iain Dunning)
  - Added an AutoModel feature to build Solver models more easily.
  - Added improved detection of decision variables when the obj fn cell has dependents
- Non-Linearity Checker
  - We now check that the solution given by CBC gives the expected LHS and objective function values when loaded into the s/sheet
  - We also provide a more extensive nonlinearity check that the user cxan run if the model appears to be non-linear; this can highlight non-linearities on the model
- Support for Models with Constraints on Other Sheets
  - Our View Model will now show constraints on sheets other than the active sheet
- Formulae in the Right Hand Side
  - OpenSolver can how handle a constraint with a formula (such as "=2*B1") entered as the right hand side
- Test Problem Bank. We now have a suite of test problems that we use for testing OpenSolver.

### Changed
- Better Model and Range Checking
  - Excel allows ranges (eg for the decsion variables) that count individual cells multiple times; we now 'fix' such ranges
  - We check for merged cells in the decision variables, and handle them correctly (allowing them if possible; Solver doesn't allow any)
  - We have improved our model error reporting (which we now think is more useful than Solver's)
  - We check for constraints that don't vary with the decision variables; if we find them, we check that they are satisfied, and if not report this explicitly to the user
  - Better checking of s/sheets that contain errors in model cells.
  - We now require all cells in constraints to contain numeric values. For example, a blank RHS gives an error (even tho' Solver allows this, but sometimes puts in zero's)

### Fixed
- OpenSolver used to crash when checking constraints that did not vary with the decision variables; this has been fixed

## v0.982 - 2010-08-17
### Added
- Handle larger problems with more than 32,000 variables and/or constraints. However, such models will be very slow to build

## v0.98 - 2010-07-16
### Fixed
- Bug fixes associated with quick solves (one GUI related, one that fixes the handling of multi-area ranges, and checks that the user is on the same sheet and workbook as that used to initialise the quick solve)
- Improvements so that OpenSolver dynamically resizes its arrays to handle large problems (assuming everything fits in memory).

## v0.95 - 2010-06-06
### Added
- Added new commands to (1) Solve LP relaxation, and (2) to open CBC command line. Also added improved support for cancelling long CBC runs.
-
### Fixed
- Better checking of parameters
- Better handling of Escape during long CBC runs (no DoEvents now, and a new dialog).
- Fixed bug in the “Last open model in CBC” where OpenSolver was waiting for CBC to close (but still allowing events to be handled, including sheet edits etc.)
- Added a fix for sheet names with spaces, and for the definition of parameters.

## Initial Version - 2010-05-17
Our first public release.
