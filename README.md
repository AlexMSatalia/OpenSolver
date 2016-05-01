# OpenSolver

http://www.opensolver.org

OpenSolver is an Excel add-in that extends Excel’s built-in Solver with a more powerful Linear Programming solver. OpenSolver provides the following features:

* OpenSolver uses the excellent, Open Source, COIN-OR CBC optimization engine, to quickly solve large Linear and Integer problems.
* Compatible with your existing Solver models, so there is no need to change your spreadsheets
* No artificial limits on the size of problem you can solve
* OpenSolver is free, open source software licensed under the GPL.

As well as providing a replacement optimization engine, OpenSolver offers:

* A built-in model visualizer that highlights your model’s decision variables, objective and constraints directly on your spreadsheet
* A fast QuickSolve mode that makes it much faster to re-solve your model after making changes to a right hand side
* An optional model building tool that analyses your spreadsheet, and then fills in the Solver dialog automatically

OpenSolver has been developed for Excel 2007 and 2010 running on Windows. It should work with these or later Excel versions.

OpenSolver is being developed by Andrew Mason in the Department of Engineering Science at the University of Auckland. 

OpenSolver is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version.  License copyright years may be listed using range notation (e.g. 2011-2016) indicating that every year in the range, inclusive, is a copyrightable year that would otherwise be listed individually.

OpenSolver uses the open source COIN-OR CBC optimization engine. CBC is released as open source code under the Eclipse Public License (EPL). It is available from the COIN-OR initiative. The CBC code has been written by primarily by John J. Forrest, and is maintained by Ted Ralphs.

The COIN-OR COUENNE and BONMIN solvers are also included. These are also released under the Eclipse Public License (EPL).

OpenSolver also uses the NOMAD non-linear solver which is released under the GNU Lesser General Public License.

Please see the included license files for more details.


## Installation

1. Download the `OpenSolver.zip file` from www.opensolver.org
2. Extract the files to a convenient location
3. Double click on `OpenSolver.xlam`
4. If asked, give Excel permissions to run OpenSolver

The OpenSolver commands will then appear under Excel’s Data tab.

OpenSolver will be available until you quit Excel. If you wish to make OpenSolver always available in Excel, the files from the zip all need to be copied into the Excel add-in directory. This is typically either: 

* `C:\Documents and Settings\”user name”\Application Data\Microsoft\Addins\`
* `C:\Users\"user name"\AppData\Roaming\Microsoft\Addins\`

## Using OpenSolver

The OpenSolver commands appear under Excel’s Data tab.

OpenSolver works with your existing Solver models. It does not replace Solver which you can still use to build your models. However, if you prefer, you can use OpenSolver's AutoModel feature which will attempt to enter all the data into Solver automatically. It does this by looking for an objective function in a row containing the word `min` or `max` (or similar). It then looks for cells containing `<=`, `=` and `>=` that define constraints.

Once you have built your model, you should check it using OpenSolver’s `Show/Hide Model` button.

## Using CBC (or Gurobi)

To solve the model, first turn on the `Assume Linear Model` option, or choose the Simplex engine. Then click OpenSolver’s `Solve` button. OpenSolver then analyses your spreadsheet to extract the optimization model, which is then written to a file and passed to the CBC optimization engine to solve. The result is then read in, and automatically loaded back into your spreadsheet. A dialog is shown only if errors occur.

After solving, OpenSolver does a quick check that your model is a linear one in the sense that the objective and constraints behave as expected when the optimal solution is loaded into the sheet. If it is not, OpenSolver shows an alert, and can then do a detailed linearity analysis.

The Solver tolerance and maximum time options are respected by OpenSolver. If you turn on Solver’s `Show Iteration Results`, OpenSolver will briefly display the CBC output during the solution process (which is typically very fast).

## Using NOMAD

To solve the model set it up in the model dialogue window. NOMAD works better if it is passed bounds which can be added as constraints in the model dialogue as well. You can then set the starting value by putting the values in the variable cells. Then click OpenSolver's Solve button.  OpenSolver then analyses your spreadsheet to extract the optimization model, which is then passed to the NOMAD solver. NOMAD will then pass new variable values to excel and update the spreadsheet before taking back the new objective and constraint values. This will continue until it finds a solution. The result is then passed back, and automatically loaded into your spreadsheet.NOMAD, like many non-linear solvers, cannot guarantee optimality but it does find the best solution that it can. NOMAD solves using the precompiled dlls `OpenSolverNomadDll.dll` or `OpenSolverNomadDll64.dll` that come with your OpenSolver installation.

## More information

Please visit http://www.opensolver.org for more information.

