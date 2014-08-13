VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CQuickSolve"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' OpenSolver
' Copyright Andrew Mason 2010
' http://www.OpenSolver.org
' This software is distributed under the terms of the GNU General Public License
'
' This file is part of OpenSolver.
'
' OpenSolver is free software: you can redistribute it and/or modify
' it under the terms of the GNU General Public License as published by
' the Free Software Foundation, either version 3 of the License, or
' (at your option) any later version.
'
' OpenSolver is distributed in the hope that it will be useful,
' but WITHOUT ANY WARRANTY; without even the implied warranty of
' MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
' GNU General Public License for more details.
'
' You should have received a copy of the GNU General Public License
' along with OpenSolver.  If not, see <http://www.gnu.org/licenses/>.
'
' v0.97: Removed a bug in handling of multi-area ranges; code now uses:
'   for each c in range
' and does not index cells within the range directly

' v1.0 - updated to handle all constraints as ranges to increase speed. Kathleen Gilbert 01.2011
Option Explicit

' THis class defines how we update a model when the values inside a parameter range are altered.
' Specifically, it records how the RHS of a model built from Solver data vary as the values
' in a range of 'parameter' cells vary

' 20110928 AJM: Converted from a() to SparseA() to handle larger problems

Private NumRows As Long
Private NumParams As Long
Private OriginalParamValues() As Double
Private OriginalRHS() As Double
' Private a() As Double
Private SparseA() As CIndexedCoeffs

Private CurrentParameterRange As Range


Function AnalyseParameters(OpenSolver As COpenSolver, sheetName As String, _
                NumConstraints As Long, ByRef RHS() As Double, _
                ByRef LHSRange() As Range, ByRef RHSRange() As Range, _
                ByRef RHSFormula() As String, _
                ByRef RHSType() As SolverInputType, ByRef LHSOriginalValues As Variant, ByRef RHSOriginalValues As Variant) As Boolean

          ' This builds a model of the form RHS = InitialRHS + A_param * (Param - InitialParam)
          ' where b = solver model constraint RHS values
          '       b_param = value of b when each parameter is 0
          '       A_param is a matrix
          '       Param is the vector of the current parameter values
          ' Returns true if successful
          ' Note: Will update arguments LHSValue and RHSValue to the current values of these cells on the spreadsheet
          ' This remembers the parameter range for which we have built the model for checking later
          
26450     On Error GoTo errorHandler
26460     Set CurrentParameterRange = Nothing
          
          ' Remember the current number of rows and original RHS values
          Dim i As Long, oldCalculationMode As Long
          
26470     NumRows = OpenSolver.GetNumRows
          
26480     oldCalculationMode = Application.Calculation

26490     If NumRows > 0 Then ReDim OriginalRHS(NumRows) As Double
26500     For i = 1 To NumRows
26510         OriginalRHS(i) = RHS(i)
26520     Next i
          
          ' Find the Parameter range, and get the numer of parameters
          Dim ParamRange As Range
26530     If Not OpenSolverModule.CheckModelHasParameterRange Then
26540         Exit Function   ' Error message already shown
26550     End If
26560     Set ParamRange = OpenSolverModule.RemoveRangeOverlap(Range(sheetName & OpenSolverModule.GetParamRangeName))
          
          Dim BadCell As Range
26570     If Not OpenSolverModule.CheckRangeContainsNoAmbiguousMergedCells(ParamRange, BadCell) Then
26580         MsgBox "QuickSolve Initialisation Failed: The parameter " & BadCell.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & " is inaccessible as it is within the merged range " & BadCell.MergeArea.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ".", , "OpenSolver" & sOpenSolverVersion & " Error"
26590         GoTo ExitSub
26600     End If
          
26610     NumParams = ParamRange.Count
          
          ' Remember the original parameter values
          ' Warning: Do not use ParamRange(i) as this fails with more than 1 area; using "for each" seems to work
          ' TODO: THis is very slow; speed it up
          Dim c As Range
26620     If NumParams > 0 Then ReDim OriginalParamValues(NumParams) As Double
26630     i = 1
26640     For Each c In ParamRange
26650         OriginalParamValues(i) = c.Value2
26660         i = i + 1
26670     Next c
          
          ' Get current values for the LHS cell and RHS cell of the Excel constraints
          ' Note: If LHSIsRange(i)=false, then LHSValue(i) already contains the fixed constant value given in the constraint
          
          ' Look at how the LHSRange and RHSRange values now change as we perturb the parameters, and hence deduce how the OpenSolver model's RHS values alter
26680     Application.Calculation = xlCalculationManual
26690     Application.ScreenUpdating = False
          
          Dim param As Long, row As Long, NonZeros As Long, constraint As Long
          Dim CurrentLHSValues  As Variant, CurrentRHSValues As Variant
26700     param = 0
          
26710     If NumRows > 0 And NumParams > 0 Then ReDim SparseA(NumRows) ' ReDim a(NumRows, NumParams) As Double
26720     For row = 1 To NumRows
26730         Set SparseA(row) = New CIndexedCoeffs
26740     Next row
         
26750     For Each c In ParamRange
26760         If c.HasArray Then
26770             MsgBox "QuickSolve Initialisation Failed: OpenSolver cannot perform quick solves for parameters that contain array formulae.", , "OpenSolver" & sOpenSolverVersion & " Error"
26780             GoTo ExitSub
26790         End If
26800         param = param + 1
26810         If param Mod 10 = 0 Then Application.StatusBar = "OpenSolver: Analysing Parameters... " & param & "/" & NumParams & " parameters"
              Dim OldFormula As String
26820         On Error Resume Next
26830         OldFormula = c.Formula
26840         If Err.Number <> 0 Then
                  ' There was an error changing a paramater cell; inform the user
26850             MsgBox "QuickSolve Initialisation Failed: An error (" & Err.Number & " " & Err.Description & IIf(Erl = 0, "", " at line " & Erl) & ") occurred while changing the value in parameter cell " & c.AddressLocal(RowAbsolute:=False, ColumnAbsolute:=False) & ".", , "OpenSolver" & sOpenSolverVersion & " Error"
26860             GoTo ExitSub
26870         End If
              ' Change the parameter value, and see hopw the RHS's change
26880         On Error GoTo errorHandler
26890         c.Value2 = OriginalParamValues(param) + 1
26900         Application.Calculate
26910         row = 1
              Dim j As Long, coeff As Double
26920         For constraint = 1 To NumConstraints
                  ' Check to see what is different and add rows to sparsea
26930             If Not LHSRange(constraint) Is Nothing Then ' skip Binary and Integer constraints
                      Dim OriginalLHSValues As Variant
                      Dim OriginalRHSValues As Variant
26940                 OriginalLHSValues = LHSOriginalValues(constraint)
26950                 OriginalRHSValues = RHSOriginalValues(constraint)
                      ' Get current value(s) for LHS and RHS of this constraint off the sheet. LHS is always an array (even if 1x1)
26960                 OpenSolver.GetCurrentConstraintValues constraint, CurrentLHSValues, CurrentRHSValues
                      'TODO Put in better error trapping for errors that occur on the s/sheet
26970                 For i = 1 To UBound(OriginalLHSValues, 1) ' number of rows
26980                     For j = 1 To UBound(OriginalLHSValues, 2) ' number of columns
26990                         If RHSType(constraint) = SolverInputType.MultiCellRange Then
27000                             coeff = -(CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) + (CurrentRHSValues(i, j) - OriginalRHSValues(i, j))
                                  'a(Row, param)  = coeff
27010                             If coeff <> 0 Then
27020                                 SparseA(row).Add param, coeff
27030                                 NonZeros = NonZeros + 1
27040                             End If
27050                         ElseIf RHSType(constraint) = SolverInputType.SingleCellRange Or RHSType(constraint) = SolverInputType.Formula Then
27060                             coeff = -(CurrentLHSValues(i, j) - OriginalLHSValues(i, j)) + (CurrentRHSValues - OriginalRHSValues)
27070                             If coeff <> 0 Then
27080                                 SparseA(row).Add param, coeff
27090                                 NonZeros = NonZeros + 1
27100                             End If
27110                         Else 'the rhs of this constraint is a constant
27120                             coeff = -(CurrentLHSValues(i, j) - OriginalLHSValues(i, j))
27130                             If coeff <> 0 Then
27140                                 SparseA(row).Add param, coeff
27150                                 NonZeros = NonZeros + 1
27160                             End If
27170                         End If
27180                         row = row + 1
27190                     Next j
27200                 Next i
                      ' End If
27210             End If
27220         Next constraint
27230         c.Formula = OldFormula
27240     Next c
          
27250     If NonZeros = 0 Then
27260         MsgBox "QuickSolve Initialisation Failed: The model's constraints do not depend on the QuickSolve parameter cells; please try again. (The QuickSolve process is used for repeatedly solving a model after making changes to the parameter cells, where these changes result in changes to the model's constraints.)", , "OpenSolver" & sOpenSolverVersion & " Error"
27270         GoTo ExitSub
27280     End If
          
27290     Set CurrentParameterRange = ParamRange
27300     AnalyseParameters = True    ' success
          
ExitSub:
          ' On Error GoTo 0
27310     Application.StatusBar = False ' Resume normal status bar behaviour
27320     Application.ScreenUpdating = True
27330     Application.Calculation = oldCalculationMode
              
27340     Exit Function

errorHandler:
27350     If Err.Number = 18 Then
27360         If MsgBox("You have pressed the Escape key. Do you wish to cancel?", _
                         vbCritical + vbYesNo + vbDefaultButton1, _
                         "OpenSolver: User Interrupt Occured...") = vbNo Then
27370             Resume 'continue on from where error occured
27380         Else
27390             Resume ExitSub
27400         End If
27410     End If
27420     MsgBox "OpenSolver encountered error " & Err.Number & ":" & vbCrLf & Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")") & vbCrLf & "Source = " & Err.Source, , "OpenSolver Code Error"
27430     Resume ExitSub
27440     Resume
End Function

Public Function UpdateModelRHS(sheetName As String, ByRef RHS() As Double) As Boolean
          ' Given the new Parameter values on the worksheet, calculate a new set of model RHS values
          
          ' Find the Parameter range on the given (current) sheet, and check that this is the range we are set up for
          Dim ParamRange As Range
27450     On Error Resume Next
27460     Set ParamRange = Range(sheetName & GetParamRangeName)
27470     If Err.Number <> 0 Then
27480         MsgBox "Error: No Parameter range could be found on the worksheet.", , "OpenSolver" & sOpenSolverVersion & " Error"
27490         UpdateModelRHS = False
27500         Exit Function
27510     End If
27520     On Error GoTo 0
27530     If (CurrentParameterRange.Address <> ParamRange.Address) Or (CurrentParameterRange.Worksheet.Name <> ParamRange.Worksheet.Name) Then
27540         MsgBox "QuickSolve Failed: You need to initialize the QuickSolve engine for the current parameters.", title:="OpenSolver" & sOpenSolverVersion & " Error"
27550         UpdateModelRHS = False
27560         Exit Function
27570     End If
27580     If NumParams <> ParamRange.Count Then
27590         MsgBox "QuickSolve Failed: The number of values in the Parameter range has changed, and so the model update can not be completed.", , "OpenSolver" & sOpenSolverVersion & " Error"
27600         UpdateModelRHS = False
27610         Exit Function
27620     End If
          
          Dim ParamChanges() As Double
27630     If NumParams > 0 Then ReDim ParamChanges(NumParams) As Double
          Dim param As Long, c As Range
27640     param = 1
27650     For Each c In ParamRange
27660         ParamChanges(param) = c.Value2 - OriginalParamValues(param)
27670         param = param + 1
27680     Next c
          
          Dim row As Long, value As Double, i As Long, coeff As Double
27690     For row = 1 To NumRows
27700         value = OriginalRHS(row)
              'For param = 1 To NumParams
              '    If a(Row, param) <> 0 Then
              '        Value = Value + a(Row, param) * ParamChanges(param)
              '    End If
              'Next param
27710         With SparseA(row)
27720             For i = 1 To .Count
27730                 param = .index(i)
27740                 coeff = .Coefficient(i)
27750                 value = value + coeff * ParamChanges(param)
27760             Next i
27770         End With
27780         RHS(row) = value
27790     Next row
          
27800     UpdateModelRHS = True    ' success

End Function
