VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================
Private Sub Workbook_AddinInstall()
10        AlterMenuItems (True)
End Sub

Private Sub Workbook_AddinUninstall()
20        AlterMenuItems (False)
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
30        AlterMenuItems (False)
End Sub

Private Sub Workbook_Open()
40        AlterMenuItems (True)
End Sub
'====================================================================
' Excel 2003 Menu Code
' Provided by Paul Becker of Eclipse Engineering (www.eclipseeng.com)
'====================================================================

Sub AlterMenuItems(AddItems As Boolean)
    Dim NeedToAdd As Boolean
    NeedToAdd = Application.Version = "11.0"
#If Mac Then
    NeedToAdd = True
#End If
    If NeedToAdd Then
        If AddItems Then
            Call OpenSolverMain.AddMenuItems
        Else
            Call OpenSolverMain.DelMenuItems
        End If
    End If
End Sub

'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================
Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    If MsgBox("Save VBA Code?", vbYesNo, "Before you go") = vbYes Then
        ExportVBACode
    End If

End Sub

Public Sub ExportVBACode()

    With New CodeExporter
        Set .SourceWorkbook = ThisWorkbook
        .ExportTo exportPath
    End With

End Sub

Private Function exportPath() As String

Dim pathName As String

    pathName = ThisWorkbook.Name & ".src"

    With New FileSystemObject
        With .GetFolder(ThisWorkbook.Path)
            On Error Resume Next
            exportPath = .SubFolders(pathName).Path
            If Err.Number <> 0 Then
                .SubFolders.Add pathName
                exportPath = .SubFolders(pathName).Path
            End If
            On Error GoTo 0
        End With
    End With

End Function
'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================

