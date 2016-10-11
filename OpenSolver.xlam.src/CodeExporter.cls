VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CodeExporter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================
Private wkbk As Excel.Workbook
Private exportFolder As String
Private fso As Object 'FileSystemObject

Public Property Set SourceWorkbook(book As Excel.Workbook)
1        Set wkbk = book
End Property

Public Sub ExportTo(where As String)
         ' Modification JWD: We need late binding on FSO here to avoid a dependence on Microsoft Scripting Runtime
1        Set fso = CreateObject("Scripting.FileSystemObject")
         
2        If Not fso.FolderExists(where) Then
3            fso.CreateFolder where
4        End If

5        exportFolder = where

6        WriteCode

End Sub

Private Sub WriteCode()

      Dim c As VBComponent
      Dim curMod As VBIDE.CodeModule

1        For Each c In wkbk.VBProject.VBComponents
         
2            Set curMod = c.CodeModule
             
3            c.Export fso.BuildPath(exportFolder, ExportFileNameFor(c))
             
4        Next

End Sub

Private Function ExportFileNameFor(codeComponent As VBComponent) As String

         Dim extension As String

1        Select Case codeComponent.Type
             Case vbext_ct_ClassModule
2                extension = "cls"
3            Case vbext_ct_MSForm
4                extension = "frm"
5            Case vbext_ct_StdModule
6                extension = "bas"
7            Case vbext_ct_Document
8                extension = "sht"
9        End Select

10       ExportFileNameFor = codeComponent.Name & "." & extension

End Function
'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================

