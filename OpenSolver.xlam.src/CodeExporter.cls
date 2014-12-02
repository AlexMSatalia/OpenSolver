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
6021     Set wkbk = book
End Property

Public Sub ExportTo(where As String)
         ' Modification JWD: We need late binding on FSO here to avoid a dependence on Microsoft Scripting Runtime
6022     Set fso = CreateObject("Scripting.FileSystemObject")
         
6023     If Not fso.FolderExists(where) Then
6024         fso.CreateFolder where
6025     End If

6026     exportFolder = where

6027     WriteCode

End Sub

Private Sub WriteCode()

      Dim c As VBComponent
      Dim curMod As VBIDE.CodeModule

6028     For Each c In wkbk.VBProject.VBComponents
         
6029         Set curMod = c.CodeModule
             
6030         c.Export fso.BuildPath(exportFolder, ExportFileNameFor(c))
             
6031     Next

End Sub

Private Function IsExportable(codeComponent As VBComponent) As Boolean
         
6032     Select Case codeComponent.Type
             Case vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_StdModule
6033             IsExportable = True
6034         Case Else
6035             IsExportable = True
6036     End Select

End Function

Private Function ExportFileNameFor(codeComponent As VBComponent) As String

      Dim extension As String

6037     Select Case codeComponent.Type
             Case vbext_ct_ClassModule
6038             extension = "cls"
6039         Case vbext_ct_MSForm
6040             extension = "frm"
6041         Case vbext_ct_StdModule
6042             extension = "bas"
6043         Case vbext_ct_Document
6044             extension = "sht"
6045     End Select

6046     ExportFileNameFor = codeComponent.Name & "." & extension

End Function
'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================

