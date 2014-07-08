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
Private fso As Object

Public Property Set SourceWorkbook(book As Excel.Workbook)
   Set wkbk = book
End Property

Public Sub ExportTo(where As String)
   ' Modification JWD: We need late binding on FSO here to avoid a dependence on Microsoft Scripting Runtime
   Set fso = CreateObject("Scripting.FileSystemObject")
   
   If Not fso.FolderExists(where) Then
       fso.CreateFolder where
   End If

   exportFolder = where

   WriteCode

End Sub

Private Sub WriteCode()

Dim c As VBComponent
Dim curMod As VBIDE.CodeModule

   For Each c In wkbk.VBProject.VBComponents
   
       Set curMod = c.CodeModule
       
       c.Export fso.BuildPath(exportFolder, ExportFileNameFor(c))
       
   Next

End Sub

Private Function IsExportable(codeComponent As VBComponent) As Boolean
   
   Select Case codeComponent.Type
       Case vbext_ct_ClassModule, vbext_ct_MSForm, vbext_ct_StdModule
           IsExportable = True
       Case Else
           IsExportable = True
   End Select

End Function

Private Function ExportFileNameFor(codeComponent As VBComponent) As String

Dim extension As String

   Select Case codeComponent.Type
       Case vbext_ct_ClassModule
           extension = "cls"
       Case vbext_ct_MSForm
           extension = "frm"
       Case vbext_ct_StdModule
           extension = "bas"
       Case vbext_ct_Document
           extension = "cls"
   End Select

   ExportFileNameFor = codeComponent.Name & "." & extension

End Function
'====================================================================
' Excel Code Exporter by Mike Woodhouse
' Released under MIT License at https://xlvbadevtools.codeplex.com/
'====================================================================

