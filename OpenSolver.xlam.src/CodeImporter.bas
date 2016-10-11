Attribute VB_Name = "CodeImporter"
' Code Importer
' Use the RemoveVBACode first to delete all modules, then use ImportVBACode to load all the modules from the
' source directory
' It seems we need to run these as two separate macros, as otherwise any modules that have been used recently
' are not deleted until the macro exits. Running the Remove and Import separately ensures that Remove is successful

Public Sub RemoveVBACode()
          'Delete all modules/Userforms from the ActiveWorkbook other than CodeImporter and ThisWorkbook
          
          Dim VBComp As VBComponent
1         For Each VBComp In ThisWorkbook.VBProject.VBComponents
2             If VBComp.Type = vbext_ct_Document Or _
                  (VBComp.Type = vbext_ct_StdModule And VBComp.Name = "CodeImporter") Then
                  'Thisworkbook or worksheet module, or CodeImporter
                  'We do nothing
3             Else
4                 ThisWorkbook.VBProject.VBComponents.Remove VBComp
5             End If
6         Next VBComp
End Sub
Public Sub ImportVBACode()
1           If MsgBox("Import VBA Code? Make sure you have run the remove macro first. " & _
                      "Remember that this doesn't affect ThisWorkbook or CodeImporter files, so if there are changes " & _
                      "to those files you will need to merge manually.", _
                      vbYesNo, "Import") = vbYes Then
2               ImportModules ThisWorkbook.FullName & ".src"
3           End If
End Sub

Private Sub ImportModules(importFolder As String)
          Dim objFile As Object 'Scripting.File
          Dim FileName As String
          Dim cmpComponents As VBIDE.VBComponents
          Dim c As CodeModule
              
          Dim fso As Object
1         Set fso = CreateObject("Scripting.FileSystemObject")
2         If fso.GetFolder(importFolder).Files.Count = 0 Then
3            MsgBox "There are no files to import"
4            Exit Sub
5         End If

6         Set cmpComponents = ThisWorkbook.VBProject.VBComponents
          
          ''' Import all the code modules in the specified path
          ''' to the ActiveWorkbook.
7         For Each objFile In fso.GetFolder(importFolder).Files
8             FileName = objFile.Name
9             If (fso.GetExtensionName(FileName) = "cls") Or _
                  (fso.GetExtensionName(FileName) = "frm") Or _
                  (fso.GetExtensionName(FileName) = "bas") Then
10                    If FileName <> "CodeImporter.bas" Then
11                        cmpComponents.Import objFile.Path
                          
                          ' VBA sometimes inserts newlines on import, so we look for this and trim them
12                        Set c = cmpComponents(Left(CStr(FileName), Len(FileName) - 4)).CodeModule
13                        If c.CountOfLines > 0 Then
14                            Do While c.lines(1, 1) = vbNullString
15                                c.DeleteLines 1, 1
16                            Loop
17                        End If
18                    End If
19            End If

20        Next objFile
          
End Sub

