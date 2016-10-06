Attribute VB_Name = "CodeImporter"
' Code Importer
' Use the RemoveVBACode first to delete all modules, then use ImportVBACode to load all the modules from the
' source directory
' It seems we need to run these as two separate macros, as otherwise any modules that have been used recently
' are not deleted until the macro exits. Running the Remove and Import separately ensures that Remove is successful

Public Sub RemoveVBACode()
    'Delete all modules/Userforms from the ActiveWorkbook other than CodeImporter and ThisWorkbook
    
    Dim VBComp As VBComponent
    For Each VBComp In ThisWorkbook.VBProject.VBComponents
        If VBComp.Type = vbext_ct_Document Or _
            (VBComp.Type = vbext_ct_StdModule And VBComp.Name = "CodeImporter") Then
            'Thisworkbook or worksheet module, or CodeImporter
            'We do nothing
        Else
            ThisWorkbook.VBProject.VBComponents.Remove VBComp
        End If
    Next VBComp
End Sub
Public Sub ImportVBACode()
      If MsgBox("Import VBA Code? Make sure you have run the remove macro first. " & _
                "Remember that this doesn't affect ThisWorkbook or CodeImporter files, so if there are changes " & _
                "to those files you will need to merge manually.", _
                vbYesNo, "Import") = vbYes Then
          ImportModules ThisWorkbook.FullName & ".src"
      End If
End Sub

Private Sub ImportModules(importFolder As String)
    Dim objFile As Object 'Scripting.File
    Dim FileName As String
    Dim cmpComponents As VBIDE.VBComponents
    Dim c As CodeModule
        
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.GetFolder(importFolder).Files.Count = 0 Then
       MsgBox "There are no files to import"
       Exit Sub
    End If

    Set cmpComponents = ThisWorkbook.VBProject.VBComponents
    
    ''' Import all the code modules in the specified path
    ''' to the ActiveWorkbook.
    For Each objFile In fso.GetFolder(importFolder).Files
        FileName = objFile.Name
        If (fso.GetExtensionName(FileName) = "cls") Or _
            (fso.GetExtensionName(FileName) = "frm") Or _
            (fso.GetExtensionName(FileName) = "bas") Then
                If FileName <> "CodeImporter.bas" Then
                    cmpComponents.Import objFile.Path
                    
                    ' VBA sometimes inserts newlines on import, so we look for this and trim them
                    Set c = cmpComponents(Left(CStr(FileName), Len(FileName) - 4)).CodeModule
                    If c.CountOfLines > 0 Then
                        Do While c.lines(1, 1) = vbNullString
                            c.DeleteLines 1, 1
                        Loop
                    End If
                End If
        End If

    Next objFile
    
End Sub

