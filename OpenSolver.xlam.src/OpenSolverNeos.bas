Attribute VB_Name = "OpenSolverNeos"
Option Explicit
Public OutgoingMessage As String
Public NeosResult As String

Private Const NEOS_ADDRESS = "http://www.neos-server.org:3332"
Private Const NEOS_RESULT_FILE = "neosresult.txt"
Private Const NEOS_SCRIPT_FILE = "NeosClient.py"

Function CallNEOS(ModelFilePathName As String, Solver As String, Optional MinimiseUserInteraction As Boolean = False) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

6806      ' Import file as continuous string
          Open ModelFilePathName For Input As #1
6807          OutgoingMessage = Input$(LOF(1), 1)
6808      Close #1
           
          ' Wrap in XML for AMPL on NEOS
6809      WrapAMPLForNEOS OutgoingMessage, Solver
           
          Dim errorString As String
          If MinimiseUserInteraction Then
              CallNEOS = SolveOnNeos(OutgoingMessage, errorString)
          Else
              frmCallingNeos.Show
              CallNEOS = NeosResult
              errorString = frmCallingNeos.Tag
          End If
          If Len(errorString) > 0 Then
              If errorString = "Aborted" Then
                  Err.Raise OpenSolver_UserCancelledError, Description:="NEOS solve was aborted"
              Else
                  Err.Raise OpenSolver_NeosError, Description:=errorString
              End If
          End If
   
          ' Dump the whole NEOS response to log file
          Dim logPath As String
          logPath = GetTempFilePath("log1.tmp")
          Open logPath For Output As #1
              Print #1, CallNEOS
          Close #1

ExitFunction:
          Close #1
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverNeos", "CallNEOS") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

Public Function SolveOnNeos(message As String, errorString As String) As String
          On Error GoTo ErrorHandler

          Dim result As String, jobNumber As String, Password As String
6820      result = SubmitNeosJob(message, jobNumber, Password)
          
6825      If jobNumber = "0" Then Err.Raise OpenSolver_NeosError, Description:="An error occured when sending file to NEOS."
          
          frmCallingNeos.Tag = "Running"
          
          ' Loop until job is done
          Dim time As Long, Done As Boolean
6835      time = 0
          Done = False
6836      While Done = False
              If frmCallingNeos.Tag = "Cancelled" Then GoTo Aborted
              
              UpdateStatusBar "OpenSolver: Solving model on NEOS... Time Elapsed: " & time & " seconds"
              
              result = GetNeosJobStatus(jobNumber, Password)
6843          If result = "Done" Then
6844              Done = True
6845          ElseIf result <> "Waiting" And result <> "Running" Then
6846              Err.Raise OpenSolver_NeosError, Description:="An error occured when sending file to NEOS. Neos returned: " & result
6848          Else
6849              SleepSeconds 1
6850              time = time + 1
6853          End If
6854      Wend
          
6861      SolveOnNeos = GetNeosResult(jobNumber, Password)
          
ExitFunction:
          Exit Function

ErrorHandler:
' We CANNOT raise an error in this function.
' It is sometimes called with a form as a conduit, which means that errors can't propogate back to the main thread.
' Instead, set the error string, which IS passed back to the main thread by the form.
          If Not ReportError("OpenSolverNeos", "SolveOnNeos_Windows") Then Resume
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then GoTo Aborted
          errorString = OpenSolverErrorHandler.ErrMsg
          GoTo ExitFunction
          
Aborted:
          SolveOnNeos = "NEOS solve was aborted"
          errorString = "Aborted"
          Exit Function
End Function

#If Mac Then  ' Define on Mac only
Private Function SendToNeos_Mac(method As String, param1 As String, Optional param2 As String) As String
' Mac doesn't have ActiveX so can't use MSXML.
' It does have python by default, so we can use python's xmlrpclib to contact NEOS instead.
' We delegate all interaction to the NeosClient.py script file.
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim SolverPath As String, NeosClientDir As String
    NeosClientDir = JoinPaths(JoinPaths(ThisWorkbook.Path, SolverDir), SolverDirMac)
    GetExistingFilePathName NeosClientDir, NEOS_SCRIPT_FILE, SolverPath
    SolverPath = MakePathSafe(SolverPath)
    system ("chmod +x " & SolverPath)

    Dim SolutionFilePathName As String
    SolutionFilePathName = GetTempFilePath(NEOS_RESULT_FILE)
    DeleteFileAndVerify SolutionFilePathName

    Dim LogFilePathName As String
    LogFilePathName = GetTempFilePath("log1.tmp")
    DeleteFileAndVerify LogFilePathName
    LogFilePathName = MakePathSafe(LogFilePathName)
    
    If Not RunExternalCommand(SolverPath & " " & method & " " & MakePathSafe(SolutionFilePathName) & " " & param1 & " " & param2, LogFilePathName) Then
        Err.Raise OpenSolver_NeosError, Description:="Error while contacting NEOS"
    End If
    
    ' Read in results from file
    Open SolutionFilePathName For Input As #1
    SendToNeos_Mac = Input$(LOF(1), 1)
    Close #1

ExitFunction:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverNeos", "SendToNeos_Mac") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

#Else  ' Define on Windows only
Private Function SendToNeos_Windows(message As String) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    ' Late binding so we don't need to add the reference to MSXML, causing a crash on Mac
    Dim objSvrHTTP As Object 'MSXML2.ServerXMLHTTP
    Set objSvrHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    objSvrHTTP.Open "POST", NEOS_ADDRESS, False
    objSvrHTTP.send message
    SendToNeos_Windows = objSvrHTTP.responseText

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverNeos", "SendToNeos_Windows") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function
#End If

Private Function GetNeosResult(jobNumber As String, Password As String) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    #If Mac Then
        GetNeosResult = SendToNeos_Mac("read", jobNumber, Password)
        If left(GetNeosResult, 6) = "Error:" Then
            Err.Raise OpenSolver_NeosError, Description:="An error occured when sending file to NEOS. Neos returned: " & GetNeosResult
        End If
    #Else
        GetNeosResult = DecodeBase64(GetXmlTagValue(SendToNeos_Windows(MakeNeosMethodCall("getFinalResults", jobNumber, Password)), "base64"))
    #End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverNeos", "GetNeosResult") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Function SubmitNeosJob(message As String, ByRef jobNumber As String, ByRef Password As String) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    #If Mac Then
        ' Create the job file
        Dim ModelFilePathName As String
        ModelFilePathName = GetTempFilePath("job.xml")
        DeleteFileAndVerify ModelFilePathName
    
        Open ModelFilePathName For Output As #1
        Print #1, message
        Close #1
        
        SubmitNeosJob = SendToNeos_Mac("send", MakePathSafe(ModelFilePathName))
        
        Dim openingParen As String, closingParen As String
        openingParen = InStr(SubmitNeosJob, "jobNumber = ") + Len("jobNumber = ")
        closingParen = InStr(openingParen, SubmitNeosJob, Chr(10))
        jobNumber = Mid(SubmitNeosJob, openingParen, closingParen - openingParen)
        
        openingParen = InStr(SubmitNeosJob, "password = ") + Len("password = ")
        closingParen = InStr(openingParen, SubmitNeosJob, Chr(10))
        Password = Mid(SubmitNeosJob, openingParen, closingParen - openingParen)
        
    #Else
        ' Clean message up
        message = Replace(message, "<", "&lt;")
        message = Replace(message, ">", "&gt;")
    
        SubmitNeosJob = SendToNeos_Windows(MakeNeosMethodCall("submitJob", StringValue:=message))
        
        jobNumber = GetXmlTagValue(SubmitNeosJob, "int")
        Password = GetXmlTagValue(SubmitNeosJob, "string")
    #End If

ExitFunction:
    Close #1
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverNeos", "SubmitNeosJob") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Function GetNeosJobStatus(jobNumber As String, Password As String) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    #If Mac Then
        GetNeosJobStatus = SendToNeos_Mac("check", jobNumber, Password)
    #Else
        GetNeosJobStatus = GetXmlTagValue(SendToNeos_Windows(MakeNeosMethodCall("getJobStatus", jobNumber, Password)), "string")
    #End If

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverNeos", "GetNeosJobStatus") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Private Function GetXmlTagValue(message As String, Tag As String) As String
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler

    Dim openingParen As Long, closingParen As Long
    openingParen = InStr(message, "<" & Tag & ">")
    closingParen = InStr(message, "</" & Tag & ">")
    
    Dim TagLength As Long
    TagLength = Len(Tag) + 2
    GetXmlTagValue = Mid(message, openingParen + TagLength, closingParen - openingParen - TagLength)

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("OpenSolverNeos", "GetXmlTagValue") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

' Code by Tim Hastings
Private Function DecodeBase64(ByVal strData As String) As String
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler

          Dim objXML As Object 'MSXML2.DOMDocument
          Dim objNode As Object 'MSXML2.IXMLDOMElement
        
          ' Help from MSXML
6950      Set objXML = CreateObject("MSXML2.DOMDocument")
6951      Set objNode = objXML.createElement("b64")
6952      objNode.DataType = "bin.base64"
6953      objNode.Text = strData
6954      DecodeBase64 = Stream_BinaryToString(objNode.nodeTypedValue)
        
          ' Clean up
6955      Set objNode = Nothing
6956      Set objXML = Nothing

ExitFunction:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Function

ErrorHandler:
          If Not ReportError("OpenSolverNeos", "DecodeBase64") Then Resume
          RaiseError = True
          GoTo ExitFunction
End Function

' Code by Tim Hastings
Function Stream_BinaryToString(Binary)
           Dim RaiseError As Boolean
           RaiseError = False
           On Error GoTo ErrorHandler
          
           Const adTypeText = 2
           Const adTypeBinary = 1
           
           'Create Stream object
           Dim BinaryStream 'As New Stream
6957       Set BinaryStream = CreateObject("ADODB.Stream")
           
           'Specify stream type - we want To save binary data.
6958       BinaryStream.Type = adTypeBinary
           
           'Open the stream And write binary data To the object
6959       BinaryStream.Open
6960       BinaryStream.Write Binary
           
           'Change stream type To text/string
6961       BinaryStream.Position = 0
6962       BinaryStream.Type = adTypeText
           
           'Specify charset For the output text (unicode) data.
6963       BinaryStream.Charset = "us-ascii"
           
           'Open the stream And get text/string data from the object
6964       Stream_BinaryToString = BinaryStream.ReadText
6965       Set BinaryStream = Nothing

ExitFunction:
           If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
           Exit Function

ErrorHandler:
           If Not ReportError("OpenSolverNeos", "Stream_BinaryToString") Then Resume
           RaiseError = True
           GoTo ExitFunction
End Function

Sub WrapAMPLForNEOS(AmplString As String, Solver As String)
' Wraps AMPL in the required XML to send to NEOS
           Dim RaiseError As Boolean
           RaiseError = False
           On Error GoTo ErrorHandler

           Dim Category As String, SolverType As String
6966       SolverType = GetNeosSolverType(Solver)
           Category = GetNeosSolverCategory(SolverType)
           
6967       AmplString = _
              "<document>" & _
                  "<category>" & Category & "</category>" & _
                  "<solver>" & SolverType & "</solver>" & _
                  "<inputType>AMPL</inputType>" & _
                  "<client></client>" & _
                  "<priority>short</priority>" & _
                  "<email></email>" & _
                  "<model><![CDATA[" & AmplString & "end]]></model>" & _
                  "<data><![CDATA[]]></data>" & _
                  "<commands><![CDATA[]]></commands>" & _
                  "<comments><![CDATA[]]></comments>" & _
              "</document>"

ExitSub:
          If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
          Exit Sub

ErrorHandler:
          If Not ReportError("OpenSolverNeos", "WrapAMPLForNEOS") Then Resume
          RaiseError = True
          GoTo ExitSub
End Sub

Function MakeNeosMethodCall(MethodName As String, Optional IntValue As String = "", Optional StringValue As String = "") As String
    MakeNeosMethodCall = "<methodCall>" & _
                         "  <methodName>" & MethodName & "</methodName>" & _
                         "  <params>" & _
                         IIf(Len(IntValue) > 0, "    " & MakeNeosParam("int", IntValue), "") & _
                         IIf(Len(StringValue) > 0, "    " & MakeNeosParam("string", StringValue), "") & _
                         "  </params>" & _
                         "</methodCall>"

End Function

Function MakeNeosParam(ParamType As String, ParamValue As String) As String
    MakeNeosParam = "<param>" & _
                    "  <value>" & _
                    "    <" & ParamType & ">" & ParamValue & "</" & ParamType & ">" & _
                    "  </value>" & _
                    "</param>"

End Function

Function GetNeosSolverCategory(SolverType As String)
    Select Case SolverType
    Case "cbc"
        GetNeosSolverCategory = "milp"
    Case "Bonmin", "Couenne"
        GetNeosSolverCategory = "minco"
    End Select
End Function
