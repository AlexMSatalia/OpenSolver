Attribute VB_Name = "SolverNeos"
Option Explicit
Public FinalMessage As String
Public NeosResult As String

Public Const NeosTermsAndConditionsLink As String = "http://www.neos-server.org/neos/termofuse.html"
Public Const NeosAdditionalSolverText As String = "Submitting a model to NEOS results in it becoming publicly available. Use of NEOS is subject to the NEOS Terms and Conditions:"

Private Const NEOS_ADDRESS = "http://www.neos-server.org:3332"
Private Const NEOS_RESULT_FILE = "neosresult.txt"
Private Const NEOS_SCRIPT_FILE = "NeosClient.py"

Public SOLVE_LOCAL As Boolean  ' Whether to use local AMPL to solve. Defaults to false

Function NeosClientScriptPath() As String
1               NeosClientScriptPath = JoinPaths(SolverDir, SolverDirMac, NEOS_SCRIPT_FILE)
End Function

Function CallNEOS(s As COpenSolver, OutgoingMessage As String) As String
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler
                
3               If SOLVE_LOCAL Then
4                   CallNEOS = CallNeos_Local(s)
5               Else
6                   CallNEOS = CallNeos_Remote(s, OutgoingMessage)
7               End If
                
8               If InStr(CallNEOS, "Error (2) in /opt/ampl/ampl -R amplin") > 0 Then
                    ' Check for any other error - the amplin error is shown for invalid parameters too
9                   s.Solver.CheckLog s
10                  RaiseGeneralError "NEOS was unable to solve the model because there was an error while running AMPL. " & _
                                      "Please let us know and send us a copy of your spreadsheet so that we can try to fix this error."
11              End If
                    
                
ExitFunction:
12              If RaiseError Then RethrowError
13              Exit Function

ErrorHandler:
14              If Not ReportError("SolverNeos", "CallNeos") Then Resume
15              RaiseError = True
16              GoTo ExitFunction
End Function

Function CallNeos_Local(s As COpenSolver) As String
                Dim RaiseError As Boolean
1               RaiseError = False
2               On Error GoTo ErrorHandler
                
                Dim NeosSolver As ISolverNeos
3               Set NeosSolver = s.Solver
4               If Len(NeosSolver.OptionFile) <> 0 Then
                    Dim OptionsFilePath As String
5                   GetTempFilePath NeosSolver.OptionFile, OptionsFilePath
6                   ParametersToOptionsFile OptionsFilePath, s.SolverParameters
7               End If
             
                Dim ScriptFilePathName As String
8               GetTempFilePath "ampl_local" & ScriptExtension, ScriptFilePathName
                
                Dim FileSolver As ISolverFile
9               Set FileSolver = s.Solver
                
                Dim SolveCommand As String
          #If Mac Then
10                  SolveCommand = SolveCommand & _
                                   "source ~/.bashrc" & ScriptSeparator & _
                                   "source ~/.zshrc" & ScriptSeparator
          #End If
11              SolveCommand = SolveCommand & "ampl " & MakePathSafe(GetModelFilePath(FileSolver))
12              CreateScriptFile ScriptFilePathName, SolveCommand
                
13              CallNeos_Local = ExecCapture(SolveCommand, s.LogFilePathName, GetTempFolder())
                
ExitFunction:
14              If RaiseError Then RethrowError
15              Exit Function

ErrorHandler:
16              If Not ReportError("SolverNeos", "CallNeos_Local") Then Resume
17              RaiseError = True
18              GoTo ExitFunction
End Function

Function CallNeos_Remote(s As COpenSolver, OutgoingMessage As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim InteractiveStatus As Boolean
3         InteractiveStatus = Application.Interactive
          
          Dim NeosSolver As ISolverNeos
4         Set NeosSolver = s.Solver
          Dim OptionsFileString As String
5         If Len(NeosSolver.OptionFile) <> 0 Then
6             OptionsFileString = ParametersToOptionsFileString(s.SolverParameters)
7         End If
             
          ' Wrap in XML for AMPL on NEOS
8         FinalMessage = WrapMessageForNEOS(OutgoingMessage, NeosSolver, OptionsFileString)
          
          Dim errorString As String
9         If s.MinimiseUserInteraction Then
10            CallNeos_Remote = SolveOnNeos(FinalMessage, errorString)
11        Else
              Dim frmCallingNeos As FCallingNeos
12            Set frmCallingNeos = New FCallingNeos
              
13            Application.Interactive = True
14            frmCallingNeos.Show
15            Application.Interactive = InteractiveStatus
              
16            CallNeos_Remote = NeosResult
17            errorString = frmCallingNeos.Tag
18            Unload frmCallingNeos
19        End If
20        If Len(errorString) > 0 Then
21            If errorString = "Aborted" Then
22                RaiseUserCancelledError
23            Else
24                RaiseGeneralError errorString
25            End If
26        End If
              
          ' Dump the whole NEOS response to log file
27        Open s.LogFilePathName For Output As #1
28            Print #1, CallNeos_Remote
29        Close #1

ExitFunction:
30        Application.Interactive = InteractiveStatus
31        Close #1
32        If RaiseError Then RethrowError
33        Exit Function

ErrorHandler:
34        If Not ReportError("SolverNeos", "CallNeos_Remote") Then Resume
35        RaiseError = True
36        GoTo ExitFunction
End Function

Public Function SolveOnNeos(message As String, errorString As String, Optional frmCallingNeos As FCallingNeos = Nothing) As String
1         On Error GoTo ErrorHandler

          Dim result As String, jobNumber As String, Password As String
2         result = SubmitNeosJob(message, jobNumber, Password)
          
3         If jobNumber = "0" Then RaiseGeneralError "An error occured when sending file to NEOS."
          
4         If Not frmCallingNeos Is Nothing Then frmCallingNeos.Tag = "Running"
          
          ' Loop until job is done
          Dim StartTime As Single, Done As Boolean
5         StartTime = Timer()
6         Done = False
7         While Done = False
8             If Not frmCallingNeos Is Nothing Then
9                 If frmCallingNeos.Tag = "Cancelled" Then GoTo Aborted
10            End If
              
11            UpdateStatusBar "OpenSolver: Solving model on NEOS... Time Elapsed: " & Int(Timer() - StartTime) & " seconds"
12            DoEvents
              
13            result = GetNeosJobStatus(jobNumber, Password)
14            If result = "Done" Then
15                Done = True
16            ElseIf result <> "Waiting" And result <> "Running" Then
17                RaiseGeneralError "An error occured while waiting for NEOS. NEOS returned: " & result
18            Else
19                mSleep 5000  ' 5 seconds
20                DoEvents
21            End If
22        Wend
          
23        SolveOnNeos = GetNeosResult(jobNumber, Password)
          
ExitFunction:
24        Exit Function

ErrorHandler:
' We CANNOT raise an error in this function.
' It is sometimes called with a form as a conduit, which means that errors can't propogate back to the main thread.
' Instead, set the error string, which IS passed back to the main thread by the form.
25        If Not ReportError("OpenSolverNeos", "SolveOnNeos_Windows") Then Resume
26        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then GoTo Aborted
27        errorString = OpenSolverErrorHandler.ErrMsg
28        GoTo ExitFunction
          
Aborted:
29        SolveOnNeos = "NEOS solve was aborted"
30        errorString = "Aborted"
31        Exit Function
End Function

#If Mac Then  ' Define on Mac only
Private Function SendToNeos_Mac(method As String, Optional param1 As String, Optional param2 As String) As String
      ' Mac doesn't have ActiveX so can't use MSXML.
      ' It does have python by default, so we can use python's xmlrpclib to contact NEOS instead.
      ' We delegate all interaction to the NeosClient.py script file.
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim SolverPath As String, NeosClientDir As String
3         NeosClientDir = JoinPaths(SolverDir, SolverDirMac)
4         GetExistingFilePathName NeosClientDir, NEOS_SCRIPT_FILE, SolverPath
5         SolverPath = MakePathSafe(SolverPath)
6         Exec "chmod +x " & SolverPath

          Dim SolutionFilePathName As String
7         GetTempFilePath NEOS_RESULT_FILE, SolutionFilePathName
8         DeleteFileAndVerify SolutionFilePathName

          Dim LogFilePathName As String
9         If GetLogFilePath(LogFilePathName) Then DeleteFileAndVerify LogFilePathName
          
10        If Not Exec("python " & SolverPath & " " & method & " " & MakePathSafe(SolutionFilePathName) & " " & param1 & " " & param2) Then
11            RaiseGeneralError "Unknown error while contacting NEOS"
12        End If
          
          ' Read in results from file
13        Open SolutionFilePathName For Input As #1
14        SendToNeos_Mac = Input$(LOF(1), 1)
15        Close #1
          
16        If Left(SendToNeos_Mac, 6) = "Error:" And method <> "ping" Then
17            RaiseGeneralError "An error occured while solving on NEOS. NEOS returned: " & SendToNeos_Mac
18        End If

ExitFunction:
19        Close #1
20        If RaiseError Then RethrowError
21        Exit Function

ErrorHandler:
22        If Not PingNeos() Then
23            Err.Description = "OpenSolver could not establish a connection to NEOS. Check your internet connection and try again. If this error message persists, NEOS may be down."
24            Err.Number = OpenSolver_Error
25        End If
26        If Not ReportError("SolverNeos", "SendToNeos_Mac") Then Resume
27        RaiseError = True
28        GoTo ExitFunction
End Function

#Else  ' Define on Windows only
Private Function SendToNeos_Windows(message As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          ' Late binding so we don't need to add the reference to MSXML, causing a crash on Mac
          Dim XmlHttpReq As Object  ' MSXML2.ServerXMLHTTP
3         Set XmlHttpReq = CreateObject("MSXML2.ServerXMLHTTP")
          
4         XmlHttpReq.Open "POST", NEOS_ADDRESS, False
5         XmlHttpReq.send message
6         SendToNeos_Windows = XmlHttpReq.responseText

ExitFunction:
7         Set XmlHttpReq = Nothing
8         If RaiseError Then RethrowError
9         Exit Function

ErrorHandler:
10        If GetXmlTagValue(message, "methodName") <> "ping" Then
11            If Not PingNeos() Then
12                Err.Description = "OpenSolver could not establish a connection to NEOS. Check your internet connection and try again. If this error message persists, NEOS may be down."
13                Err.Number = OpenSolver_Error
14            End If
15            If Not ReportError("SolverNeos", "SendToNeos_Windows") Then Resume
16        End If
17        RaiseError = True
18        GoTo ExitFunction
End Function
#End If

Private Function GetNeosResult(jobNumber As String, Password As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

    #If Mac Then
3             GetNeosResult = SendToNeos_Mac("read", jobNumber, Password)
    #Else
4             GetNeosResult = DecodeBase64(GetXmlTagValue(SendToNeos_Windows(MakeNeosMethodCall("getFinalResults", jobNumber, Password)), "base64"))
    #End If

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("SolverNeos", "GetNeosResult") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

Private Function SubmitNeosJob(message As String, ByRef jobNumber As String, ByRef Password As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

    #If Mac Then
              ' Create the job file
              Dim ModelFilePathName As String
3             GetTempFilePath "job.xml", ModelFilePathName
4             DeleteFileAndVerify ModelFilePathName
          
5             Open ModelFilePathName For Output As #1
6             Print #1, message
7             Close #1
              
8             SubmitNeosJob = SendToNeos_Mac("send", MakePathSafe(ModelFilePathName))
              
              Dim openingParen As String, closingParen As String
9             openingParen = InStr(SubmitNeosJob, "jobNumber = ") + Len("jobNumber = ")
10            closingParen = InStr(openingParen, SubmitNeosJob, Chr(10))
11            jobNumber = Mid(SubmitNeosJob, openingParen, closingParen - openingParen)
              
12            openingParen = InStr(SubmitNeosJob, "password = ") + Len("password = ")
13            closingParen = InStr(openingParen, SubmitNeosJob, Chr(10))
14            Password = Mid(SubmitNeosJob, openingParen, closingParen - openingParen)
              
    #Else
              ' Clean message up
15            message = Replace(message, "<", "&lt;")
16            message = Replace(message, ">", "&gt;")
          
17            SubmitNeosJob = SendToNeos_Windows(MakeNeosMethodCall("submitJob", StringValue:=message))
              
18            jobNumber = GetXmlTagValue(SubmitNeosJob, "int")
19            Password = GetXmlTagValue(SubmitNeosJob, "string")
    #End If

ExitFunction:
20        Close #1
21        If RaiseError Then RethrowError
22        Exit Function

ErrorHandler:
23        If Not ReportError("SolverNeos", "SubmitNeosJob") Then Resume
24        RaiseError = True
25        GoTo ExitFunction
End Function

Private Function GetNeosJobStatus(jobNumber As String, Password As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

    #If Mac Then
3             GetNeosJobStatus = SendToNeos_Mac("check", jobNumber, Password)
    #Else
4             GetNeosJobStatus = GetXmlTagValue(SendToNeos_Windows(MakeNeosMethodCall("getJobStatus", jobNumber, Password)), "string")
    #End If

ExitFunction:
5         If RaiseError Then RethrowError
6         Exit Function

ErrorHandler:
7         If Not ReportError("SolverNeos", "GetNeosJobStatus") Then Resume
8         RaiseError = True
9         GoTo ExitFunction
End Function

Private Function GetXmlTagValue(message As String, Tag As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim openingParen As Long, closingParen As Long
3         openingParen = InStr(message, "<" & Tag & ">")
4         closingParen = InStr(message, "</" & Tag & ">")
          
          Dim TagLength As Long
5         TagLength = Len(Tag) + 2
6         GetXmlTagValue = Mid(message, openingParen + TagLength, closingParen - openingParen - TagLength)

ExitFunction:
7         If RaiseError Then RethrowError
8         Exit Function

ErrorHandler:
9         If Not ReportError("SolverNeos", "GetXmlTagValue") Then Resume
10        RaiseError = True
11        GoTo ExitFunction
End Function

' Code by Tim Hastings
Private Function DecodeBase64(ByVal strData As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler

          Dim objXML As Object 'MSXML2.DOMDocument
          Dim objNode As Object 'MSXML2.IXMLDOMElement
        
          ' Help from MSXML
3         Set objXML = CreateObject("MSXML2.DOMDocument")
4         Set objNode = objXML.createElement("b64")
5         objNode.DataType = "bin.base64"
6         objNode.Text = strData
7         DecodeBase64 = Stream_BinaryToString(objNode.nodeTypedValue)
        
          ' Clean up
8         Set objNode = Nothing
9         Set objXML = Nothing

ExitFunction:
10        If RaiseError Then RethrowError
11        Exit Function

ErrorHandler:
12        If Not ReportError("SolverNeos", "DecodeBase64") Then Resume
13        RaiseError = True
14        GoTo ExitFunction
End Function

' Code by Tim Hastings
Function Stream_BinaryToString(Binary)
           Dim RaiseError As Boolean
1          RaiseError = False
2          On Error GoTo ErrorHandler
          
           Const adTypeText = 2
           Const adTypeBinary = 1
           
           'Create Stream object
           Dim BinaryStream 'As New Stream
3          Set BinaryStream = CreateObject("ADODB.Stream")
           
           'Specify stream type - we want To save binary data.
4          BinaryStream.Type = adTypeBinary
           
           'Open the stream And write binary data To the object
5          BinaryStream.Open
6          BinaryStream.Write Binary
           
           'Change stream type To text/string
7          BinaryStream.Position = 0
8          BinaryStream.Type = adTypeText
           
           'Specify charset For the output text (unicode) data.
9          BinaryStream.Charset = "us-ascii"
           
           'Open the stream And get text/string data from the object
10         Stream_BinaryToString = BinaryStream.ReadText
11         Set BinaryStream = Nothing

ExitFunction:
12         If RaiseError Then RethrowError
13         Exit Function

ErrorHandler:
14         If Not ReportError("SolverNeos", "Stream_BinaryToString") Then Resume
15         RaiseError = True
16         GoTo ExitFunction
End Function

Function WrapMessageForNEOS(message As String, NeosSolver As ISolverNeos, Optional OptionsFileString As String) As String
      ' Wraps AMPL in the required XML to send to NEOS
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
3         WrapMessageForNEOS = _
              WrapInTag( _
                  WrapInTag(NeosSolver.NeosSolverCategory, "category") & _
                  WrapInTag(NeosSolver.NeosSolverName, "solver") & _
                  WrapInTag(GetInputType(NeosSolver), "inputType") & _
                  WrapInTag(vbNullString, "client") & _
                  WrapInTag("short", "priority") & _
                  WrapInTag(vbNullString, "email") & _
                  WrapInTag(message, "model", True) & _
                  WrapInTag(vbNullString, "data", True) & _
                  WrapInTag(vbNullString, "commands", True) & _
                  IIf(Len(OptionsFileString) > 0, WrapInTag(OptionsFileString & vbNewLine, "options", True), "") & _
                  WrapInTag(vbNullString, "comments", True), _
              "document")

ExitSub:
4         If RaiseError Then RethrowError
5         Exit Function

ErrorHandler:
6         If Not ReportError("SolverNeos", "WrapAMPLForNEOS") Then Resume
7         RaiseError = True
8         GoTo ExitSub
End Function

Function WrapInTag(value As String, TagName As String, Optional AddCData As Boolean = False) As String
1         WrapInTag = "<" & TagName & ">" & _
                          IIf(AddCData, "<![CDATA[", vbNullString) & _
                              value & _
                          IIf(AddCData, "]]>", vbNullString) & _
                      "</" & TagName & ">"
End Function

Function MakeNeosMethodCall(MethodName As String, Optional IntValue As String = vbNullString, Optional StringValue As String = vbNullString) As String
1         MakeNeosMethodCall = WrapInTag( _
                                   WrapInTag(MethodName, "methodName") & _
                                   WrapInTag( _
                                       IIf(Len(IntValue) > 0, MakeNeosParam("int", IntValue), vbNullString) & _
                                       IIf(Len(StringValue) > 0, MakeNeosParam("string", StringValue), vbNullString), _
                                   "params"), _
                               "methodCall")

End Function

Function MakeNeosParam(ParamType As String, ParamValue As String) As String
1         MakeNeosParam = WrapInTag( _
                              WrapInTag( _
                                  WrapInTag(ParamValue, ParamType), _
                              "value"), _
                          "param")

End Function

Private Function GetInputType(Solver As ISolver)
1         If TypeOf Solver Is ISolverFileAMPL Then
2             GetInputType = "AMPL"
3         End If
End Function

Function PingNeos() As Boolean
1         On Error GoTo CantAccess
          
          Dim status As String
    #If Mac Then
2             status = SendToNeos_Mac("ping")
    #Else
3             status = GetXmlTagValue(SendToNeos_Windows(MakeNeosMethodCall("ping")), "string")
    #End If
          Const AliveMessage As String = "NeosServer is alive"
4         PingNeos = (Left(status, Len(AliveMessage)) = AliveMessage)
5         Exit Function
          
CantAccess:
6         PingNeos = False
End Function
