Attribute VB_Name = "OpenSolverNeos"
Option Explicit
Public OutgoingMessage As String
Public NeosResult As String

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
#If Mac Then
6810      SolveOnNeos = SolveOnNeos_Mac(message, errorString)
#Else
6811      SolveOnNeos = SolveOnNeos_Windows(message, errorString)
#End If
End Function

Private Function SolveOnNeos_Windows(message As String, errorString As String) As String
          On Error GoTo ErrorHandler

          ' Server name
          Dim txtURL As String
6815      txtURL = "http://www.neos-server.org:3332"
          
          ' Late binding so we don't need to add the reference to MSXML, causing a crash on Mac
          Dim objSvrHTTP As Object 'MSXML2.ServerXMLHTTP
6816      Set objSvrHTTP = CreateObject("MSXML2.ServerXMLHTTP")
          
          ' Set up obj for a POST request
6817      objSvrHTTP.Open "POST", txtURL, False
          
          ' Clean message up
6818      message = Replace(message, "<", "&lt;")
6819      message = Replace(message, ">", "&gt;")
          
          ' Set up message as XML
6820      message = "<methodCall>" & _
                    "    <methodName>submitJob</methodName>" & _
                    "    <params>" & _
                    "        <param>" & _
                    "            <value>" & _
                    "                <string>" & message & "</string>" & _
                    "            </value>" & _
                    "        </param>" & _
                    "    </params>" & _
                    "</methodCall>"
          
          ' Send Message to NEOS
6821      objSvrHTTP.send (message)
          
          ' Extract Job Number
          Dim openingParen As String, closingParen As String, jobNumber As String
6822      openingParen = InStr(objSvrHTTP.responseText, "<int>")
6823      closingParen = InStr(objSvrHTTP.responseText, "</int>")
6824      jobNumber = Mid(objSvrHTTP.responseText, openingParen + Len("<int>"), closingParen - openingParen - Len("<int>"))
          
6825      If jobNumber = 0 Then
6826          errorString = "An error occured when sending file to NEOS."
6827          GoTo ExitFunction
6828      End If
          
          ' Extract Password
          Dim Password As String
6829      openingParen = InStr(objSvrHTTP.responseText, "<string>")
6830      closingParen = InStr(objSvrHTTP.responseText, "</string>")
6831      Password = Mid(objSvrHTTP.responseText, openingParen + Len("<string>"), closingParen - openingParen - Len("<string>"))
          
          ' Set up Job Status message for XML
          Dim Done As Boolean, result As String
6832      message = "<methodCall>" & _
                    "    <methodName>getJobStatus</methodName>" & _
                    "    <params>" & _
                    "        <param>" & _
                    "            <value>" & _
                    "                <int>" & jobNumber & "</int>" & _
                    "            </value>" & _
                    "            <value>" & _
                    "                <string>" & Password & "</string>" & _
                    "            </value>" & _
                    "        </param>" & _
                    "    </params>" & _
                    "</methodCall>"
6833      Done = False
          
          frmCallingNeos.Tag = "Running"
          
          ' Loop until job is done
          Dim time As Long
6835      time = 0
6836      While Done = False
              If frmCallingNeos.Tag = "Cancelled" Then GoTo Aborted
              
6837          DoEvents
              
              ' Reset obj
6838          objSvrHTTP.Open "POST", txtURL, False
              
              ' Send message
6839          objSvrHTTP.send (message)
              
              ' Extract answer
6840          openingParen = InStr(objSvrHTTP.responseText, "<string>")
6841          closingParen = InStr(objSvrHTTP.responseText, "</string>")
6842          result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
              
              ' Evaluate result
6843          If result = "Done" Then
6844              Done = True
6845          ElseIf result <> "Waiting" And result <> "Running" Then
6846              errorString = "An error occured when sending file to NEOS. Neos returned: " & result
6847              GoTo ExitFunction
6848          Else
6849              sleep 1000
6850              time = time + 1
6851              Application.StatusBar = "OpenSolver: Solving model on NEOS... Time Elapsed: " & time & " seconds"
6852              DoEvents
6853          End If
6854      Wend
          
          ' Set up final message for XML
6856      message = "<methodCall>" & _
                    "    <methodName>getFinalResults</methodName>" & _
                    "    <params>" & _
                    "        <param>" & _
                    "            <value>" & _
                    "                <int>" & jobNumber & "</int>" & _
                    "            </value>" & _
                    "        </param>" & _
                    "        <param>" & _
                    "            <value>" & _
                    "                <string>" & Password & "</string>" & _
                    "            </value>" & _
                    "        </param>" & _
                    "    </params>" & _
                    "</methodCall>"
          
          ' Reset obj
6857      objSvrHTTP.Open "POST", txtURL, False
          
6858      objSvrHTTP.send (message)
          
          ' Extract Result
6859      openingParen = InStr(objSvrHTTP.responseText, "<base64>")
6860      closingParen = InStr(objSvrHTTP.responseText, "</base64>")
6861      result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
          
          ' The message is returned from NEOS in base 64
6862      SolveOnNeos_Windows = DecodeBase64(result)
          
ExitFunction:
          Exit Function

ErrorHandler:
' We CANNOT raise an error in these functions.
' They are sometimes called with a form as a conduit, which means that errors can't propogate back to the main thread.
' Instead, set the error string, which IS passed back to the main thread by the form.
          If Not ReportError("OpenSolverNeos", "SolveOnNeos_Windows") Then Resume
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then GoTo Aborted
          errorString = OpenSolverErrorHandler.ErrMsg
          GoTo ExitFunction
          
Aborted:
          SolveOnNeos_Windows = "NEOS solve was aborted"
          errorString = "Aborted"
          Exit Function
End Function

#If Mac Then
Private Function SolveOnNeos_Mac(message As String, errorString As String) As String
          ' Mac doesn't have ActiveX so can't use MSXML.
          ' It does have python by default, so we can use python's xmlrpclib to contact NEOS instead.
          ' We delegate all interaction to the NeosClient.py script file.
          Dim RaiseError As Boolean
          RaiseError = False
          On Error GoTo ErrorHandler
          
          ' For some reason this status bar doesn't stick unless we update the screen
6868      Application.ScreenUpdating = True
6869      Application.StatusBar = "OpenSolver: Solving model on NEOS... "
6870      Application.ScreenUpdating = False
          
          Dim ModelFilePathName As String
6871      ModelFilePathName = GetTempFilePath("job.xml")
6872      DeleteFileAndVerify ModelFilePathName

          ' Create the job file
6874      Open ModelFilePathName For Output As #1
6875      Print #1, message
6876      Close #1
          
          ' Set up commands for NeosClient
          ' NeosClient call is of the form: NeosClient.py <method> <neosresult.txt> <extra params> >> <logfile>
6878      ModelFilePathName = MakePathSafe(ModelFilePathName)
          
          Dim SolverPath As String, NeosClientDir As String
6879      NeosClientDir = JoinPaths(ThisWorkbook.Path, SolverDir)
6880      NeosClientDir = JoinPaths(NeosClientDir, SolverDirMac)
          
6881      GetExistingFilePathName NeosClientDir, "NeosClient.py", SolverPath
6882      SolverPath = MakePathSafe(SolverPath)
6883      system ("chmod +x " & SolverPath)
          
          Dim SolutionFilePathName As String
6884      SolutionFilePathName = GetTempFilePath("neosresult.txt")
6885      DeleteFileAndVerify SolutionFilePathName

          Dim LogFilePathName As String
6886      LogFilePathName = GetTempFilePath("log1.tmp")
6887      DeleteFileAndVerify LogFilePathName
6888      LogFilePathName = MakePathSafe(LogFilePathName)

6889      Application.ScreenUpdating = True
6890      Application.Cursor = xlWait ' doesn't seem to work on mac

          ' Run NeosClient.py->send
          Dim result As Boolean
6891      result = RunExternalCommand(SolverPath & " send " & MakePathSafe(SolutionFilePathName) & " " & ModelFilePathName, LogFilePathName)
6892      If Not result Then GoTo ContactError
          
          ' Read in job number and password
          Dim jobNumber As String, Password As String
6896      Open SolutionFilePathName For Input As #1
6897          Line Input #1, message
6898          jobNumber = Mid(message, Len("jobNumber = ") + 1)
6899          Line Input #1, message
6900          Password = Mid(message, Len("password = ") + 1)
6901      Close #1
6903      DeleteFileAndVerify SolutionFilePathName

          ' Loop until job is done
          Dim time As Long, Done As Boolean
6904      Done = False
6905      time = 0
6906      While Done = False
              If frmCallingNeos.Tag = "Cancelled" Then GoTo Aborted
              
              ' Run NeosClient.py->check
6907          result = RunExternalCommand(SolverPath & " check " & MakePathSafe(SolutionFilePathName) & " " & jobNumber & " " & Password, LogFilePathName)
6908          If Not result Then GoTo ContactError
6911          Open SolutionFilePathName For Input As #1
6912              Line Input #1, message
6913          Close #1
6914          DeleteFileAndVerify SolutionFilePathName

              ' Evaluate result
6915          If message = "Done" Then
6916              Done = True
6917          ElseIf message <> "Waiting" And message <> "Running" Then
6918              GoTo ContactError
6919          Else
6920              sleep 1
6921              time = time + 1
6922              Application.StatusBar = "OpenSolver: Solving model on NEOS... Time Elapsed: " & time & " seconds"
6923              DoEvents
6924          End If
6925      Wend
          
6926      Application.Cursor = xlDefault
6927      Application.ScreenUpdating = False
          
          ' Run NeosClient.py->check
6928      result = RunExternalCommand(SolverPath & " read " & MakePathSafe(SolutionFilePathName) & " " & jobNumber & " " & Password, LogFilePathName)
6929      If Not result Then GoTo ContactError
          
          ' Read in results from file
6933      Open SolutionFilePathName For Input As #1
6934          message = Input$(LOF(1), 1)
6935      Close #1
          
6937      If left(message, 6) = "Error:" Then
              errorString = "An error occured when sending file to NEOS. Neos returned: " & message
6938          GoTo ExitFunction
6939      End If
6940      SolveOnNeos_Mac = message

ExitFunction:
6941      Exit Function

ErrorHandler:
' We CANNOT raise an error in these functions.
' They are sometimes called with a form as a conduit, which means that errors can't propogate back to the main thread.
' Instead, set the error string, which IS passed back to the main thread by the form.
          If Not ReportError("OpenSolverNeos", "SolveOnNeos_Windows") Then Resume
          If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then GoTo Aborted
          errorString = OpenSolverErrorHandler.ErrMsg
          GoTo ExitFunction
          
Aborted:
          SolveOnNeos_Mac = "NEOS solve was aborted"
          errorString = "Aborted"
          GoTo ExitFunction
          
ContactError:
          errorString = "An error occured contacting NEOS"
          GoTo ExitFunction
End Function
#End If

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

Function GetNeosSolverCategory(SolverType As String)
    Select Case SolverType
    Case "cbc"
        GetNeosSolverCategory = "milp"
    Case "Bonmin", "Couenne"
        GetNeosSolverCategory = "minco"
    End Select
End Function
