Attribute VB_Name = "OpenSolverNeos"
Option Explicit
Function CallNEOS(ModelFilePathName As String, Solver As String, errorString As String) As String
          ' Import file as continuous string
          Dim message As String
6805      On Error GoTo ErrHandler
6806      Open ModelFilePathName For Input As #1
6807          message = Input$(LOF(1), 1)
6808      Close #1
           
          ' Wrap in XML for AMPL on NEOS
6809      WrapAMPLForNEOS message, Solver
           
#If Mac Then
6810      CallNEOS = CallNEOS_Mac(message, errorString)
#Else
6811      CallNEOS = CallNEOS_Windows(message, errorString)
#End If
6812      Exit Function
          
ErrHandler:
6813      Close #1
6814      Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
           
End Function

Private Function CallNEOS_Windows(message As String, errorString As String)
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
6820      message = "<methodCall><methodName>submitJob</methodName><params><param><value><string>" _
             & message & "</string></value></param></params></methodCall>"
          
          ' Send Message to NEOS
6821      objSvrHTTP.send (message)
          
          ' Extract Job Number
          Dim openingParen As String, closingParen As String, jobNumber As String
6822      openingParen = InStr(objSvrHTTP.responseText, "<int>")
6823      closingParen = InStr(objSvrHTTP.responseText, "</int>")
6824      jobNumber = Mid(objSvrHTTP.responseText, openingParen + Len("<int>"), closingParen - openingParen - Len("<int>"))
          
6825      If jobNumber = 0 Then
6826          MsgBox "An error occured when sending file to NEOS."
6827          GoTo ExitSub
6828      End If
          
          ' Extract Password
          Dim Password As String
6829      openingParen = InStr(objSvrHTTP.responseText, "<string>")
6830      closingParen = InStr(objSvrHTTP.responseText, "</string>")
6831      Password = Mid(objSvrHTTP.responseText, openingParen + Len("<string>"), closingParen - openingParen - Len("<string>"))
          
          ' Set up Job Status message for XML
          Dim Done As Boolean, result As String
6832      message = "<methodCall><methodName>getJobStatus</methodName><params><param><value><int>" _
             & jobNumber & "</int></value><value><string>" & Password & _
             "</string></value></param></params></methodCall>"
6833      Done = False
          
6834      CallingNeos.Show False
          
          ' Loop until job is done
          Dim time As Long
6835      time = 0
6836      While Done = False
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
6846              MsgBox "An error occured when sending file to NEOS. Neos returned: " & result
6847              GoTo ExitSub
6848          Else
6849              Application.Wait (Now + TimeValue("0:00:01"))
6850              time = time + 1
6851              Application.StatusBar = "OpenSolver: Solving model on NEOS... Time Elapsed: " & time & " seconds"
6852              DoEvents
6853          End If
6854      Wend
          
6855      CallingNeos.Hide
          
          ' Set up final message for XML
6856      message = "<methodCall><methodName>getFinalResults</methodName><params><param><value><int>" & _
                    jobNumber & "</int></value></param><param><value><string>" & Password & _
                    "</string></value></param></params></methodCall>"
          
          ' Reset obj
6857      objSvrHTTP.Open "POST", txtURL, False
          
6858      objSvrHTTP.send (message)
          
          ' Extract Result
6859      openingParen = InStr(objSvrHTTP.responseText, "<base64>")
6860      closingParen = InStr(objSvrHTTP.responseText, "</base64>")
6861      result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
          
          ' The message is returned from NEOS in base 64
6862      CallNEOS_Windows = DecodeBase64(result)
          
6863      Exit Function
                
ExitSub:
6864      errorString = "ExitSub"
6865      Exit Function
errorHandler:
6866      errorString = "Error while contacting NEOS."

End Function

Private Function CallNEOS_Mac(message As String, errorString As String)
          ' Mac doesn't have ActiveX so can't use MSXML.
          ' It does have python by default, so we can use python's xmlrpclib to contact NEOS instead.
          ' We delegate all interaction to the NeosClient.py script file.
          Dim errorPrefix As String
6867      errorPrefix = "Sending model to NEOS"
          
          ' For some reason this status bar doesn't stick unless we update the screen
6868      Application.ScreenUpdating = True
6869      Application.StatusBar = "OpenSolver: Solving model on NEOS... "
6870      Application.ScreenUpdating = False
          
          Dim ModelFilePathName As String
6871      ModelFilePathName = GetTempFilePath("job.xml")
6872      DeleteFileAndVerify ModelFilePathName, errorPrefix, "Unable to delete the job file: " & ModelFilePathName
          
          ' Create the job file
6873      On Error GoTo ErrHandler
6874      Open ModelFilePathName For Output As #1
6875      Print #1, message
6876      Close #1
6877      On Error GoTo 0
          
          ' Set up commands for NeosClient
          ' NeosClient call is of the form: NeosClient.py <method> <neosresult.txt> <extra params> >> <logfile>
6878      ModelFilePathName = QuotePath(ConvertHfsPath(ModelFilePathName))
          
          Dim SolverPath As String, NeosClientDir As String
6879      NeosClientDir = JoinPaths(ThisWorkbook.Path, SolverDir)
6880      NeosClientDir = JoinPaths(NeosClientDir, SolverDirMac)
          
6881      GetExistingFilePathName NeosClientDir, "NeosClient.py", SolverPath
6882      SolverPath = QuotePath(ConvertHfsPath(SolverPath))
6883      system ("chmod +x " & SolverPath)
          
          Dim SolutionFilePathName As String
6884      SolutionFilePathName = GetTempFilePath("neosresult.txt")
6885      DeleteFileAndVerify SolutionFilePathName, errorPrefix, "Unable to delete the solution file: " & SolutionFilePathName

          Dim LogFilePathName As String
6886      LogFilePathName = GetTempFilePath("log1.tmp")
6887      DeleteFileAndVerify LogFilePathName, errorPrefix, "Unable to delete the log file: " & LogFilePathName
6888      LogFilePathName = " >> " & QuotePath(ConvertHfsPath(LogFilePathName))

          ' Mac doesn't support modeless forms
          'CallingNeos.Show False
6889      Application.ScreenUpdating = True
6890      Application.Cursor = xlWait ' doesn't seem to work on mac

          ' Run NeosClient.py->send
          Dim result As Boolean
6891      result = OSSolveSync(SolverPath & " send " & QuotePath(ConvertHfsPath(SolutionFilePathName)) & " " & ModelFilePathName, "", "", LogFilePathName)
6892      If Not result Then
6893          GoTo NEOSError
6894      End If
          
          ' Read in job number and password
          Dim jobNumber As String, Password As String
6895      On Error GoTo ErrHandler
6896      Open SolutionFilePathName For Input As #1
6897          Line Input #1, message
6898          jobNumber = Mid(message, Len("jobNumber = ") + 1)
6899          Line Input #1, message
6900          Password = Mid(message, Len("password = ") + 1)
6901      Close #1
6902      On Error GoTo 0
6903      DeleteFileAndVerify SolutionFilePathName, errorPrefix, "Unable to delete the solution file: " & SolutionFilePathName

          ' Loop until job is done
          Dim time As Long, Done As Boolean
6904      Done = False
6905      time = 0
6906      While Done = False
              ' Run NeosClient.py->check
6907          result = OSSolveSync(SolverPath & " check " & QuotePath(ConvertHfsPath(SolutionFilePathName)) & " " & jobNumber & " " & Password, "", "", LogFilePathName)
6908          If Not result Then
6909              GoTo NEOSError
6910          End If
6911          Open SolutionFilePathName For Input As #1
6912              Line Input #1, message
6913          Close #1
6914          DeleteFileAndVerify SolutionFilePathName, errorPrefix, "Unable to delete the solution file: " & SolutionFilePathName

              ' Evaluate result
6915          If message = "Done" Then
6916              Done = True
6917          ElseIf message <> "Waiting" And message <> "Running" Then
6918              GoTo NEOSError
6919          Else
6920              Application.Wait (Now + TimeValue("0:00:01"))
6921              time = time + 1
6922              Application.StatusBar = "OpenSolver: Solving model on NEOS... Time Elapsed: " & time & " seconds"
6923              DoEvents
6924          End If
6925      Wend
          
6926      Application.Cursor = xlDefault
6927      Application.ScreenUpdating = False
          'CallingNeos.Hide
          
          ' Run NeosClient.py->check
6928      result = OSSolveSync(SolverPath & " read " & QuotePath(ConvertHfsPath(SolutionFilePathName)) & " " & jobNumber & " " & Password, "", "", LogFilePathName)
6929      If Not result Then
6930          GoTo NEOSError
6931      End If
          
          ' Read in results from file
6932      On Error GoTo ErrHandler
6933      Open SolutionFilePathName For Input As #1
6934          message = Input$(LOF(1), 1)
6935      Close #1
6936      On Error GoTo 0
          
6937      If left(message, 6) = "Error:" Then
6938          GoTo NEOSError
6939      End If
6940      CallNEOS_Mac = message
6941      Exit Function
          
ErrHandler:
6942      Close #1
6943      errorString = Err.Description
6944      CallNEOS_Mac = ""
6945      Exit Function
          
NEOSError:
6946      Close #1
6947      errorString = "Error contacting NEOS."
6948      CallNEOS_Mac = ""
6949      MsgBox "An error occured when sending file to NEOS. Neos returned: " & message
End Function

' Code by Tim Hastings
Private Function DecodeBase64(ByVal strData As String) As String
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
End Function

' Code by Tim Hastings
Function Stream_BinaryToString(Binary)
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
6961       BinaryStream.position = 0
6962       BinaryStream.Type = adTypeText
           
           'Specify charset For the output text (unicode) data.
6963       BinaryStream.Charset = "us-ascii"
           
           'Open the stream And get text/string data from the object
6964       Stream_BinaryToString = BinaryStream.ReadText
6965       Set BinaryStream = Nothing
End Function

Sub WrapAMPLForNEOS(AmplString As String, Solver As String)
      ' Wraps AMPL in the required XML to send to NEOS
           Dim Category As String, SolverType As String
6966       GetNeosValues Solver, Category, SolverType
           
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
End Sub
