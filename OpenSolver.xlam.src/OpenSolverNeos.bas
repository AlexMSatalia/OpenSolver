Attribute VB_Name = "OpenSolverNeos"
Option Explicit
Function CallNEOS(ModelFilePathName As String, Solver As String, errorString As String) As String
    ' Import file as continuous string
    Dim message As String
    On Error GoTo ErrHandler
    Open ModelFilePathName For Input As #1
        message = Input$(LOF(1), 1)
    Close #1
     
    ' Wrap in XML for AMPL on NEOS
    WrapAMPLForNEOS message, Solver
     
#If Mac Then
    CallNEOS = CallNEOS_Mac(message, errorString)
#Else
    CallNEOS = CallNEOS_Windows(message, errorString)
#End If
    Exit Function
    
ErrHandler:
    Close #1
    Err.Raise Err.Number, Err.Source, Err.Description & IIf(Erl = 0, "", " (at line " & Erl & ")")
     
End Function

Private Function CallNEOS_Windows(message As String, errorString As String)
    ' Server name
    Dim txtURL As String
    txtURL = "http://www.neos-server.org:3332"
    
    ' Late binding so we don't need to add the reference to MSXML, causing a crash on Mac
    Dim objSvrHTTP As Object 'MSXML2.ServerXMLHTTP
    Set objSvrHTTP = CreateObject("MSXML2.ServerXMLHTTP")
    
    ' Set up obj for a POST request
    objSvrHTTP.Open "POST", txtURL, False
    
    ' Clean message up
    message = Replace(message, "<", "&lt;")
    message = Replace(message, ">", "&gt;")
    
    ' Set up message as XML
    message = "<methodCall><methodName>submitJob</methodName><params><param><value><string>" _
       & message & "</string></value></param></params></methodCall>"
    
    ' Send Message to NEOS
    objSvrHTTP.send (message)
    
    ' Extract Job Number
    Dim openingParen As String, closingParen As String, jobNumber As String
    openingParen = InStr(objSvrHTTP.responseText, "<int>")
    closingParen = InStr(objSvrHTTP.responseText, "</int>")
    jobNumber = Mid(objSvrHTTP.responseText, openingParen + Len("<int>"), closingParen - openingParen - Len("<int>"))
    
    If jobNumber = 0 Then
        MsgBox "An error occured when sending file to NEOS."
        GoTo ExitSub
    End If
    
    ' Extract Password
    Dim Password As String
    openingParen = InStr(objSvrHTTP.responseText, "<string>")
    closingParen = InStr(objSvrHTTP.responseText, "</string>")
    Password = Mid(objSvrHTTP.responseText, openingParen + Len("<string>"), closingParen - openingParen - Len("<string>"))
    
    ' Set up Job Status message for XML
    Dim Done As Boolean, result As String
    message = "<methodCall><methodName>getJobStatus</methodName><params><param><value><int>" _
       & jobNumber & "</int></value><value><string>" & Password & _
       "</string></value></param></params></methodCall>"
    Done = False
    
    CallingNeos.Show False
    
    ' Loop until job is done
    Dim time As Long
    time = 0
    While Done = False
        DoEvents
        
        ' Reset obj
        objSvrHTTP.Open "POST", txtURL, False
        
        ' Send message
        objSvrHTTP.send (message)
        
        ' Extract answer
        openingParen = InStr(objSvrHTTP.responseText, "<string>")
        closingParen = InStr(objSvrHTTP.responseText, "</string>")
        result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
        
        ' Evaluate result
        If result = "Done" Then
            Done = True
        ElseIf result <> "Waiting" And result <> "Running" Then
            MsgBox "An error occured when sending file to NEOS. Neos returned: " & result
            GoTo ExitSub
        Else
            Application.Wait (Now + TimeValue("0:00:01"))
            time = time + 1
            Application.StatusBar = "OpenSolver: Solving model on NEOS... Time Elapsed: " & time & " seconds"
            DoEvents
        End If
    Wend
    
    CallingNeos.Hide
    
    ' Set up final message for XML
    message = "<methodCall><methodName>getFinalResults</methodName><params><param><value><int>" & _
              jobNumber & "</int></value></param><param><value><string>" & Password & _
              "</string></value></param></params></methodCall>"
    
    ' Reset obj
    objSvrHTTP.Open "POST", txtURL, False
    
    objSvrHTTP.send (message)
    
    ' Extract Result
    openingParen = InStr(objSvrHTTP.responseText, "<base64>")
    closingParen = InStr(objSvrHTTP.responseText, "</base64>")
    result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
    
    ' The message is returned from NEOS in base 64
    CallNEOS_Windows = DecodeBase64(result)
    
    Exit Function
          
ExitSub:
    errorString = "ExitSub"
    Exit Function
errorHandler:
    errorString = "Error while contacting NEOS."

End Function

Private Function CallNEOS_Mac(message As String, errorString As String)
    ' Mac doesn't have ActiveX so can't use MSXML.
    ' It does have python by default, so we can use python's xmlrpclib to contact NEOS instead.
    ' We delegate all interaction to the NeosClient.py script file.
    Dim errorPrefix As String
    errorPrefix = "Sending model to NEOS"
    
    Dim ModelFilePathName As String
    ModelFilePathName = GetTempFilePath("job.xml")
    DeleteFileAndVerify ModelFilePathName, errorPrefix, "Unable to delete the job file: " & ModelFilePathName
    
    ' Create the job file
    On Error GoTo ErrHandler
    Open ModelFilePathName For Output As #1
    Print #1, message
    Close #1
    On Error GoTo 0
    
    ' Set up commands for NeosClient
    ' NeosClient call is of the form: NeosClient.py <job.xml> <neosresult.txt> > <logfile>
    ModelFilePathName = QuotePath(ConvertHfsPath(ModelFilePathName))
    
    Dim SolverPath As String
    GetExistingFilePathName ThisWorkbook.Path, "NeosClient.py", SolverPath
    SolverPath = QuotePath(ConvertHfsPath(SolverPath))
    system ("chmod +x " & SolverPath)
    
    Dim SolutionFilePathName As String
    SolutionFilePathName = GetTempFilePath("neosresult.txt")
    DeleteFileAndVerify SolutionFilePathName, errorPrefix, "Unable to delete the solution file: " & SolutionFilePathName

    Dim LogFilePathName As String
    LogFilePathName = GetTempFilePath("log1.tmp")
    DeleteFileAndVerify LogFilePathName, errorPrefix, "Unable to delete the log file: " & LogFilePathName
    LogFilePathName = " > " & QuotePath(ConvertHfsPath(LogFilePathName))

    ' Mac doesn't support modal forms
    'CallingNeos.Show False
    Application.Cursor = xlWait

    ' Run NeosClient.py
    Dim result As Boolean
    result = OSSolveSync(SolverPath & " " & ModelFilePathName & " " & QuotePath(ConvertHfsPath(SolutionFilePathName)), "", "", LogFilePathName)
    If Not result Then
        GoTo NEOSError
    End If
    
    Application.Cursor = xlDefault
    
    'CallingNeos.Hide
    
    ' Read in results from file
    On Error GoTo ErrHandler
    Open SolutionFilePathName For Input As #1
        message = Input$(LOF(1), 1)
    Close #1
    On Error GoTo 0
    
    If left(message, 6) = "Error:" Then
        GoTo NEOSError
    End If
    CallNEOS_Mac = message
    Exit Function
    
ErrHandler:
    Close #1
    errorString = Err.Description
    CallNEOS_Mac = ""
    Exit Function
    
NEOSError:
    Close #1
    errorString = "Error contacting NEOS."
    CallNEOS_Mac = ""
    MsgBox "An error occured when sending file to NEOS. Neos returned: " & message
End Function

' Code by Tim Hastings
Private Function DecodeBase64(ByVal strData As String) As String
    Dim objXML As Object 'MSXML2.DOMDocument
    Dim objNode As Object 'MSXML2.IXMLDOMElement
  
    ' Help from MSXML
    Set objXML = CreateObject("MSXML2.DOMDocument")
    Set objNode = objXML.createElement("b64")
    objNode.DataType = "bin.base64"
    objNode.Text = strData
    DecodeBase64 = Stream_BinaryToString(objNode.nodeTypedValue)
  
    ' Clean up
    Set objNode = Nothing
    Set objXML = Nothing
End Function

' Code by Tim Hastings
Function Stream_BinaryToString(Binary)
     Const adTypeText = 2
     Const adTypeBinary = 1
     
     'Create Stream object
     Dim BinaryStream 'As New Stream
     Set BinaryStream = CreateObject("ADODB.Stream")
     
     'Specify stream type - we want To save binary data.
     BinaryStream.Type = adTypeBinary
     
     'Open the stream And write binary data To the object
     BinaryStream.Open
     BinaryStream.Write Binary
     
     'Change stream type To text/string
     BinaryStream.Position = 0
     BinaryStream.Type = adTypeText
     
     'Specify charset For the output text (unicode) data.
     BinaryStream.Charset = "us-ascii"
     
     'Open the stream And get text/string data from the object
     Stream_BinaryToString = BinaryStream.ReadText
     Set BinaryStream = Nothing
End Function

Sub WrapAMPLForNEOS(AmplString As String, Solver As String)
' Wraps AMPL in the required XML to send to NEOS
     Dim Category As String, SolverType As String
     GetNeosValues Solver, Category, SolverType
     
     AmplString = _
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
