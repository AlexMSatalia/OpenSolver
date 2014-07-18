Attribute VB_Name = "OpenSolverNeos"
Option Explicit
Function CallNEOS(ModelFilePathName As String, Solver As String, errorString As String) As String
     Dim objSvrHTTP As Object 'MSXML2.ServerXMLHTTP
     Dim message As String, txtURL As String
     Dim Done As Boolean, result As String
     Dim openingParen As String, closingParen As String, jobNumber As String, Password As String, solutionFile As String, solution As String
     Dim i As Integer, LinearSolveStatusString As String, var As Long
     
     ' Server name
     txtURL = "http://www.neos-server.org:3332"
     Set objSvrHTTP = CreateObject("MSXML2.ServerXMLHTTP")
     
     ' Set up obj for a POST request
     objSvrHTTP.Open "POST", txtURL, False
     
     ' Import file as continuous string
     Open ModelFilePathName For Input As #1
         message = Input$(LOF(1), 1)
     Close #1
     
     ' Wrap in XML for AMPL on NEOS
     WrapAMPLForNEOS message, Solver
     
     ' Clean message up
     message = Replace(message, "<", "&lt;")
     message = Replace(message, ">", "&gt;")
     
     ' Set up message as XML
     message = "<methodCall><methodName>submitJob</methodName><params><param><value><string>" _
        & message & "</string></value></param></params></methodCall>"
     
     ' Send Message to NEOS
     objSvrHTTP.send (message)
     
     ' Extract Job Number
     openingParen = InStr(objSvrHTTP.responseText, "<int>")
     closingParen = InStr(objSvrHTTP.responseText, "</int>")
     jobNumber = Mid(objSvrHTTP.responseText, openingParen + Len("<int>"), closingParen - openingParen - Len("<int>"))
     
     If jobNumber = 0 Then
         MsgBox "An error occured when sending file to NEOS."
         GoTo ExitSub
     End If
     
     ' Extract Password
     openingParen = InStr(objSvrHTTP.responseText, "<string>")
     closingParen = InStr(objSvrHTTP.responseText, "</string>")
     Password = Mid(objSvrHTTP.responseText, openingParen + Len("<string>"), closingParen - openingParen - Len("<string>"))
     
     ' Set up Job Status message for XML
     message = "<methodCall><methodName>getJobStatus</methodName><params><param><value><int>" _
        & jobNumber & "</int></value><value><string>" & Password & _
        "</string></value></param></params></methodCall>"
     Done = False
     
     CallingNeos.Show False
     
     ' Loop until job is done
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
         End If
     Wend
     
     CallingNeos.Hide
     
     ' Set up final message for XML
     message = "<methodCall><methodName>getFinalResults</methodName><params><param><value><int>" _
        & jobNumber & "</int></value></param><param><value><string>" & Password & _
        "</string></value></param></params></methodCall>"
     
     ' Reset obj
     objSvrHTTP.Open "POST", txtURL, False
    
     objSvrHTTP.send (message)
     
     ' Extract Result
     openingParen = InStr(objSvrHTTP.responseText, "<base64>")
     closingParen = InStr(objSvrHTTP.responseText, "</base64>")
     result = Mid(objSvrHTTP.responseText, openingParen + 8, closingParen - openingParen - 8)
     
     ' The message is returned from NEOS in base 64
     CallNEOS = DecodeBase64(result)
     
15870          Exit Function
          
ExitSub:
15880     errorString = "ExitSub"
15890     Exit Function
errorHandler:
15900     errorString = "Error contacting NEOS."

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
