Attribute VB_Name = "SolverSolveEngine"
Option Explicit

' Solve Engine constants

Public Const postMethod As String = "POST"
Public Const getMethod As String = "GET"
Public Const putMethod As String = "PUT"
Public Const deleteMethod As String = "DELETE"

Public Const strJobStatusCompleted As String = "completed"
Public Const strJobStatusFailed As String = "failed"
Public Const strJobStatusTranslating As String = "translating"
Public Const strJobStatusStarting As String = "starting"
Public Const strJobStatusStarted As String = "started"
Public Const strJobStatusQueued As String = "queued"
Public Const strJobStatusTimeOut As String = "timeout"
Public Const strJobStatusStopped As String = "stopped"
Public Const strJobStatusCreated As String = "created"

Public Const SolveEngineServer As String = "https://solve.satalia.com/api/v2"
Public Const TrackingId As String = "d5721e0943f8ce3d5a2ee6ec534ca402c5a129a3"

Public Const strFileName As String = "openSolverProblem.lp"
Public Const strAuthTokenFile As String = "AuthToken.txt"
Public Const strLogFileName As String = "SolveEngineLog.txt"

Public Const strJsonKeyId As String = "id"
Public Const strJsonKeyJobId As String = "job_id"
Public Const strJsonKeyStatus As String = "status"
Public Const strJsonKeyResult As String = "result"
Public Const strJsonKeyObjValue As String = "objective_value"
Public Const strJsonKeyVariables As String = "variables"
Public Const strJsonKeyCode As String = "code"
Public Const strJsonKeyMsg As String = "message"
Public Const strJsonKeyValue As String = "value"
Public Const strJsonKeyName As String = "name"
Public Const strJsonKeyProblems As String = "problems"
Public Const strJsonKeyData As String = "data"

Public Const strUrlJobs As String = "/jobs/"
Public Const strUrlInput As String = "/input"
Public Const strUrlResults As String = "/results"
Public Const strUrlStatus As String = "/status"
Public Const strUrlSchedule As String = "/schedule"
Public Const strUrlStop As String = "/stop"

Public Const strSolveStatusSEOptimal As String = "optimal"
Public Const strSolveStatusSEInfeasible As String = "infeasible"
Public Const strSolveStatusSEInterrupted As String = "interrupted"
Public Const strSolveStatusSEFailed As String = "failed"
Public Const strSolveStatusSEUnbounded As String = "unbounded"

' Name of save location for API key
Private Const SolveEngineRegName = "SolveEngineApiKey"
' Attributes that the form needs to get/set
Public SolveEngineFinalResponse As String
Public SolveEngineLpModel As String
Public SolveEngineLogPath As String

Public Function CallSolveEngine(ByVal s As COpenSolver, lpModel As String) As String
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
              
          Dim InteractiveStatus As Boolean
3         InteractiveStatus = Application.Interactive
          
4         SolveEngineLpModel = lpModel
5         SolveEngineLogPath = s.LogFilePathName
          
          Dim errorString As String
6         If s.MinimiseUserInteraction Then
              ' We are running quietly so call solve engine directly
7             SolveEngineFinalResponse = SolveOnSolveEngine(SolveEngineLpModel, SolveEngineLogPath, errorString)
8         Else
              ' We are running interactively so show the form
              Dim frmSolveEngine As FSolveEngine
9             Set frmSolveEngine = New FSolveEngine
              
10            Application.Interactive = True
11            frmSolveEngine.Show
12            Application.Interactive = InteractiveStatus

13            errorString = frmSolveEngine.Tag
14        End If
15        If Len(errorString) > 0 Then
16            If errorString = "Aborted" Then
17                RaiseUserCancelledError
18            Else
19                RaiseGeneralError errorString
20            End If
21        End If
22        CallSolveEngine = SolveEngineFinalResponse
              
ExitFunction:
23        If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
24        Exit Function

ErrorHandler:
25        If Not ReportError("SolverSolveEngine", "CallSolveEngine") Then Resume
26        RaiseError = True
27        GoTo ExitFunction
End Function

' Accessing the saved API key in the registry
Public Function GetSolveEngineApiKey() As String
1         GetSolveEngineApiKey = GetSetting(OpenSolverRegName, PreferencesRegName, SolveEngineRegName, VALUE_IF_MISSING)
End Function
Public Sub SaveSolveEngineApiKey(ApiKey As String)
1         SaveSetting OpenSolverRegName, PreferencesRegName, SolveEngineRegName, ApiKey
End Sub
Public Sub DeleteSolveEngineApiKey()
1         DeleteSetting OpenSolverRegName, PreferencesRegName, SolveEngineRegName
End Sub

Function GetSolveEngineApiKeyOrPrompt() As String
      ' Gets the saved API key and prompts user if no key is saved
          Dim ApiKey As String
1         ApiKey = GetSolveEngineApiKey()
          
2         If ApiKey = VALUE_IF_MISSING Then
3             ApiKey = PromptSolveEngineApiKey()
4         End If
                  
5         GetSolveEngineApiKeyOrPrompt = ApiKey
End Function

Function PromptSolveEngineApiKey() As String
      ' Prompt user to enter API key
          Dim ApiKey As String
1         ApiKey = Application.InputBox( _
              prompt:="Please enter your Satalia SolveEngine API key.", _
              Type:=2, _
              Title:="SolveEngine API Key")
          
2         If ApiKey <> "False" Then SaveSolveEngineApiKey (ApiKey)
          
3         PromptSolveEngineApiKey = ApiKey
End Function

Public Function SolveEngineRequest(method As String, URL As String, ApiKey As String, Optional body As String) As Dictionary
      ' Make an HTTP request to the SolveEngine endpoint and return the result as a JSON dict
          Dim strResp As String
    #If Mac Then
1             strResp = SolveEngineRequest_Mac(method, URL, ApiKey, body)
    #Else
2             strResp = SolveEngineRequest_Win(method, URL, ApiKey, body)
    #End If
3         Set SolveEngineRequest = ParseJson(strResp)
End Function
#If Mac Then
    Private Function SolveEngineRequest_Mac(method As String, URL As String, ApiKey As String, Optional body As String) As String
              Dim strCommand As String
1             strCommand = "/usr/bin/curl " & _
                  "--request " & method & _
                  " --header " & """Authorization:api-key " & ApiKey & """" & _
                  " --header " & """X-tracking-id:" & TrackingId & """" & _
                  IIf(Len(body) > 0, " --data '" & body & "'", "") & _
                  " " & URL
            
2             SolveEngineRequest_Mac = ExecCapture(strCommand)
    End Function
#Else
    Private Function SolveEngineRequest_Win(method As String, URL As String, ApiKey As String, Optional body As String) As String
              Dim xmlRequ As Object
1             Set xmlRequ = CreateObject("MSXML2.ServerXMLHTTP")
              
2             With xmlRequ
3                 .Open method, URL
4                 .setRequestHeader "Authorization", "api-key " & ApiKey
5                 .setRequestHeader "X-tracking-id", TrackingId
                  
6                 If Len(body) > 0 Then .send body Else .send
              
7                 SolveEngineRequest_Win = .responseText
8             End With
    End Function
#End If

Function GetErrorMessage(resp As Dictionary) As String
      ' Get error message from response when the response is not as we expected
1         On Error GoTo BadResponse
          Dim code As String, msg As String
2         code = resp(strJsonKeyCode)
3         msg = resp(strJsonKeyMsg)
          
4         GetErrorMessage = msg
5         Exit Function
          
BadResponse:
6         RaiseGeneralError "Malformed error response: " & vbNewLine & ConvertToJson(resp, 2)
End Function

Public Function SolveOnSolveEngine(lpModel As String, LogPath As String, errorString As String, Optional frmSolveEngine As FSolveEngine = Nothing) As String
' Solve the given LP model on the solve engine
1         On Error GoTo ErrorHandler

    ' Load the API key
    Dim ApiKey As String
2         ApiKey = LoadApiKey()
3         AppendFile LogPath, "Auth token: " & ApiKey
    
    ' Build the message to send problem to Solve Engine
    Dim problemData As String
4         problemData = BuildProblem(lpModel)
    
    Dim jobId As String
5         CheckIfCancel frmSolveEngine, ApiKey, jobId
6         UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Sending model to the SolveEngine..."
    
    ' Send the problem and get the job id
7         jobId = CreateJob(ApiKey, problemData)
8         AppendFile LogPath, "Job ID: " & jobId

9         CheckIfCancel frmSolveEngine, ApiKey, jobId

    ' Start the job
10        ScheduleJob ApiKey, jobId

    ' Wait for the job to finish
11        WaitForAnswer ApiKey, jobId, LogPath, frmSolveEngine
    
12        CheckIfCancel frmSolveEngine, ApiKey, jobId

    ' Get the final results once the job is complete
    Dim FinalResponse As Dictionary
13        Set FinalResponse = GetResults(ApiKey, jobId)
    
14        AppendFile LogPath, ConvertToJson(FinalResponse, 2)
    
15        SolveOnSolveEngine = ConvertToJson(FinalResponse)

ExitFunction:
16        Exit Function
    
ErrorHandler:
      ' We CANNOT raise an error in this function.
      ' It is sometimes called with a form as a conduit, which means that errors can't propogate back to the main thread.
      ' Instead, set the error string, which IS passed back to the main thread by the form.
17        If Not ReportError("SolverSolveEngine", "SolveOnSolveEngine") Then Resume
18        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then GoTo Aborted
19        errorString = OpenSolverErrorHandler.ErrMsg
20        GoTo ExitFunction
          
Aborted:
21        SolveOnSolveEngine = "SolveEngine solve was aborted"
22        errorString = "Aborted"
23        Exit Function
End Function

Sub UpdateStatus(frmSolveEngine As FSolveEngine, message As String)
      ' Update the status of the solve on the status bar, and on the form if it is showing
1         UpdateStatusBar message
2         If Not frmSolveEngine Is Nothing Then
3             frmSolveEngine.UpdateStatus message
4         End If
End Sub

Private Sub CheckIfCancel(frmSolveEngine As FSolveEngine, ApiKey As String, jobId As String)
      ' Check if we need to cancel the job
1         If Not frmSolveEngine Is Nothing Then
2             Interaction.DoEvents 'important in order to catch a click on Cancel button
              
3             If frmSolveEngine.shouldCancel Then
4                 CancelJob ApiKey, jobId
5             End If
6         End If
End Sub

Private Sub CancelJob(ApiKey As String, jobId As String)
      ' Send a request to cancel the job
          Dim resp As Dictionary
1         Set resp = SolveEngineRequest( _
              deleteMethod, _
              SolveEngineServer & strUrlJobs & jobId & strUrlStop, _
              ApiKey)
          
2         If resp.Count > 0 Then
              Dim msg As String
3             msg = GetErrorMessage(resp)
4             If Len(msg) > 0 Then
5                 RaiseGeneralError "The job could not be cancelled. The response was: " & msg
6             Else
7                 RaiseUserCancelledError
8             End If
9         End If
End Sub

Private Function LoadApiKey() As String
      ' Gets the API Key
          Dim RaiseError As Boolean
1         RaiseError = False
2         On Error GoTo ErrorHandler
          
          Dim ApiKey As String
3         ApiKey = GetSolveEngineApiKeyOrPrompt()
          
4         If ApiKey = "False" Then
5             RaiseUserCancelledError
6         ElseIf Len(ApiKey) = 0 Then
7             DeleteSolveEngineApiKey
8             RaiseUserError "No SolveEngine API key was given, so the solve was aborted."
9         End If
          
10        LoadApiKey = ApiKey

ExitFunction:
11        If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
12        Exit Function

ErrorHandler:
13        If Not ReportError("SolveEngineRequest", "IdentifyClient") Then Resume
14        RaiseError = True
15        GoTo ExitFunction
End Function

Public Function BuildProblem(lpModel As String) As String
          ' Encode LP model as base64 and remove weird spaces
          Dim encodedLpModel As String
1         encodedLpModel = Replace(EncodeBase64(lpModel), Chr(10), "")

          Dim problem As Dictionary
2         Set problem = New Dictionary
3         problem.Add strJsonKeyName, strFileName
4         problem.Add strJsonKeyData, encodedLpModel
          
          Dim problems As Collection
5         Set problems = New Collection
6         problems.Add problem
          
          Dim problemDataDict As Dictionary
7         Set problemDataDict = New Dictionary
8         problemDataDict.Add strJsonKeyProblems, problems
          
9         BuildProblem = ConvertToJson(problemDataDict)
End Function

Private Function CreateJob(ApiKey As String, problemData As String) As String
      ' Send the job data to solve engine
          Dim resp As Dictionary
1         Set resp = SolveEngineRequest( _
              postMethod, _
              SolveEngineServer & strUrlJobs, _
              ApiKey, _
              problemData)

          Dim jobId As String
2         If resp.Exists(strJsonKeyId) Then
3             jobId = resp(strJsonKeyId)
4         End If

5         If Len(jobId) = 0 Then
              Dim msg As String
6             msg = GetErrorMessage(resp)
7             Select Case LCase(msg)
              Case "invalid api key":
8                 RaiseUserError "Invalid API key, please check at https://solve.satalia.com/api"
9             Case "unauthorized":
10                RaiseGeneralError "The user is unauthorized."
11            Case Else:
12                RaiseGeneralError "Unknown error while creating the job: " & msg
13            End Select
14        End If
          
15        CreateJob = jobId
End Function

Private Sub ScheduleJob(ApiKey As String, jobId As String)
      ' Start job on solve engine
          Dim body As Dictionary
1         Set body = New Dictionary
2         body.Add strJsonKeyId, jobId

          Dim resp As Dictionary
3         Set resp = SolveEngineRequest( _
              postMethod, _
              SolveEngineServer & strUrlJobs & jobId & strUrlSchedule, _
              ApiKey, _
              ConvertToJson(body))

4         If resp.Count > 0 Then
5             RaiseGeneralError "Unexpected response while scheduling job: " & GetErrorMessage(resp)
6         End If
End Sub

Private Sub WaitForAnswer(ApiKey As String, jobId As String, LogPath As String, frmSolveEngine As FSolveEngine)
      ' Wait for job to complete
          Dim resp As Dictionary, status As String, timeElapsed As Long

1         Do
2             Set resp = SolveEngineRequest( _
                  getMethod, _
                  SolveEngineServer & strUrlJobs & jobId & strUrlStatus, _
                  ApiKey)

3             status = GetStatus(resp)
4             AppendFile LogPath, status

5             Select Case status
                  Case strJobStatusCompleted, strJobStatusFailed:
6                     Exit Do
7                 Case strJobStatusTranslating:
8                     UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Translating problem..."
9                 Case strJobStatusStarted, strJobStatusStarting:
10                    UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Solving in progress, " & _
                                                   "time elapsed " & timeElapsed & "s..."
11                Case strJobStatusQueued:
12                    UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Waiting in SolveEngine queue..."
13                Case strJobStatusCreated:
14                    UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Waiting in SolveEngine queue..."
15                Case Else:
16                    RaiseGeneralError "Unknown status: " & status
17            End Select
                  
18            CheckIfCancel frmSolveEngine, ApiKey, jobId
              
19            mSleep 1000 ' Wait 1 second
20            timeElapsed = timeElapsed + 1
21        Loop

22        UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Loading solution..."
End Sub

Private Function GetStatus(resp As Dictionary) As String
          Dim status As String
1         If resp.Exists(strJsonKeyStatus) Then
2             status = resp(strJsonKeyStatus)
3         Else
4             RaiseGeneralError "Unexpected response while checking status of job: " & GetErrorMessage(resp)
5         End If
          
6         GetStatus = status
End Function

Private Function GetResults(ApiKey As String, jobId As String) As Dictionary
      ' Get final results
          Dim resp As Dictionary
1         Set resp = SolveEngineRequest( _
              getMethod, _
              SolveEngineServer & strUrlJobs & jobId & strUrlResults, _
              ApiKey)
              
2         Set GetResults = resp
End Function
