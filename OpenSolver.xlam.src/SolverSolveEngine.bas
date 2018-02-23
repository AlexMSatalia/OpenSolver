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
    RaiseError = False
    On Error GoTo ErrorHandler
        
    Dim InteractiveStatus As Boolean
    InteractiveStatus = Application.Interactive
    
    SolveEngineLpModel = lpModel
    SolveEngineLogPath = s.LogFilePathName
    
    Dim errorString As String
    If s.MinimiseUserInteraction Then
        ' We are running quietly so call solve engine directly
        SolveEngineFinalResponse = SolveOnSolveEngine(SolveEngineLpModel, SolveEngineLogPath, errorString)
    Else
        ' We are running interactively so show the form
        Dim frmSolveEngine As FSolveEngine
        Set frmSolveEngine = New FSolveEngine
        
        Application.Interactive = True
        frmSolveEngine.Show
        Application.Interactive = InteractiveStatus

        errorString = frmSolveEngine.Tag
    End If
    If Len(errorString) > 0 Then
        If errorString = "Aborted" Then
            RaiseUserCancelledError
        Else
            RaiseGeneralError errorString
        End If
    End If
    CallSolveEngine = SolveEngineFinalResponse
        
ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("SolverSolveEngine", "CallSolveEngine") Then Resume
    RaiseError = True
    GoTo ExitFunction
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
    ApiKey = GetSolveEngineApiKey()
    
    If ApiKey = VALUE_IF_MISSING Then
        ApiKey = PromptSolveEngineApiKey()
    End If
            
    GetSolveEngineApiKeyOrPrompt = ApiKey
End Function

Function PromptSolveEngineApiKey() As String
' Prompt user to enter API key
    Dim ApiKey As String
    ApiKey = Application.InputBox( _
        prompt:="Please enter your Satalia SolveEngine API key.", _
        Type:=2, _
        Title:="SolveEngine API Key")
    
    If ApiKey <> "False" Then SaveSolveEngineApiKey (ApiKey)
    
    PromptSolveEngineApiKey = ApiKey
End Function

Public Function SolveEngineRequest(method As String, URL As String, ApiKey As String, Optional body As String) As Dictionary
' Make an HTTP request to the SolveEngine endpoint and return the result as a JSON dict
    Dim strResp As String
    #If Mac Then
        strResp = SolveEngineRequest_Mac(method, URL, ApiKey, body)
    #Else
        strResp = SolveEngineRequest_Win(method, URL, ApiKey, body)
    #End If
    Set SolveEngineRequest = ParseJson(strResp)
End Function
#If Mac Then
    Private Function SolveEngineRequest_Mac(method As String, URL As String, ApiKey As String, Optional body As String) As String
        Dim strCommand As String
        strCommand = "/usr/bin/curl " & _
            "--request " & method & _
            " --header " & """Authorization:api-key " & ApiKey & """" & _
            " --header " & """X-tracking-id:" & TrackingId & """" & _
            IIf(Len(body) > 0, " --data '" & body & "'", "") & _
            " " & URL
      
        SolveEngineRequest_Mac = ExecCapture(strCommand)
    End Function
#Else
    Private Function SolveEngineRequest_Win(method As String, URL As String, ApiKey As String, Optional body As String) As String
        Dim xmlRequ As Object
        Set xmlRequ = CreateObject("MSXML2.ServerXMLHTTP")
        
        With xmlRequ
            .Open method, URL
            .setRequestHeader "Authorization", "api-key " & ApiKey
            .setRequestHeader "X-tracking-id", TrackingId
            
            If Len(body) > 0 Then .send body Else .send
        
            SolveEngineRequest_Win = .responseText
        End With
    End Function
#End If

Function GetErrorMessage(resp As Dictionary) As String
' Get error message from response when the response is not as we expected
    On Error GoTo BadResponse
    Dim code As String, msg As String
    code = resp(strJsonKeyCode)
    msg = resp(strJsonKeyMsg)
    
    GetErrorMessage = msg
    Exit Function
    
BadResponse:
    RaiseGeneralError "Malformed error response: " & vbNewLine & ConvertToJson(resp, 2)
End Function

Public Function SolveOnSolveEngine(lpModel As String, LogPath As String, errorString As String, Optional frmSolveEngine As FSolveEngine = Nothing) As String
' Solve the given LP model on the solve engine
    On Error GoTo ErrorHandler

    ' Load the API key
    Dim ApiKey As String
    ApiKey = LoadApiKey()
    AppendFile LogPath, "Auth token: " & ApiKey
    
    ' Build the message to send problem to Solve Engine
    Dim problemData As String
    problemData = BuildProblem(lpModel)
    
    Dim jobId As String
    CheckIfCancel frmSolveEngine, ApiKey, jobId
    UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Sending model to the SolveEngine..."
    
    ' Send the problem and get the job id
    jobId = CreateJob(ApiKey, problemData)
    AppendFile LogPath, "Job ID: " & jobId

    CheckIfCancel frmSolveEngine, ApiKey, jobId

    ' Start the job
    ScheduleJob ApiKey, jobId

    ' Wait for the job to finish
    WaitForAnswer ApiKey, jobId, LogPath, frmSolveEngine
    
    CheckIfCancel frmSolveEngine, ApiKey, jobId

    ' Get the final results once the job is complete
    Dim FinalResponse As Dictionary
    Set FinalResponse = GetResults(ApiKey, jobId)
    
    AppendFile LogPath, ConvertToJson(FinalResponse, 2)
    
    SolveOnSolveEngine = ConvertToJson(FinalResponse)

ExitFunction:
    Exit Function
    
ErrorHandler:
      ' We CANNOT raise an error in this function.
      ' It is sometimes called with a form as a conduit, which means that errors can't propogate back to the main thread.
      ' Instead, set the error string, which IS passed back to the main thread by the form.
25        If Not ReportError("SolverSolveEngine", "SolveOnSolveEngine") Then Resume
26        If OpenSolverErrorHandler.ErrNum = OpenSolver_UserCancelledError Then GoTo Aborted
27        errorString = OpenSolverErrorHandler.ErrMsg
28        GoTo ExitFunction
          
Aborted:
29        SolveOnSolveEngine = "SolveEngine solve was aborted"
30        errorString = "Aborted"
31        Exit Function
End Function

Sub UpdateStatus(frmSolveEngine As FSolveEngine, message As String)
' Update the status of the solve on the status bar, and on the form if it is showing
    UpdateStatusBar message
    If Not frmSolveEngine Is Nothing Then
        frmSolveEngine.UpdateStatus message
    End If
End Sub

Private Sub CheckIfCancel(frmSolveEngine As FSolveEngine, ApiKey As String, jobId As String)
' Check if we need to cancel the job
    If Not frmSolveEngine Is Nothing Then
        Interaction.DoEvents 'important in order to catch a click on Cancel button
        
        If frmSolveEngine.shouldCancel Then
            CancelJob ApiKey, jobId
        End If
    End If
End Sub

Private Sub CancelJob(ApiKey As String, jobId As String)
' Send a request to cancel the job
    Dim resp As Dictionary
    Set resp = SolveEngineRequest( _
        deleteMethod, _
        SolveEngineServer & strUrlJobs & jobId & strUrlStop, _
        ApiKey)
    
    If resp.Count > 0 Then
        Dim msg As String
        msg = GetErrorMessage(resp)
        If Len(msg) > 0 Then
            RaiseGeneralError "The job could not be cancelled. The response was: " & msg
        Else
            RaiseUserCancelledError
        End If
    End If
End Sub

Private Function LoadApiKey() As String
' Gets the API Key
    Dim RaiseError As Boolean
    RaiseError = False
    On Error GoTo ErrorHandler
    
    Dim ApiKey As String
    ApiKey = GetSolveEngineApiKeyOrPrompt()
    
    If ApiKey = "False" Then
        RaiseUserCancelledError
    ElseIf Len(ApiKey) = 0 Then
        DeleteSolveEngineApiKey
        RaiseUserError "No SolveEngine API key was given, so the solve was aborted."
    End If
    
    LoadApiKey = ApiKey

ExitFunction:
    If RaiseError Then Err.Raise OpenSolverErrorHandler.ErrNum, Description:=OpenSolverErrorHandler.ErrMsg
    Exit Function

ErrorHandler:
    If Not ReportError("SolveEngineRequest", "IdentifyClient") Then Resume
    RaiseError = True
    GoTo ExitFunction
End Function

Public Function BuildProblem(lpModel As String) As String
    ' Encode LP model as base64 and remove weird spaces
    Dim encodedLpModel As String
    encodedLpModel = Replace(EncodeBase64(lpModel), Chr(10), "")

    Dim problem As Dictionary
    Set problem = New Dictionary
    problem.Add strJsonKeyName, strFileName
    problem.Add strJsonKeyData, encodedLpModel
    
    Dim problems As Collection
    Set problems = New Collection
    problems.Add problem
    
    Dim problemDataDict As Dictionary
    Set problemDataDict = New Dictionary
    problemDataDict.Add strJsonKeyProblems, problems
    
    BuildProblem = ConvertToJson(problemDataDict)
End Function

Private Function CreateJob(ApiKey As String, problemData As String) As String
' Send the job data to solve engine
    Dim resp As Dictionary
    Set resp = SolveEngineRequest( _
        postMethod, _
        SolveEngineServer & strUrlJobs, _
        ApiKey, _
        problemData)

    Dim jobId As String
    If resp.Exists(strJsonKeyId) Then
        jobId = resp(strJsonKeyId)
    End If

    If Len(jobId) = 0 Then
        Dim msg As String
        msg = GetErrorMessage(resp)
        Select Case LCase(msg)
        Case "invalid api key":
            RaiseUserError "Invalid API key, please check at https://solve.satalia.com/api"
        Case "unauthorized":
            RaiseGeneralError "The user is unauthorized."
        Case Else:
            RaiseGeneralError "Unknown error while creating the job: " & msg
        End Select
    End If
    
    CreateJob = jobId
End Function

Private Sub ScheduleJob(ApiKey As String, jobId As String)
' Start job on solve engine
    Dim body As Dictionary
    Set body = New Dictionary
    body.Add strJsonKeyId, jobId

    Dim resp As Dictionary
    Set resp = SolveEngineRequest( _
        postMethod, _
        SolveEngineServer & strUrlJobs & jobId & strUrlSchedule, _
        ApiKey, _
        ConvertToJson(body))

    If resp.Count > 0 Then
        RaiseGeneralError "Unexpected response while scheduling job: " & GetErrorMessage(resp)
    End If
End Sub

Private Sub WaitForAnswer(ApiKey As String, jobId As String, LogPath As String, frmSolveEngine As FSolveEngine)
' Wait for job to complete
    Dim resp As Dictionary, status As String, timeElapsed As Long

    Do
        Set resp = SolveEngineRequest( _
            getMethod, _
            SolveEngineServer & strUrlJobs & jobId & strUrlStatus, _
            ApiKey)

        status = GetStatus(resp)
        AppendFile LogPath, status

        Select Case status
            Case strJobStatusCompleted, strJobStatusFailed:
                Exit Do
            Case strJobStatusTranslating:
                UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Translating problem..."
            Case strJobStatusStarted, strJobStatusStarting:
                UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Solving in progress, " & _
                                             "time elapsed " & timeElapsed & "s..."
            Case strJobStatusQueued:
                UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Waiting in SolveEngine queue..."
            Case strJobStatusCreated:
                UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Waiting in SolveEngine queue..."
            Case Else:
                RaiseGeneralError "Unknown status: " & status
        End Select
            
        CheckIfCancel frmSolveEngine, ApiKey, jobId
        
        mSleep 1000 ' Wait 1 second
        timeElapsed = timeElapsed + 1
    Loop

    UpdateStatus frmSolveEngine, "Solving model on SolveEngine: Loading solution..."
End Sub

Private Function GetStatus(resp As Dictionary) As String
    Dim status As String
    If resp.Exists(strJsonKeyStatus) Then
        status = resp(strJsonKeyStatus)
    Else
        RaiseGeneralError "Unexpected response while checking status of job: " & GetErrorMessage(resp)
    End If
    
    GetStatus = status
End Function

Private Function GetResults(ApiKey As String, jobId As String) As Dictionary
' Get final results
    Dim resp As Dictionary
    Set resp = SolveEngineRequest( _
        getMethod, _
        SolveEngineServer & strUrlJobs & jobId & strUrlResults, _
        ApiKey)
        
    Set GetResults = resp
End Function
