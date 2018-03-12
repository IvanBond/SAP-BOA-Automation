VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Option Compare Text

Rem Introduction
    ' Author: Ivan Bondarenko
    ' Initial release: 2017-01
    ' https://bondarenkoivan.wordpress.com
    ' http://excel.city
    ' https://linkedin.com/in/bondarenkoivan/en

    ' Big Thanks to Andres Merino for ideas and testing!

    ' update 2017-05-24
    ' now macro handles cases with SAPSetVariables only
    ' no need to refresh twice in such cases
     
    ' update 2017-08-01
    ' release of version with additional actions
    ' added actions: 'Email', 'Save As', 'Refresh All', 'Email on Success', 'Run Macro'
    ' planned: Publish to PowerBI, Update PP+ThinkCell, Update PowerPoint Tables
    
    ' update 2017-08-13
    ' fixed serious issue with subsequent refresh
    ' callback now used only for first refresh of DS
    ' every next refresh macro uses SAPSetVariable is DS is active
    ' if not active - firstly refesh it, then set filters
    ' reason: Disable / Enable BOA is not stable action, it crashes Excel too often.
    ' What else? Many small fixes, renamed all objects
    ' now prefix for tables and named ranges is common: CP_ (from Control Panel)
    '
    
    ' update 2017-08-19
    ' improved logic of FillHeaders sub
    
    ' update 2017-09-18
    ' Run Actions after each Scope - new parameter
    ' many minor changes
    
    ' update 2017-11-26
    ' new parameters for some actions
    ' named ranges for relative dates
    
    ' update 2018-03-12
    ' Email on Success renamed to Email
    ' minor fixes
    ' added comments
    
    ' declaration of function that is used in Function CreatePath
    ' possibly will be used for logging
    
Rem Declaration
    #If VBA7 Then
        Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib _
            "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
        Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
    #Else
        Private Declare Function MakeSureDirectoryPathExists Lib _
            "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
        Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
    #End If

    Const ForAppending = 8 ' for OpenTextFile
    Const ForReading = 1
    
    Const sTableScopesName = "CP_SCOPES"
    Const sTableDataSourcesName = "CP_DATA_SOURCES"
    Const sTableVariablesName = "CP_VARIABLES"
    Const STableActions = "CP_ACTIONS"
    
    Public Logs_Enabled As Boolean
    Public LogFile As String

    Public objFSO As Object
    Public objLog As Object
    Dim arrDS ' array for data sources
    Dim currentDS As String ' current data source ID
    Dim currentDS_index As Long ' index in array arrDS
    Dim bManualRefresh As Boolean
    Dim bBOAEnabled As Boolean ' var to store initial state of BOA
    ' Enable / Disable Boa causes Excel to crash when Excel starts with BOA activated
    ' however, works fine when BOA is not enabled by default
    ' In addition, subsequent refresh doesn't trigger CallBack
    Dim bBeforeFirstPromptsDisplayRegistered As Boolean
    Dim bBeforeFirstPromptsDisplaySubExecuted As Boolean
    Dim dicRefreshedDataSources ' dictionary with refreshed datasources
    ' for refreshed data sources callback BeforeFirstPromptsDisplay won't work
    
    Dim currentActionRowId As Long
    Dim bCallActionFromDStable As Boolean

    Dim bMultiLogonMode As Boolean
    ' special mode when refresh is done data source be data source
    ' and after each data source macro run SAPLogoff
    ' to be sure that next data source will be refreshed with defined language
    ' and defined user

    Enum enumDataSource ' used to reference to elements of arrDS array
        DS_ID = 0
        DS_Refresh = 1
        DS_Language = 2
        DS_System = 3
        DS_User = 4
        DS_MacrosAfter = 5
        DS_Client = 6
        DS_CalculateAfter = 7
        DS_Row = 8
        DS_Sheet = 9
        ' Repeat Refresh can be changed by macro / formula, so do not keep it in this array
        DS_Crosstab = 10
        ' Clear Crosstab - is used by action "Clear Crosstabs", or after all actions
        DS_FillHeaders = 11
        DS_CallActions = 12
    End Enum

    Enum enumMailImportance
        Low = 0
        Normal = 1
        High = 2
    End Enum

    Dim dummy As Integer
Rem End of declaration

Function Manual_Refresh() As Boolean
    bManualRefresh = True
    Manual_Refresh = Refresh
End Function

Private Sub Worksheet_Activate()
    If Me.Names("CP_PASSWORDS_PATH").RefersToRange.Value = vbNullString Then
        Me.Names("CP_PASSWORDS_PATH").RefersToRange.Value = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\Passwords.txt"
    End If
    If Me.Names("CP_GENERAL_USER").RefersToRange.Value = vbNullString Then
        Me.Names("CP_GENERAL_USER").RefersToRange.Value = LCase(Environ("username"))
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If Me.Names("CP_PASSWORDS_PATH").RefersToRange.Value = vbNullString Then
        Me.Names("CP_PASSWORDS_PATH").RefersToRange.Value = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\Passwords.txt"
    End If
    If Me.Names("CP_GENERAL_USER").RefersToRange.Value = vbNullString Then
        Me.Names("CP_GENERAL_USER").RefersToRange.Value = LCase(Environ("username"))
    End If
End Sub

Private Sub DefineGlobalVariables()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Left(ThisWorkbook.Path, 4) <> "http" Then
'        Logs_Enabled = True ' TODO - add cell with parameter
        LogFile = ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".log"
    Else
        LogFile = Environ("temp") & "\" & ThisWorkbook.Name & ".log"
    End If
End Sub

Function Refresh() As Boolean
    Dim StartPoint As Double
    Dim cell_scope As Range
    Dim i As Long
    Dim m As Integer
    Dim arrMacrosBefore
    Dim arrMacrosAfter
    Dim arrMacrosAfterDSRefresh
    Dim arrActions
    Dim Attempt As Long
    Dim bYetAnotherAttempt As Boolean
    Dim bSAPSetVariablesOnly As Boolean ' when scope's filters don't contain SAPSetFilter command
    Dim lScopeRow As Long
    Dim iRet As Integer
    Dim InitialCalculationState As Long
    Dim bScopeFailed  As Boolean
    
    On Error GoTo ErrHandler
    
    If Me.ListObjects("CP_SCOPES").DataBodyRange Is Nothing Then
        If bManualRefresh Then
            MsgBox "No scopes defined!", vbCritical
        End If
        Exit Function
    End If
    
    StartPoint = Now()
    InitialCalculationState = Application.Calculation
    Application.Calculation = xlCalculationManual
    ThisWorkbook.Windows(1).Activate ' otherwise Logon fails, if another wb is active
    
    Me.Names("CP_LAST_REPORT_DATE").RefersToRange.Value = _
        Me.Names("CP_REPORT_DATE").RefersToRange.Value
    
    If IsEmpty(dicRefreshedDataSources) Then
        Set dicRefreshedDataSources = CreateObject("Scripting.Dictionary")
        dicRefreshedDataSources.comparemode = 1 ' TextCompare
    End If
    
    Application.StatusBar = "Enabling BOA..."
    Call EnableBOA ' does nothing if addin is already enabled
    
    Application.StatusBar = "Defining global variables..."
    Call DefineGlobalVariables
        
    Application.StatusBar = "Executing macros before Refresh..."
    arrMacrosBefore = Split(Me.Names("CP_MACROS_BEFORE").RefersToRange.Value, ",")
    
    For m = 0 To UBound(arrMacrosBefore)
        On Error Resume Next
        Application.StatusBar = "Executing macro " & Trim(arrMacrosBefore(m)) & "..."
        Application.Run Trim(arrMacrosBefore(m))
        If Err.Number <> 0 Then
            Debug.Print Now, "Couldn't run macro " & Trim(arrMacrosBefore(m))
            GoTo ErrHandler
        End If
        On Error GoTo 0
        
        ' ? do we need to calc after each macro? - macro owner decides
    Next m
    Application.StatusBar = vbNullString
    
    On Error GoTo ErrHandler
    ' initial calculate
    With Application
        .StatusBar = "Initial calculation..."
        .Calculate
        .StatusBar = vbNullString
    End With
    
    For lScopeRow = 1 To Me.ListObjects("CP_SCOPES").DataBodyRange.Rows.Count
        
        Me.ListObjects("CP_SCOPES").DataBodyRange.Calculate
        
        ' if 'Refresh All Scopes' mode is enabled - change main Scope cell
        If Me.Names("CP_REFRESH_ALL_SCOPES").RefersToRange.Value = "Y" Then
            If Me.ListObjects("CP_SCOPES").ListColumns("Enabled?").DataBodyRange.Cells(lScopeRow, 1).Value <> "Y" Then
                GoTo Next_Scope
            Else
                Application.EnableEvents = False
                Me.Names("CP_SCOPE").RefersToRange.Value = _
                    "'" & CStr(Me.ListObjects("CP_SCOPES").ListColumns("Scope").DataBodyRange.Cells(lScopeRow, 1).Value)
                Application.EnableEvents = True
                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": calculating..."
                Application.Calculate ' just in case
                Application.StatusBar = vbNullString
            End If
        Else
        ' leave Scope cell as is
        End If
        
        ' refresh screen
        If bManualRefresh Then
            Application.ScreenUpdating = True
            WaitSeconds 1
            Application.ScreenUpdating = False
        End If
        
        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Collecting data sources..."
        Call GetArrDS  ' populate array with data sources
        Application.StatusBar = vbNullString
        
        On Error Resume Next
        Debug.Print Now, "# Data sources: "; UBound(arrDS)
        If Err.Number <> 0 Then
            Err.Clear
            GoTo Skipped_Data_Sources
        End If
        On Error GoTo ErrHandler
                
        ' old version - check if any Data Sources defined
'        If WorksheetFunction.CountIf(Me.ListObjects(sTableDataSourcesName).ListColumns("Scope").DataBodyRange, _
'            CStr(Me.Names("CP_SCOPE").RefersToRange.Value)) = 0 Then
'            ' no data sources defined for scope
'            If bManualRefresh Then
'                MsgBox "No Data Sources specified for scope '" & Me.Names("CP_SCOPE").RefersToRange.Value & "'." & vbCrLf _
'                    & "Disable Scope or add Data Sources.", vbExclamation
'            Else
'                Debug.Print Now, "No Data Sources specified for scope " & Me.Names("CP_SCOPE").RefersToRange.Value
'            End If
'            bScopeFailed = True
'            GoTo Next_Scope
'        End If
        ' No check for 'data sources defined' to support scenario, when we need only Actions
        ' this is not for refresh of data sources, but rather for misuse of this tool
        ' e.g. define list of Scopes and turn on "Refresh All Scopes"
        ' after each Scope macro can execute actions, like refresh + Save As
        ' it allows to generate files from one worksheet
        
        
        ' when common connection parameters are used - enough to logon once
        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Logging On..."
        If Not LogOn_DataSources Then
            Debug.Print Now, "Couldn't LogOn"
            If bManualRefresh Then
                Stop
            End If
            GoTo ErrHandler
        End If
        Application.StatusBar = vbNullString
        
        ' refresh all enabled data sources in the specified scope using specified order
        For i = 1 To UBound(arrDS, 1)
            ' array elements are ordered
            ' macro refreshes only those that enabled
            If arrDS(i, enumDataSource.DS_Refresh) = "Y" Then
                Attempt = 0 ' attempt to refresh in case of delay (if no new data)
                bYetAnotherAttempt = False
                bSAPSetVariablesOnly = IsSAPSetVariablesOnly(CStr(arrDS(i, enumDataSource.DS_ID)))
                currentDS = CStr(arrDS(i, enumDataSource.DS_ID))
                currentDS_index = i ' index in array, not a row of sheet!

                ' Jump to new refresh attempt if RepeatRefresh is enabled
Start_Refresh:
                On Error GoTo ErrHandler
                
                ' if another attempt - we have to logon again
                ' as for safety macro does Logoff (avoid kick from server)
                If bYetAnotherAttempt Then
                    Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Logging On..."
                    If Not LogOn_DataSources Then
                        Debug.Print Now, "Couldn't LogOn"
                        If bManualRefresh Then
                            Stop
                        End If
                        GoTo ErrHandler
                    End If
                    Application.StatusBar = vbNullString
                End If
                
                ' if within one Scope we should refresh several data sources with different languages
                ' e.g. everything is available in local language except one data source, which should be refreshed in EN
                If bMultiLogonMode Then
                    If i = 1 Then ' if first data source
                    ' logon with specified Language / User / Client
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Logging On " & arrDS(i, enumDataSource.DS_ID)
                        If Run("SAPLogon", arrDS(i, enumDataSource.DS_ID), arrDS(i, enumDataSource.DS_Client), arrDS(i, enumDataSource.DS_User), _
                                GetPassword(CStr(arrDS(i, enumDataSource.DS_System)), CStr(arrDS(i, enumDataSource.DS_User))), _
                                arrDS(i, enumDataSource.DS_Language)) = 0 Then
                            
                            Debug.Print Now, "Couldn't LogOn to " & CStr(arrDS(i, enumDataSource.DS_System))
                            If bManualRefresh Then
                                Stop
                            End If
                            GoTo ErrHandler
                        End If
                    Else
                        ' if one of previous connection parameters differ from current
                        ' in the end of prev cycle macro logged off - check end of loop area for MultiLogon case
                        If arrDS(i, enumDataSource.DS_Language) <> arrDS(i - 1, enumDataSource.DS_Language) Or _
                            arrDS(i, enumDataSource.DS_System) <> arrDS(i - 1, enumDataSource.DS_System) Or _
                            arrDS(i, enumDataSource.DS_Client) <> arrDS(i - 1, enumDataSource.DS_Client) Or _
                            arrDS(i, enumDataSource.DS_User) <> arrDS(i - 1, enumDataSource.DS_User) Then
                            
                            Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Logging On " & arrDS(i, enumDataSource.DS_ID)
                            If Run("SAPLogon", arrDS(i, enumDataSource.DS_ID), arrDS(i, enumDataSource.DS_Client), arrDS(i, enumDataSource.DS_User), _
                                    GetPassword(CStr(arrDS(i, enumDataSource.DS_System)), CStr(arrDS(i, enumDataSource.DS_User))), _
                                    arrDS(i, enumDataSource.DS_Language)) = 0 Then
                                
                                Debug.Print Now, "Couldn't LogOn to " & CStr(arrDS(i, enumDataSource.DS_System))
                                If bManualRefresh Then
                                    Stop
                                End If
                                GoTo ErrHandler
                            End If
                        End If
                        
                    End If ' if first data source in array
                Else
                    ' not bMultiLogonMode
                    ' logon was done during LogOn_DataSources
                End If ' MultiLogon check
                
                ' Logon is done at this step
                ' after Logon - Initial Refresh (for SAPSetFilter)
                If Not bSAPSetVariablesOnly Then
                    ' SAPSetVariableOnly is Scope-depending
                    If Not Application.Run("SAPGetProperty", "IsDataSourceActive", CStr(arrDS(i, enumDataSource.DS_ID))) Then
                        ' if SAPSetFilter is used then we need "Initial Refresh" before ApplyScopeFilters
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                        If Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID)) = 0 Then
                            Debug.Print Now, "Couldn't Refresh " & CStr(arrDS(i, enumDataSource.DS_ID))
                            GoTo ErrHandler
                        End If
                        Application.StatusBar = vbNullString
                        bBeforeFirstPromptsDisplaySubExecuted = False ' if callback was called - do not track it here
                        
                        ' add to dictionary
                        If Not dicRefreshedDataSources.exists(CStr(arrDS(i, enumDataSource.DS_ID))) Then
                            dicRefreshedDataSources.Add CStr(arrDS(i, enumDataSource.DS_ID)), CStr(arrDS(i, enumDataSource.DS_ID))
                        End If
                    End If
                End If ' not bSAPSetVariablesOnly
                
                ' just in case, if user set pause
                ' without it VBA cannot use com-addin's API
                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                Call Run("SAPExecuteCommand", "PauseVariableSubmit", "Off") ' unpause submit of variables
                Call Run("SAPSetRefreshBehaviour", "On") ' triggers RefreshData
                Application.StatusBar = vbNullString
                        
                ' set filters / variables
                ' AnotherAttempt - is a second or third attempt after successful first refresh (where filters were set)
                ' first attempt is Scope-depending. It is not the same as first refresh of data source within session
                If Not bYetAnotherAttempt Then
                    ' note:
                    ' if first attempt (not the same as first refresh of data source) - we have to apply filters / variables
                    ' note: if second / third ... attempt - no need to set variables again
                    
                    ' if active - just apply scope filters
                    If Application.Run("SAPGetProperty", "IsDataSourceActive", CStr(arrDS(i, enumDataSource.DS_ID))) Then
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & " - Applying filters to " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                        ' during ApplyScopeFilters macro will
                        ' set PauseVariableSubmit to On, and SAPSetRefreshBehaviour to Off
                        If Not ApplyScopeFilters(CStr(Me.Names("CP_SCOPE").RefersToRange.Value), CStr(arrDS(i, enumDataSource.DS_ID))) Then
                            Debug.Print Now, "Couldn't Apply filters for " & CStr(arrDS(i, enumDataSource.DS_ID))
                            If bManualRefresh Then
                                Stop
                            End If
                        End If
                        
                        ' unpause
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                        Call Run("SAPExecuteCommand", "PauseVariableSubmit", "Off") ' unpause submit of variables
                        Call Run("SAPSetRefreshBehaviour", "On") ' triggers RefreshData
                        Application.StatusBar = vbNullString
                    Else
                        ' DS is not active - probably after logoff / logon or the very first refresh
                        ' the very first refresh of data source (callback) OR consequent refresh
                        If dicRefreshedDataSources.exists(CStr(arrDS(i, enumDataSource.DS_ID))) Then
                            ' callback will not work for consequent refresh
                            ' as DS is not Active - we have to refresh it
                            ' in old version: we used Disable BOA + Enable BOA to reset callback,
                            ' but it crashes Excel in most of the cases
                            ' new method can lead to double refresh, but no other stable choice
                            Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                            If Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID)) = 0 Then
                                Debug.Print Now, "Couldn't Refresh " & CStr(arrDS(i, enumDataSource.DS_ID))
                                GoTo ErrHandler
                            End If
                            ' no need to add to the dictionary
                            bBeforeFirstPromptsDisplaySubExecuted = False
                            
                            ' after refresh we can set filters
                            Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Applying filters to " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                            ' during ApplyScopeFilters macro will
                            ' set PauseVariableSubmit to On, and SAPSetRefreshBehaviour to Off
                            If Not ApplyScopeFilters(CStr(Me.Names("CP_SCOPE").RefersToRange.Value), CStr(arrDS(i, enumDataSource.DS_ID))) Then
                                Debug.Print Now, "Couldn't Apply filters for " & CStr(arrDS(i, enumDataSource.DS_ID))
                                If bManualRefresh Then
                                    Stop
                                End If
                            End If
                            
                            ' unpause
                            Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                            Call Run("SAPExecuteCommand", "PauseVariableSubmit", "Off") ' unpause submit of variables
                            Call Run("SAPSetRefreshBehaviour", "On") ' triggers RefreshData
                            Application.StatusBar = vbNullString
                        Else
                            ' doesn't exists in dictionary
                            
                            If bSAPSetVariablesOnly Then
                                ' the very first refresh - use callback, it should work
                                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Registering callback BeforeFirstPromptsDisplay..."
                                Call Callback_BeforeFirstPromptsDisplay_Reg
                                
                                ' just refresh and callback will set filters
                                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                                iRet = Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID))
                                Debug.Print Now, "Refresh", CStr(arrDS(i, enumDataSource.DS_ID)), iRet
                                
                                ' check whether callback was called
                                If bBeforeFirstPromptsDisplayRegistered And (Not bBeforeFirstPromptsDisplaySubExecuted) Then
                                    ' variables were not applied
                                    Debug.Print Now, "Callback didn't work"
                                    ' still have to apply filters - only using double refresh
                                    ' this never happend to me, but in case of issue of BOA...
                                    If bManualRefresh Then
                                        Stop
                                    End If
                                End If

                                ' unregistering is not necessary, but just in case. It takes no time
                                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Unregistering callback BeforeFirstPromptsDisplay..."
                                Call Callback_BeforeFirstPromptsDisplay_UnReg
                                Application.StatusBar = vbNullString
                            Else
                                ' have SAPSetFilter
                                ' if it is first attempt - call ApplyScopeFilters
                                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Applying filters to " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                            
                                ' during ApplyScopeFilters macro will
                                ' set PauseVariableSubmit to On, and SAPSetRefreshBehaviour to Off
                                If Not ApplyScopeFilters(CStr(Me.Names("CP_SCOPE").RefersToRange.Value), CStr(arrDS(i, enumDataSource.DS_ID))) Then
                                    Debug.Print Now, "Couldn't Apply filters for " & CStr(arrDS(i, enumDataSource.DS_ID))
                                    GoTo ErrHandler
                                End If
                                Application.StatusBar = vbNullString
                                
                                ' unpause
                                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                                Call Run("SAPExecuteCommand", "PauseVariableSubmit", "Off") ' unpause submit of variables
                                Call Run("SAPSetRefreshBehaviour", "On") ' triggers RefreshData
                                Application.StatusBar = vbNullString
                            End If ' If bSAPSetVariablesOnly Then
                        End If ' if dicRefreshedDataSources.Exists( cstr(arrDS(i, enumDataSource.DS_ID)) )  then
                    End If ' if Application.Run("SAPGetProperty", "IsDataSourceActive", CStr(arrDS(i, enumDataSource.DS_ID))) then
                Else
                    ' 2nd, 3rd... attempt
                    ' just refresh, all filters were set for first refresh
                    Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Refreshing " & arrDS(i, enumDataSource.DS_ID) & "...(" & i & "/" & UBound(arrDS) & ")"
                    Call Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID))
                    Application.StatusBar = vbNullString
                End If ' If Not bYetAnotherAttempt Then
                
                bBeforeFirstPromptsDisplaySubExecuted = False
                
                ' since program is here - DS was refreshed
                If Not dicRefreshedDataSources.exists(CStr(arrDS(i, enumDataSource.DS_ID))) Then
                    ' register fact of refresh - add to dictionary
                    dicRefreshedDataSources.Add CStr(arrDS(i, enumDataSource.DS_ID)), CStr(arrDS(i, enumDataSource.DS_ID))
                End If
                
                If bMultiLogonMode Then
                    ' if next DS has different connection parameters - logoff
                    If i < UBound(arrDS) Then
                        If arrDS(i, enumDataSource.DS_Language) <> arrDS(i + 1, enumDataSource.DS_Language) Or _
                            arrDS(i, enumDataSource.DS_System) <> arrDS(i + 1, enumDataSource.DS_System) Or _
                            arrDS(i, enumDataSource.DS_Client) <> arrDS(i + 1, enumDataSource.DS_Client) Or _
                            arrDS(i, enumDataSource.DS_User) <> arrDS(i + 1, enumDataSource.DS_User) Then
                            
                            Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Logging Off..."
                            Application.EnableEvents = False ' for some reason LogOff triggers SelectionChange event
                            Call LogOffBOA
                            Application.EnableEvents = True
                            Application.StatusBar = vbNullString
                        End If
                    End If
                End If
                
                ' **************************** Extra actions for data source ****************************
                
                ' Fill Headers
                If arrDS(i, enumDataSource.DS_FillHeaders) = "Y" Then
                    Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Filling headers for " & CStr(arrDS(i, enumDataSource.DS_ID))
                    Call FillHeaders(CStr(arrDS(i, enumDataSource.DS_Crosstab)))
                    Application.StatusBar = vbNullString
                End If
                
                ' calculate if user specified this parameter
                If arrDS(i, enumDataSource.DS_CalculateAfter) = "Y" Then
                    Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Calculating after " & CStr(arrDS(i, enumDataSource.DS_ID))
                    'Application.CalculationInterruptKey = xlNoKey
                    Application.Calculate ' trigger calculation of formulas. Possibly, next report has filter based on the refreshed report
                    'Application.CalculationInterruptKey = xlAnyKey
                    Application.StatusBar = vbNullString
                End If
                
                Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Executing Macros After " & CStr(arrDS(i, enumDataSource.DS_ID))
                If arrDS(i, enumDataSource.DS_MacrosAfter) <> vbNullString Then
                    arrMacrosAfterDSRefresh = Split(arrDS(i, enumDataSource.DS_MacrosAfter), ",")
                    
                    For m = 0 To UBound(arrMacrosAfterDSRefresh)
                        On Error Resume Next
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Executing macro " & Trim(arrMacrosAfterDSRefresh(m)) & "..."
                        Application.Run Trim(arrMacrosAfterDSRefresh(m))
                        If Err.Number <> 0 Then
                            Debug.Print Now, "Couldn't run macro " & Trim(arrMacrosAfterDSRefresh(m))
                            GoTo ErrHandler
                        End If
                        On Error GoTo 0
                    Next m ' next macro
                End If ' If arrDS(i, enumDataSource.DS_MacrosAfter) <> vbNullString Then
                Application.StatusBar = vbNullString
                
                ' Handle Delay scenario
                Me.Names("CP_DELAY").RefersToRange.Calculate
                ' user could call a sub, which change CONTROL_PANEL_DELAY cell value
                ' check it - Check if no new data
                If Me.Names("CP_DELAY").RefersToRange.Value = "Y" Then
                    ' check if have attempts
                    ' starts from 0 for each data source
                    Attempt = Attempt + 1
                    If Attempt > Me.Names("CP_ATTEMPTS").RefersToRange.Value Then
                        Debug.Print Now, "Attempts limit exceed. " & arrDS(i, enumDataSource.DS_ID) & " has no new data."
                        GoTo ErrHandler
                    Else
                        ' if trigger - wait
                        ' log off to prevent kick from server
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Logging Off..."
                        Application.EnableEvents = False ' for some reason LogOff triggers SelectionChange event
                        Call LogOffBOA
                        Application.EnableEvents = True
                        
                        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Delay till " & Format((Now() + TimeValue("00:" & Right("0" & _
                                CStr(Me.Names("CP_DELAY_FOR").RefersToRange.Value), 2) & ":00")), "yyyy-MM-dd hh:mm:ss")
                        
                        Debug.Print Now, "Delay till " & Format((Now() + TimeValue("00:" & Right("0" & _
                                CStr(Me.Names("CP_DELAY_FOR").RefersToRange.Value), 2) & ":00")), "yyyy-MM-dd hh:mm:ss")
                                                                
                        WaitSeconds 60 * Me.Names("CP_DELAY_FOR").RefersToRange.Value
                        
                        Application.StatusBar = vbNullString
                        
                        bYetAnotherAttempt = True
                        ' refresh same data source again - and make same check again
                        GoTo Start_Refresh
                    End If ' If Attempt > Me.Names("CP_ATTEMPTS").RefersToRange.Value Then
                End If ' If Me.Names("CP_DELAY").RefersToRange.Value = "Y" Then
                
                ' Call Action
                ' call specified actions
                If CStr(arrDS(i, enumDataSource.DS_CallActions)) <> vbNullString Then
                    bCallActionFromDStable = True
                    arrActions = Split(CStr(arrDS(i, enumDataSource.DS_CallActions)), ",")
                    For m = LBound(arrActions) To UBound(arrActions)
                        Call Actions_Handler(Trim(arrActions(m)))
                    Next m
                    bCallActionFromDStable = False
                End If
                
                ' if Repeat refresh
                With Me.ListObjects(sTableDataSourcesName)
                    ' calculate to get state of Repeat Refresh
                    Me.Cells(arrDS(i, enumDataSource.DS_Row), .ListColumns("Repeat Refresh").Range.Column).Calculate
                    If Me.Cells(arrDS(i, enumDataSource.DS_Row), .ListColumns("Repeat Refresh").Range.Column).Value = "Y" Then
                        
                        GoTo Start_Refresh
                    End If
                End With
                    
            End If ' If arrDS(i, enumDataSource.DS_Refresh) = "Y" Then
            
Next_Data_Source:
        Next i ' next Data source

Skipped_Data_Sources:
        ' if 'Refresh All Scopes' mode is disabled - then quit from loop
        If Me.Names("CP_REFRESH_ALL_SCOPES").RefersToRange.Value <> "Y" Then
            Exit For ' Next lScopeRow
        End If
        
Next_Scope:

        If Not bScopeFailed Then
            If Me.Names("CP_RUN_ACTIONS_EACH_SCOPE").RefersToRange.Value = "Y" Then
                Application.StatusBar = "Actions after " & Me.Names("CP_SCOPE").RefersToRange.Value
                Call Actions_Handler
            End If
        End If
        
        Application.StatusBar = Me.Names("CP_SCOPE").RefersToRange.Value & ": Completed"
    Next lScopeRow
    
    arrMacrosAfter = Split(Me.Names("CP_MACROS_AFTER").RefersToRange.Value, ",")
    For m = 0 To UBound(arrMacrosAfter)
        Application.StatusBar = "Executing macros after Refresh..."
        On Error Resume Next
        Run Trim(arrMacrosAfter(m))
        If Err.Number <> 0 Then
            Debug.Print Now, "Couldn't run macro " & arrMacrosAfterDSRefresh(m)
            GoTo ErrHandler
        End If
        On Error GoTo 0
    Next m
    
    ' resulting calculation
    Application.StatusBar = "Final calculation..."
    Application.Calculate
    Application.StatusBar = vbNullString
    
    Me.Names("CP_LAST_REFRESH_DATETIME").RefersToRange.Value = Now()
    
    Debug.Print Now, "Refresh time: " & Format(Now() - StartPoint, "hh:mm:ss")
    
    ' Call Write_Log("END # Overall execution time # " & _
        Round((Now() - StartPoint) * 3600 * 24 / 60, 0) & "m " & _
        FormatNumber(Round((Now() - StartPoint) * 3600 * 24, 0) Mod 60, 0) & "s")
    
    ' if Refresh for All scopes enabled and Actions after each scope
    ' Actions Handler was called after last enabled scope
    ' no need to call it again
    If Me.Names("CP_REFRESH_ALL_SCOPES").RefersToRange.Value = "Y" And _
        Me.Names("CP_RUN_ACTIONS_EACH_SCOPE").RefersToRange.Value = "Y" Then
        ' skip
    Else
        Call Actions_Handler
    End If
    
    Refresh = True
            
Exit_Function:
    On Error Resume Next
    
    Erase arrDS
    
    If Not bManualRefresh Then
        Application.StatusBar = "Logging off..."
        Call LogOffBOA ' just in case
        
        Application.StatusBar = "Disabling BOA..."
        ' Disabling is needed for automation
        ' otherwise Excel application cannot be closed
        ' BO Analysis addin requires something...
        Call DisableBOA
        Application.StatusBar = vbNullString
    Else
        'Application.StatusBar = "Logging off..."
        'Call Run("SAPLogOff", True) ' re-connect
        'Application.StatusBar = vbNullString
    End If
    
    With Application
        .Cursor = xlDefault
        .EnableEvents = True
        .ScreenUpdating = True ' just in case
        .StatusBar = "Final calculation..." ' just in case
        .Calculation = InitialCalculationState
        .StatusBar = vbNullString
    End With
    
    Exit Function
    
ErrHandler:
    ' send email?
    If Err.Number <> 0 Then
        Debug.Print Now, "Refresh", Err.Number & ": " & Err.Description
        Err.Clear
        Application.Cursor = xlDefault
        Application.ScreenUpdating = True
        If bManualRefresh Then Stop
    End If
    GoTo Exit_Function
    Resume
End Function

Private Function IsBOAEnabled() As Boolean
    Dim addIn As COMAddIn
    On Error Resume Next
    For Each addIn In Application.COMAddIns
        If addIn.progID = "SapExcelAddIn" Then
            If addIn.Connect = True Then
                IsBOAEnabled = True
                Exit Function
            End If
        End If
    Next
End Function

Private Sub EnableBOA()
    Dim addIn As COMAddIn
    On Error Resume Next
    For Each addIn In Application.COMAddIns
        If addIn.progID = "SapExcelAddIn" Then
            If addIn.Connect = False Then
                addIn.Connect = True
                Exit Sub
            End If
        End If
    Next
End Sub

Private Sub DisableBOA()
    Dim addIn As COMAddIn
    On Error Resume Next
    For Each addIn In Application.COMAddIns
        If addIn.progID = "SapExcelAddIn" Then
            addIn.Connect = False
            Exit Sub
        End If
    Next
End Sub

Private Function LogOn_DataSources() As Boolean
    Dim arrSystems
    Dim arrDataSources
    Dim arrLanguages
    Dim i         As Long
    Dim r         As Long
    Dim sLanguage As String
    Dim sUser     As String
    Dim sSystem   As String
    Dim sClient   As String
    Dim dicSystems As Object
    Dim Key
    Dim n As Long
    
    On Error GoTo ErrHandler
    
    Set dicSystems = CreateObject("Scripting.Dictionary")
    dicSystems.comparemode = 1 ' TextCompare
    
    With Me.ListObjects(sTableDataSourcesName)
        If .DataBodyRange Is Nothing Then
            LogOn_DataSources = True
            Exit Function
        Else
            ' check if any data sources in scope
            ' number of data sources in Scope
            n = WorksheetFunction.CountIf(Me.ListObjects(sTableDataSourcesName).ListColumns("Scope").DataBodyRange, _
                Me.Names("CP_SCOPE").RefersToRange.Value)
            
            If n = 0 Then
                Debug.Print Now, "No data sources defined for scope " & Me.Names("CP_SCOPE").RefersToRange.Value
                If bManualRefresh Then
                    Stop
                End If
            End If
            
            ' check if several languages in Scope
            For r = 1 To .DataBodyRange.Rows.Count
                If CStr(.ListColumns("Scope").DataBodyRange.Cells(r, 1).Value) = _
                                CStr(Me.Names("CP_SCOPE").RefersToRange.Value) Then
                    ' Connection parameters - System, Client, Language
                    ' in 99.9% cases workbook use sames values for all three
                    ' however, sometimes we can meet workbooks with data from different systems
                    ' or queries that uses different languages
                    ' if so - I call it Multi-Logon mode
                    ' as it requires to logoff from prev session and logon with new parameters
                    
                    If sLanguage = vbNullString And .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                    ' on first step - remember value in variable
                        sLanguage = IIf(.ListColumns("Language").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                            Me.Names("CP_GENERAL_LANGUAGE").RefersToRange.Value, _
                                            .ListColumns("Language").DataBodyRange.Cells(r, 1).Value)
                    Else
                        ' second and further steps - compare each row with value in variable
                        If sLanguage <> IIf(.ListColumns("Language").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                            Me.Names("CP_GENERAL_LANGUAGE").RefersToRange.Value, _
                                            .ListColumns("Language").DataBodyRange.Cells(r, 1).Value) And _
                                .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                            bMultiLogonMode = True
                            LogOn_DataSources = True
                            Exit Function
                        End If
                    End If ' empty language
                    
                    If sUser = vbNullString And .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                        ' on first step - remember value in variable
                        sUser = IIf(.ListColumns("User").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("CP_GENERAL_USER").RefersToRange.Value, _
                                    .ListColumns("User").DataBodyRange.Cells(r, 1).Value)
                    Else
                        ' second and further steps - compare each row with value in variable
                        If sUser <> IIf(.ListColumns("User").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("CP_GENERAL_USER").RefersToRange.Value, _
                                    .ListColumns("User").DataBodyRange.Cells(r, 1).Value) And _
                                .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                            bMultiLogonMode = True
                            LogOn_DataSources = True
                            Exit Function
                        End If
                    End If
                    
                    ' check if different clients used in Scope
                    If sClient = vbNullString And .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                        ' on first step - remember value in variable
                        sClient = IIf(.ListColumns("Client").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("CP_GENERAL_CLIENT").RefersToRange.Value, _
                                    .ListColumns("Client").DataBodyRange.Cells(r, 1).Value)
                    Else
                        ' second and further steps - compare each row with value in variable
                        If sClient <> IIf(.ListColumns("Client").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("CP_GENERAL_CLIENT").RefersToRange.Value, _
                                    .ListColumns("Client").DataBodyRange.Cells(r, 1).Value) And _
                                .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                            bMultiLogonMode = True
                            LogOn_DataSources = True
                            Exit Function
                        End If
                    End If
                    
                    ' Dictionary of Systems
                    sSystem = IIf(.ListColumns("System").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("CP_GENERAL_SYSTEM").RefersToRange.Value, _
                                        .ListColumns("System").DataBodyRange.Cells(r, 1).Value)
                    If Not dicSystems.exists(sSystem & "+" & sClient) Then
                        dicSystems.Add sSystem & "+" & sClient, .ListColumns("Data Source").DataBodyRange.Cells(r, 1).Value
                    End If
                
                End If ' row in Scope, System and Language is not null
            Next r ' row of table 'data sources'
            
            ' if here - same language in all rows
            bMultiLogonMode = False ' just in case
            ' only one language is used - that defined in General settings
            ' call usual logon - for any data source
            
            ' we checked that Scope contains same User and same Language
            ' it means that we can logon to all systems at once
            For Each Key In dicSystems.keys
                'Debug.Print Now, CStr(dicSystems.Item(Key)), Application.Run("SAPGetProperty", "IsDataSourceActive", CStr(dicSystems.Item(Key)))
                
                ' 'arrDS(1, enumDataSource.DS_ID),
                If Run("SapLogon", _
                    CStr(dicSystems.Item(Key)), _
                    Split(Key, "+")(1), _
                    sUser, _
                    GetPassword(CStr(Split(Key, "+")(0)), _
                        sUser), _
                    sLanguage) = 0 Then
                    
                    Debug.Print Now, "Couldn't log on to the system '" & CStr(dicSystems.Item(Key)) & "'"
                    If bManualRefresh Then
                        Stop
                    End If
                    Exit Function ' leave Log_On_Systems as False
                Else
                    LogOn_DataSources = True
                    Exit Function
                End If
            Next Key ' of dicSystems - dictionary with list of systems in Scope
            
        End If ' if table 'data source' is not empty
    End With

Exit_Sub:
    Exit Function
ErrHandler:
    Debug.Print Now, "LogOn Data Sources", Err.Number, Err.Description
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    If bManualRefresh Then Stop
    
    GoTo Exit_Sub
    Resume
End Function

Private Function GetPassword(system As String, username As String) As String
    ' Expected txt file with structure
    ' SYSTEM;USERNAME;PASSWORD
    ' Sample:
    ' P01;ivan;12345
    ' P01; ivan; 23456
    ' spaces will be trimmed, case doesn't matter (only in password)
    
    Dim objFile As Object
    Dim strLine As String
    
    On Error GoTo ErrHandler
    If objFSO Is Nothing Then Set objFSO = CreateObject("Scripting.FileSystemObject")
    
    On Error Resume Next
    Set objFile = objFSO.OpenTextFile(Me.Names("CP_PASSWORDS_PATH").RefersToRange.Value, ForReading)
    If Err.Number <> 0 Then
        Err.Clear
        Application.EnableEvents = False
        Me.Cells(13, 5).Value = "File with passwords not found!"
        Me.Cells(13, 5).Font.Color = 3
        GoTo Exit_Function
    End If
    Application.EnableEvents = False
    Me.Cells(13, 5).Value = vbNullString
    Me.Cells(13, 5).ClearFormats
        
    On Error GoTo ErrHandler
    Do Until objFile.AtEndOfStream
        strLine = objFile.ReadLine
        If InStr(1, strLine, system, vbTextCompare) > 0 And _
            InStr(1, strLine, username, vbTextCompare) > 0 Then
            GetPassword = Trim(Mid(strLine, InStr(InStr(strLine, ";") + 1, strLine, ";") + 1))
            Exit Do ' password found
        End If
    Loop
        
    objFile.Close

Exit_Function:
    Set objFSO = Nothing
    Set objFile = Nothing
    Exit Function

ErrHandler:
    Debug.Print Now, "GetPassword", Err.Number & ": " & Err.Description
    Err.Clear
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    ' do not stop when file not found
    'If bManualRefresh Then Stop
    
    GoTo Exit_Function
    Resume
End Function

Private Function Get_Array_From_Range(rng_ref)
    Dim rng_tmp As Range
    Dim arr_tmp
    If Left(rng_ref, 2) <> "'[" Then ' if not external link
        ThisWorkbook.Windows(1).Activate ' Range fails if another workbook is activated
    End If
    Set rng_tmp = Application.Range(rng_ref)
    If rng_tmp.Cells.Count = 1 Then
        ReDim arr_tmp(1 To 1, 1 To 1)
        arr_tmp(1, 1) = rng_tmp.Value
    Else
        arr_tmp = Application.Range(rng_ref)
    End If
    Get_Array_From_Range = arr_tmp
End Function

Private Sub Write_Log(Msg As String)
    
    If Logs_Enabled Then
        On Error Resume Next
        If objFSO Is Nothing Then _
            Set objFSO = CreateObject("Scripting.FileSystemObject")
        If objLog Is Nothing Then _
            Set objLog = CreateObject("Scripting.FileSystemObject")
        
        ' CreatePath LogsFolder ' in case we need to create special path
        If Not objFSO.FileExists(LogFile) Then objFSO.CreateTextFile LogFile
        Set objLog = objFSO.OpenTextFile(LogFile, ForAppending)
        objLog.WriteLine (Now() & "# " & Msg)
        objLog.Close
    End If
End Sub

Private Function ApplyScopeFilters(scope As String, data_source As String) As Boolean
    Dim arr_var_scope
    Dim arr_var_ds
    Dim arr_var_command
    Dim arr_var_field
    Dim arr_var_value
    Dim arr_var_ref_type
    Dim i As Long
    Dim variable_row As Long
    Dim arrVariableValue
    Dim sVariableValue As String
    
    On Error GoTo ErrHandler
    
    If Me.ListObjects(sTableVariablesName).DataBodyRange Is Nothing Then
        ApplyScopeFilters = True
        Exit Function
    End If
    
    ' define arrays with Variables table columns - faster then read each cell
    With Me.ListObjects(sTableVariablesName)
        ' if sure that more than 1 variable
        ' explanation: http://www.cpearson.com/excel/ArraysAndRanges.aspx
        ' arr_var_ds = .ListColumns("Data Source ID").DataBodyRange.Value
        ' if sure that Excel runs only one workbook - local address
        ' arr_var_ds = Get_Array_From_Range("'" & Split(.ListColumns("Data Source ID").DataBodyRange.Address(External:=True), "]")(1))
        
        arr_var_scope = Get_Array_From_Range(.ListColumns("Scope").DataBodyRange.Address(external:=True))
        arr_var_ds = Get_Array_From_Range(.ListColumns("Data Source").DataBodyRange.Address(external:=True))
        arr_var_command = Get_Array_From_Range(.ListColumns("Command").DataBodyRange.Address(external:=True))
        arr_var_field = Get_Array_From_Range(.ListColumns("Field").DataBodyRange.Address(external:=True))
        arr_var_value = Get_Array_From_Range(.ListColumns("Value").DataBodyRange.Address(external:=True))
        arr_var_ref_type = Get_Array_From_Range(.ListColumns("Reference Type").DataBodyRange.Address(external:=True))
    End With
    
    ' loop over rows of Variables table
    ' check Scope in each row
    
    Call Run("SAPSetRefreshBehaviour", "Off")
    
    ' two loops, because SetFilters should go before SetVariables (due to PauseVariableSubmit behavior)
    ' Set Filters
    For variable_row = 1 To UBound(arr_var_scope)
        If arr_var_scope(variable_row, 1) = scope And arr_var_ds(variable_row, 1) = data_source Then
            If arr_var_command(variable_row, 1) = "SAPSetFilter" Then
                If arr_var_ref_type(variable_row, 1) = "Range" Then
                    
                    ' easier then INPUT_STRING_AS_ARRAY?
                    ' in this way we can control length of resulting variable, as there is limitaion equal to ~max possible chars in standard Excel cell
                    ' ~32000 (including spaces and delimiters)
                    On Error Resume Next
                    Debug.Print Application.Range(CStr(arr_var_value(variable_row, 1))).Rows.Count
                    If Err.Number <> 0 Then
                        Debug.Print Now, "Range '" & CStr(arr_var_value(variable_row, 1)) & "' wasn't found"
                        If bManualRefresh Then
                            Stop
                        Else
                            GoTo ErrHandler
                        End If
                    End If
                    On Error GoTo ErrHandler
                    
                    arrVariableValue = Get_Array_From_Range(Application.Range(CStr(arr_var_value(variable_row, 1))).Address(external:=True))
                    For i = 1 To UBound(arrVariableValue)
                        sVariableValue = sVariableValue & "; " & arrVariableValue(i, 1)
                    Next i
                    sVariableValue = Mid(sVariableValue, 3)
                    
                    '    sVariableValue = "1; 2; 3; .... 90000; ... "
                    '    sVariableValue = "!1; !2; !3; .... !90000; ... "
                Else
                    sVariableValue = CStr(arr_var_value(variable_row, 1))
                End If
                
                If Run(arr_var_command(variable_row, 1), _
                        arr_var_ds(variable_row, 1), _
                        arr_var_field(variable_row, 1), _
                        sVariableValue) = 0 Then
                    
                    Debug.Print Now, "Couldn't apply filter for variable " & CStr(arr_var_field(variable_row, 1)) & " : " & sVariableValue
                    Exit Function
                Else
                    Application.EnableEvents = False
                    ' if successful - keep values
                    Me.ListObjects(sTableVariablesName).ListColumns("Last Refresh").DataBodyRange.Cells(variable_row, 1).Value = Now()
                    Me.ListObjects(sTableVariablesName).ListColumns("Last Used Value").DataBodyRange.Cells(variable_row, 1).Value = "'" & sVariableValue
                    Application.EnableEvents = True
                End If ' successful command
                sVariableValue = vbNullString
            End If ' var command
        End If ' var scope and data source
    Next variable_row
    
    ' Set Variables
    Call Run("SAPExecuteCommand", "PauseVariableSubmit", "On") ' stop re-calculation after Variable change
    
    For variable_row = 1 To UBound(arr_var_ds)
        If arr_var_scope(variable_row, 1) = scope And arr_var_ds(variable_row, 1) = data_source Then
            If arr_var_command(variable_row, 1) = "SAPSetVariable" Then
                If arr_var_ref_type(variable_row, 1) = "Range" Then
                    
                    On Error Resume Next
                    Debug.Print Application.Range(CStr(arr_var_value(variable_row, 1))).Rows.Count
                    If Err.Number <> 0 Then
                        Debug.Print Now, "Range '" & CStr(arr_var_value(variable_row, 1)) & "' wasn't found"
                        If bManualRefresh Then
                            Stop
                        Else
                            GoTo ErrHandler
                        End If
                    End If
                    On Error GoTo ErrHandler
                    
                    ' easier then INPUT_STRING_AS_ARRAY?
                    arrVariableValue = Get_Array_From_Range(Application.Range(CStr(arr_var_value(variable_row, 1))).Address(external:=True))
                    For i = 1 To UBound(arrVariableValue)
                        sVariableValue = sVariableValue & "; " & arrVariableValue(i, 1)
                    Next i
                    sVariableValue = Mid(sVariableValue, 3)
                    
                    '    sVariableValue = "1; 2; 3; .... 90000; ... "
                    '    sVariableValue = "!1; !2; !3; .... !90000; ... "
                Else
                    sVariableValue = CStr(arr_var_value(variable_row, 1))
                End If
                                                
                If Run(arr_var_command(variable_row, 1), _
                        arr_var_field(variable_row, 1), _
                        sVariableValue, _
                        "INPUT_STRING", _
                        arr_var_ds(variable_row, 1)) = 0 Then
                    
                    Debug.Print Now, "Couldn't apply filter for variable " & CStr(arr_var_field(variable_row, 1)) & " : " & sVariableValue
                    GoTo Exit_Sub
                Else
                    Application.EnableEvents = False
                    ' if successful - keep values
                    Me.ListObjects(sTableVariablesName).ListColumns("Last Refresh").DataBodyRange.Cells(variable_row, 1).Value = Now()
                    Me.ListObjects(sTableVariablesName).ListColumns("Last Used Value").DataBodyRange.Cells(variable_row, 1).Value = "'" & sVariableValue
                    Application.EnableEvents = True
                End If
                sVariableValue = vbNullString
            End If ' var command
        End If ' var scope and data source
    Next variable_row
    
    ApplyScopeFilters = True
Exit_Sub:
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    Exit Function

ErrHandler:
    Debug.Print Now, "Apply Filters", Err.Number, Err.Description
    Err.Clear
    
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    If bManualRefresh Then Stop
    
    GoTo Exit_Sub
    Resume ' for debug
End Function

Private Sub GetArrDS()
    ' sub collects all active data sources
    Dim n As Long
    Dim i As Long
    Dim k As Long
    
    On Error GoTo ErrHandler
    
    If Me.ListObjects(sTableDataSourcesName).DataBodyRange Is Nothing Then Exit Sub
    If Me.Names("CP_SCOPE").RefersToRange.Value = vbNullString Then Exit Sub
    
    ' just sort by Order, if no Order - default order of execution
    Call SortDataSources
    
    ' number of data sources in Scope
    n = WorksheetFunction.CountIf(Me.ListObjects(sTableDataSourcesName).ListColumns("Scope").DataBodyRange, _
            Me.Names("CP_SCOPE").RefersToRange.Value)
    
    ' n must be equal to MAX value in column Order for corresponding Scope
    k = 1
    ReDim arrDS(n, 12) ' size must be equal to number of parameters, check enumDataSource
    With Me.ListObjects(sTableDataSourcesName)
        For i = 1 To .DataBodyRange.Rows.Count
            ' fill array according to order of execution
            
            ' expected correct value in column Order
            If CStr(.ListColumns("Scope").DataBodyRange.Cells(i, 1).Value) = _
                CStr(Me.Names("CP_SCOPE").RefersToRange.Value) Then
                
                arrDS(k, enumDataSource.DS_ID) = _
                    .ListColumns("Data Source").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_Refresh) = _
                    .ListColumns("Refresh?").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_MacrosAfter) = _
                    .ListColumns("Macros After").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_CalculateAfter) = _
                    .ListColumns("Calculate After").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_Row) = _
                    .ListColumns("Scope").DataBodyRange.Cells(i, 1).Row
                
                arrDS(k, enumDataSource.DS_Sheet) = _
                    .ListColumns("Sheet").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_Crosstab) = _
                    .ListColumns("Crosstab").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_CallActions) = _
                    .ListColumns("Call Actions").DataBodyRange.Cells(i, 1).Value
                
                arrDS(k, enumDataSource.DS_FillHeaders) = _
                    .ListColumns("Fill Headers").DataBodyRange.Cells(i, 1).Value
                
                ' if empty settings - use General options
                arrDS(k, enumDataSource.DS_System) = _
                    IIf(.ListColumns("System").DataBodyRange.Cells(i, 1).Value = vbNullString, _
                        Me.Names("CP_GENERAL_SYSTEM").RefersToRange.Value, _
                            .ListColumns("System").DataBodyRange.Cells(i, 1).Value)
                
                arrDS(k, enumDataSource.DS_Language) = _
                    IIf(.ListColumns("Language").DataBodyRange.Cells(i, 1) = vbNullString, _
                        Me.Names("CP_GENERAL_LANGUAGE").RefersToRange.Value, _
                            .ListColumns("Language").DataBodyRange.Cells(i, 1).Value)
                    
                arrDS(k, enumDataSource.DS_User) = _
                    IIf(.ListColumns("User").DataBodyRange.Cells(i, 1) = vbNullString, _
                        Me.Names("CP_GENERAL_USER").RefersToRange.Value, _
                            .ListColumns("User").DataBodyRange.Cells(i, 1).Value)
            
                arrDS(k, enumDataSource.DS_Client) = _
                    IIf(.ListColumns("Client").DataBodyRange.Cells(i, 1) = vbNullString, _
                        Me.Names("CP_GENERAL_CLIENT").RefersToRange.Value, _
                            .ListColumns("Client").DataBodyRange.Cells(i, 1).Value)
            
                k = k + 1
            End If
        Next i
    End With ' DATA SOURCES table
    
Exit_Sub:
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "GetArrDS", Err.Number & ": " & Err.Description
    Err.Clear
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume 'debug
End Sub

Sub LogOffBOA()
    On Error Resume Next
    Call Run("SAPLogOff", False) ' disconnect
End Sub

Private Function CreatePath(NewPath As String) As Boolean
    ' Function possibly will be used for logging
    ' create path to Logs
    Dim sPath As String
    'Add a trailing slash if none
    sPath = NewPath & IIf(Right$(NewPath, 1) = "\", "", "\")

    'Call API
    If MakeSureDirectoryPathExists(sPath) <> 0 Then
        'No errors, return True
        CreatePath = True
    End If

End Function

' ***********************************************************************************
' *                             COLLECTION OF VARIABLES                             *
' *                                                                                 *
' ***********************************************************************************

' Author: Ivan Bondarenko
' Release date: 2017-01
' https://bondarenkoivan.wordpress.com
' https://linkedin.com/in/bondarenkoivan/en

' + inspiration from https://blogs.sap.com/2017/02/03/analysis-for-office-variables-and-filters-via-vba/

Private Sub Collect_Variables()
' sub will add data sources and variables into existing tables
    
    Dim arrVar
    Dim arrFilters
    Dim arrDimensions
    Dim r
    Dim r_ds
    Dim i As Long
    Dim Var, Fil, Dimen
    Dim sPassword As String
    Dim sSystem   As String
    Dim order As Long
    Dim WhatToCollectResponse As String
    Dim bFirstCollection As Boolean
    
AskWhatToCollect:
    WhatToCollectResponse = InputBox("Please, specify what you want to collect:" & vbCrLf _
        & "1 - Variables" & vbCrLf _
        & "2 - Field (dimension) & Measure Filters" & vbCrLf _
        & "3 - both", "What to collect?", "1")
    
    If WhatToCollectResponse = vbNullString Then
        Exit Sub
    ElseIf WhatToCollectResponse <> "1" And WhatToCollectResponse <> "2" And WhatToCollectResponse <> "3" Then
        GoTo AskWhatToCollect
    End If
    
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    'On Error Resume Next
    'ThisWorkbook.Sheets("Result").ListObjects(sTableVariablesName).DataBodyRange.Rows.Delete
    'ThisWorkbook.Sheets("Result").ListObjects(sTableDataSourcesName).DataBodyRange.Rows.Delete
    'On Error GoTo ErrHandler
    
    ' If Application.Workbooks.Count = 1 Then Exit Sub
    ' Call LogOffBOA
    'Call DefineNamedRangeScope
    
    Application.StatusBar = "Enabling BOA..."
    Call EnableBOA
    Application.StatusBar = vbNullString
    ThisWorkbook.Activate
    
    Call ListObjectDeleteEmptyRows(Me.ListObjects(sTableDataSourcesName))
    Call ListObjectDeleteEmptyRows(Me.ListObjects(sTableVariablesName))
    
    If Me.ListObjects(sTableDataSourcesName).DataBodyRange Is Nothing Then
        bFirstCollection = True
    End If
    
    Call GetListOfDS ' get list of Data Sources
    
    order = 1
    If IsArray(arrDS) Then
        ' for each DS in workbook - collect variables
        'For i = 1 To UBound(arrDS, 1)
        For i = LBound(arrDS, 1) To UBound(arrDS, 1)
            If arrDS(i, enumDataSource.DS_ID) = vbNullString Then GoTo Next_DS
            
            If Not Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, enumDataSource.DS_ID)) Then
                
                Application.StatusBar = "Logging On..."
                Call Application.Run("SapLogon", _
                    CStr(arrDS(i, enumDataSource.DS_ID)), _
                    Me.Names("CP_GENERAL_CLIENT").RefersToRange.Value, _
                    Me.Names("CP_GENERAL_USER").RefersToRange.Value, _
                    GetPassword(Me.Names("CP_GENERAL_SYSTEM").RefersToRange.Value, _
                        Me.Names("CP_GENERAL_USER").RefersToRange.Value), _
                    Me.Names("CP_GENERAL_LANGUAGE").RefersToRange.Value)
                
                Application.StatusBar = "Refreshing..."
                Run "SAPExecuteCommand", "Refresh"
                
            End If
            
            If Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, enumDataSource.DS_ID)) Then
                Application.StatusBar = "Collecting variables and filters..."
                Application.EnableEvents = False
                Set r_ds = Me.ListObjects(sTableDataSourcesName).ListRows.Add(AlwaysInsert:=False)
                
                If bFirstCollection Then
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Scope").Index) = "Main"
                End If
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Sheet").Index) = arrDS(i, enumDataSource.DS_Sheet) ' Sheet Name
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Data Source").Index) = arrDS(i, enumDataSource.DS_ID)
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Crosstab").Index) = arrDS(i, enumDataSource.DS_Crosstab)
                
                ' hyperlink to corresponding Crosstab
                Me.Hyperlinks.Add Anchor:=r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Crosstab").Index), _
                    Address:="", SubAddress:=arrDS(i, enumDataSource.DS_Crosstab), _
                    TextToDisplay:=arrDS(i, enumDataSource.DS_Crosstab)
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Data Source Name").Index) = _
                    Application.Run("SapGetSourceInfo", arrDS(i, enumDataSource.DS_ID), "DataSourceName")
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Query").Index) = _
                    Application.Run("SapGetSourceInfo", arrDS(i, enumDataSource.DS_ID), "QueryTechName")
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("System").Index) = _
                    Application.Run("SapGetSourceInfo", arrDS(i, enumDataSource.DS_ID), "System")
                                                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Refresh?").Index) = "Y"
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Order").Index) = order
                order = order + 1
                Application.EnableEvents = True
                
                If WhatToCollectResponse = "1" Or WhatToCollectResponse = "3" Then
                    ' Get Variables
                    arrVar = Application.Run("SAPListOfVariables", arrDS(i, enumDataSource.DS_ID), "INPUT_STRING", "PROMPTS")
                    
                    If IsArray(arrVar) Then
                        On Error Resume Next
                        Debug.Print arrVar(1, 1)
                        If Err.Number <> 0 Then
                            Err.Clear
                            On Error GoTo 0
                            ' 1-dimensional - only 1 variable
                            Application.EnableEvents = False
                            Set r = Me.ListObjects(sTableVariablesName).ListRows.Add(AlwaysInsert:=False)
                            
                            If bFirstCollection Then
                                r.Range.Cells(1, r.Parent.ListColumns("Scope").Index) = "Main"
                            End If
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source").Index) = arrDS(i, enumDataSource.DS_ID)
                            r.Range.Cells(1, r.Parent.ListColumns("Command").Index) = "SAPSetVariable"
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source Sheet").Index) = arrDS(i, enumDataSource.DS_Sheet)
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Field").Index) = _
                                    Application.Run("SAPGetVariable", arrDS(i, enumDataSource.DS_ID), arrVar(1), "TECHNICALNAME")
                                
                            r.Range.Cells(1, r.Parent.ListColumns("Field Name").Index) = arrVar(1)
                            r.Range.Cells(1, r.Parent.ListColumns("Value").Index) = "'" & arrVar(2)
                            Application.EnableEvents = True
                        Else
                            ' two-dimentional
                            For Var = LBound(arrVar) To UBound(arrVar)
                                Application.EnableEvents = False
                                Set r = Me.ListObjects(sTableVariablesName).ListRows.Add(AlwaysInsert:=False)
                                
                                If bFirstCollection Then
                                    r.Range.Cells(1, r.Parent.ListColumns("Scope").Index) = "Main"
                                End If
                            
                                r.Range.Cells(1, r.Parent.ListColumns("Data Source").Index) = arrDS(i, enumDataSource.DS_ID)
                                r.Range.Cells(1, r.Parent.ListColumns("Command").Index) = "SAPSetVariable"
                                r.Range.Cells(1, r.Parent.ListColumns("Data Source Sheet").Index) = arrDS(i, enumDataSource.DS_Sheet)
                                
                                r.Range.Cells(1, r.Parent.ListColumns("Field").Index) = _
                                        Application.Run("SAPGetVariable", arrDS(i, enumDataSource.DS_ID), arrVar(Var, 1), "TECHNICALNAME")
                                    
                                r.Range.Cells(1, r.Parent.ListColumns("Field Name").Index) = arrVar(Var, 1)
                                r.Range.Cells(1, r.Parent.ListColumns("Value").Index) = "'" & arrVar(Var, 2)
                                Application.EnableEvents = True
                            Next Var
                        End If
                        
                    Else ' If IsArray(arrVar) Then
                        ' if no variables in Data Source
                        Application.EnableEvents = False
                        Set r = Me.ListObjects(sTableVariablesName).ListRows.Add(AlwaysInsert:=False)
                        
                        If bFirstCollection Then
                            r.Range.Cells(1, r.Parent.ListColumns("Scope").Index) = "Main"
                        End If
                        
                        r.Range.Cells(1, r.Parent.ListColumns("Data Source").Index) = arrDS(i, enumDataSource.DS_ID)
                        r.Range.Cells(1, r.Parent.ListColumns("Command").Index) = "SAPSetVariable"
                        r.Range.Cells(1, r.Parent.ListColumns("Data Source Sheet").Index) = arrDS(i, enumDataSource.DS_Sheet)
                        
                        r.Range.Cells(1, r.Parent.ListColumns("Field").Index) = ""
                        r.Range.Cells(1, r.Parent.ListColumns("Field Name").Index) = "Not applicable"
                        r.Range.Cells(1, r.Parent.ListColumns("Value").Index) = ""
                        Application.EnableEvents = True
                    End If
                End If ' WhatToCollectResponse = 1 or 3
                
                If WhatToCollectResponse = "2" Or WhatToCollectResponse = "3" Then
                    ' Get Filters
                    ' arrFilters = Application.Run("SAPListOfEffectiveFilters", arrDS(i, enumDataSource.DS_ID), "INPUT_STRING")
                    ' Effective filters also includes variables
                    
                    ' dynamic filters includes dimensions and measures
                    arrFilters = Application.Run("SAPListOfDynamicFilters", arrDS(i, enumDataSource.DS_ID), "INPUT_STRING")
                    arrDimensions = Application.Run("SAPListOfDimensions", arrDS(i, enumDataSource.DS_ID))
                    
                    If IsArray(arrFilters) Then
                        On Error Resume Next
                        Debug.Print arrFilters(1, 1)
                        If Err.Number <> 0 Then
                            Err.Clear
                            On Error GoTo 0
                            ' 1-dimensional - only 1 variable
                            Application.EnableEvents = False
                            Set r = Me.ListObjects(sTableVariablesName).ListRows.Add(AlwaysInsert:=False)
                            
                            If bFirstCollection Then
                                r.Range.Cells(1, r.Parent.ListColumns("Scope").Index) = "Main"
                            End If
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source").Index) = arrDS(i, enumDataSource.DS_ID)
                            r.Range.Cells(1, r.Parent.ListColumns("Command").Index) = "SAPSetFilter"
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source Sheet").Index) = arrDS(i, enumDataSource.DS_Sheet)
                            
                            For Dimen = LBound(arrDimensions) To UBound(arrDimensions)
                                 If arrDimensions(Dimen, 2) = arrFilters(1) Then
                                    r.Range.Cells(1, r.Parent.ListColumns("Field").Index) = arrDimensions(Dimen, 1) ' technical name
                                    Exit For
                                 End If
                            Next Dimen
                                
                            r.Range.Cells(1, r.Parent.ListColumns("Field Name").Index) = arrFilters(1)
                            r.Range.Cells(1, r.Parent.ListColumns("Value").Index) = "'" & arrFilters(2)
                            Application.EnableEvents = True
                        Else
                            ' 2-deminsional
                            For Fil = LBound(arrFilters) To UBound(arrFilters)
                                If arrFilters(Fil, 1) <> "Measures" Then
                                    Application.EnableEvents = False
                                    Set r = Me.ListObjects(sTableVariablesName).ListRows.Add(AlwaysInsert:=False)
                                    
                                    If bFirstCollection Then
                                        r.Range.Cells(1, r.Parent.ListColumns("Scope").Index) = "Main"
                                    End If
                                    
                                    r.Range.Cells(1, r.Parent.ListColumns("Data Source").Index) = arrDS(i, enumDataSource.DS_ID)
                                    r.Range.Cells(1, r.Parent.ListColumns("Command").Index) = "SAPSetFilter"
                                    r.Range.Cells(1, r.Parent.ListColumns("Data Source Sheet").Index) = arrDS(i, enumDataSource.DS_Sheet)
                                    
                                    For Dimen = LBound(arrDimensions) To UBound(arrDimensions)
                                         If arrDimensions(Dimen, 2) = arrFilters(Fil, 1) Then
                                            r.Range.Cells(1, r.Parent.ListColumns("Field").Index) = arrDimensions(Dimen, 1) ' technical name
                                            Exit For
                                         End If
                                    Next Dimen
                                        
                                    r.Range.Cells(1, r.Parent.ListColumns("Field Name").Index) = arrFilters(Fil, 1)
                                    r.Range.Cells(1, r.Parent.ListColumns("Value").Index) = "'" & arrFilters(Fil, 2)
                                    Application.EnableEvents = False
                                End If ' if measures
                            Next Fil
                                            
                        End If
                    Else
                    
                    End If ' If IsArray(arrFilters)
                End If
            Else ' If Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, enumDataSource.DS_ID)) Then
                ' write that DS is not active
                                
            End If ' isDataSourceActive
Next_DS:
        Next i ' next DS in arrDS
    Else
        
    
    End If ' If IsArray(arrDS) Then

    Application.EnableEvents = False
    Me.Range(sTableDataSourcesName & "[#All]").RemoveDuplicates Header:=xlYes, _
    Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19)
        
    Me.Range(sTableVariablesName & "[#All]").RemoveDuplicates Header:=xlYes, _
     Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
    
    Call ListObjectDeleteEmptyRows(Me.ListObjects(sTableDataSourcesName))
    Call ListObjectDeleteEmptyRows(Me.ListObjects(sTableVariablesName))
        
Exit_Sub:
    On Error Resume Next
    Erase arrDS
    Application.Cursor = xlDefault
    
    Application.StatusBar = vbNullString
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "Collect_Variables", Err.Number & ": " & Err.Description
    Err.Clear
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume ' for debug
End Sub

' ***********************************************************************************
' *                             COLLECTION OF VARIABLES END                         *
' *                                                                                 *
' ***********************************************************************************

Private Sub GetListOfDS()
    ' Sub for "collect variables" part
    Dim tmpCrossTabs
    Dim i As Long
    
    ' works for active workbook
    tmpCrossTabs = Application.Run("SAPListOf", "CROSSTABS")
    
    On Error Resume Next
    If Not IsArray(tmpCrossTabs) Then
        MsgBox "Cannot find any BOA Data Source in this workbook." & vbCrLf _
            & "Such issue was noticed when BO Analysis addin cannot initialize itself." & vbCrLf _
            & "If you are sure that have Data Sources in this workbook - " & vbCrLf _
            & "save this file, close Excel completely, check that no Excel process in Task Manager," & vbCrLf _
            & "open your file and try again."
        Exit Sub
    End If
    Debug.Print tmpCrossTabs(1, 1) ' check if it is 2-dim array (when only 1 DS - response is 1-dim array)
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' only 1 dimension
        ReDim arrDS(0, 12)
        arrDS(0, enumDataSource.DS_ID) = tmpCrossTabs(3) ' data source ID

        On Error Resume Next
        arrDS(0, enumDataSource.DS_Sheet) = ActiveWorkbook.Names("SAP" & tmpCrossTabs(1)).RefersToRange.Parent.Name ' worksheet name
        ' not 100% precise method, can fail if user renamed NamedRange for crosstab
        arrDS(0, enumDataSource.DS_Crosstab) = "SAP" & tmpCrossTabs(1)
        Err.Clear
        On Error GoTo 0
        
    Else
        Err.Clear
        On Error GoTo 0
        ' arrDS = tmpCrossTabs
        ReDim arrDS(UBound(tmpCrossTabs, 1), 12)
        
        For i = 1 To UBound(tmpCrossTabs, 1)
            arrDS(i, enumDataSource.DS_ID) = tmpCrossTabs(i, 3) ' data source ID
            
            On Error Resume Next
            arrDS(i, enumDataSource.DS_Sheet) = ThisWorkbook.Names("SAP" & tmpCrossTabs(i, 1)).RefersToRange.Parent.Name ' worksheet name
            ' not 100% precise method, can fail if user renamed NamedRange for crosstab
            arrDS(i, enumDataSource.DS_Crosstab) = "SAP" & tmpCrossTabs(i, 1)
            Err.Clear
            On Error GoTo 0
        Next i
        
    End If

End Sub

Private Sub ListObjectDeleteEmptyRows(lo As ListObject)
    Dim i As Long
    Dim col As ListColumn
    Dim bNonEmptyFound As Boolean
    If lo.DataBodyRange Is Nothing Then Exit Sub
    i = 1
    ' TODO: replace by loop that ignores columns with hasformula
    'Do While i <= lo.ListRows.Count
    '    If WorksheetFunction.CountA(lo.ListRows(i).Range) = 0 Then
    '        lo.ListRows(i).Delete
    '        i = i - 1
    '    End If
    '    i = i + 1
    'Loop
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Do While i <= lo.ListRows.Count
        For Each col In lo.ListColumns
            If col.DataBodyRange.HasFormula Then
                ' skip
            Else
                If lo.DataBodyRange.Cells(i, col.Index).Value <> vbNullString Then
                    bNonEmptyFound = True
                    Exit For
                End If
            End If
        Next col
        
        If Not bNonEmptyFound Then
            lo.ListRows(i).Delete
            i = i - 1
        End If
        
        i = i + 1
        bNonEmptyFound = False
    Loop
    
End Sub

Private Sub Callback_BeforeFirstPromptsDisplay(Optional dummy As Variant)
    ' callback is triggered on Data Source Refresh
    ' at than point we know Scope and 'currentDS' stores DS ID
    Dim arr_var_scope
    Dim arr_var_ds
    Dim arr_var_command
    Dim arr_var_field
    Dim arr_var_value
    Dim arr_var_ref_type
    Dim i As Long
    Dim variable_row As Long
    Dim arrVariableValue
    Dim sVariableValue As String
    Dim scope As String
        
    Debug.Print Now, "Callback_BeforeFirstPromptsDisplay"
    scope = CStr(Me.Names("CP_SCOPE").RefersToRange.Value)
    
    ' define arrays with Variables table columns - faster then read each cell
    With Me.ListObjects(sTableVariablesName)
        arr_var_scope = Get_Array_From_Range(.ListColumns("Scope").DataBodyRange.Address(external:=True))
        arr_var_ds = Get_Array_From_Range(.ListColumns("Data Source").DataBodyRange.Address(external:=True))
        arr_var_command = Get_Array_From_Range(.ListColumns("Command").DataBodyRange.Address(external:=True))
        arr_var_field = Get_Array_From_Range(.ListColumns("Field").DataBodyRange.Address(external:=True))
        arr_var_value = Get_Array_From_Range(.ListColumns("Value").DataBodyRange.Address(external:=True))
        arr_var_ref_type = Get_Array_From_Range(.ListColumns("Reference Type").DataBodyRange.Address(external:=True))
    End With
              
    ' Set Variables
    For variable_row = 1 To UBound(arr_var_ds)
        If arr_var_scope(variable_row, 1) = scope And arr_var_ds(variable_row, 1) = currentDS Then
            If arr_var_command(variable_row, 1) = "SAPSetVariable" Then
                If arr_var_ref_type(variable_row, 1) = "Range" Then
                    
                    ' easier then INPUT_STRING_AS_ARRAY?
                    arrVariableValue = Get_Array_From_Range(Application.Range(CStr(arr_var_value(variable_row, 1))).Address(external:=True))
                    For i = 1 To UBound(arrVariableValue)
                        sVariableValue = sVariableValue & "; " & arrVariableValue(i, 1)
                    Next i
                    sVariableValue = Mid(sVariableValue, 3)
                    
                    '    sVariableValue = "1; 2; 3; .... 90000; ... "
                    '    sVariableValue = "!1; !2; !3; .... !90000; ... "
                Else
                    sVariableValue = CStr(arr_var_value(variable_row, 1))
                End If
                
                Me.ListObjects(sTableVariablesName).ListColumns("Last Refresh").DataBodyRange.Cells(variable_row, 1).Value = Now()
                Me.ListObjects(sTableVariablesName).ListColumns("Last Used Value").DataBodyRange.Cells(variable_row, 1).Value = "'" & sVariableValue
                
                If Run(arr_var_command(variable_row, 1), _
                        arr_var_field(variable_row, 1), _
                        sVariableValue, _
                        "INPUT_STRING", _
                        arr_var_ds(variable_row, 1)) = 0 Then
                    
                    Debug.Print Now, "Couldn't apply filter for variable " & str(arr_var_field(variable_row, 1)) & " : " & sVariableValue
                    GoTo Exit_Sub
                End If
                sVariableValue = vbNullString
            End If ' var command
        End If ' var scope and data source
    Next variable_row
    
    bBeforeFirstPromptsDisplaySubExecuted = True
Exit_Sub:
    Exit Sub

ErrHandler:
    Debug.Print Now, "Callback_BeforeFirstPromptsDisplay", Err.Number & ": " & Err.Description
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Sub

Private Sub Callback_BeforeFirstPromptsDisplay_Reg()
    Dim iRet As Integer
    iRet = Application.Run("SAPExecuteCommand", "RegisterCallback", "BeforeFirstPromptsDisplay", "ControlPanel.Callback_BeforeFirstPromptsDisplay")
    Debug.Print Now, "Callback_BeforeFirstPromptsDisplay_Reg", iRet
    bBeforeFirstPromptsDisplayRegistered = (iRet = 1)
End Sub

Private Sub Callback_BeforeFirstPromptsDisplay_UnReg()
    Dim iRet As Integer
    iRet = Application.Run("SAPExecuteCommand", "UnRegisterCallback", "BeforeFirstPromptsDisplay", "ControlPanel.Callback_BeforeFirstPromptsDisplay")
    Debug.Print Now, "Callback_BeforeFirstPromptsDisplay_UnReg", iRet
End Sub

Private Function IsSAPSetVariablesOnly(ds As String) As Boolean
    IsSAPSetVariablesOnly = True ' by default
    With Me.ListObjects(sTableVariablesName)
        If Not .DataBodyRange Is Nothing Then
            IsSAPSetVariablesOnly = (WorksheetFunction.CountIfs(.ListColumns("Scope").DataBodyRange, Me.Names("CP_SCOPE").RefersToRange.Value, _
                    .ListColumns("Data Source").DataBodyRange, ds, _
                    .ListColumns("Command").DataBodyRange, "SAPSetFilter") = 0)
        End If
    End With
End Function

Private Function IsRepeatRefreshOn(ds As String) As Boolean
    With Me.ListObjects(sTableDataSourcesName)
        If Not .DataBodyRange Is Nothing Then
            IsRepeatRefreshOn = (WorksheetFunction.CountIfs(.ListColumns("Scope").DataBodyRange, Me.Names("CP_SCOPE").RefersToRange.Value, _
                    .ListColumns("Data Source").DataBodyRange, ds, _
                    .ListColumns("Repeat Refresh"), "Y") = 1)
        End If
    End With
End Function

' http://www.fmsinc.com/microsoftaccess/modules/examples/avoiddoevents.asp
Private Sub WaitSeconds(intSeconds As Integer)
    Dim datTime As Date
    Dim sStatusBarInitial As String
    Dim k As Integer
    Dim bScreenUpdatingInitial As Boolean
    Dim CursorInitial As Double
    
    With Application
        bScreenUpdatingInitial = .ScreenUpdating
        CursorInitial = .Cursor
        sStatusBarInitial = IIf(.StatusBar <> False, .StatusBar, vbNullString)
        
        .ScreenUpdating = False
        .Cursor = xlWait
    End With
    
    datTime = DateAdd("s", intSeconds, Now)
    Do
        ' 255 chars is limit of status bar
        If Len(Application.StatusBar) + Len(CStr(CStr(intSeconds - k)) & "...") > 255 Then
            Application.StatusBar = Left(Left(sStatusBarInitial, 255 - Len(CStr(intSeconds - k) & "...") - 1) _
                & " " & CStr(intSeconds - k) & "...", 255)
        Else
            Application.StatusBar = sStatusBarInitial & " " & CStr(intSeconds - k) & "..."
        End If
        ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
        Sleep 1000
        If Me.Names("CP_USE_DOEVENTS").RefersToRange.Value = "Y" Then
            DoEvents
        End If
        k = k + 1
    Loop Until Now >= datTime

    With Application
        .StatusBar = sStatusBarInitial
        .ScreenUpdating = bScreenUpdatingInitial
        .Cursor = CursorInitial
    End With
End Sub

' **************************** ACTIONS HANDLER ****************************
' *************************************************************************
Private Sub Actions_Trigger()
    ' test
    'Actions_Handler_Manual
    Call Actions_Handler("3")
End Sub

Function Actions_Handler_Manual(Optional ActionID As String) As Boolean
    bManualRefresh = True
    Call Actions_Handler
End Function

Function Actions_Handler(Optional ActionID As String) As Boolean
    ' go row by row of actions table, read parameters and call corresponding procedures
    Dim row_id As Long
    Dim sPath As String
    Dim rSearchResult As Range
    Dim i As Long
    
    On Error GoTo ErrHandler
    Debug.Print Now, "Actions Handler " & ActionID
    
    With Me.ListObjects(STableActions)
    
        If .DataBodyRange Is Nothing Then Exit Function
        
        Me.UsedRange.Calculate ' re-calc all formulas on ControlPanel
        
        If ActionID <> vbNullString Then
            ' range.find doesn't work when table is hidden
            'Set rSearchResult = .ListColumns("Action ID").DataBodyRange.Find(ActionID, , xlValues, xlWhole)
            For i = 1 To .DataBodyRange.Rows.Count
                If CStr(.ListColumns("Action ID").DataBodyRange.Cells(i, 1).Value) = ActionID Then
                    Set rSearchResult = .DataBodyRange.Rows(i)
                    Exit For
                End If
            Next i
            
            If Not rSearchResult Is Nothing Then
                currentActionRowId = rSearchResult.Row - .HeaderRowRange.Row
                Call Action_Handler(currentActionRowId)
                Exit Function
            End If
        End If
        
        For row_id = 1 To .DataBodyRange.Rows.Count
            currentActionRowId = row_id
            ' skip disabled actions
            If .ListColumns("Enabled").DataBodyRange.Cells(row_id, 1).Value = "Y" Then
                Call Action_Handler(currentActionRowId)
            End If
        Next row_id
    End With ' Me.ListObjects(STableActions)
    
    Actions_Handler = True
    
Exit_Sub:

    Exit Function
    
ErrHandler:
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function Action_Handler(row_id As Long) As Boolean
    On Error GoTo ErrHandler
    Debug.Print Now, "Action Handler " & row_id
    
    With Me.ListObjects(STableActions)
        If .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Save As & Email" Then
            ' prepare file for email
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Save As & Email"
            Call Process_Email(row_id)
            
        ElseIf .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Refresh All" Then
            ' special function as update is complex operation
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Refresh All"
            Call RefreshWorkbook
        
        ElseIf .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Save As" Then
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Save As"
            Call FileSave_Handler(row_id)
            
        ElseIf .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Email" Then
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Email"
            If .ListColumns("Email Method").DataBodyRange.Cells(row_id, 1).Value = "SMTP" Then
                Call Send_EMail_CDO(.ListColumns("Email From").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email To").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Subject").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Body").DataBodyRange.Cells(row_id, 1).Value, _
                                    vbNullString, _
                                    .ListColumns("Email Importancy").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email CC").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email BCC").DataBodyRange.Cells(row_id, 1).Value)
                                    
            ElseIf .ListColumns("Email Method").DataBodyRange.Cells(row_id, 1).Value = "Outlook" Then
                Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Email"
                
                Call Send_Email_Outlook(.ListColumns("Email To").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Subject").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Body").DataBodyRange.Cells(row_id, 1).Value, _
                                    vbNullString, _
                                    .ListColumns("Email Importancy").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email CC").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email BCC").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email From").DataBodyRange.Cells(row_id, 1).Value)
            End If
        
        ElseIf .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Run Macro" Then
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Run Macro"
            On Error Resume Next
            Application.Run .ListColumns("Macro Name").DataBodyRange.Cells(row_id, 1).Value
            Debug.Print Now, "Action_Handler - Run Macro", Err.Number, Err.Description
            Err.Clear
            On Error GoTo ErrHandler
        
        ElseIf .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Clear Crosstabs" Then
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Clear Crosstabs"
            Call ClearCrosstabs
            
        ElseIf .ListColumns("Action").DataBodyRange.Cells(row_id, 1).Value = "Save" Then
            Application.StatusBar = CStr(Me.Names("CP_SCOPE").RefersToRange.Value) & ": Save"
            ' always keep backup of your file !!!
            ThisWorkbook.Save
        
        ElseIf 1 = 2 Then
        
        End If
    End With
    
Exit_Sub:
    Application.StatusBar = vbNullString
    Exit Function
    
ErrHandler:
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    GoTo Exit_Sub
    Resume ' for debug
End Function

Sub Process_Email(row_id As Long)
    Dim sPath As String
    
    On Error GoTo ErrHandler
    Debug.Print Now, "Process_Email"
    
    ' firstly we have to save file, then attach it to email
    sPath = FileSave_Handler(row_id) ' saves file and return path
    
    If sPath <> vbNullString Then
        With Me.ListObjects(STableActions)
            If .ListColumns("Email Method").DataBodyRange.Cells(row_id, 1).Value = "SMTP" Then
                Call Send_EMail_CDO(.ListColumns("Email From").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email To").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Subject").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Body").DataBodyRange.Cells(row_id, 1).Value, _
                                    sPath, _
                                    .ListColumns("Email Importancy").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email CC").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email BCC").DataBodyRange.Cells(row_id, 1).Value)
                                     
            ElseIf .ListColumns("Email Method").DataBodyRange.Cells(row_id, 1).Value = "Outlook" Then
                Call Send_Email_Outlook(.ListColumns("Email To").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Subject").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email Body").DataBodyRange.Cells(row_id, 1).Value, _
                                    sPath, _
                                    .ListColumns("Email Importancy").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email CC").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email BCC").DataBodyRange.Cells(row_id, 1).Value, _
                                    .ListColumns("Email From").DataBodyRange.Cells(row_id, 1).Value)
            End If
        End With
    Else
    ' write log - failed to save file
    End If

Exit_Sub:

    Exit Sub
    
ErrHandler:
    ' write log - failed to send email
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    GoTo Exit_Sub
    Resume ' for debug
End Sub

Function FileSave_Handler(row_id As Long) As String
     ' returns file path when successful of -1 if not
    Dim sExtension As String
    Dim sFolder As String
    Dim sFileName As String
    Dim sPath As String
    
    ' Possible cases
    ' 1. File Format is not provided - use current workbook's format
    '   2 if no sheets - send entire workbook
    '    3. if no Folder - use %temp%
    '       4 if no file name provided - use name of current workbook
    '           5 if 'add current datetime' - add
    '       else - use provided file name
    '           if  'add current datetime' - add
    '   else - list of sheets - create workbook with provided sheets
    ' else - depends on provided File Format
    
    On Error GoTo ErrHandler
    Debug.Print Now, "FileSave_Handler"
    FileSave_Handler = FileSave_XOR_Folder(row_id)
    Exit Function
    
Exit_Sub:

    Exit Function
    
ErrHandler:
    ' write log
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    FileSave_Handler = vbNullString
        
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function FileSave_XOR_Folder(row_id As Long) As String
    Dim sFolder As String
    On Error GoTo ErrHandler

    With Me.ListObjects(STableActions)
        If .ListColumns("Folder / Workspace").DataBodyRange.Cells(row_id, 1).Value = vbNullString Then
            ' no Folder - use %temp% directory
            sFolder = Environ("temp") & "\"
        Else
            sFolder = .ListColumns("Folder / Workspace").DataBodyRange.Cells(row_id, 1).Value
            If Left(sFolder, 1) = "%" And Right(sFolder, 1) = "%" Then
                If sFolder = "%temp%" Then
                    sFolder = Environ("temp") & "\"
                Else
                    sFolder = CreateObject("WScript.Shell").SpecialFolders(Replace(sFolder, "%", vbNullString)) & "\"
                End If
            End If
            If Left(sFolder, 4) = "http" Then
                If Right(sFolder, 1) <> "/" Then: sFolder = sFolder & "/"
            Else
                ' not http folder
                If Right(sFolder, 1) <> "\" Then: sFolder = sFolder & "\"
            End If
        End If ' check Folder / Workspace
    End With
    
    'If Left(sFolder, 4) = "http" Then
    '    sFolder = WorksheetFunction.EncodeURL(sFolder)
    'End If
    
    FileSave_XOR_Folder = FileSave_XOR_FileName(row_id, sFolder)
    
Exit_Sub:

    Exit Function
    
ErrHandler:
    ' write log
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    FileSave_XOR_Folder = vbNullString
        
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function FileSave_XOR_FileName(row_id As Long, sFolder As String) As String
    Dim sFileName As String
    
    On Error GoTo ErrHandler

    With Me.ListObjects(STableActions)
        If .ListColumns("File Name").DataBodyRange.Cells(row_id, 1).Value = vbNullString Then
            ' if no file name provided - use name of current workbook
            sFileName = Left(ThisWorkbook.Name, (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
        Else
            ' use provided name
            sFileName = ReplaceIllegalChar(.ListColumns("File Name").DataBodyRange.Cells(row_id, 1).Value)
        End If ' File Name
    End With
    
    FileSave_XOR_FileName = FileSave_XOR_DateTime(row_id, sFolder, sFileName)
Exit_Sub:

    Exit Function
    
ErrHandler:
    ' write log
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    FileSave_XOR_FileName = vbNullString
        
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function FileSave_XOR_DateTime(row_id As Long, sFolder As String, sFileName As String) As String
    Dim sFileNameNew As String
    
    On Error GoTo ErrHandler
    
    With Me.ListObjects(STableActions)
        If .ListColumns("Add Current DateTime").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
            ' if 'add current datetime' - add
            sFileNameNew = sFileName & " " & Format(Now(), .ListColumns("Add Current DateTime").DataBodyRange.Cells(row_id, 1).Value)
        Else
            sFileNameNew = sFileName
        End If ' Add Current DateTime
    End With
    
    FileSave_XOR_DateTime = FileSave_XOR_Extension(row_id, sFolder, sFileNameNew)
        
Exit_Sub:
    
    Exit Function
    
ErrHandler:
    ' write log
    Debug.Print Now, Err.Number, Err.Description
    Err.Clear
    If bManualRefresh Then Stop
    
    FileSave_XOR_DateTime = vbNullString
        
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function FileSave_XOR_Extension(row_id As Long, sFolder As String, sFileName As String) As String
    Dim sPath As String
    Dim sTmp_path As String
    Dim sTmp_path_xlsx As String
    Dim tmp_wb As Workbook
    Dim sExtension As String
    Dim sh As Worksheet
    Dim i As Integer
    Dim arr
    Dim sResulting_FileFormat As String
    Dim sCSVfiles As String
    Dim sh_id As Integer
    Dim bFound As Boolean
    Dim sFileNameNew As String
    Dim bReadOnlyRecommended As Boolean
    
    On Error GoTo ErrHandler
    Debug.Print Now, "FileSave_XOR_Extension"
    
    ' define parameters
    With Me.ListObjects(STableActions)
        If .ListColumns("File Format").DataBodyRange.Cells(row_id, 1).Value = vbNullString Then
            ' take extension of current workbook
            sExtension = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
        Else
            sExtension = "." & .ListColumns("File Format").DataBodyRange.Cells(row_id, 1).Value
        End If
        
        bReadOnlyRecommended = (.ListColumns("ReadOnly Recommended").DataBodyRange.Cells(row_id, 1).Value = "Y")
    End With
    
    'If Left(sFolder, 4) = "http" Then
    '    sFileName = WorksheetFunction.EncodeURL(sFileName)
    'End If
    
    ' target path
    sPath = sFolder & sFileName & LCase(sExtension)
    
    ' we cannot modify current workbook for 'SaveAs' and 'Send' actions
    ' must save as new one - ALWAYS
    ' this is the idea of additional actions
    ' only action 'Save' will save current workbook inplace
    ' for 'Saved As' workbook can be requested additional changes
    ' such as 'Delete Sheets', 'Save Sheets', 'Sheets for Formulas to Values'
    ' in addition, depending on 'File Format' we must use proper savinig method, especially PDF and CSV
    
    ' simplest way
    ' if no modification requested, and same extension as current wb
    'SaveCopyAs doesn't support save to HTTP
    If Left(sFolder, 4) <> "http" Then
        If sExtension = Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1)) Then
            With Me.ListObjects(STableActions)
                If .ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value = vbNullString And _
                   .ListColumns("Delete Sheets").DataBodyRange.Cells(row_id, 1).Value = vbNullString And _
                   .ListColumns("Delete WB Queries").DataBodyRange.Cells(row_id, 1).Value = vbNullString And _
                   .ListColumns("Sheets for Formulas to Values").DataBodyRange.Cells(row_id, 1).Value = vbNullString And _
                   .ListColumns("Save Without Macro").DataBodyRange.Cells(row_id, 1).Value = vbNullString And _
                   (Not bReadOnlyRecommended) Then
                    ' if no modifications then just save
                    'SaveCopyAs doesn't support save to HTTP
                    ThisWorkbook.SaveCopyAs sPath
                    FileSave_XOR_Extension = sPath
                    Exit Function
                End If
            End With
        End If
    End If
    
    ' if method is "Copy/Paste" - applied only to xl* formats
    With Me.ListObjects(STableActions)
        If .ListColumns("Saving Method").DataBodyRange.Cells(row_id, 1).Value = "Copy/Paste" Then
            If .ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
                ' copy/paste specified sheets to a new workbook
                arr = Split(.ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value, ",")
                
                Application.DisplayAlerts = False
                Application.ScreenUpdating = False
                ' transfer worksheets by Copy / Paste
                Set tmp_wb = Application.Workbooks.Add
                For Each sh In ThisWorkbook.Sheets
                    For i = LBound(arr) To UBound(arr)
                        If Trim(arr(i)) = sh.Name Then
                            sh.Activate
                            Call CopySheetValues(tmp_wb, sh.Name)
                            Exit For  ' loop through array
                        End If
                    Next i
                Next sh
                
                ' delete default worksheets from resulting workbook
                sh_id = 1
                Do While sh_id <= tmp_wb.Sheets.Count
                    For i = LBound(arr) To UBound(arr)
                        If Trim(arr(i)) = tmp_wb.Sheets(sh_id).Name Then
                            bFound = True
                            Exit For
                        End If
                    Next i
                    
                    If Not bFound Then
                        tmp_wb.Sheets(sh_id).Visible = xlSheetVisible
                        tmp_wb.Sheets(sh_id).Delete
                        sh_id = sh_id - 1
                    End If
                    sh_id = sh_id + 1
                    bFound = False
                Loop
                
                ' jump to last action
                GoTo Saving_Workbook
            Else
            ' nothing is specified - save all sheets ?
            ' not supported so far
            End If
            
        End If ' If .ListColumns("Saving Method").DataBodyRange.Cells(row_id, 1).Value = "Copy/Paste" Then
    End With
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    ' kind of GUID :), enough for such DIY solution
    sTmp_path = Environ("temp") & "\" & sFileName & " " & Format(Now(), "yyyy-MM-dd hh-mm-ss") & _
        Right(ThisWorkbook.Name, Len(ThisWorkbook.Name) - (InStrRev(ThisWorkbook.Name, ".", -1, vbTextCompare) - 1))
    ThisWorkbook.SaveCopyAs sTmp_path
    Application.EnableEvents = False ' prevent macro in Workbook_Open
    Application.Calculation = xlCalculationManual
    ' open workbook
    Set tmp_wb = Workbooks.Open(sTmp_path, UpdateLinks:=False)
    
    ' if save without macro - just save as xlsx
    If Me.ListObjects(STableActions).ListColumns("Save Without Macro").DataBodyRange.Cells(row_id, 1).Value = "Y" Then
        sTmp_path_xlsx = Environ("temp") & "\" & sFileName & " " & Format(Now(), "yyyy-MM-dd hh-mm-ss") & ".xlsx"
        tmp_wb.SaveAs sTmp_path_xlsx
    End If
    
    With Me.ListObjects(STableActions)
        If sExtension = ".CSV" Then
            If .ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
                arr = Split(.ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value, ",")
                
                sCSVfiles = vbNullString
                For Each sh In tmp_wb.Sheets
                    For i = LBound(arr) To UBound(arr)
                        If Trim(arr(i)) = sh.Name Then
                            sh.Activate
                            
                            ' if only one sheet
                            If LBound(arr) = UBound(arr) Then
                                If .ListColumns("File Name").DataBodyRange.Cells(row_id, 1).Value = vbNullString Then
                                    sFileNameNew = sFileName & " " & ReplaceIllegalChar(sh.Name)
                                Else
                                    sFileNameNew = sFileName
                                End If
                            Else
                            ' many sheets
                            ' then have to add Sheet Name to file name provided
                                sFileNameNew = sFileName & " " & ReplaceIllegalChar(sh.Name)
                                If Left(sFolder, 4) = "http" Then
                                    'sFileNameNew = WorksheetFunction.EncodeURL(sFileNameNew)
                                    sFileNameNew = sFileNameNew
                                End If
                            End If
                            
                            Application.DisplayAlerts = False
                            Application.ScreenUpdating = False
                            
                            ' xlCSVUTF8 exists only in newest versions of Excel 2016
                            ' clear formats
                            sh.Cells.ClearFormats
                            tmp_wb.SaveAs sFolder & sFileNameNew & ".csv", FileFormat:=xlCSVUTF8, CreateBackup:=False
                            sCSVfiles = sCSVfiles & """" & sFolder & sFileNameNew & ".csv" & """" & "+"
                            Exit For
                        End If
                    Next i
                Next sh
            End If
            
            
            If .ListColumns("Combine CSV files").DataBodyRange.Cells(row_id, 1).Value = "Y" Then
                ' files can be combined by using shell command
                ' copy 1.csv + 2.csv + ... N.csv Resulting.csv
                        
                If Left(sFolder, 4) <> "http" Then
                    sCSVfiles = Left(sCSVfiles, Len(sCSVfiles) - 1) ' remove last +
                    'Call Shell("cmd.exe /S /K " & "copy " & sCSVfiles & " " & """" & sFolder & sFileName & ".csv" & """", vbNormalFocus)
                    Call Shell("cmd.exe /S /C " & "copy " & sCSVfiles & " " & """" & sFolder & sFileName & ".csv" & """", vbNormalFocus)
                    ' https://stackoverflow.com/questions/17956651/execute-a-command-in-command-prompt-using-excel-vba
                    ' https://stackoverflow.com/questions/15951837/wait-for-shell-command-to-complete
                End If
                                
                FileSave_XOR_Extension = sFolder & sFileName & ".csv"
            Else
                ' last file
                FileSave_XOR_Extension = sFolder & sFileNameNew & ".csv"
            End If
            GoTo Close_and_Exit
        ElseIf sExtension = ".PDF" Then
            FileSave_XOR_Extension = Save_Sheets_To_PDF(tmp_wb, _
                .ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value, _
                sFolder, sFileName)
            GoTo Close_and_Exit
        End If
    End With
        
    ' ************** part of special actions with target workbook **************
    ' e.g. delete ControlPanel
    ' tmp_wb.Sheets("ControlPanel").Delete
    With Me.ListObjects(STableActions)
        
        ' transform formulas to values for specified sheets
        If .ListColumns("Sheets for Formulas to Values").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
            arr = Split(.ListColumns("Sheets for Formulas to Values").DataBodyRange.Cells(row_id, 1).Value, ",")
            
            For Each sh In tmp_wb.Sheets
                For i = LBound(arr) To UBound(arr)
                    If Trim(arr(i)) = sh.Name Then
                        sh.Activate
                        sh.UsedRange.Value = sh.UsedRange.Value
                        Exit For
                    End If
                Next i
            Next sh
        End If
        
        ' remove unspecified sheets
        If .ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
            arr = Split(.ListColumns("Save Sheets").DataBodyRange.Cells(row_id, 1).Value, ",")
            
            sh_id = 1
            Do While sh_id <= tmp_wb.Sheets.Count
                For i = LBound(arr) To UBound(arr)
                    If Trim(arr(i)) = tmp_wb.Sheets(sh_id).Name Then
                        bFound = True
                        Exit For
                    End If
                Next i
                
                If Not bFound Then
                    tmp_wb.Sheets(sh_id).Visible = xlSheetVisible
                    tmp_wb.Sheets(sh_id).Delete
                    sh_id = sh_id - 1
                End If
                sh_id = sh_id + 1
                bFound = False
            Loop
        End If
        
        ' delete specified sheets
        If .ListColumns("Delete Sheets").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
            arr = Split(.ListColumns("Delete Sheets").DataBodyRange.Cells(row_id, 1).Value, ",")
            
            sh_id = 1
            Do While sh_id <= tmp_wb.Sheets.Count
                For i = LBound(arr) To UBound(arr)
                    If Trim(arr(i)) = tmp_wb.Sheets(sh_id).Name Then
                        tmp_wb.Sheets(sh_id).Visible = xlSheetVisible
                        tmp_wb.Sheets(sh_id).Delete
                        sh_id = sh_id - 1
                        Exit For
                    End If
                Next i
                sh_id = sh_id + 1
            Loop
        End If
        
        ' delete wb queries in target workbook
        If .ListColumns("Delete WB Queries").DataBodyRange.Cells(row_id, 1).Value <> vbNullString Then
            Call DeleteQueriesAndConnections(tmp_wb)
        End If
    End With
    
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
Saving_Workbook:
    If sExtension = ".xlsx" Then
        sResulting_FileFormat = xlOpenXMLWorkbook
    ElseIf sExtension = ".xlsb" Then
        sResulting_FileFormat = xlExcel12
    ElseIf sExtension = ".xlsm" Then
        sResulting_FileFormat = xlOpenXMLWorkbookMacroEnabled
'    ElseIf sExtension = ".CSV" Then
'        sResulting_FileFormat = xlCSV
    End If
    
    tmp_wb.SaveAs sPath, FileFormat:=sResulting_FileFormat, ReadOnlyRecommended:=bReadOnlyRecommended

    FileSave_XOR_Extension = sPath
    
Close_and_Exit:
        
Exit_Sub:
    Application.Calculation = xlCalculationAutomatic
    On Error Resume Next
    Application.DisplayAlerts = False
    tmp_wb.Close ' in case of error in tmp_wb
    Kill sTmp_path
    If Me.ListObjects(STableActions).ListColumns("Save Without Macro").DataBodyRange.Cells(row_id, 1).Value = "Y" Then
        Kill sTmp_path_xlsx
    End If
    ' just in case
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Exit Function
    
ErrHandler:
    ' write log
    Debug.Print Now, "FileSave_XOR_Extension", Err.Number, Err.Description; ""
    Err.Clear
    If bManualRefresh Then Stop
    FileSave_XOR_Extension = vbNullString
    
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function Save_Sheets_To_PDF(Wb As Workbook, sSheets As String, sFolder As String, sFileName As String)
' sSheets expected as comma separated list
    Dim arr
    Dim i As Integer
    
    On Error GoTo ErrHandler
    Debug.Print Now, "Save_Sheets_To_PDF"
    If sSheets <> vbNullString Then
        arr = Split(sSheets, ",")
        
        Wb.Activate
        Wb.Sheets(Trim(arr(LBound(arr)))).Select Replace:=True
        For i = LBound(arr) To UBound(arr)
            Wb.Sheets(Trim(arr(i))).Select Replace:=False
        Next i
        
        ActiveSheet.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=sFolder & sFileName & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    Else
        Wb.Activate
        ActiveWorkbook.ExportAsFixedFormat Type:=xlTypePDF, _
            Filename:=sFolder & sFileName & ".pdf", _
            Quality:=xlQualityStandard, _
            IncludeDocProperties:=False, _
            IgnorePrintAreas:=False, _
            OpenAfterPublish:=False
    End If
    
    Save_Sheets_To_PDF = sFolder & sFileName & ".pdf"
Exit_Sub:

    Exit Function
    
ErrHandler:
    ' write log
    Debug.Print Now, "Save_Sheets_To_PDF", Err.Number, Err.Description
    Save_Sheets_To_PDF = vbNullString
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume ' for debug
End Function

Function Send_EMail_CDO(sFrom As String, _
                    sRecipients As String, _
                    sSubject As String, _
                    Optional sMessage As String, _
                    Optional sAttachmentPath As String, _
                    Optional Importance As String = "Normal", _
                    Optional sCC As String, _
                    Optional sBCC As String)
                                    
    Dim iMsg As Object
    Dim iConf As Object
    Dim strbody As String
    Dim sSendUsing As String
    Dim sAuthentication As String
    Dim Flds
    Dim szServer As String
        
    'https://www.experts-exchange.com/questions/23044027/CDO-Message-sendusing-and-smtpauthenticate.html
    Const cdoSendUsingPickup = 1 'Send message using the local SMTP service pickup directory.
    Const cdoSendUsingPort = 2 'Send the message using the network (SMTP over the network).
    
    Const cdoAnonymous = 0 'Do not authenticate
    Const cdoBasic = 1 'basic (clear-text) authentication
    Const cdoNTLM = 2 'NTLM
 
    Dim oMyMail As Object
    Dim objShell As Object
    
    On Error GoTo ErrHandler
    
    If Me.Names("CP_STMP_SERVER").RefersToRange.Value = vbNullString Then
        If bManualRefresh Then
            MsgBox "Please, provide SMTP server!", vbExclamation
            ThisWorkbook.Activate
            Me.Names("CP_STMP_SERVER").RefersToRange.Select
            Application.ScreenUpdating = True
            Application.Cursor = xlDefault
            Application.EnableEvents = True
            Stop
        Else
            Debug.Print Now, "STMP server is not provided! Send CDO failed."
            Exit Function
        End If
    End If
    
    Set objShell = CreateObject("WScript.Shell")
    Set oMyMail = CreateObject("CDO.Message")
    Set iConf = CreateObject("CDO.Configuration")
    Set Flds = iConf.Fields
    szServer = "http://schemas.microsoft.com/cdo/configuration/"
    
    Select Case "Network"  '[SETTINGS_SMTP_SENDUSING].Value
        Case "Network"
            sSendUsing = cdoSendUsingPort
        Case "Local Catalog"
            sSendUsing = cdoSendUsingPickup
    End Select
    
    If Me.Names("CP_SMTP_USER_NAME").RefersToRange.Value <> vbNullString Then
        sAuthentication = cdoBasic
    Else
        sAuthentication = cdoAnonymous
    End If
    
    With Flds
        .Item(szServer & "sendusing") = sSendUsing
        .Item(szServer & "smtpserver") = Me.Names("CP_STMP_SERVER").RefersToRange.Value
        .Item(szServer & "smtpserverport") = Me.Names("CP_STMP_PORT").RefersToRange.Value
        .Item(szServer & "smtpconnectiontimeout") = 100 '[SETTINGS_SMTP_TIMEOUT].Value ' quick timeout
        .Item(szServer & "smtpauthenticate") = sAuthentication
        .Item(szServer & "smtpusessl") = IIf(Me.Names("CP_USE_SSL").RefersToRange.Value = "Y", True, False)
        
        If sAuthentication = cdoBasic Then
            .Item(szServer & "sendusername") = Me.Names("CP_SMTP_USER_NAME").RefersToRange.Value
            
            .Item(szServer & "sendpassword") = GetPassword(Me.Names("CP_STMP_SERVER").RefersToRange.Value, _
                Me.Names("CP_SMTP_USER_NAME").RefersToRange.Value)
        Else
            .Item(szServer & "smtpauthenticate") = 0
        End If
        .Update
    End With
    
    With oMyMail
            Set .Configuration = iConf
            .bodypart.Charset = "utf-8"
            .To = sRecipients
            .cc = sCC
            .From = IIf(Me.Names("CP_SMTP_USER_NAME").RefersToRange.Value <> vbNullString, _
                        Me.Names("CP_SMTP_USER_NAME").RefersToRange.Value, _
                        IIf(sFrom <> vbNullString, sFrom, "BOA Automation Tool"))
                        
            .Subject = sSubject
                                    
            .htmlbody = Replace(sMessage, Chr(10), "<br>")
            
            '& vbCrLf & _
                vbCrLf & _
                ThisWorkbook.Name & vbCrLf & _
                objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
            
            '.CreateMHTMLBody "file://c|/temp/test.htm"
            
            If sAttachmentPath <> vbNullString Then
                .AddAttachment sAttachmentPath
            End If
            
            .Send
    End With
    
    Send_EMail_CDO = True

Exit_Sub:
    Set oMyMail = Nothing
    Set iConf = Nothing
    Set Flds = Nothing
    Exit Function

ErrHandler:
    ' write_log...
    Debug.Print Now, "Send_EMail_CDO", Err.Number, Err.Description
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Function

Function Send_Email_Outlook(sRecipients As String, _
                    sSubject As String, _
                    Optional sMessage As String, _
                    Optional sAttachmentPath As String, _
                    Optional Importance As String = "Normal", _
                    Optional sCC As String, _
                    Optional sBCC As String, _
                    Optional sFrom As String)
                    
    Dim oOutlook As Object
    Dim oMyMail As Object
    Dim objShell As Object
    
    Const olMailItem = 0
    
    ' get or create outlook
    On Error Resume Next
    Set oOutlook = GetObject(, "Outlook.Application")
    Err.Clear
    
    On Error GoTo ErrHandler
    If oOutlook Is Nothing Then
        Set oOutlook = CreateObject("Outlook.Application")
    End If
    
    Set objShell = CreateObject("WScript.Shell")
    Set oMyMail = oOutlook.CreateItem(olMailItem)
    
    With oMyMail
        If sFrom <> vbNullString Then
            .SentOnBehalfOfName = sFrom
        End If
    
        .To = Replace(sRecipients, ",", ";")
        
        If sCC <> vbNullString Then
            .cc = Replace(sCC, ",", ";")
        End If
        If sBCC <> vbNullString Then
            .Bcc = Replace(sBCC, ",", ";")
        End If
        
        .Subject = sSubject
        .BodyFormat = 2 ' olFormatHTML
        
        .htmlbody = sMessage
        '& vbCrLf & _
            vbCrLf & _
            ThisWorkbook.Name & vbCrLf & _
            objShell.ExpandEnvironmentStrings("%COMPUTERNAME%")
        
        If sAttachmentPath <> vbNullString Then
            .Attachments.Add sAttachmentPath
        End If
        
        ' Normal importance is default
        Select Case Importance
            Case "High"
                .Importance = enumMailImportance.High
            Case "Low"
                .Importance = enumMailImportance.Low
        End Select
        
        .Send
    End With
    
    Send_Email_Outlook = True

Exit_Sub:
    Set oMyMail = Nothing
    Set oOutlook = Nothing
    Exit Function

ErrHandler:
    ' write_log...
    Debug.Print Now, "Send_Email_Outlook", Err.Number, Err.Description
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Function

Function ReplaceIllegalChar(strIn As String) As String
' https://www.experts-exchange.com/questions/28025657/Vba-Code-Eliminate-Illegal-Characters-from-a-filename.html
' adjusted to own char list
    Dim j As Integer
    Dim varStr As String, xStr As String
    varStr = strIn
    For j = 1 To Len(varStr)
       Select Case Asc(Mid(varStr, j, 1))
            ' excuded resticted characters for OneDrive
            
            Case 32, 38, 39, 43, 45, 46, 48 To 57, 61, 64 To 91, 93 To 123, 125 To 126, 130 To 142, 145 To 151, 153 To 156, 158 To 255
            xStr = xStr & Mid(varStr, j, 1)
       Case Else
            xStr = xStr & "_"
       End Select
    Next
    ReplaceIllegalChar = xStr
End Function

Function RefreshWorkbook(Optional Wb As Workbook) As Boolean
    Dim cnct As Variant
    Dim slc As SlicerCache
    Dim BeforeAction
    Dim target_wb As Workbook
    Dim bCubeFormulasFound As Boolean
    Dim bScreenUpdatingInitial As Boolean
    Dim bEnableEventsInitial As Boolean
    Dim CalcModeInitial As Double
    Dim CursorStateInitial As Double
    
    On Error GoTo ErrHandler
    Debug.Print Now, "Updating connections..."
    
    With Application
        bScreenUpdatingInitial = .ScreenUpdating
        bEnableEventsInitial = .EnableEvents
        CalcModeInitial = .Calculation
        CursorStateInitial = .Cursor
        
        ' switch everything off
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .Cursor = xlWait
    End With
    
    If Wb Is Nothing Then
        Set target_wb = ThisWorkbook
    Else
        Set target_wb = Wb
    End If
    
    On Error Resume Next
    If IsError(target_wb.Model.ModelTables.Count) Then
        ' cannot access model
        ' do nothing
    Else
        If target_wb.Model.ModelTables.Count > 0 Then
            Application.StatusBar = "Initializing Data Model..."
            target_wb.Model.Initialize
            WaitSeconds 5
        End If
    End If
    
    Err.Clear
    On Error GoTo ErrHandler
    
    ' deny background refresh
    ' ToThink - probably worth to restore initial settings
    ' however, if workbook is done for Power Refresh solution, it should not contain "background" connections
    ' create 2D array, restore settings after update
    Application.StatusBar = "Switching off background refresh..."
    For Each cnct In target_wb.Connections
        Select Case cnct.Type
            Case xlConnectionTypeODBC
                cnct.ODBCConnection.BackgroundQuery = False
            Case xlConnectionTypeOLEDB
                cnct.OLEDBConnection.BackgroundQuery = False
        End Select
    Next cnct
    
    Application.StatusBar = "Refreshing Data Model and Connections..."
    target_wb.RefreshAll
    WaitSeconds 1
    Application.CalculateUntilAsyncQueriesDone
    WaitSeconds 1
    
    For Each cnct In target_wb.Connections
        Select Case cnct.Type
            Case xlConnectionTypeODBC
                Do While cnct.ODBCConnection.Refreshing
                    WaitSeconds 1
                Loop
            Case xlConnectionTypeOLEDB
                Do While cnct.OLEDBConnection.Refreshing
                    WaitSeconds 1
                Loop
        End Select
    Next cnct
    
    Application.StatusBar = "Calculating after connections refresh..."
    Application.Calculate
    Application.CalculateUntilAsyncQueriesDone
    WaitSeconds 1
    
    Application.StatusBar = "Checking existence of cube formulas..."
    bCubeFormulasFound = IsWBHasCubeFormulas(target_wb)
    
    ' update cache after Model refresh
    ' ignore all possible errors with slicers
    On Error Resume Next
    Application.StatusBar = "Updating slicers..."
    For Each slc In target_wb.SlicerCaches
        slc.ClearManualFilter
        slc.ClearAllFilters
        'slc.ClearDateFilter
    Next slc
    Err.Clear
    On Error GoTo ErrHandler
    ' if needed, slicer default value can be set in BeforeSave event of target workbook, or in custom macro
    
    If bCubeFormulasFound Then
        ' wait for refresh of cube formulas
        If target_wb.SlicerCaches.Count > 0 Then
            Application.StatusBar = "Calculating after slicers refresh..."
            Application.Calculate
            Application.CalculateUntilAsyncQueriesDone
        End If
        
        Application.StatusBar = "Waiting for cube formulas..."
        WaitSeconds 20
    End If
    
    If Not Application.CalculationState = xlDone Then
        ' infinite loop can be trully infinite
        ' so just delay
        Application.StatusBar = "Waiting for application to calculate..."
        WaitSeconds 5
    End If
        
    RefreshWorkbook = True

Exit_Function:
    On Error Resume Next
    
    ' restore initial state
    With Application
        .ScreenUpdating = bScreenUpdatingInitial
        .EnableEvents = bEnableEventsInitial
        .Calculation = CalcModeInitial
        .Cursor = CursorStateInitial
        .StatusBar = vbNullString
    End With
    
    Exit Function
    
ErrHandler:
    Debug.Print Now, "Update Connections", Err.Number, Err.Description, Application.StatusBar
    
    If bManualRefresh Then
        Application.Cursor = xlDefault
        Stop
    End If
    
    Err.Clear
    GoTo Exit_Function
    Resume ' for debug purpose
End Function

Private Function IsWBHasCubeFormulas(Optional Wb As Workbook) As Boolean
    Dim sh As Worksheet
    Dim cell As Range
    Dim bFound As Boolean
    Dim bScreenUpdatingInitial As Boolean
    Dim bEnableEventsInitial As Boolean
    Dim CalcModeInitial As Integer
    Dim rngFormulas As Range
    
    On Error GoTo ErrHandler
    
    If Wb Is Nothing Then
        Set Wb = ThisWorkbook ' ActiveWorkbook ' alternatively
    End If
    
    With Application
        bScreenUpdatingInitial = .ScreenUpdating
        bEnableEventsInitial = .EnableEvents
        CalcModeInitial = .Calculation
        
        ' switch everything off
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    For Each sh In Wb.Sheets
        'Debug.Print sh.Name
        
        Err.Clear
        On Error Resume Next
        Set rngFormulas = sh.Cells.SpecialCells(xlCellTypeFormulas)
        bFound = (Err.Number = 0) ' no error, means SpecialCells returned non-empty range
        Err.Clear
        On Error GoTo ErrHandler
        
        ' if result of SpecialCells was non-empty - check formulas
        If bFound Then
            For Each cell In rngFormulas
                'Debug.Print cell.Formula
                If Left(cell.Formula, 5) = "=CUBE" Then
                    IsWBHasCubeFormulas = True
                    GoTo Exit_Function
                End If
            Next cell
        End If
    Next sh
    
Exit_Function:
    On Error Resume Next
    
    ' restore initial state
    With Application
        .ScreenUpdating = bScreenUpdatingInitial
        .EnableEvents = bEnableEventsInitial
        .Calculation = CalcModeInitial
    End With
    
    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "IsWBHasCubeFormulas", Err.Number & ": " & Err.Description
        Err.Clear
    End If
    
    GoTo Exit_Function
    Resume ' for debug purpose
End Function

Private Sub ClearCrosstabs()
    Dim r As Long
    
    On Error GoTo ErrHandler
    
    If bCallActionFromDStable Then
        ' do not go through all rows
        ' clear only current ds row
        If Not ThisWorkbook.Names(CStr(arrDS(currentDS_index, enumDataSource.DS_Crosstab))).RefersToRange Is Nothing Then
            With ThisWorkbook.Names(CStr(arrDS(currentDS_index, enumDataSource.DS_Crosstab))).RefersToRange
                    If .Rows.Count > 1 Then
                        .Offset(1, 0).Resize(.Rows.Count - 1, _
                                             .Columns.Count).Clear
                    End If
            End With
        End If
    End If
    
    With Me.ListObjects(sTableDataSourcesName)
        If .DataBodyRange Is Nothing Then
            Exit Sub
        Else
            For r = 1 To .DataBodyRange.Rows.Count
                If CStr(.ListColumns("Scope").DataBodyRange.Cells(r, 1).Value) = _
                                CStr(Me.Names("CP_SCOPE").RefersToRange.Value) Then
                    
                    'If .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" And
                    If .ListColumns("Clear Crosstab").DataBodyRange.Cells(r, 1).Value = "Y" Then
                        ' get range of crosstab
                        With ThisWorkbook.Names(.ListColumns("Crosstab").DataBodyRange.Cells(r, 1).Value).RefersToRange
                            If .Rows.Count > 1 Then
                                .Offset(1, 0).Resize(.Rows.Count - 1, _
                                                     .Columns.Count).Clear
                                
                            End If
                        End With
                    End If
                End If
            Next r ' row of Data Sources table
        End If
    End With

Exit_Sub:
    
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "ClearCrosstabs", Err.Number, Err.Description
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Sub

Private Sub FillHeaders(sCrosstab As String)
    Dim i As Long
    Dim col As Integer
    Dim bStyleMode As Boolean
    Dim bMultiHeaderMode As Boolean
    
    On Error GoTo ErrHandler
    If ThisWorkbook.Names(sCrosstab).RefersToRange Is Nothing Then Exit Sub
    
    With ThisWorkbook.Names(sCrosstab).RefersToRange
        If .Cells(1, 1).Style.Name = "SAPDimensionCell" Then
            bStyleMode = True
        End If ' check style in first cell
        
        If .Rows.Count > 1 Then
            If bStyleMode Then
                If .Cells(2, 1).Style.Name = "SAPDimensionCell" Then
                    bMultiHeaderMode = True
                End If
            Else
                If .Cells(1, 1).Value = vbNullString Then
                    bMultiHeaderMode = True
                End If
            End If
        End If
        
        ' check first cell is empty - Show Scaling Factor is enabled
        For i = 1 To .Rows.Count
            ' when found non-empty cell in 1st column
            If .Cells(i, 1) <> vbNullString Then
                ' fill columns
                For col = 2 To .Columns.Count
                    ' when found empty cell - copy value from prior cell + ' Name'
                    If .Cells(i, col).Value = vbNullString Then
                        ' when we know style - it is safe method, we don't fill Scaling Factors
                        If bStyleMode Then
                            If .Cells(i, col).Style.Name = "SAPDimensionCell" Then
                                .Cells(i, col).Value = .Cells(i, col - 1).Value & " Name"
                            ElseIf .Cells(i, col).Style.Name = "SAPMemberCell" Then
                                ' not repeated values of dimensions
                                .Cells(i, col).Value = .Cells(i, col - 1).Value
                            End If
                        Else
                        '   when we don't know style - have to check
                            If bMultiHeaderMode Then
                                ' do not fill if there is something in prev row (possibly header)
                                If .Cells(i - 1, col).Value = vbNullString Then
                                    .Cells(i, col).Value = .Cells(i, col - 1).Value & " Name"
                                End If
                            Else
                                ' only one row - fill all columns
                                .Cells(i, col).Value = .Cells(i, col - 1).Value & " Name"
                            End If
                        End If
                    End If ' If .Cells(i, col).Value = vbNullString Then
                Next col
                
                Exit For ' exit when first found line is filled
            Else
            ' first cell is null - fill only when we know cell's style
                If bStyleMode Then
                    For col = 2 To .Columns.Count
                        If .Cells(i, col).Value = vbNullString And _
                           .Cells(i, col).Style.Name = "SAPMemberCell" Then
                                ' for measure headers - fill members of dimensions
                            .Cells(i, col).Value = .Cells(i, col - 1).Value
                        End If
                    Next col
                End If
            End If ' if first row is not null
        Next i
            
    End With

Exit_Sub:
    
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "FillHeaders", Err.Number, Err.Description
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Sub

Private Sub CopySheetValues(target_wb As Workbook, shName As String)
    Dim sh As Worksheet
    
    On Error GoTo ErrHandler
    With ThisWorkbook.Sheets(shName)
        .Activate
        .Range(.Cells(1, 1), .UsedRange.SpecialCells(xlCellTypeLastCell)).Copy
    End With
        
    On Error Resume Next
    target_wb.Sheets(shName).Activate
    If Err.Number <> 0 Then
        Err.Clear
        Set sh = target_wb.Sheets.Add ' (after:=target_wb.Sheets.Count)
        sh.Name = shName
    End If
    
    On Error GoTo ErrHandler
    With target_wb.Sheets(shName)
        .Activate
        With .[A1]
            .PasteSpecial xlPasteValues ' just values
            .PasteSpecial xlPasteColumnWidths
            .PasteSpecial xlPasteFormats ' formats of cells - fill, borders, + Conditional Formatting etc.
            .PasteSpecial xlPasteComments
        End With
    End With
    
    Application.CutCopyMode = False
    Set sh = Nothing

Exit_Sub:
    
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "CopySheetValues", Err.Number, Err.Description
    Err.Clear
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Sub


Private Sub DeleteQueriesAndConnections(target_wb As Workbook)
    
    On Error Resume Next
    Application.Cursor = xlWait
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    Do While target_wb.Queries.Count > 0
        target_wb.Queries(1).Delete
    Loop
    
    Do While target_wb.Connections.Count > 0
        If target_wb.Connections(1).Name <> "ThisWorkbookDataModel" Then
            target_wb.Connections(1).Delete
        Else
            If target_wb.Connections.Count = 1 Then Exit Do
            target_wb.Connections(2).Delete
        End If
    Loop
    
End Sub

Private Sub SortDataSources()
    On Error GoTo ErrHandler
    Application.EnableEvents = False
    With Me.ListObjects(sTableDataSourcesName)
        .Sort.SortFields.Clear
        
        .Sort.SortFields.Add Key:=.ListColumns("Scope Order").Range, _
                             SortOn:=xlSortOnValues, _
                             order:=xlAscending, _
                             DataOption:=xlSortNormal
    
        .Sort.SortFields.Add Key:=.ListColumns("Order").Range, _
                             SortOn:=xlSortOnValues, _
                             order:=xlAscending, _
                             DataOption:=xlSortNormal
                             
        .Sort.SortFields.Add Key:=.ListColumns("Sheet").Range, _
                             SortOn:=xlSortOnValues, _
                             order:=xlAscending, _
                             DataOption:=xlSortNormal
                             
        .Sort.SortFields.Add Key:=.ListColumns("Data Source Name").Range, _
                             SortOn:=xlSortOnValues, _
                             order:=xlAscending, _
                             DataOption:=xlSortNormal
                             
        With .Sort
            .Header = xlYes
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End With

Exit_Sub:
    
    Exit Sub
    
ErrHandler:
    Debug.Print Now, "SortDataSources", Err.Number, Err.Description
    Err.Clear
    Application.ScreenUpdating = True
    Application.Cursor = xlDefault
    Application.EnableEvents = True
    
    If bManualRefresh Then Stop
    GoTo Exit_Sub
    Resume
End Sub

' *************************************************** EXPERIMENTAL PROCEDURES ***************************************************
' *******************************************************************************************************************************
Sub FakeCheckIfNoData(Optional t As String)
    ' use this sub to test how Delay Functionality works
    ' put FakeCheckIfNoData into MacrosAfter settings in data source table
    If dummy = 0 Then
        Me.Names("CP_DELAY").RefersToRange.Value = "Y"
        dummy = 1
    Else
        Me.Names("CP_DELAY").RefersToRange.Value = "N"
    End If
End Sub

Sub FakeTurnRepeatRefreshOn(Optional t As String)
    ' use this sub to test how Repeat Refresh works
    ' put FakeTurnRepeatRefreshOn into MacrosAfter settings in data source table
    If dummy = 0 Then
        Me.Cells(arrDS(currentDS_index, enumDataSource.DS_Row), _
                Me.ListObjects(sTableDataSourcesName).ListColumns("Repeat Refresh").Range.Column).Value = "Y"
        dummy = 1
    Else
        Me.Cells(arrDS(currentDS_index, enumDataSource.DS_Row), _
                Me.ListObjects(sTableDataSourcesName).ListColumns("Repeat Refresh").Range.Column).Value = vbNullString
    End If
End Sub

Private Sub Style_Remover()
    Dim i As Long
    Application.ScreenUpdating = True
    'Debug.Print ThisWorkbook.Styles.Count
    
    On Error Resume Next
    For i = 1 To 1000000
        If Not ThisWorkbook.Styles(i).BuiltIn Then
            ThisWorkbook.Styles(i).Delete
            If Err.Number = 0 Then
                i = i - 1
            End If
            Err.Clear
        End If
    Next i

    Application.ScreenUpdating = True
End Sub

Private Sub PrepareForPublishing()
    ' clean tables and cells with parameters
    '
    Dim i As Long
    
    Call ListObjectDeleteEmptyRows(Me.ListObjects(sTableDataSourcesName))
    
    On Error Resume Next
    Me.ListObjects(sTableDataSourcesName).DataBodyRange.Delete
    Me.ListObjects(sTableVariablesName).DataBodyRange.Delete
    
    With Me.ListObjects(STableActions).ListColumns("Enabled").DataBodyRange
        For i = 1 To .Rows.Count
            .Cells(i, 1).Value = "N"
            Me.ListObjects(STableActions).ListColumns("Email From").DataBodyRange.Cells(i, 1).Value = vbNullString
            Me.ListObjects(STableActions).ListColumns("Email To").DataBodyRange.Cells(i, 1).Value = vbNullString
            Me.ListObjects(STableActions).ListColumns("Email CC").DataBodyRange.Cells(i, 1).Value = vbNullString
        Next i
    End With
    
    Me.Names("CP_LAST_REPORT_DATE").RefersToRange.Value = vbNullString
    Me.Names("CP_LAST_REFRESH_DATETIME").RefersToRange.Value = vbNullString
    Me.Names("CP_PASSWORDS_PATH").RefersToRange.Value = vbNullString
    Me.Names("CP_GENERAL_USER").RefersToRange.Value = vbNullString
    Me.Names("CP_STMP_SERVER").RefersToRange.Value = vbNullString
    Me.Names("CP_STMP_PORT").RefersToRange.Value = vbNullString
    Me.Names("CP_USE_SSL").RefersToRange.Value = vbNullString
    Me.Names("CP_SMTP_USER_NAME").RefersToRange.Value = vbNullString
    Me.Names("CP_STMP_PORT").RefersToRange.Value = vbNullString
    
    Call Style_Remover
    
End Sub
 
Private Sub test()
    Debug.Print Me.PivotTables.Count
End Sub

Private Sub CreateNames()
    
    Const ReportDateCell = "CP_REPORT_DATE"
    
    On Error Resume Next
    
    ' names must be visible to be used in formulas by end-user
    '    Me.Names("CP_DATE_RDM_START_END_DDMMYYYY").Visible = True
    
    ' Report Date Year
    Me.Names.Add Name:="CP_DATE_RDY", RefersTo:="=YEAR( CP_REPORT_DATE )"
    ' Prior to 'Report Date Month' Year
    Me.Names.Add Name:="CP_DATE_PRDM_YYYY", RefersTo:="=YEAR( EOMONTH( CP_REPORT_DATE, -1) )"
    
    ' Report Date Month
    Call CreateNamesWithFormats("RDM", "EOMONTH( CP_REPORT_DATE, -1 )+1", "EOMONTH( CP_REPORT_DATE, 0 )")
        
    ' Prior to Report Date Month
    Call CreateNamesWithFormats("PRDM", "EOMONTH( CP_REPORT_DATE, -2 )+1", "EOMONTH( CP_REPORT_DATE, -1 )")
    
    ' Prev to Prior to Report Date Month
    Call CreateNamesWithFormats("PPRDM", "EOMONTH( CP_REPORT_DATE, -3 )+1", "EOMONTH( CP_REPORT_DATE, -2 )")
        
    ' Last Two Months Excluding = LTM_EXC
    Call CreateNamesWithFormats("LTM_EXC", "EOMONTH( CP_REPORT_DATE, -3 )+1", "EOMONTH( CP_REPORT_DATE, -1 )")
        
    ' RDM in PY - Same RD month in Prior Year
    Call CreateNamesWithFormats("RDM_IN_PY", "EOMONTH( CP_REPORT_DATE, -13 )+1", "EOMONTH( CP_REPORT_DATE, -12 )")
        
    ' PRDM in PY, 'Prior to RDM' in Prior Year
    Call CreateNamesWithFormats("PRDM_IN_PY", "EOMONTH( CP_REPORT_DATE, -14 )+1", "EOMONTH( CP_REPORT_DATE, -13 )")
    
    ' R12m RDM INC - Rolling 12 months including RDM
    Call CreateNamesWithFormats("R12M_RDM_INC", "EOMONTH( CP_REPORT_DATE, -12 )+1", "EOMONTH( CP_REPORT_DATE, 0 )")
    
    ' R12m RDM EXC - Rolling 12 months to RDM excluding
    Call CreateNamesWithFormats("R12M_RDM_EXC", "EOMONTH( CP_REPORT_DATE, -13 )+1", "EOMONTH( CP_REPORT_DATE, -1 )")
    
    ' R4m RDM INC - Rolling 4 months to RDM including
    Call CreateNamesWithFormats("R4M_RDM_INC", "EOMONTH( CP_REPORT_DATE, -4 )+1", "EOMONTH( CP_REPORT_DATE, 0 )")
        
    ' R4m RDM EXC - Rolling 4 months to RDM including
    Call CreateNamesWithFormats("R4M_RDM_EXC", "EOMONTH( CP_REPORT_DATE, -5 )+1", "EOMONTH( CP_REPORT_DATE, -1 )")
        
    ' CP_DATE_YTD_RDM - YTD to RDM including
    Call CreateNamesWithFormats("YTD_RDM", "DATE( YEAR(CP_REPORT_DATE),1,1)", "EOMONTH( CP_REPORT_DATE, 0 )")
        
    ' CP_DATE_YTD_PRDM - YTD calculated for 'Prior RDM'
    Call CreateNamesWithFormats("YTD_PRDM", "DATE( YEAR(EOMONTH(CP_REPORT_DATE,-1)),1,1)", "EOMONTH( CP_REPORT_DATE, -1 )")
            
    ' YTD PY to RD month - YTD calculated for 'RDM in PY'
    Call CreateNamesWithFormats("YTD_RDM_IN_PY", "DATE( YEAR(CP_REPORT_DATE)-1,1,1)", "EOMONTH( CP_REPORT_DATE, -12 )")
    
    ' YTD PY to prior RD month - YTD calculated for 'PRDM in PY'
    Call CreateNamesWithFormats("YTD_PRDM_IN_PY", "DATE( YEAR(EOMONTH(CP_REPORT_DATE,-1))-1,1,1)", "EOMONTH( CP_REPORT_DATE, -13 )")
    
    ' PY_CY_TO_RDM including. E.g. 01.2016 - 10.2017
    Call CreateNamesWithFormats("PY_CY_TO_RDM", "DATE( YEAR(CP_REPORT_DATE)-1,1,1)", "EOMONTH( CP_REPORT_DATE, 0 )")
    
    ' PY_CY_TO_PRDM including. E.g. 01.2016 - 09.2017
    Call CreateNamesWithFormats("PY_CY_TO_PRDM", "DATE( YEAR(EOMONTH(CP_REPORT_DATE,-1))-1,1,1)", "EOMONTH( CP_REPORT_DATE, -1 )")
    
    ' PY_R12M_RDM, R12m in PY to 'RDM in PY', 11.2015 - 10.2016
    Call CreateNamesWithFormats("PY_R12M_RDM", "EOMONTH( CP_REPORT_DATE, -24 )+1", "EOMONTH( CP_REPORT_DATE, -12 )")
    
    ' PY_R12M_PRDM, R12m in PY to 'RDM in PY', 10.2015 - 09.2016
    Call CreateNamesWithFormats("PY_R12M_PRDM", "EOMONTH( CP_REPORT_DATE, -25 )+1", "EOMONTH( CP_REPORT_DATE, -13 )")
    
    ' PY_R4M_RDM, R4m in PY for RDM, 07.2016 - 10.2016
    Call CreateNamesWithFormats("PY_R4M_RDM", "EOMONTH( CP_REPORT_DATE, -12-4 )+1", "EOMONTH( CP_REPORT_DATE, -12 )")
    
    ' PY_R4M_PRDM, R4m in PY for PRDM, 06.2016 - 09.2016
    Call CreateNamesWithFormats("PY_R4M_PRDM", "EOMONTH( CP_REPORT_DATE, -13-4 )+1", "EOMONTH( CP_REPORT_DATE, -13 )")
    
    ' PRDY_FULL, Prior to Report Date Year - Full, 01.2016 - 12.2016
    Call CreateNamesWithFormats("PRDY_FULL", "DATE( YEAR(CP_REPORT_DATE)-1,1,1)", "DATE( YEAR(CP_REPORT_DATE)-1,12,31)")
    
    ' RDY_FULL, Report Date Year Full, 01.2017 - 12.2017
    Call CreateNamesWithFormats("RDY_FULL", "DATE( YEAR(CP_REPORT_DATE),1,1)", "DATE( YEAR(CP_REPORT_DATE),12,31)")
    
End Sub

Private Sub CreateNamesWithFormats(sID As String, sStartFormula As String, sEndFormula As String)
    Call CreateName(sID, "MMYYYY", sStartFormula, sEndFormula)
    Call CreateName(sID, "0MMYYYY", sStartFormula, sEndFormula)
    Call CreateName(sID, "DDMMYYYY", sStartFormula, sEndFormula)
End Sub

Private Sub CreateName(sID As String, sFormat As String, sStartFormula As String, sEndFormula As String)
    Const sPartStart = "START"
    Const sPartEnd = "END"
    Const sPart = "START_END"
    Dim sFullID As String
    Dim sFullIDStart As String
    Dim sFullIDEnd As String
    Dim sFullIDStartDate As String
    Dim sFullIDEndDate As String
    
    Const sMainPrefix = "CP_DATE"
    
    On Error Resume Next
    
    ' As Date
    sFullIDStartDate = sMainPrefix & "_" & sID & "_" & sPartStart & "_DATE"
    Me.Names.Add Name:=sFullIDStartDate, RefersTo:="=" & sStartFormula
    
    sFullIDEndDate = sMainPrefix & "_" & sID & "_" & sPartEnd & "_DATE"
    Me.Names.Add Name:=sFullIDEndDate, RefersTo:="=" & sEndFormula
    
    ' FORMATTED
    ' START
    sFullIDStart = sMainPrefix & "_" & sID & "_" & sPartStart & "_" & sFormat
    Me.Names.Add Name:=sFullIDStart, _
        RefersTo:="=TEXT( " & sFullIDStartDate & ", CP_DATE_FORMAT_" & sFormat & " )"
    ' END
    sFullIDEnd = sMainPrefix & "_" & sID & "_" & sPartEnd & "_" & sFormat
    Me.Names.Add Name:=sFullIDEnd, _
        RefersTo:="=TEXT( " & sFullIDEndDate & ", CP_DATE_FORMAT_" & sFormat & " )"
    ' START - END
    sFullID = sMainPrefix & "_" & sID & "_" & sPart & "_" & sFormat
    Me.Names.Add Name:=sFullID, RefersTo:="=" & sFullIDStart & "& "" - "" & " & sFullIDEnd
    
End Sub


