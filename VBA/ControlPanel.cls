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

' Author: Ivan Bondarenko
' Initial release: 2017-01
' https://bondarenkoivan.wordpress.com
' https://linkedin.com/in/bondarenkoivan/en
' Big Thanks to Andres Merino for ideas and testing!

' update 2017-05-24
' now macro handles cases with SAPSetVariables only
' no need to refresh twice in such cases

' declaration of function that is used in Function CreatePath
' possibly will be used for logging
#If VBA7 Then
    Private Declare PtrSafe Function MakeSureDirectoryPathExists Lib _
        "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
#Else
    Private Declare Function MakeSureDirectoryPathExists Lib _
        "IMAGEHLP.DLL" (ByVal DirPath As String) As Long
#End If

Const ForAppending = 8 ' for OpenTextFile
Const ForReading = 1

Public Logs_Enabled As Boolean
Public LogFile As String

Public objFSO As Object
Public objLog As Object
Dim arrDS ' array for data sources
Dim currentDS As String ' current data source

Dim bMultiLogonMode As Boolean
' special mode when refresh is done data source be data source
' and after each data source macro run SAPLogoff
' to be sure that next data source will be refreshed with defined language
' and defined user

Dim NamedRangesScope ' can be Workbook or Worksheet

Enum enumDataSource ' used to reference to elements of arrDS array
    DS_ID = 0
    DS_Refresh = 1
    DS_Language = 2
    DS_System = 3
    DS_User = 4
    DS_MacrosAfter = 5
    DS_Client = 6
    DS_CalculateAfter = 7
End Enum

Dim dummy As Integer
'

Private Sub DefineGlobalVariables()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Left(ThisWorkbook.Path, 4) <> "http" Then ' disable log for workbooks on SharePoint
        Logs_Enabled = True ' TODO - add cell with parameter
        LogFile = ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".log"
    End If
    
    Call GetArrDS  ' populate array with data sources
        
End Sub

'Private Sub DefineNamedRangeScope()
'    Set NamedRangesScope = Me
'    ' all needed named ranged are in ControlPanel sheet scope
'
'    'On Error Resume Next
'    ' if user copy ControlPanel several times and do not delete #REF named ranges after deletion of ControlPanel
'    ' Excel makes ControlPanel's named ranges in scope of worksheet
'    'Debug.Print Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Address
'    'If Err.Number = 0 Then
'    '    Set NamedRangesScope = Me
'    'Else
'    '    Set NamedRangesScope = ThisWorkbook
'    'End If
'End Sub

Function Refresh() As Boolean
    Dim StartPoint As Double
    Dim cell_scope As Range
    Dim i As Long
    Dim m As Integer
    Dim arrMacrosBefore
    Dim arrMacrosAfter
    Dim arrMacrosAfterDSRefresh
    Dim Attempt As Long
    Dim bYetAnotherAttempt As Boolean
    Dim bSAPSetVariablesOnly As Boolean ' when scope filters contains no SAPSetFilter command
    
    On Error GoTo ErrHandler
    
    StartPoint = Now()
    ThisWorkbook.Windows(1).Activate ' otherwise Logon fails, if another wb is active
    
    Application.StatusBar = "Defining global variables..."
    Call DefineGlobalVariables
'    Call DefineNamedRangeScope
    
    Application.StatusBar = "Enabling BOA..."
    Call EnableBOA
    
    ' when only one language is used -
    ' it is enough to logon only once with defined language
    
    If Not IsArray(arrDS) Then
        MsgBox "Please, add at least one data source in table DATA_SOURCES!", vbExclamation
        'Exit Sub
        ' remain it false
        GoTo Exit_Function
    End If
    
    Application.StatusBar = "Executing macros before Refresh..."
    arrMacrosBefore = Split(Me.Names("MACROS_BEFORE").RefersToRange.Value, ",")
    For m = 0 To UBound(arrMacrosBefore)
        On Error Resume Next
        Application.StatusBar = "Executing macro " & Trim(arrMacrosBefore(i)) & "..."
        Run Trim(arrMacrosBefore(i))
        If Err.Number <> 0 Then
            Debug.Print "Couldn't run macro " & Trim(arrMacrosBefore(i))
            GoTo ErrHandler
        End If
        On Error GoTo 0
        
        ' ? do we need to calc after each macro? - macro owner decides
        ' if needed - user will have to add it in his/her macro
        'Application.StatusBar = "Calculating..."
        'Application.Calculate
    Next m
    
    Application.StatusBar = "Log On..."
    If Not LogOn_DataSources Then
        Debug.Print "Couldn't LogOn"
        GoTo ErrHandler
    End If
    
    ' initial calculate
    Application.StatusBar = "Calculating..."
    Application.Calculate
    Application.StatusBar = vbNullString
    
    ' refresh all enabled data sources in specified order
    For i = 1 To UBound(arrDS, 1)
        ' array elements are ordered
        ' macro refreshes only those that enabled
        If arrDS(i, 1) = "Y" Then
            Attempt = 0 ' attempt to refresh in case of delay (if no new data)
            bYetAnotherAttempt = False
            bSAPSetVariablesOnly = IsSAPSetVariablesOnly(CStr(arrDS(i, enumDataSource.DS_ID)))
            currentDS = CStr(arrDS(i, enumDataSource.DS_ID))
            
Start_Refresh:
            On Error GoTo ErrHandler
                                    
            ' if another attempt - we have to logon again
            If bYetAnotherAttempt Then
                Application.StatusBar = "Log On..."
                If Not LogOn_DataSources Then
                    Debug.Print "Couldn't LogOn"
                    GoTo ErrHandler
                End If
            End If
            
            ' if within one Scope we should refresh several data sources with different languages
            ' e.g. everything is available in local language except one data source, which should be refreshed in EN
            If bMultiLogonMode Then
                If i = 1 Then ' if first data source
                ' logon with specified Language
                    Application.StatusBar = "Logging On " & arrDS(i, enumDataSource.DS_ID)
                    If Run("SAPLogon", arrDS(i, enumDataSource.DS_ID), "100", arrDS(i, enumDataSource.DS_User), _
                            GetPassword(CStr(arrDS(i, enumDataSource.DS_System)), CStr(arrDS(i, enumDataSource.DS_User))), _
                            arrDS(i, enumDataSource.DS_Language)) = 0 Then
                        
                        Debug.Print "Couldn't LogOn to " & CStr(arrDS(i, enumDataSource.DS_System))
                        GoTo ErrHandler
                        ' couldn't logon
                    End If
                Else
                    ' if previous language differs from current
                    ' in the end of prev cycle macro logged off
                    If arrDS(i, enumDataSource.DS_Language) <> arrDS(i - 1, enumDataSource.DS_Language) Then
                        Application.StatusBar = "Logging On " & arrDS(i, enumDataSource.DS_ID)
                        If Run("SAPLogon", arrDS(i, enumDataSource.DS_ID), "100", arrDS(i, enumDataSource.DS_User), _
                                GetPassword(CStr(arrDS(i, enumDataSource.DS_System)), CStr(arrDS(i, enumDataSource.DS_User))), _
                                arrDS(i, enumDataSource.DS_Language)) = 0 Then
                            
                            Debug.Print "Couldn't LogOn to " & CStr(arrDS(i, enumDataSource.DS_System))
                            GoTo ErrHandler
                        End If
                    End If
                End If
                
                ' after Logon - Initial Refresh (for SAPSetFilter)
                If Not bSAPSetVariablesOnly Then
                    ' if SAPSetFilter is used then we need "Initial Refresh" before ApplyScopeFilters
                    Application.StatusBar = "Refreshing " & arrDS(i, enumDataSource.DS_ID)
                    If Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID)) = 0 Then
                        Debug.Print "Couldn't Refresh " & CStr(arrDS(i, enumDataSource.DS_ID))
                        GoTo ErrHandler
                    End If
                End If ' not bSAPSetVariablesOnly
            Else
                ' not bMultiLogonMode
                ' logon was done during LogOn_DataSources
                
                If Not bSAPSetVariablesOnly Then
                    ' if SAPSetFilter is used then we need "Initial Refresh" before ApplyScopeFilters
                    Application.StatusBar = "Refreshing " & arrDS(i, enumDataSource.DS_ID)
                    If Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID)) = 0 Then
                        ' debug
                        Debug.Print "Couldn't Refresh " & CStr(arrDS(i, enumDataSource.DS_ID))
                        GoTo ErrHandler
                        ' write log - couldn't refresh Data source XXX
                    Else
                        Application.StatusBar = vbNullString
                    End If
                End If ' not bSAPSetVariablesOnly
                
            End If
            
            ' set filters / variables
            If Not bYetAnotherAttempt Then
                ' if first run - we have to apply filters / variables
                If bSAPSetVariablesOnly Then
                    ' if bSAPSetVariablesOnly - then Variables will be set by corresponding callback
                    Application.StatusBar = "Registering callback BeforeFirstPromptsDisplay..."
                    Call Callback_BeforeFirstPromptsDisplay_Reg
                    
                    ' The Default refresh - only SAPSetVariables, first refresh
                    Application.StatusBar = "Refreshing data source " & arrDS(i, enumDataSource.DS_ID) & "..."
                    Call Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID))
                    
                    Application.StatusBar = "Unregistering callback BeforeFirstPromptsDisplay..."
                    Call Callback_BeforeFirstPromptsDisplay_UnReg
                Else
                    ' if it is first Refresh - call ApplyScopeFilters
                    Application.StatusBar = "Applying filters to data source " & arrDS(i, enumDataSource.DS_ID)
                    
                    ' during ApplyScopeFilters macro will
                    ' set PauseVariableSubmit to On, and SAPSetRefreshBehaviour to Off
                    If Not ApplyScopeFilters(Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Value, CStr(arrDS(i, enumDataSource.DS_ID))) Then
                        Debug.Print "Couldn't Apply filters for " & CStr(arrDS(i, enumDataSource.DS_ID))
                        GoTo ErrHandler
                    End If
                End If ' bSAPSetVariablesOnly
            End If
            
            ' for all next attempts no need to set Variables or Filters - just refresh
            ' refresh report
            If bSAPSetVariablesOnly Then
                If bYetAnotherAttempt Then
                    Application.StatusBar = "Refreshing data source " & arrDS(i, enumDataSource.DS_ID) & "..."
                    Call Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID))
                End If
            Else
            ' if SAPSetFilter, then ApplyScopeFilters was executed, therefore we have to turn On default addin behavior
                Application.StatusBar = "Turning On RefreshBehavior for data source " & arrDS(i, enumDataSource.DS_ID)
                Call Run("SAPExecuteCommand", "PauseVariableSubmit", "Off") ' unpause submit of variables
                Call Run("SAPSetRefreshBehaviour", "On") ' trigger RefreshData
            End If
            
            If bMultiLogonMode Then
                ' if next DS has different language - logoff
                If i < UBound(arrDS) Then
                    If arrDS(i, enumDataSource.DS_Language) <> arrDS(i + 1, enumDataSource.DS_Language) Then
                        Application.StatusBar = "Logging Off..."
                        Application.EnableEvents = False ' for some reason LogOff triggers SelectionChange event
                        Call LogOffBOA
                        Application.EnableEvents = True
                    End If
                End If
            End If
        End If
        
        Application.StatusBar = "Executing Macros After " & arrDS(i, enumDataSource.DS_ID)
        If arrDS(i, enumDataSource.DS_MacrosAfter) <> vbNullString Then
            arrMacrosAfterDSRefresh = Split(arrDS(i, enumDataSource.DS_MacrosAfter), ",")
            For m = 0 To UBound(arrMacrosAfterDSRefresh)
                On Error Resume Next
                Application.StatusBar = "Executing macro " & Trim(arrMacrosAfterDSRefresh(m)) & "..."
                Application.Run "ControlPanel." & Trim(arrMacrosAfterDSRefresh(m))
                If Err.Number <> 0 Then
                    Debug.Print "Couldn't run macro " & Trim(arrMacrosAfterDSRefresh(m))
                    GoTo ErrHandler
                End If
                On Error GoTo 0
                'Application.StatusBar = "Calculating..."
                'Application.Calculate
                'Application.StatusBar = vbNullString
                
                ' user could call a sub, which change CONTROL_PANEL_DELAY cell value
                ' check it - Check if no new data
                If Me.Names("CONTROL_PANEL_DELAY").RefersToRange.Value = "Y" Then
                    ' check if have attempts
                    ' starts from 0 for each data source
                    Attempt = Attempt + 1
                    If Attempt > Me.Names("CONTROL_PANEL_ATTEMPTS").RefersToRange.Value Then
                        Debug.Print "Attempts limit exceed. " & arrDS(i, enumDataSource.DS_ID) & " has no new data."
                        GoTo ErrHandler
                    Else
                        ' if trigger - wait
                        ' log off to prevent kick from server
                        Application.StatusBar = "Logging Off..."
                        Application.EnableEvents = False ' for some reason LogOff triggers SelectionChange event
                        Call LogOffBOA
                        Application.EnableEvents = True
                        
                        'Application.StatusBar = "Delay starting from " & Format(Now(), "yyyy-MM-dd hh:mm") & " for " _
                            & CStr(me.Names("CONTROL_PANEL_DELAY_FOR").RefersToRange.Value) & " min"
                        
                        Application.StatusBar = "Delay till " & Format((Now() + TimeValue("00:" & Right("0" & _
                                CStr(Me.Names("CONTROL_PANEL_DELAY_FOR").RefersToRange.Value), 2) & ":00")), "yyyy-MM-dd hh:mm:ss")
                        
                        Debug.Print "Delay till " & Format((Now() + TimeValue("00:" & Right("0" & _
                                CStr(Me.Names("CONTROL_PANEL_DELAY_FOR").RefersToRange.Value), 2) & ":00")), "yyyy-MM-dd hh:mm:ss")
                                
                        Application.Wait (Now() + TimeValue("00:" & Right("0" & _
                                CStr(Me.Names("CONTROL_PANEL_DELAY_FOR").RefersToRange.Value), 2) & ":00"))
                        bYetAnotherAttempt = True
                        ' refresh same data source again - and make same check again
                        GoTo Start_Refresh
                        
                    End If
                End If
                
            Next m ' next macro
        End If
        
        ' calculate if user specified this parameter
        If arrDS(i, enumDataSource.DS_CalculateAfter) = "Y" Then
            Application.StatusBar = "Calculating..."
            'Application.CalculationInterruptKey = xlNoKey
            Application.Calculate ' trigger calculation of formulas. Possibly, next report has filter based on the refreshed report
            'Application.CalculationInterruptKey = xlAnyKey
            Application.StatusBar = vbNullString
        End If
    Next i ' next Data source
    
    Application.StatusBar = "Logging Off..."
    
    Application.EnableEvents = False
    Call LogOffBOA
    Application.EnableEvents = True
    
    arrMacrosAfter = Split(Me.Names("MACROS_AFTER").RefersToRange.Value, ",")
    
    Me.Names("LAST_REFRESH_DATETIME").RefersToRange.Value = Now()
    
    Application.StatusBar = "Executing macros after Refresh..."
    
    For m = 0 To UBound(arrMacrosAfter)
        On Error Resume Next
        Run Trim(arrMacrosAfter(i))
        If Err.Number <> 0 Then
            Debug.Print "Couldn't run macro " & arrMacrosAfterDSRefresh(m)
            GoTo ErrHandler
        End If
        On Error GoTo 0
    Next m
    
    ' resulting calculation
    Application.StatusBar = "Calculating..."
    Application.Calculate
    Application.StatusBar = vbNullString
    
    ' Call Write_Log("END # Overall execution time # " & _
        Round((Now() - StartPoint) * 3600 * 24 / 60, 0) & "m " & _
        FormatNumber(Round((Now() - StartPoint) * 3600 * 24, 0) Mod 60, 0) & "s")
        
    Refresh = True
            
Exit_Function:
    On Error Resume Next
    Application.StatusBar = "Disabling BOA..."
    
    Call LogOffBOA ' just in case
    
    'Call DisableBOA ' crashes Excel from time to time

    Application.EnableEvents = True
    Application.ScreenUpdating = True ' just in case
    Application.StatusBar = vbNullString
    
    Exit Function
    
ErrHandler:
    ' send email?
    If Err.Number <> 0 Then
        Debug.Print Err.Number & ": " & Err.Description
        Err.Clear
    End If
    GoTo Exit_Function
    Resume
End Function

Private Sub EnableBOA()
    Dim addIn As COMAddIn

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
    
    Set dicSystems = CreateObject("Scripting.Dictionary")
    dicSystems.comparemode = 1 ' TextCompare
    
    With Me.ListObjects("DATA_SOURCES")
        If .DataBodyRange Is Nothing Then
            Exit Function
        Else
            ' check if several languages in Scope
            For r = 1 To .DataBodyRange.Rows.Count
                If .ListColumns("Scope").DataBodyRange.Cells(r, 1).Value = _
                                Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Value Then
                    
                    If sLanguage = vbNullString And .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                    ' on first step - remember value in variable
                        sLanguage = IIf(.ListColumns("Language").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                            Me.Names("GENERAL_LANGUAGE").RefersToRange.Value, _
                                            .ListColumns("Language").DataBodyRange.Cells(r, 1).Value)
                    Else
                    ' second and further steps - compare each row with value in variable
                        If sLanguage <> IIf(.ListColumns("Language").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                            Me.Names("GENERAL_LANGUAGE").RefersToRange.Value, _
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
                                    Me.Names("GENERAL_USER").RefersToRange.Value, _
                                    .ListColumns("User").DataBodyRange.Cells(r, 1).Value)
                    Else
                    ' second and further steps - compare each row with value in variable
                        If sUser <> IIf(.ListColumns("User").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("GENERAL_USER").RefersToRange.Value, _
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
                                    Me.Names("GENERAL_CLIENT").RefersToRange.Value, _
                                    .ListColumns("Client").DataBodyRange.Cells(r, 1).Value)
                    Else
                    ' second and further steps - compare each row with value in variable
                        If sClient <> IIf(.ListColumns("Client").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("GENERAL_CLIENT").RefersToRange.Value, _
                                    .ListColumns("Client").DataBodyRange.Cells(r, 1).Value) And _
                                .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                            bMultiLogonMode = True
                            LogOn_DataSources = True
                            Exit Function
                        End If
                    End If
                    
                    ' Dictionary of Systems
                    sSystem = IIf(.ListColumns("System").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    Me.Names("GENERAL_SYSTEM").RefersToRange.Value, _
                                        .ListColumns("System").DataBodyRange.Cells(r, 1).Value)
                    'sClient = IIf(.ListColumns("Client").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    me.Names("GENERAL_CLIENT").RefersToRange.Value, _
                                        .ListColumns("Client").DataBodyRange.Cells(r, 1).Value)
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
                
                ' 'arrDS(1, enumDataSource.DS_ID),
                If Run("SapLogon", _
                    CStr(dicSystems.Item(Key)), _
                    Split(Key, "+")(1), _
                    sUser, _
                    GetPassword(CStr(Split(Key, "+")(0)), _
                        sUser), _
                    sLanguage) = 0 Then
                    
                    Debug.Print "Couldn't log on to the system " & .ListColumns("System").DataBodyRange.Cells(i, 1).Value
                    Exit Function ' leave Log_On_Systems as False
                Else
                    LogOn_DataSources = True
                    Exit Function
                End If
            Next Key ' of dicSystems - dictionary with list of systems in Scope
            
        End If ' if table 'data source' is not empty
    End With
    
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
    
    Set objFile = objFSO.OpenTextFile(Me.Names("PASSWORDS_PATH").RefersToRange.Value, ForReading)
    
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
    Debug.Print Err.Number & ": " & Err.Description
    Err.Clear
    On Error GoTo 0
    GoTo Exit_Function
    Resume
End Function

Private Function Get_Array_From_Range(rng_ref)
    Dim rng_tmp As Range
    Dim arr_tmp
    If Left(rng_ref, 2) <> "'[" Then ' if not external link
        ThisWorkbook.Windows(1).Activate ' Range fails if another workbook is activated, that doesn't contain rng_reference
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
    
    ' define arrays with Variables table columns - faster then read each cell
    With Me.ListObjects("VARIABLES")
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
                
                Me.ListObjects("VARIABLES").ListColumns("Last Refresh").DataBodyRange.Cells(variable_row, 1).Value = Now()
                Me.ListObjects("VARIABLES").ListColumns("Last Used Value").DataBodyRange.Cells(variable_row, 1).Value = "'" & sVariableValue
                
                If Run(arr_var_command(variable_row, 1), _
                        arr_var_ds(variable_row, 1), _
                        arr_var_field(variable_row, 1), _
                        sVariableValue) = 0 Then
                    
                    Debug.Print "Couldn't apply filter for variable " & CStr(arr_var_field(variable_row, 1)) & " : " & sVariableValue
                    Exit Function
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
                
                Me.ListObjects("VARIABLES").ListColumns("Last Refresh").DataBodyRange.Cells(variable_row, 1).Value = Now()
                Me.ListObjects("VARIABLES").ListColumns("Last Used Value").DataBodyRange.Cells(variable_row, 1).Value = "'" & sVariableValue
                
                If Run(arr_var_command(variable_row, 1), _
                        arr_var_field(variable_row, 1), _
                        sVariableValue, _
                        "INPUT_STRING", _
                        arr_var_ds(variable_row, 1)) = 0 Then
                    
                    Debug.Print "Couldn't apply filter for variable " & Str(arr_var_field(variable_row, 1)) & " : " & sVariableValue
                    GoTo Exit_sub
                End If
                sVariableValue = vbNullString
            End If ' var command
        End If ' var scope and data source
    Next variable_row
    
    ApplyScopeFilters = True
Exit_sub:
    Exit Function

ErrHandler:
    Err.Clear
    On Error GoTo 0
    GoTo Exit_sub
End Function

Private Sub GetArrDS()
    ' sub collects all active data sources
    Dim n As Long
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    If Me.ListObjects("DATA_SOURCES").DataBodyRange Is Nothing Then Exit Sub
    If Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Value = vbNullString Then Exit Sub
    
    If WorksheetFunction.CountIf(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange, "=") > 0 Then
        Me.ListObjects("DATA_SOURCES").ListColumns("Order").Range.Select
        Debug.Print "Please, provide order of execution."
        Exit Sub
    End If
    
    ' amount of data sources in Scope
    n = WorksheetFunction.CountIf(Me.ListObjects("DATA_SOURCES").ListColumns("Scope").DataBodyRange, _
            Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Value)
    
    ReDim arrDS(n, 7) ' size must be equal to number of parameters
    For i = 1 To Me.ListObjects("DATA_SOURCES").DataBodyRange.Rows.Count
        ' fill array according to order of execution
    
        If Me.ListObjects("DATA_SOURCES").ListColumns("Scope").DataBodyRange.Cells(i, 1).Value = _
            Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Value Then
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_ID) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Data Source").DataBodyRange.Cells(i, 1)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_Refresh) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Refresh?").DataBodyRange.Cells(i, 1)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_MacrosAfter) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Macros After").DataBodyRange.Cells(i, 1)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_CalculateAfter) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Calculate After").DataBodyRange.Cells(i, 1)
            
            ' if empty settings - use General options
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_System) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("System").DataBodyRange.Cells(i, 1).Value = vbNullString, _
                    Me.Names("GENERAL_SYSTEM").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("System").DataBodyRange.Cells(i, 1).Value)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_Language) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("Language").DataBodyRange.Cells(i, 1) = vbNullString, _
                    Me.Names("GENERAL_LANGUAGE").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("Language").DataBodyRange.Cells(i, 1).Value)
                
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_User) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("User").DataBodyRange.Cells(i, 1) = vbNullString, _
                    Me.Names("GENERAL_USER").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("User").DataBodyRange.Cells(i, 1).Value)
        
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Client").DataBodyRange.Cells(i, 1), enumDataSource.DS_Client) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("Client").DataBodyRange.Cells(i, 1) = vbNullString, _
                    Me.Names("GENERAL_CLIENT").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("Client").DataBodyRange.Cells(i, 1).Value)
        
        End If
    Next i

Exit_sub:
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number & ": " & Err.Description
    ' Stop
    Err.Clear
    On Error GoTo 0
    GoTo Exit_sub
    Resume 'debug
End Sub

Private Sub Worksheet_Activate()
    'If IsEmpty(me) Then
    '    DefineNamedRangeScope
    'End If
    If Me.Names("PASSWORDS_PATH").RefersToRange.Value = vbNullString Then
        Me.Names("PASSWORDS_PATH").RefersToRange.Value = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\Passwords.txt"
    End If
    If Me.Names("GENERAL_USER").RefersToRange.Value = vbNullString Then
        Me.Names("GENERAL_USER").RefersToRange.Value = LCase(Environ("username"))
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    'If IsEmpty(me) Then
    '    DefineNamedRangeScope
    'End If
    If Me.Names("PASSWORDS_PATH").RefersToRange.Value = vbNullString Then
        Me.Names("PASSWORDS_PATH").RefersToRange.Value = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\Passwords.txt"
    End If
    If Me.Names("GENERAL_USER").RefersToRange.Value = vbNullString Then
        Me.Names("GENERAL_USER").RefersToRange.Value = LCase(Environ("username"))
    End If
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

Sub Collect_Variables()
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
    'ThisWorkbook.Sheets("Result").ListObjects("VARIABLES").DataBodyRange.Rows.Delete
    'ThisWorkbook.Sheets("Result").ListObjects("DATA_SOURCES").DataBodyRange.Rows.Delete
    'On Error GoTo ErrHandler
    
    ' If Application.Workbooks.Count = 1 Then Exit Sub
    ' Call LogOffBOA
    'Call DefineNamedRangeScope
    
    Application.StatusBar = "Enabling BOA..."
    Call EnableBOA
    Application.StatusBar = vbNullString
    ThisWorkbook.Activate
    
    Call ListObjectDeleteEmptyRows(Me.ListObjects("DATA_SOURCES"))
    Call ListObjectDeleteEmptyRows(Me.ListObjects("VARIABLES"))
    
    Call GetListOfDS ' get list of Data Sources
    
    order = 1
    If IsArray(arrDS) Then
        ' for each DS in workbook - collect variables
        'For i = 1 To UBound(arrDS, 1)
        For i = LBound(arrDS, 1) To UBound(arrDS, 1)
            If arrDS(i, enumDataSource.DS_ID) = vbNullString Then GoTo Next_DS
            
            If Not Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, enumDataSource.DS_ID)) Then
                
                Application.StatusBar = "Log On..."
                Call Application.Run("SapLogon", _
                    CStr(arrDS(i, enumDataSource.DS_ID)), _
                    Me.Names("GENERAL_CLIENT").RefersToRange.Value, _
                    Me.Names("GENERAL_USER").RefersToRange.Value, _
                    GetPassword(Me.Names("GENERAL_SYSTEM").RefersToRange.Value, _
                        Me.Names("GENERAL_USER").RefersToRange.Value), _
                    Me.Names("GENERAL_LANGUAGE").RefersToRange.Value)
                
                Application.StatusBar = "Refreshing..."
                Run "SAPExecuteCommand", "Refresh" ' , arrDS(i, 1)
                
            End If
            
            If Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, enumDataSource.DS_ID)) Then
                Application.StatusBar = "Collecting variables and filters..."
                'Set r_ds = Me.ListObjects("DATA_SOURCES").ListRows.Add(AlwaysInsert:=True)
                Set r_ds = Me.ListObjects("DATA_SOURCES").ListRows.Add(AlwaysInsert:=False)
                
                If Me.Names("LIST_SCOPES").RefersToRange.Cells.Count = 1 Then
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Scope").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                        "'" & Me.Names("LIST_SCOPES").RefersToRange.Value
                End If
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Sheet").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = arrDS(i, 1) ' Sheet Name
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Data Source").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = arrDS(i, enumDataSource.DS_ID)
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Data Source Name").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                    Application.Run("SapGetSourceInfo", arrDS(i, enumDataSource.DS_ID), "DataSourceName")
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Query").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                    Application.Run("SapGetSourceInfo", arrDS(i, enumDataSource.DS_ID), "QueryTechName")
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("System").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                    Application.Run("SapGetSourceInfo", arrDS(i, enumDataSource.DS_ID), "System")
                
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Refresh?").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = "Y"
                r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Order").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = order
                order = order + 1
                
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
                            'Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=True)
                            Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=False)
                            
                            If Me.Names("LIST_SCOPES").RefersToRange.Cells.Count = 1 Then
                                r.Range.Cells(1, r.Parent.ListColumns("Scope").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                    "'" & Me.Names("LIST_SCOPES").RefersToRange.Value
                            End If
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, enumDataSource.DS_ID)
                            r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetVariable"
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Field").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                    Application.Run("SAPGetVariable", arrDS(i, enumDataSource.DS_ID), arrVar(1), "TECHNICALNAME")
                                
                            r.Range.Cells(1, r.Parent.ListColumns("Field Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrVar(1)
                            r.Range.Cells(1, r.Parent.ListColumns("Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = "'" & arrVar(2)
                        Else
                            ' two-dimentional
                            For Var = LBound(arrVar) To UBound(arrVar)
                                'Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=True)
                                Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=False)
                                
                                If Me.Names("LIST_SCOPES").RefersToRange.Cells.Count = 1 Then
                                    r.Range.Cells(1, r.Parent.ListColumns("Scope").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                        "'" & Me.Names("LIST_SCOPES").RefersToRange.Value
                                End If
                            
                                r.Range.Cells(1, r.Parent.ListColumns("Data Source").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, enumDataSource.DS_ID)
                                r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetVariable"
                                
                                r.Range.Cells(1, r.Parent.ListColumns("Field").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                        Application.Run("SAPGetVariable", arrDS(i, enumDataSource.DS_ID), arrVar(Var, 1), "TECHNICALNAME")
                                    
                                r.Range.Cells(1, r.Parent.ListColumns("Field Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrVar(Var, 1)
                                r.Range.Cells(1, r.Parent.ListColumns("Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = "'" & arrVar(Var, 2)
                            Next Var
                        End If
                        
                    Else ' If IsArray(arrVar) Then
                        ' if no variables in Data Source
                        ' Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=True)
                        Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=False)
                        
                        If Me.Names("LIST_SCOPES").RefersToRange.Cells.Count = 1 Then
                            r.Range.Cells(1, r.Parent.ListColumns("Scope").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                "'" & Me.Names("LIST_SCOPES").RefersToRange.Value
                        End If
                        
                        r.Range.Cells(1, r.Parent.ListColumns("Data Source").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, enumDataSource.DS_ID)
                        r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetVariable"
                        r.Range.Cells(1, r.Parent.ListColumns("Field").DataBodyRange.Column - r.Parent.Range.Column + 1) = ""
                        r.Range.Cells(1, r.Parent.ListColumns("Field Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = "Not applicable"
                        r.Range.Cells(1, r.Parent.ListColumns("Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = ""
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
                            ' Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=True)
                            Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=False)
                            
                            If Me.Names("LIST_SCOPES").RefersToRange.Cells.Count = 1 Then
                                r.Range.Cells(1, r.Parent.ListColumns("Scope").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                    "'" & Me.Names("LIST_SCOPES").RefersToRange.Value
                            End If
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, enumDataSource.DS_ID)
                            r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetFilter"
                            
                            For Dimen = LBound(arrDimensions) To UBound(arrDimensions)
                                 If arrDimensions(Dimen, 2) = arrFilters(1) Then
                                    r.Range.Cells(1, r.Parent.ListColumns("Field").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDimensions(Dimen, 1) ' technical name
                                    Exit For
                                 End If
                            Next Dimen
                                
                            r.Range.Cells(1, r.Parent.ListColumns("Field Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrFilters(1)
                            r.Range.Cells(1, r.Parent.ListColumns("Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = "'" & arrFilters(2)
                            
                        Else
                            ' 2-deminsional
                            For Fil = LBound(arrFilters) To UBound(arrFilters)
                                If arrFilters(Fil, 1) <> "Measures" Then
                                    'Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=True)
                                    Set r = Me.ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=False)
                                    
                                    If Me.Names("LIST_SCOPES").RefersToRange.Cells.Count = 1 Then
                                        r.Range.Cells(1, r.Parent.ListColumns("Scope").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                            "'" & Me.Names("LIST_SCOPES").RefersToRange.Value
                                    End If
                                    
                                    r.Range.Cells(1, r.Parent.ListColumns("Data Source").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, enumDataSource.DS_ID)
                                    r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetFilter"
                                    
                                    For Dimen = LBound(arrDimensions) To UBound(arrDimensions)
                                         If arrDimensions(Dimen, 2) = arrFilters(Fil, 1) Then
                                            r.Range.Cells(1, r.Parent.ListColumns("Field").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDimensions(Dimen, 1) ' technical name
                                            Exit For
                                         End If
                                    Next Dimen
                                        
                                    r.Range.Cells(1, r.Parent.ListColumns("Field Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrFilters(Fil, 1)
                                    r.Range.Cells(1, r.Parent.ListColumns("Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = "'" & arrFilters(Fil, 2)
                                End If ' if measures
                            Next Fil
                                            
                        End If
                    Else
                    
                    End If ' If IsArray(arrFilters)
                End If
            Else ' If Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, enumDataSource.DS_ID)) Then
                ' write that DS is not active
                
                ' MsgBox "Please, refresh data source " & arrDS(i, 1)
            End If ' isDataSourceActive
Next_DS:
        Next i ' next DS in arrDS
    Else
        
    
    End If ' If IsArray(arrDS) Then

    Me.Range("DATA_SOURCES[#All]").RemoveDuplicates Header:=xlYes, _
    Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12)
        
    Me.Range("VARIABLES[#All]").RemoveDuplicates Header:=xlYes, _
     Columns:=Array(1, 2, 3, 4, 5, 6, 7, 8, 9)
    
    Call ListObjectDeleteEmptyRows(Me.ListObjects("DATA_SOURCES"))
    Call ListObjectDeleteEmptyRows(Me.ListObjects("VARIABLES"))
    
    Application.StatusBar = "Logging Off..."
    
    Application.EnableEvents = False
    Call LogOffBOA
    Application.EnableEvents = True
    
Exit_sub:
    On Error Resume Next
    Erase arrDS
    Application.Cursor = xlDefault
    'Application.StatusBar = "Disabling BOA..."
    'Call DisableBOA
    
    Application.StatusBar = vbNullString
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    On Error GoTo 0
    Application.ScreenUpdating = True
    Debug.Print Err.Number & ": " & Err.Description
    Application.Cursor = xlDefault
    GoTo Exit_sub
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
            & "Such issue were noticed when BO Analysis addin cannot initialize itself." & vbCrLf _
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
        ReDim arrDS(0, 1)
        arrDS(0, enumDataSource.DS_ID) = tmpCrossTabs(3) ' data source ID

        On Error Resume Next
        arrDS(0, 1) = ActiveWorkbook.Names("SAP" & tmpCrossTabs(1)).RefersToRange.Parent.Name ' worksheet name
        ' not 100% precise method, can fail if user rename NamedRange
        Err.Clear
        On Error GoTo 0
        
    Else
        Err.Clear
        On Error GoTo 0
        ' arrDS = tmpCrossTabs
        ReDim arrDS(UBound(tmpCrossTabs, 1), 1)
        
        For i = 1 To UBound(tmpCrossTabs, 1)
            arrDS(i, enumDataSource.DS_ID) = tmpCrossTabs(i, 3) ' data source ID
            
            On Error Resume Next
            arrDS(i, 1) = ThisWorkbook.Names("SAP" & tmpCrossTabs(i, 1)).RefersToRange.Parent.Name ' worksheet name
            ' not 100% precise method, can fail if user rename NamedRange
            Err.Clear
            On Error GoTo 0
        Next i
        
    End If

End Sub

Private Sub ListObjectDeleteEmptyRows(lo As ListObject)
    Dim i As Long
    If lo.DataBodyRange Is Nothing Then Exit Sub
    i = 1
    Do While i <= lo.ListRows.Count
        If WorksheetFunction.CountA(lo.ListRows(i).Range) = 0 Then
            lo.ListRows(i).Delete
            i = i - 1
        End If
        i = i + 1
    Loop
End Sub

Sub FakeCheckIfNoData()
' use this sub to test how Delay Functionality works
    If dummy = 0 Then
        Me.Names("CONTROL_PANEL_DELAY").RefersToRange.Value = "Y"
        dummy = 1
    Else
        Me.Names("CONTROL_PANEL_DELAY").RefersToRange.Value = "N"
    End If
End Sub

Sub Callback_BeforeFirstPromptsDisplay(Optional dummy As Variant)
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
    scope = Me.Names("CONTROL_PANEL_SCOPE").RefersToRange.Value
    
    ' define arrays with Variables table columns - faster then read each cell
    With Me.ListObjects("VARIABLES")
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
                
                Me.ListObjects("VARIABLES").ListColumns("Last Refresh").DataBodyRange.Cells(variable_row, 1).Value = Now()
                Me.ListObjects("VARIABLES").ListColumns("Last Used Value").DataBodyRange.Cells(variable_row, 1).Value = "'" & sVariableValue
                
                If Run(arr_var_command(variable_row, 1), _
                        arr_var_field(variable_row, 1), _
                        sVariableValue, _
                        "INPUT_STRING", _
                        arr_var_ds(variable_row, 1)) = 0 Then
                    
                    Debug.Print "Couldn't apply filter for variable " & Str(arr_var_field(variable_row, 1)) & " : " & sVariableValue
                    GoTo Exit_sub
                End If
                sVariableValue = vbNullString
            End If ' var command
        End If ' var scope and data source
    Next variable_row
        
Exit_sub:
    Exit Sub

ErrHandler:
    Err.Clear
    On Error GoTo 0
    GoTo Exit_sub
End Sub

Private Sub Callback_BeforeFirstPromptsDisplay_Reg()
    Dim iRet As Integer
    Call Application.Run("SAPExecuteCommand", "RegisterCallback", "BeforeFirstPromptsDisplay", "ControlPanel.Callback_BeforeFirstPromptsDisplay")
    Debug.Print Now, "Callback_BeforeFirstPromptsDisplay_Reg", iRet
End Sub

Private Sub Callback_BeforeFirstPromptsDisplay_UnReg()
    Dim iRet As Integer
    iRet = Application.Run("SAPExecuteCommand", "UnRegisterCallback", "BeforeFirstPromptsDisplay", "ControlPanel.Callback_BeforeFirstPromptsDisplay")
    Debug.Print Now, "Callback_BeforeFirstPromptsDisplay_UnReg", iRet
End Sub

Private Function IsSAPSetVariablesOnly(ds As String) As Boolean
    IsSAPSetVariablesOnly = True ' by default
    With Me.ListObjects("VARIABLES")
        If Not .DataBodyRange Is Nothing Then
            IsSAPSetVariablesOnly = (WorksheetFunction.CountIfs(.ListColumns("Data Source").DataBodyRange, ds, _
                    .ListColumns("Command").DataBodyRange, "SAPSetFilter") = 0)
        End If
    End With
End Function


