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
End Enum
'

Private Sub DefineGlobalVariables()
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Left(ThisWorkbook.Path, 4) <> "http" Then ' disable log for workbooks on SharePoint
        Logs_Enabled = True ' TODO - add cell with parameter
        LogFile = ThisWorkbook.Path & "\" & ThisWorkbook.Name & ".log"
    End If
    
    Call GetArrDS  ' populate array with data sources
End Sub

Function Refresh() As Boolean
    Dim StartPoint As Double
    Dim cell_scope As Range
    Dim i As Long
    Dim m As Integer
    Dim arrMacrosBefore
    Dim arrMacrosAfter
    Dim arrMacrosAfterDSRefresh
    
    On Error GoTo ErrHandler
    
    StartPoint = Now()
    ThisWorkbook.Windows(1).Activate ' otherwise Logon fails, if another wb is active
    
    Application.StatusBar = "Defining global variables..."
    Call DefineGlobalVariables
    
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
    arrMacrosBefore = Split(ThisWorkbook.Names("MACROS_BEFORE").RefersToRange.Value, ",")
    For m = 0 To UBound(arrMacrosBefore)
        On Error Resume Next
        Application.StatusBar = "Executing macro " & Trim(arrMacrosBefore(i)) & "..."
        Run Trim(arrMacrosBefore(i))
        If Err.Number <> 0 Then
            Debug.Print "Couldn't run macro " & Trim(arrMacrosBefore(i))
            GoTo ErrHandler
        End If
        On Error GoTo 0
        Application.StatusBar = "Calculating..."
        Application.Calculate
    Next m
    
    Application.StatusBar = "Logging On..."
    If Not LogOn_DataSources Then
        Err.Raise 21, "Refresh", "Couldn't LogOn"
        GoTo ErrHandler
    End If
    
    ' initial calculate
    Application.StatusBar = "Calculating..."
    Application.Calculate
    
    ' refresh all enabled data sources in specified order
    For i = 1 To UBound(arrDS, 1)
        ' array elements are ordered
        ' macro refreshes only those that enabled
        If arrDS(i, 1) = "Y" Then
            
            ' if within one Scope we should refresh several data sources with different languages
            ' e.g. everything is available in local language except one data source, which should be refreshed in EN
            If bMultiLogonMode Then
                If i = 1 Then
                ' logon with specified Language
                    Application.StatusBar = "Logging On " & arrDS(i, enumDataSource.DS_ID)
                    If Run("SAPLogon", arrDS(i, enumDataSource.DS_ID), "100", arrDS(i, enumDataSource.DS_User), _
                            GetPassword(CStr(arrDS(i, enumDataSource.DS_System)), CStr(arrDS(i, enumDataSource.DS_User))), _
                            arrDS(i, enumDataSource.DS_Language)) = 0 Then
                        
                        Err.Raise 22, "Refresh", "Couldn't LogOn to " & CStr(arrDS(i, enumDataSource.DS_System))
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
                            Err.Raise 23, "Refresh", "Couldn't LogOn to " & CStr(arrDS(i, enumDataSource.DS_System))
                            GoTo ErrHandler
                        End If
                    End If
                End If
                
                Application.StatusBar = "Refreshing " & arrDS(i, enumDataSource.DS_ID)
                If Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID)) = 0 Then
                    Err.Raise 24, "Refresh", "Couldn't Refresh " & CStr(arrDS(i, enumDataSource.DS_ID))
                    GoTo ErrHandler
                End If
            Else
                ' logon was done during Log_On_Systems
                Application.StatusBar = "Refreshing " & arrDS(i, enumDataSource.DS_ID)
                If Run("SAPExecuteCommand", "Refresh", arrDS(i, enumDataSource.DS_ID)) = 0 Then
                    ' debug
                    Err.Raise 25, "Refresh", "Couldn't Refresh " & CStr(arrDS(i, enumDataSource.DS_ID))
                    GoTo ErrHandler
                    ' write log - couldn't refresh Data source XXX
                End If
            End If
            
            Application.StatusBar = "Applying filters to data source " & arrDS(i, enumDataSource.DS_ID)
            If Not ApplyScopeFilters(ThisWorkbook.Names("SCOPE").RefersToRange.Value, CStr(arrDS(i, enumDataSource.DS_ID))) Then
                Err.Raise 28, "Refresh", "Couldn't Apply filters for " & CStr(arrDS(i, enumDataSource.DS_ID))
                GoTo ErrHandler
            End If

            ' refresh report
            Application.StatusBar = "Refreshing data source " & arrDS(i, enumDataSource.DS_ID)
            Call Run("SAPExecuteCommand", "PauseVariableSubmit", "Off") ' unpause submit of variables
            Call Run("SAPSetRefreshBehaviour", "On") ' trigger RefreshData
            
            Application.StatusBar = "Calculating..."
            Application.Calculate ' trigger calculation of formulas. Possibly, next report has filter based on refreshed report
            
            If bMultiLogonMode Then
                ' if next DS has different language - logoff
                If i < UBound(arrDS) Then
                    If arrDS(i, enumDataSource.DS_Language) <> arrDS(i + 1, enumDataSource.DS_Language) Then
                        Application.StatusBar = "Logging Off..."
                        Application.EnableEvents = False ' for some reason LogOff triggers SelectionChange event
                        Call Run("SAPLogOff", False) ' disconnect
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
                Run Trim(arrMacrosAfterDSRefresh(m))
                If Err.Number <> 0 Then
                    Debug.Print "Couldn't run macro " & Trim(arrMacrosAfterDSRefresh(m))
                    Err.Raise 26, "Refresh", "Couldn't run macro " & Trim(arrMacrosAfterDSRefresh(m))
                    GoTo ErrHandler
                End If
                On Error GoTo 0
                Application.StatusBar = "Calculating..."
                Application.Calculate
            Next m ' next macro
        End If
        
        Application.StatusBar = "Calculating..."
        Application.Calculate
    Next i ' next Data source
    
    Application.StatusBar = "Logging Off..."
    
    Application.EnableEvents = False
    Call Run("SAPLogOff", False) ' disconnect
    Application.EnableEvents = True
    
    arrMacrosAfter = Split(ThisWorkbook.Names("MACROS_AFTER").RefersToRange.Value, ",")
    
    ThisWorkbook.Names("LAST_REFRESH_DATETIME").RefersToRange.Value = Now()
    
    Application.StatusBar = "Executing macros after Refresh..."
    
    For m = 0 To UBound(arrMacrosAfter)
        On Error Resume Next
        Run Trim(arrMacrosAfter(i))
        If Err.Number <> 0 Then
            Debug.Print "Couldn't run macro " & arrMacrosAfterDSRefresh(m)
            Err.Raise 27, "Refresh", "Couldn't run macro " & arrMacrosAfterDSRefresh(m)
            GoTo ErrHandler
        End If
        On Error GoTo 0
    Next m
    
    Application.StatusBar = "Calculating..."
    Application.Calculate
    
    ' Call Write_Log("END # Overall execution time # " & _
        Round((Now() - StartPoint) * 3600 * 24 / 60, 0) & "m " & _
        FormatNumber(Round((Now() - StartPoint) * 3600 * 24, 0) Mod 60, 0) & "s")
        
    Refresh = True
    
    
    
Exit_Function:
    Application.StatusBar = "Disabling BOA..."
    Call DisableBOA

    Application.EnableEvents = True
    Application.ScreenUpdating = True ' just in case
    Application.StatusBar = vbNullString
    
    Exit Function
    
ErrHandler:
    ' send email?
    Debug.Print Err.Number & ": " & Err.Description
    On Error GoTo 0
    Resume Exit_Function
    Resume
End Function

Private Sub EnableBOA()
    Dim addin As COMAddIn

    For Each addin In Application.COMAddIns
        If addin.progID = "SapExcelAddIn" Then
            If addin.Connect = False Then
                addin.Connect = True
                Exit Sub
            End If
        End If
    Next
    
End Sub

Private Sub DisableBOA()
    Dim addin As COMAddIn

    For Each addin In Application.COMAddIns
        If addin.progID = "SapExcelAddIn" Then
            addin.Connect = False
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
                                ThisWorkbook.Names("SCOPE").RefersToRange.Value Then
                    
                    If sLanguage = vbNullString And .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                    ' on first step - remember value in variable
                        sLanguage = IIf(.ListColumns("Language").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                            ThisWorkbook.Names("GENERAL_LANGUAGE").RefersToRange.Value, _
                                            .ListColumns("Language").DataBodyRange.Cells(r, 1).Value)
                    Else
                    ' second and further steps - compare each row with value in variable
                        If sLanguage <> IIf(.ListColumns("Language").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                            ThisWorkbook.Names("GENERAL_LANGUAGE").RefersToRange.Value, _
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
                                    ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value, _
                                    .ListColumns("User").DataBodyRange.Cells(r, 1).Value)
                    Else
                    ' second and further steps - compare each row with value in variable
                        If sUser <> IIf(.ListColumns("User").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value, _
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
                                    ThisWorkbook.Names("GENERAL_CLIENT").RefersToRange.Value, _
                                    .ListColumns("Client").DataBodyRange.Cells(r, 1).Value)
                    Else
                    ' second and further steps - compare each row with value in variable
                        If sClient <> IIf(.ListColumns("Client").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    ThisWorkbook.Names("GENERAL_CLIENT").RefersToRange.Value, _
                                    .ListColumns("Client").DataBodyRange.Cells(r, 1).Value) And _
                                .ListColumns("Refresh?").DataBodyRange.Cells(r, 1).Value = "Y" Then
                            bMultiLogonMode = True
                            LogOn_DataSources = True
                            Exit Function
                        End If
                    End If
                    
                    ' Dictionary of Systems
                    sSystem = IIf(.ListColumns("System").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    ThisWorkbook.Names("GENERAL_SYSTEM").RefersToRange.Value, _
                                        .ListColumns("System").DataBodyRange.Cells(r, 1).Value)
                    'sClient = IIf(.ListColumns("Client").DataBodyRange.Cells(r, 1).Value = vbNullString, _
                                    ThisWorkbook.Names("GENERAL_CLIENT").RefersToRange.Value, _
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
    
    Set objFile = objFSO.OpenTextFile(ThisWorkbook.Names("PASSWORDS_PATH").RefersToRange.Value, ForReading)
    
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
    Resume Exit_Function
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
    Resume Exit_sub
End Function

Private Sub GetArrDS()
    ' sub collects all active data sources
    Dim n As Long
    Dim i As Long
    
    On Error GoTo ErrHandler
    
    If Me.ListObjects("DATA_SOURCES").DataBodyRange Is Nothing Then Exit Sub
    If ThisWorkbook.Names("SCOPE").RefersToRange.Value = vbNullString Then Exit Sub
    
    If WorksheetFunction.CountIf(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange, "=") > 0 Then
        Me.ListObjects("DATA_SOURCES").ListColumns("Order").Range.Select
        Debug.Print "Please, provide order of execution."
        Exit Sub
    End If
    
    ' amount of data sources in Scope
    n = WorksheetFunction.CountIf(Me.ListObjects("DATA_SOURCES").ListColumns("Scope").DataBodyRange, _
            ThisWorkbook.Names("SCOPE").RefersToRange.Value)
    
    ReDim arrDS(n, 6)
    For i = 1 To Me.ListObjects("DATA_SOURCES").DataBodyRange.Rows.Count
        ' fill array according to order of execution
    
        If Me.ListObjects("DATA_SOURCES").ListColumns("Scope").DataBodyRange.Cells(i, 1).Value = _
            ThisWorkbook.Names("SCOPE").RefersToRange.Value Then
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_ID) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Data Source").DataBodyRange.Cells(i, 1)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_Refresh) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Refresh?").DataBodyRange.Cells(i, 1)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_MacrosAfter) = _
                Me.ListObjects("DATA_SOURCES").ListColumns("Macros After").DataBodyRange.Cells(i, 1)
                
            ' if empty - use General options
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_System) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("System").DataBodyRange.Cells(i, 1).Value = vbNullString, _
                    ThisWorkbook.Names("GENERAL_SYSTEM").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("System").DataBodyRange.Cells(i, 1).Value)
            
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_Language) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("Language").DataBodyRange.Cells(i, 1) = vbNullString, _
                    ThisWorkbook.Names("GENERAL_LANGUAGE").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("Language").DataBodyRange.Cells(i, 1).Value)
                
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Order").DataBodyRange.Cells(i, 1), enumDataSource.DS_User) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("User").DataBodyRange.Cells(i, 1) = vbNullString, _
                    ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("User").DataBodyRange.Cells(i, 1).Value)
        
            arrDS(Me.ListObjects("DATA_SOURCES").ListColumns("Client").DataBodyRange.Cells(i, 1), enumDataSource.DS_Client) = _
                IIf(Me.ListObjects("DATA_SOURCES").ListColumns("Client").DataBodyRange.Cells(i, 1) = vbNullString, _
                    ThisWorkbook.Names("GENERAL_CLIENT").RefersToRange.Value, _
                        Me.ListObjects("DATA_SOURCES").ListColumns("Client").DataBodyRange.Cells(i, 1).Value)
        
        End If
    Next i

Exit_sub:
    Exit Sub
    
ErrHandler:
    Debug.Print Err.Number & ": " & Err.Description
    ' Stop
    Resume Exit_sub
    Resume 'debug
End Sub

Private Sub Worksheet_Activate()
    If ThisWorkbook.Names("PASSWORDS_PATH").RefersToRange.Value = vbNullString Then
        ThisWorkbook.Names("PASSWORDS_PATH").RefersToRange.Value = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\Passwords.txt"
    End If
    If ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value = vbNullString Then
        ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value = LCase(Environ("username"))
    End If
End Sub

Private Sub Worksheet_SelectionChange(ByVal Target As Range)
    If ThisWorkbook.Names("PASSWORDS_PATH").RefersToRange.Value = vbNullString Then
        ThisWorkbook.Names("PASSWORDS_PATH").RefersToRange.Value = CreateObject("WScript.Shell").SpecialFolders("mydocuments") & "\Passwords.txt"
    End If
    If ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value = vbNullString Then
        ThisWorkbook.Names("GENERAL_USER").RefersToRange.Value = LCase(Environ("username"))
    End If
End Sub

Sub LogOffBOA()
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
