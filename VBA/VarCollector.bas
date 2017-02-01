Attribute VB_Name = "VarCollector"
Option Explicit
Option Base 1

' Author: Ivan Bondarenko
' Release date: 2017-01
' bondarenko.ivan@me.com
' https://bondarenkoivan.wordpress.com
' https://linkedin.com/in/bondarenkoivan/en

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Public arrDS

Sub Collect_Variables()
    Dim wb As Workbook
    
    Dim arrVar
    Dim r
    Dim r_ds
    Dim i As Long
    Dim Var
    Dim sPassword As String
    Dim sSystem   As String
    
    Application.Cursor = xlWait
    Application.ScreenUpdating = False
    
    On Error Resume Next
    ThisWorkbook.Sheets("Result").ListObjects("VARIABLES").DataBodyRange.Rows.Delete
    ThisWorkbook.Sheets("Result").ListObjects("DATA_SOURCES").DataBodyRange.Rows.Delete
    
    On Error GoTo ErrHandler
    
    If Application.Workbooks.Count = 1 Then Exit Sub
    
    Call EnableBOA
    
    For Each wb In Application.Workbooks
        If wb.Name = ThisWorkbook.Name Then GoTo Next_WB
        wb.Activate
        
        Call GetListOfDS
        
        If IsArray(arrDS) Then
            ' for each DS in workbook - collect variables
            For i = LBound(arrDS, 1) To UBound(arrDS, 1)
                
                If Not Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, 1)) Then
                    
                    Run "SAPExecuteCommand", "Refresh" ' , arrDS(i, 1)
                    
                End If
                
                If Application.Run("SAPGetProperty", "IsDataSourceActive", arrDS(i, 1)) Then
                                
                    Set r_ds = ThisWorkbook.Sheets("Result").ListObjects("DATA_SOURCES").ListRows.Add(AlwaysInsert:=True)
                    
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Workbook").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = wb.Name
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Sheet").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = arrDS(i, 2)
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Data Source").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = arrDS(i, 1)
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Data Source Name").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                        Application.Run("SapGetSourceInfo", arrDS(i, 1), "DataSourceName")
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("Query").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                        Application.Run("SapGetSourceInfo", arrDS(i, 1), "QueryTechName")
                    
                    r_ds.Range.Cells(1, r_ds.Parent.ListColumns("System").DataBodyRange.Column - r_ds.Parent.Range.Column + 1) = _
                        Application.Run("SapGetSourceInfo", arrDS(i, 1), "System")
                        
                    arrVar = Application.Run("SAPListOfVariables", arrDS(i, 1), "INPUT_STRING", ThisWorkbook.Names("DISPLAY").RefersToRange.Value)
                    
                    If IsArray(arrVar) Then
                        For Var = LBound(arrVar) To UBound(arrVar)
                            Set r = ThisWorkbook.Sheets("Result").ListObjects("VARIABLES").ListRows.Add(AlwaysInsert:=True)
                            r.Range.Cells(1, r.Parent.ListColumns("Workbook").DataBodyRange.Column - r.Parent.Range.Column + 1) = wb.Name
                            r.Range.Cells(1, r.Parent.ListColumns("Sheet").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, 2)
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, 1)
                            r.Range.Cells(1, r.Parent.ListColumns("Data Source Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                Application.Run("SapGetSourceInfo", arrDS(i, 1), "DataSourceName")
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Variable Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrVar(Var, 1)
                            r.Range.Cells(1, r.Parent.ListColumns("Variable Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = "'" & arrVar(Var, 2)
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Variable ID").DataBodyRange.Column - r.Parent.Range.Column + 1) = _
                                Application.Run("SAPGetVariable", arrDS(i, 1), arrVar(Var, 1), "TECHNICALNAME")
                            
                            r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetVariable"
                        Next Var
                    Else
                        Set r = ThisWorkbook.Sheets("Result").ListObjects("Result").ListRows.Add(AlwaysInsert:=True)
                        r.Range.Cells(1, r.Parent.ListColumns("Workbook Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = wb.Name
                        r.Range.Cells(1, r.Parent.ListColumns("Sheet Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, 2)
                        r.Range.Cells(1, r.Parent.ListColumns("Data Source ID").DataBodyRange.Column - r.Parent.Range.Column + 1) = arrDS(i, 1)
                        r.Range.Cells(1, r.Parent.ListColumns("Data Source Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = Application.Run("SapGetSourceInfo", arrDS(i, 1), "DataSourceName")
                        
                        r.Range.Cells(1, r.Parent.ListColumns("Data Source Tech. Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = Application.Run("SapGetSourceInfo", arrDS(i, 1), "QueryTechName")
                        r.Range.Cells(1, r.Parent.ListColumns("System").DataBodyRange.Column - r.Parent.Range.Column + 1) = Application.Run("SapGetSourceInfo", arrDS(i, 1), "System")
                        
                        r.Range.Cells(1, r.Parent.ListColumns("Variable Name").DataBodyRange.Column - r.Parent.Range.Column + 1) = "Not applicable"
                        r.Range.Cells(1, r.Parent.ListColumns("Variable Value").DataBodyRange.Column - r.Parent.Range.Column + 1) = ""
                        
                        r.Range.Cells(1, r.Parent.ListColumns("Variable ID").DataBodyRange.Column - r.Parent.Range.Column + 1) = ""
                        r.Range.Cells(1, r.Parent.ListColumns("Command").DataBodyRange.Column - r.Parent.Range.Column + 1) = "SAPSetVariable"
                    End If
                Else
                    ' write that DS is not active
                    
                    ' MsgBox "Please, refresh data source " & arrDS(i, 1)
                End If ' isDataSourceActive
            Next i
        Else
            
        
        End If

Next_WB:
        arrDS = ""
    Next wb
    
    ThisWorkbook.Activate

Exit_sub:
    Application.Cursor = xlDefault
    Application.ScreenUpdating = True
    Exit Sub
    
ErrHandler:
    On Error GoTo 0
    Application.ScreenUpdating = True
    Debug.Print Err.Number & ": " & Err.Description
    Application.Cursor = xlDefault
    Resume Exit_sub
    Resume ' for debug
End Sub

Sub ForceEnableBOA()
    ' Excel can crash
    Dim addIn As COMAddIn
    On Error Resume Next
    
    For Each addIn In Application.COMAddIns
        If addIn.progID = "SapExcelAddIn" Then
        'Force reconnect
            addIn.Connect = False
            addIn.Connect = True ' crashes here from time to time
        End If
    Next
End Sub

Sub EnableBOA()
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

Private Sub GetListOfDS()
    Dim tmpCrossTabs
    Dim i As Long
        
    ' works for active workbook
    tmpCrossTabs = Application.Run("SAPListOf", "CROSSTABS")
    
    On Error Resume Next
    If Not IsArray(tmpCrossTabs) Then Exit Sub
    Debug.Print tmpCrossTabs(1, 1) ' check if it is 2-dim array (when only 1 DS - response is 1-dim array)
    
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo 0
        ' only 1 dimension
        ReDim arrDS(1, 2)
        arrDS(1, 1) = tmpCrossTabs(3) ' data source ID
                
        On Error Resume Next
        arrDS(1, 2) = ActiveWorkbook.Names("SAP" & tmpCrossTabs(1)).RefersToRange.Parent.Name ' worksheet name
        ' not 100% precise method, can fail if user rename NamedRange
        Err.Clear
        On Error GoTo 0
        
    Else
        Err.Clear
        On Error GoTo 0
        ' arrDS = tmpCrossTabs
        ReDim arrDS(UBound(tmpCrossTabs, 1), 2)
        
        For i = 1 To UBound(tmpCrossTabs, 1)
            arrDS(i, 1) = tmpCrossTabs(i, 3) ' data source ID
            
            On Error Resume Next
            arrDS(i, 2) = Names("SAP" & tmpCrossTabs(i, 1)).RefersToRange.Parent.Name ' worksheet name
            ' not 100% precise method, can fail if user rename NamedRange
            Err.Clear
            On Error GoTo 0
        Next i
        
    End If

End Sub

' http://www.fmsinc.com/microsoftaccess/modules/examples/avoiddoevents.asp
Public Sub WaitSeconds(intSeconds As Integer)
  Dim datTime As Date

  datTime = DateAdd("s", intSeconds, Now)

  Do
   ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
    Sleep 100
    DoEvents
  Loop Until Now >= datTime

PROC_EXIT:
  Exit Sub
End Sub
