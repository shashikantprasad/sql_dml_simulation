Attribute VB_Name = "Main"
Option Explicit

'External data SQL query
Const id_fld_hdr As String = "record_id"
Const grp_fld_hdr As String = "group_fld"
Const str_fld_hdr As String = "str_fld"
Const num_fld_hdr As String = "num_fld"
Const date_fld_hdr As String = "date_fld"
Const externalDataQuery _
    = "select " _
        & id_fld_hdr & ", " _
        & grp_fld_hdr & ", " _
        & str_fld_hdr & ", " _
        & num_fld_hdr & ", " _
        & date_fld_hdr & " from c##xeread.table_sim_v"
Const dbConnectionString = "OLEDB;Provider=OraOLEDB.Oracle.1;Password=""password"";Persist Security Info=True;User ID=c##xeread;Data Source=localhost:1521/sample_dataset;Extended Properties="""""

Const syncedDataConnectionName As String = "synced_data_link"
Const syncedDataSheetName As String = "synced_data"
Const syncedDataTableName As String = "synced_data_tbl"

Const workDataConnectionName As String = "work_data_link"
Const workDataSheetName As String = "work_data"
Const workDataTableName As String = "work_data_tbl"

Const sqlOutputTableStyle As String = "Records"
Const sqlOutputCellStyle As String = "Data Cells"

Const defaultAnimationFPS As Integer = 12

Enum TableDiffType
    DiffNone = 0
    DiffInsert = 1
    DiffUpdate = 2
    DiffDelete = 3
End Enum

Function Delay(ms)
    Delay = Timer + ms / 1000
    While Timer < Delay: DoEvents: Wend
End Function

Sub ColorTransition(interiorObj As Interior, color, iVal, fVal, durationMillis, Optional refreshRate = defaultAnimationFPS)
    Dim steps As Integer, val, delta
    steps = 1 + (refreshRate * durationMillis / 1000)
    delta = (fVal - iVal) / steps
    interiorObj.ThemeColor = color
    For val = iVal To fVal Step delta
        interiorObj.TintAndShade = val
        Delay durationMillis / steps
    Next
End Sub

Sub formatOutput(listObj As ListObject)
    With listObj
        .ShowAutoFilterDropDown = False
        
        .Range.columnwidth = 15
        .ListColumns(str_fld_hdr).Range.columnwidth = 20
        .Range.RowHeight = 18
        
        .TableStyle = sqlOutputTableStyle
        .Range.Style = sqlOutputCellStyle
        
        .ListColumns(id_fld_hdr).DataBodyRange.NumberFormat = "General"
        .ListColumns(grp_fld_hdr).DataBodyRange.NumberFormat = "General"
        .ListColumns(str_fld_hdr).DataBodyRange.NumberFormat = "General"
        .ListColumns(num_fld_hdr).DataBodyRange.NumberFormat = "General"
        .ListColumns(date_fld_hdr).DataBodyRange.NumberFormat = "d/m/yyyy"
    End With
End Sub

Sub sortTable(listObj As ListObject, Optional sortField As String = id_fld_hdr)
    With listObj.Sort
        .SortFields.Clear
        .SortFields.Add Key:=.Parent.Parent.Range(.Parent.DisplayName & "[[#All],[" & sortField & "]]"), _
                        SortOn:=xlSortOnValues, _
                        Order:=xlAscending, _
                        DataOption:=xlSortNormal
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
End Sub

'Debug.Print "[" & err.Source & "] Error " & err.Number & ": " & err.Description

Sub setup()
Attribute setup.VB_ProcData.VB_Invoke_Func = "i\n14"
    Debug.Print "Setup started."
    Application.DisplayAlerts = False
    'On Error Resume Next
    
    With ThisWorkbook
        Debug.Print "Recreating sheets.."
        .Sheets.Add(Before:=.Sheets(1)).Name = "temp" 'To avoid attempts on deleting the only sheet
        .Sheets(syncedDataSheetName).Delete
        .Sheets(workDataSheetName).Delete
        '.Connections(syncedDataConnectionName).Delete
        .Sheets("temp").Name = syncedDataSheetName
        .Sheets.Add(After:=.Sheets(syncedDataSheetName)).Name = workDataSheetName
        
        Debug.Print "Fetching external data.."
        With .Sheets(syncedDataSheetName).ListObjects.Add( _
            SourceType:=xlSrcExternal, _
            Source:=Array(dbConnectionString), _
            Destination:=.Sheets(syncedDataSheetName).Range("$A$1")).QueryTable

            .WorkbookConnection.Name = syncedDataConnectionName
            .CommandType = xlCmdSql
            .CommandText = Array(externalDataQuery)
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = True
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh BackgroundQuery:=False 'Initial Data Fetch
            
            Dim syncedTable As ListObject
            Set syncedTable = .ListObject
            
            syncedTable.DisplayName = syncedDataTableName
            syncedTable.ShowTotals = True
            syncedTable.ListColumns(date_fld_hdr).TotalsCalculation = xlTotalsCalculationCount
            
            Call formatOutput(listObj:=syncedTable)
            Call sortTable(listObj:=syncedTable, sortField:=id_fld_hdr)
            
            Set syncedTable = Nothing
        End With
        
        .Connections.Add2 _
            Name:=workDataConnectionName, _
            Description:="Temporary connection to synced data for initial work data load", _
            connectionString:="WORKSHEET;" & .FullName, CommandText:=.Name & "!" & syncedDataTableName, _
            lCmdType:=7, CreateModelConnection:=True, ImportRelationships:=False
        
        Debug.Print "Building work data.."
        With .Sheets(workDataSheetName).ListObjects.Add( _
            SourceType:=xlSrcModel, Source:=.Connections(workDataConnectionName), _
            Destination:=.Sheets(workDataSheetName).Range("$B$2")).TableObject

            .RowNumbers = False
            .PreserveFormatting = True
            .RefreshStyle = 1
            .AdjustColumnWidth = True
            .Refresh 'Intial Data Snapshot
            
            Dim workTable As ListObject
            Set workTable = .ListObject

            workTable.Unlink 'To disable data refresh
            workTable.DisplayName = workDataTableName
            workTable.ShowTotals = True
            workTable.ListColumns(date_fld_hdr).TotalsCalculation = xlTotalsCalculationCount

            Call formatOutput(listObj:=workTable)
            Call sortTable(listObj:=workTable, sortField:=id_fld_hdr)
            
            Set workTable = Nothing
        End With
        
        .Connections(workDataConnectionName).Delete 'Delete the temporary connection
        
        Debug.Print "Saving Workbook.."
        .Save
    End With

    ThisWorkbook.Sheets(workDataSheetName).Activate
    Dim wndw As Window
    For Each wndw In ThisWorkbook.Windows
        wndw.DisplayGridlines = False
    Next
    Application.DisplayAlerts = True
    Debug.Print "Setup complete."
End Sub

Sub repeatableTask()
Attribute repeatableTask.VB_ProcData.VB_Invoke_Func = "l\n14"
    Debug.Print "Data Refresh Started."

    Dim syncedTable As ListObject, workTable As ListObject
    Set syncedTable = ThisWorkbook.Sheets(syncedDataSheetName).ListObjects(syncedDataTableName)
    Set workTable = ThisWorkbook.Sheets(workDataSheetName).ListObjects(workDataTableName)
    
    Debug.Print "Syncing External Data.."
    syncedTable.QueryTable.Refresh BackgroundQuery:=False
    
    Debug.Print "Delta (Synced data vs Work data):"
    Dim sIdx As Integer, wIdx As Integer
    wIdx = 1
    sIdx = 1
    Do While wIdx <= workTable.ListRows.Count Or sIdx <= syncedTable.ListRows.Count
        Dim delta As TableDiffType
        delta = DiffNone
        If sIdx > syncedTable.ListRows.Count Then
            delta = DiffDelete
        ElseIf wIdx > workTable.ListRows.Count Then
            delta = DiffInsert
        ElseIf syncedTable.ListColumns(id_fld_hdr).DataBodyRange(sIdx, 1).Value _
            < workTable.ListColumns(id_fld_hdr).DataBodyRange(wIdx, 1).Value Then
            delta = DiffInsert
        ElseIf syncedTable.ListColumns(id_fld_hdr).DataBodyRange(sIdx, 1).Value _
            = workTable.ListColumns(id_fld_hdr).DataBodyRange(wIdx, 1).Value Then
            delta = DiffUpdate
        ElseIf syncedTable.ListColumns(id_fld_hdr).DataBodyRange(sIdx, 1).Value _
            > workTable.ListColumns(id_fld_hdr).DataBodyRange(wIdx, 1).Value Then
            delta = DiffDelete
        End If
        
        Select Case delta
            Case DiffInsert
                Debug.Print Chr(9) & "INSERT: " & syncedTable.ListColumns(id_fld_hdr).DataBodyRange(sIdx, 1).Value
                
                workTable.ListRows.Add (wIdx)
                
                Delay 500
                Call ColorTransition( _
                    interiorObj:=workTable.ListRows(wIdx).Range.Interior, _
                    color:=xlThemeColorAccent6, _
                    iVal:=1, fVal:=0.4, _
                    durationMillis:=250 _
                )
                
                Dim i As Integer
                For i = 1 To workTable.ListColumns.Count
                    workTable.DataBodyRange(wIdx, i).Value = syncedTable.DataBodyRange(sIdx, i).Value
                Next
                
                Call ColorTransition( _
                    interiorObj:=workTable.ListRows(wIdx).Range.Interior, _
                    color:=xlThemeColorAccent6, _
                    iVal:=0.4, fVal:=1, _
                    durationMillis:=250 _
                )
                
                workTable.ListRows(wIdx).Range.Interior.ThemeColor = xlColorIndexNone
                workTable.ListRows(wIdx).Range.Interior.TintAndShade = 0
                
                wIdx = wIdx + 1
                sIdx = sIdx + 1

            Case DiffUpdate
                Dim hdr, fmt
                For Each hdr In Array(grp_fld_hdr, str_fld_hdr, num_fld_hdr, date_fld_hdr)
                    fmt = workTable.ListColumns(hdr).DataBodyRange.NumberFormat
                    
                    If Format(workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Value, fmt) _
                        <> Format(syncedTable.ListColumns(hdr).DataBodyRange(sIdx, 1).Value, fmt) Then
                        Debug.Print Chr(9) _
                            & "UPDATE: " & workTable.ListColumns(id_fld_hdr).DataBodyRange(wIdx, 1).Value _
                            & ", Field = " & hdr _
                            & ", Old = " & Format(workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Value, fmt) _
                            & ", New = " & Format(syncedTable.ListColumns(hdr).DataBodyRange(sIdx, 1).Value, fmt)
                        
                        Call ColorTransition( _
                            interiorObj:=workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Interior, _
                            color:=xlThemeColorAccent4, _
                            iVal:=1, fVal:=0.4, _
                            durationMillis:=250 _
                        )

                        workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Value _
                            = syncedTable.ListColumns(hdr).DataBodyRange(sIdx, 1).Value

                        Call ColorTransition( _
                            interiorObj:=workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Interior, _
                            color:=xlThemeColorAccent4, _
                            iVal:=0.4, fVal:=1, _
                            durationMillis:=250 _
                        )
                        
                        workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Interior.ThemeColor = xlColorIndexNone
                        workTable.ListColumns(hdr).DataBodyRange(wIdx, 1).Interior.TintAndShade = 0
                    End If
                Next
                
                wIdx = wIdx + 1
                sIdx = sIdx + 1

            Case DiffDelete
                Debug.Print Chr(9) & "DELETE: " & workTable.ListColumns(id_fld_hdr).DataBodyRange(wIdx, 1).Value
            
                Call ColorTransition( _
                    interiorObj:=workTable.ListRows(wIdx).Range.Interior, _
                    color:=xlThemeColorAccent3, _
                    iVal:=1, fVal:=0, _
                    durationMillis:=500 _
                )
    
                workTable.ListRows(wIdx).Delete
        End Select

        Delay Int(500 * Rnd)
    Loop
    
    Set syncedTable = Nothing
    Set workTable = Nothing

    Debug.Print "Data Refresh Complete."
End Sub

Sub MainProgram()
Attribute MainProgram.VB_ProcData.VB_Invoke_Func = "m\n14"
    Debug.Print "MainProgram()"
    Call setup
    Dim EndTime, refreshLoopForSeconds
    refreshLoopForSeconds = 60
    Debug.Print refreshLoopForSeconds & " seconds run begins"
    EndTime = Timer + refreshLoopForSeconds
    Do While Timer < EndTime
        Call repeatableTask
    Loop
    Debug.Print "MainProgram() End"
End Sub
