# Code:

## Class Module:

### ExtWorkBook:
```VBA
Public WB_Workbook As Workbook

Private pControlWBName As String
Private pControlSheet As String


Public Property Let ControlWBName(Value As String)
    pControlWBName = Value
End Property

Public Property Get ControlWBName() As String
    ControlWBName = pControlWBName
End Property

Public Property Let ControlSheet(Value As String)
    pControlSheet = Value
End Property

Public Property Get ControlSheet() As String
    ControlSheet = pControlSheet
End Property


Private Sub Class_Initialize()
    
    pControlSheet = "_Control_"
    pControlWBName = "WB_Name"
        
End Sub

Sub OpenQueryFile()

    Dim FSO As Scripting.FileSystemObject
    Dim TS As Scripting.TextStream
    Dim QueryString As String
    Dim FD As FileDialog
    Dim QueryFilePath As String
    Dim FD_Result As Integer
    
    
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    FD.InitialFileName = "C:\Users\Prabin.Shrestha\Documents\SQL Server Management Studio\Test_Query_1"
    FD.Title = "Pick a Query to run"
    FD.ButtonName = "Use Selected Query"
    FD.AllowMultiSelect = False
    FD.Filters.Clear
    FD.Filters.Add "SQL Querries", "*.sql"
    
    FD_Result = FD.Show
    
    If FD_Result = 0 Then Exit Sub
    
    QueryFilePath = FD.SelectedItems(1)
    
    Set FSO = New Scripting.FileSystemObject
    Set TS = FSO.OpenTextFile(QueryFilePath)
    
    QueryString = TS.ReadAll
    
    TS.Close
    
    
    TestSQLQuery QueryString, "B2"
    
    
End Sub
```

---

### SQLToExcel:
```VBA
Option Explicit


Private pSN As String
Private pRangeName As String
Private pSheetName As String
Private pSQLSource As String
Private pUpdateFlag As Boolean
Private pHeaderFlag As Boolean
Private pDatabaseName As String
Private pSQLServerName As String
Private pWBNameDir As String

Private pControlWorkbookName As String
Private pControlSheet As String
Private pControlAncor As String
Private pControlDBName As String
Private pControlSSName As String
Private pControlWBName As String

Private pCon As ADODB.Connection
Private pRS As ADODB.Recordset
Private pCellRange As Range
Private pWorkSheetControl As Worksheet

Public ProcessUpdatedSucessFlag As Boolean
Public ProcessClearSucessFlag As Boolean
Public ValidateWorkbookFlag As Boolean

Private pControlWorkbook As Workbook
Private pWB_Workbook As Workbook

Private pRangeAddress As String
Private pWBFileName As String

Public Property Let ControlWBName(Value As String)
    pControlWBName = Value
End Property

Public Property Get ControlWBName() As String
    ControlWBName = pControlWBName
End Property

Public Property Let SheetName(Value As String)
    pSheetName = Value
End Property

Public Property Get SheetName() As String
    SheetName = pSheetName
End Property

Public Property Let RangeName(Value As String)
    pRangeName = Value
End Property

Public Property Get RangeName() As String
    RangeName = pRangeName
End Property

Public Property Let DatabaseName(Value As String)
    pDatabaseName = Value
End Property

Public Property Get DatabaseName() As String
    DatabaseName = pDatabaseName
End Property

Public Property Let WBNameDir(Value As String)
    pWBNameDir = Value
End Property

Public Property Get WBNameDir() As String
    SQLServerName = pWBNameDir
End Property

Public Property Let SQLServerName(Value As String)
    pSQLServerName = Value
End Property

Public Property Get SQLServerName() As String
    SQLServerName = pSQLServerName
End Property

Public Property Let SQLSource(Value As String)
    pSQLSource = Value
End Property

Public Property Get SQLSource() As String
    SQLSource = pSQLSource
End Property

Public Property Let UpdateFlag(Value As String)
    If LCase(Value) = "yes" Then pUpdateFlag = True Else pUpdateFlag = False
    End If
End Property

Public Property Get UpdateFlag() As String
    If pUpdateFlag = True Then UpdateFlag = "YES" Else UpdateFlag = "NO"
    End If
End Property

Public Property Let HeaderFlag(Value As String)
    If LCase(Value) = "yes" Then pHeaderFlag = True Else pHeaderFlag = False
    End If
End Property

Public Property Get HeaderFlag() As String
    If pHeaderFlag = True Then HeaderFlag = "YES" Else HeaderFlag = "NO"
    End If
End Property


Private Sub subUpdateFlag(Value As String)
    If UCase(Value) = "YES" Then pUpdateFlag = True Else pUpdateFlag = False
End Sub

Private Sub subHeaderFlag(Value As String)
    If UCase(Value) = "YES" Then pHeaderFlag = True Else pHeaderFlag = False
End Sub


Public Property Let ControlSheet(Value As String)
    pControlSheet = Value
End Property


Public Property Get ControlSheet() As String
    ControlSheet = pControlSheet
End Property

Public Property Get DestWBName() As String
    DestWBName = pWBFileName
End Property

Public Property Get RangeAddress() As String
    RangeAddress = pRangeAddress
End Property

Public Property Get SNumber() As String
    SNumber = pSN
End Property


Private Function RunPass() As Boolean
    If pUpdateFlag = True And (pRangeName <> "" And pSheetName <> "" And pSQLSource <> "") Then
        RunPass = True
    Else
        RunPass = False
    End If
End Function

'Public Property Let ControlSheet(Value As String)
'    pControlAncor = Value
'End Property
'Public Property Get ControlAncor() As String
'    ControlAncor = pControlAncor
'End Property


Private Sub Class_Initialize()
    
    Set pControlWorkbook = ActiveWorkbook
    pControlWorkbookName = Application.ActiveWorkbook.FullName
    'pControlWorkbook = Application.Workbooks.Open(pControlWorkbookName)
    pControlSheet = "_Control_"
    pControlAncor = "Ancor"
    pControlDBName = "DBName"
    pControlSSName = "SQLServerName"
    pControlWBName = "WB_Name"
    
    ProcessUpdatedSucessFlag = False
    ProcessClearSucessFlag = False
    
End Sub

Public Sub ShowParameters()

    Debug.Print "Sheet Name", "Range Name", "Database", "SQL Source", "Update Flag", "Header Flag", "Control Sheet"
    Debug.Print pSheetName, pRangeName, pDatabaseName, pSQLSource, pUpdateFlag, pHeaderFlag, pControlSheet
    MsgBox pSheetName & pRangeName & pDatabaseName & pSQLSource & pUpdateFlag & pHeaderFlag & pControlSheet
    

End Sub

Private Sub ConnectSQL()

    On Error GoTo ErrorHandler
    Set pCon = New ADODB.Connection
    
    pCon.ConnectionString = _
        "Provider=SQLNCLI11;" & _
        "Server=" & pSQLServerName & ";" & _
        "Database=" & pDatabaseName & ";" & _
        "Trusted_Connection=yes;"
        
    pCon.Open
    
Exit Sub
ErrorHandler:
    MsgBox "Error: SQL Server Name/Database Name Invalid"
    MsgBox "Error: Unable to connect to SQL Server. Process Failed!"
    End
    Exit Sub
    
End Sub

Private Sub DisconnectSQL()
    pCon.Close
    Set pCon = Nothing
End Sub

     
Private Sub GetData()
    
    On Error GoTo ErrorHandler
    Set pRS = pCon.Execute(pSQLSource)

Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): SQL Source Invalid"
    Err.Raise (13)
    Exit Sub
    
End Sub
  

Private Sub UpdateCellRange()
    On Error GoTo ErrorHandler
    
    Set pCellRange = pWB_Workbook.Worksheets(pSheetName).Range(pRangeName)
    pRangeAddress = pCellRange.Address
    
Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Range Name/Sheet Name Invalid"
    Err.Raise (13)
    Exit Sub
    
End Sub


Private Sub ClearCellRange()
    On Error GoTo ErrorHandler
    pCellRange.ClearContents
    
Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Range Name/Sheet Name Invalid"
    Err.Raise (13)
    Exit Sub
End Sub

Private Sub CellRangeSelection()
    On Error GoTo ErrorHandler
    Application.GoTo pCellRange

Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Range Name/Sheet Name Invalid"
    Err.Raise (13)
    Exit Sub
    
End Sub

Private Sub Paste_Data()
    On Error GoTo ErrorHandler
    
    Dim I As Integer
    Dim F As ADODB.Field

    I = 0
 
    If pHeaderFlag = True Then
        For Each F In pRS.Fields
            pWB_Workbook.Worksheets(pSheetName).Range(pCellRange.Item(1, 1).Address).Offset(, I).Value = F.Name
            I = I + 1
        Next F
        
        pCellRange.Offset(1).CopyFromRecordset pRS
    
    Else
        pCellRange.CopyFromRecordset pRS
    End If
   
Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Unable to Paste Data"
    Err.Raise (13)
    Exit Sub
End Sub



Public Sub GetParameter(SN As Integer)
    pSN = pControlWorkbook.Worksheets(pControlSheet).Range(pControlAncor).Offset(SN).Value
    pRangeName = pControlWorkbook.Worksheets(pControlSheet).Range(pControlAncor).Offset(SN, 1).Value
    pSheetName = pControlWorkbook.Worksheets(pControlSheet).Range(pControlAncor).Offset(SN, 2).Value
    pSQLSource = pControlWorkbook.Worksheets(pControlSheet).Range(pControlAncor).Offset(SN, 3).Value
    
    subUpdateFlag (pControlWorkbook.Worksheets(pControlSheet).Range(pControlAncor).Offset(SN, 4).Value)
    subHeaderFlag (pControlWorkbook.Worksheets(pControlSheet).Range(pControlAncor).Offset(SN, 5).Value)
    
    
End Sub

Public Sub SetContolDB _
    (Optional Control_Sheet As String = "", _
    Optional Control_Ancor As String = "", _
    Optional SQLServer_Name As String = "", _
    Optional Database_Name As String = "")
    
    If Control_Sheet <> "" Then pControlSheet = Control_Sheet
    End If
    If Control_Ancor <> "" Then pControlAncor = Control_Ancor
    End If
    If Database_Name <> "" Then pControlDBName = Database_Name
    End If
    If SQLServer_Name <> "" Then pControlSSName = SQLServer_Name
    End If
    
    DBControlParameterUpdate
    
End Sub


Private Sub DBControlParameterUpdate()
    
    pSQLServerName = pControlWorkbook.Worksheets(pControlSheet).Range(pControlSSName)
    pDatabaseName = pControlWorkbook.Worksheets(pControlSheet).Range(pControlDBName)
    pWBNameDir = pControlWorkbook.Worksheets(pControlSheet).Range(pControlWBName)
    
    Set pWorkSheetControl = Worksheets(pControlSheet)
    

End Sub

Private Sub WorkbookOpen()
    On Error GoTo ErrorHandler

    Set pWB_Workbook = Workbooks.Open(pWBNameDir)
    pWBFileName = pWB_Workbook.Name
    
Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Error on Opening workbook"
    MsgBox "Error: Invalid Workbook. Process Failed!"
    End
    Exit Sub
End Sub

Private Sub WorkbookClose()
    On Error GoTo ErrorHandler

    pWB_Workbook.Save
    pWB_Workbook.Close (True)
    Set pWB_Workbook = Nothing
    
Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Workbook not saved"
    Err.Raise (13)
    Exit Sub
End Sub

Private Sub WorkbookErrorClose()
    On Error GoTo ErrorHandler

    pWB_Workbook.Close (False)
    Set pWB_Workbook = Nothing
    
Exit Sub
ErrorHandler:
    MsgBox "Error (Item no. " & pSN & "): Workbook Unable to close"
    Err.Raise (13)
    Exit Sub
End Sub

Private Sub ValidWB()
    On Error GoTo ErrorHandler
    
    ValidateWorkbookFlag = False
    DBControlParameterUpdate
    
    Set pWB_Workbook = Workbooks.Open(pWBNameDir)
    pWBFileName = pWB_Workbook.Name
    pWB_Workbook.Close (False)
    
    ValidateWorkbookFlag = True
    
    Set pWB_Workbook = Nothing

Exit Sub
ErrorHandler:
ValidateWorkbookFlag = False
Err.Raise (13)

End Sub


Private Sub CheckValidateWorkbook()
    On Error GoTo ErrorHandler
    
    ValidWB

Exit Sub
ErrorHandler:
MsgBox "Error : Error on Opening workbook/ Invalid Workbook"
End
End Sub


' MAIN SUB
' ---------------------------------------------------------------------------------------


Public Sub Process_UpdateData(SN As Integer)
    On Error GoTo ErrorHandler1
    DBControlParameterUpdate
    GetParameter SN
    ProcessUpdatedSucessFlag = False
    If RunPass() = False Then
        Exit Sub
    Else
        ConnectSQL
        On Error GoTo ErrorHandler2
        
        GetData
        WorkbookOpen
        On Error GoTo ErrorHandler3
        
        UpdateCellRange
        ClearCellRange
        Paste_Data
        WorkbookClose
        On Error GoTo ErrorHandler2
        
        DisconnectSQL
        On Error GoTo ErrorHandler1
        
        ProcessUpdatedSucessFlag = True
    End If
    
Exit Sub
ErrorHandler3:
WorkbookErrorClose
ErrorHandler2:
DisconnectSQL
ErrorHandler1:
Exit Sub

End Sub

Sub Quick_Process_UpdateData(Optional StartItem As Integer = 1, Optional EndItem As Integer = 100)
    
    Dim I As Integer
    ProcessUpdatedSucessFlag = False
    On Error GoTo ErrorHandler1
    DBControlParameterUpdate
    ConnectSQL
    On Error GoTo ErrorHandler2
    WorkbookOpen
    On Error GoTo ErrorHandler3
    
    For I = StartItem To EndItem
        GetParameter I
        
        If RunPass() = True Then
            GetData
            UpdateCellRange
            ClearCellRange
            Paste_Data
            
        End If
    Next I
    
    WorkbookClose
    On Error GoTo ErrorHandler2
    DisconnectSQL
    On Error GoTo ErrorHandler1
    ProcessUpdatedSucessFlag = True
    
    Exit Sub
ErrorHandler3:
WorkbookErrorClose
ErrorHandler2:
DisconnectSQL
ErrorHandler1:
Exit Sub

End Sub



Public Sub Process_ClearData(SN As Integer)
    On Error GoTo ErrorHandler1
    DBControlParameterUpdate
    GetParameter SN
    ProcessClearSucessFlag = False
    If RunPass() = False Then
        Exit Sub
    Else
        WorkbookOpen
        On Error GoTo ErrorHandler2
        UpdateCellRange
        ClearCellRange
        WorkbookClose
        On Error GoTo ErrorHandler1
        ProcessClearSucessFlag = True
    End If
  
Exit Sub
ErrorHandler2:
WorkbookErrorClose
ErrorHandler1:
Exit Sub

End Sub



Public Sub Quick_Process_ClearData(StartItem As Integer, EndItem As Integer)
    
    Dim I As Integer
    ProcessClearSucessFlag = False
    On Error GoTo ErrorHandler1
    DBControlParameterUpdate
    WorkbookOpen
    On Error GoTo ErrorHandler2
    
    For I = StartItem To EndItem
    
        GetParameter I
        
        If RunPass() = True Then
            UpdateCellRange
            ClearCellRange
            
        End If
    Next I
    
    WorkbookClose
    On Error GoTo ErrorHandler1
    ProcessClearSucessFlag = True
    
    Exit Sub
ErrorHandler2:
WorkbookErrorClose
ErrorHandler1:
Exit Sub
    
End Sub


Public Sub Process_GoToRange(SN As Integer)
    On Error GoTo ErrorHandler1
    DBControlParameterUpdate
    GetParameter SN
    WorkbookOpen
    
    On Error GoTo ErrorHandler2
    UpdateCellRange
    CellRangeSelection

Exit Sub
ErrorHandler2:
WorkbookErrorClose
ErrorHandler1:
Exit Sub
    
End Sub

Public Sub ValidateWorkbook()
    On Error GoTo ErrorHandler
    
    ValidWB

Exit Sub
ErrorHandler:
ValidateWorkbookFlag = False
Exit Sub

End Sub
```


## Module:

### ControlModule:
```VBA
Sub Run_Update(SN As Integer)

    Dim R As Integer
    Dim F As SQLToExcel
    
    R = MsgBox("Are you sure you want to Update?", vbYesNo)
    If R <> 6 Then Exit Sub

    Set F = New SQLToExcel
    F.Process_UpdateData (SN)
    If F.ProcessUpdatedSucessFlag = True Then
        MsgBox "Data Updated on " & F.DestWBName & " [" & F.SheetName & " " & F.RangeAddress & "] For Item No. " & F.SNumber
    End If
    Set F = Nothing

End Sub

Sub Run_Clear(SN As Integer)

    Dim R As Integer
    Dim F As SQLToExcel
    
    R = MsgBox("Are you sure you want to Clear?", vbYesNo)
    If R <> 6 Then Exit Sub
    
    Set F = New SQLToExcel
    F.Process_ClearData (SN)
    If F.ProcessClearSucessFlag = True Then
        MsgBox "Data Cleared on " & F.DestWBName & " [" & F.SheetName & " " & F.RangeAddress & "] For Item No. " & F.SNumber
    End If
    Set F = Nothing
    
End Sub

Sub Go_To_Range(ByVal SN As Integer)

    Dim F As SQLToExcel
    Set F = New SQLToExcel
    F.Process_GoToRange SN
    Set F = Nothing
    
End Sub

Sub Run_Update_Item(ByVal SN As Integer)

    Dim F As SQLToExcel
    Set F = New SQLToExcel
    F.Process_UpdateData (SN)
    Set F = Nothing
    
End Sub

Sub Run_Clear_Item(ByVal SN As Integer)

    Dim F As SQLToExcel
    Set F = New SQLToExcel
    F.Process_ClearData (SN)
    Set F = Nothing
    
End Sub


Sub Run_Update_All(Optional StartItem As Integer = 1, Optional EndItem As Integer = 100)

    Dim I As Integer
    For I = StartItem To EndItem
        Run_Update_Item I
    Next I
    
    MsgBox "Update Sucessfull"

End Sub

Sub Run_Clear_All(Optional StartItem As Integer = 1, Optional EndItem As Integer = 100)

    Dim I As Integer
    For I = StartItem To EndItem
        Run_Clear_Item I
    Next I
    
    MsgBox "Clear Sucessfull"

End Sub

Sub Run_Update_All_Q(Optional StartItem As Integer = 1, Optional EndItem As Integer = 100)

    Dim F As SQLToExcel
    Set F = New SQLToExcel
    F.Quick_Process_UpdateData StartItem, EndItem
    
    If F.ProcessUpdatedSucessFlag = True Then
        MsgBox "Update All Successfull."
    Else
        MsgBox "Update All Failed. Turn off Quick mode to debug the issue"
    End If
    Set F = Nothing

End Sub

Sub Run_Clear_All_Q(Optional StartItem As Integer = 1, Optional EndItem As Integer = 100)
    
    Dim F As SQLToExcel
    Set F = New SQLToExcel
    F.Quick_Process_ClearData StartItem, EndItem
    
    If F.ProcessClearSucessFlag = True Then
        MsgBox "Clear All Successfull."
    Else
        MsgBox "Clear All Failed. Turn off Quick mode to debug the issue"
    End If
    Set F = Nothing

End Sub

Sub Update_All_Item()
    Dim R As Integer
    R = MsgBox("Are you sure you want to Update All?", vbYesNoCancel)
    If R <> 6 Then
        Exit Sub
    End If
    If ActiveSheet.Range("QUICK_MODE").Value = "YES" Then
        Run_Update_All_Q
    Else
        Run_Update_All
    End If
End Sub

Sub Clear_All_Item()
    Dim R As Integer
    R = MsgBox("Are you sure you want to Clear All?", vbYesNoCancel)
    If R <> 6 Then
        Exit Sub
    End If
    If ActiveSheet.Range("QUICK_MODE").Value = "YES" Then
        Run_Clear_All_Q
    Else
        Run_Clear_All
    End If
End Sub

Sub Workbook_Select()

    Dim FD As FileDialog
    Dim FSO As New FileSystemObject
    Dim FD_Result As Integer
    Dim DefaultLocation As String
    
    DefaultLocation = ActiveSheet.Range("WB_Name").Value
    
    Set FD = Application.FileDialog(msoFileDialogFilePicker)
    
    If DefaultLocation = "" Then
        FD.InitialFileName = "F:\Bausch Call Planning"
    Else
        FD.InitialFileName = FSO.GetParentFolderName(DefaultLocation)
    End If
    
    FD.Title = "Select the workbook"
    FD.ButtonName = "Use this workbook"
    FD.AllowMultiSelect = False
    FD.Filters.Clear
    FD.Filters.Add "XLSX", "*.xlsx"
    FD.Filters.Add "XLS", "*.xls"
    
    FD_Result = FD.Show
    
    If FD_Result = 0 Then Exit Sub
    
    ActiveSheet.Range("WB_Name").Value = FD.SelectedItems(1)
    
End Sub

Sub Validate_workbook()

    Dim F As SQLToExcel
    Set F = New SQLToExcel
    F.ValidateWorkbook
    If F.ValidateWorkbookFlag = True Then
        MsgBox "Workbook " & F.DestWBName & " is Valid"
    Else
        MsgBox "Workbook invalid!"
        
    End If
    
End Sub

```


# Sheet

```VBA

Private Sub Worksheet_FollowHyperlink(ByVal Target As Hyperlink)
    
    Dim RowNum, ColumnNum, SN As Integer
    
    RowNum = Target.Range.Row
    ColumnNum = Target.Range.Column
    SN = RowNum - 8
    
    'MsgBox RowNum & " " & ColumnNum & " " & SN
    'Exit Sub
    
    If RowNum >= 9 And RowNum <= 108 And ColumnNum >= 8 And ColumnNum <= 10 Then
        If ColumnNum = 9 Then
            Run_Update (SN)
            Exit Sub
        End If
        If ColumnNum = 8 Then
            Run_Clear (SN)
            Exit Sub
        End If
        If ColumnNum = 10 Then
            Go_To_Range (SN)
            Exit Sub
        End If
    End If
    
End Sub


```

