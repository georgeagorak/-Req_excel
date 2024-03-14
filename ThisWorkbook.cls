VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

'After successful save of .xlsx file, .csv will be generated.
'
' - it exports to a file named as the current sheet
' - it respects the locale separation char.
' - allows to continue editing xlsm file, instead of editing the exported CSV.
' - closes a exported CSV file is is currently open.


Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    Dim MyFileName As String, MyFilePathName As String ' Vars declaration
    Dim CurrentWB As Workbook, TempWB As Workbook
    Dim CSV_fileOpened As Boolean
    Dim sht As Worksheet
    
    Found = False
    For Each sht In ThisWorkbook.Worksheets
        If sht.Name Like "*-Req" Then
            ' this worksheet ends in -Req, so we can assume it contaisn requirements
            ' so create a csv file for this sheet
            Found = True
            Set CurrentWB = ActiveWorkbook
            sht.UsedRange.Copy
    
            Set TempWB = Application.Workbooks.Add(1)
            With TempWB.Sheets(1).Range("A1")
              .PasteSpecial xlPasteValues
              .PasteSpecial xlPasteFormats
            End With ' copy all the contents of current workbook into a new temporary xlsm file (before exporting)

            'Dim Change below to "- 4"  to become compatible with .xls files
            MyFilePathName = CurrentWB.Path & "\" & Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & sht.Name & ".csv" ' get name of current work book and make it compitable for csv format
            MyFileName = Left(CurrentWB.Name, Len(CurrentWB.Name) - 5) & sht.Name & ".csv"
    
            If isFilePresent(MyFilePathName) Then ' does csv file exist...
                If IsWorkBookOpen(MyFilePathName) Then ' is csv file open...
                    Workbooks(MyFileName).Close SaveChanges:=False ' if so close it so export can be success
                End If
            End If
    
            Application.DisplayAlerts = False
            TempWB.SaveAs FileName:=MyFilePathName, FileFormat:=xlCSV, CreateBackup:=False, Local:=True ' export current workbook to csv
            TempWB.Close SaveChanges:=False
        End If
    Next
    Application.DisplayAlerts = True
    If Found = False Then
        MsgBox "Did not find any worksheets names ending with -Req to export to CSV"
    End If
End Sub

'
' Determinates if the given specified workbook is currently open.
'
Function IsWorkBookOpen(FileName As String)
    Dim ff As Long, ErrNo As Long

    On Error Resume Next
    ff = FreeFile()
    Open FileName For Input Lock Read As #ff
    Close ff
    ErrNo = Err
    On Error GoTo 0

    Select Case ErrNo
    Case 0:    IsWorkBookOpen = False
    Case 70:   IsWorkBookOpen = True
    Case Else: Error ErrNo
    End Select
End Function

Function isFilePresent(ByVal fName As String) As Boolean
'Returns TRUE if the provided name points to an existing file.
'Returns FALSE if not existing, or if it's a folder
    On Error Resume Next
    IsFile = ((GetAttr(fName) And vbDirectory) <> vbDirectory)
End Function

