Attribute VB_Name = "modData"
Option Explicit

' to create a new record, the record date (the key) has to be checked against
' existing ones, so the date has to be provided first
Sub StartNewRecord()
    Dim inputWks As Worksheet
    Dim historicalWks As Worksheet
    Dim rngDataEntry As Range
    Dim rngDates As Range
    Dim rngHistorical As Range
    Dim rngNew As Range
    Dim newDt As Date
    Dim str As String
    Dim iRow As Integer
    
    Set inputWks = wksDataEntry
    Set historicalWks = wksHistorical
    Set rngHistorical = historicalWks.Range("tblHistorical")
    Set rngDataEntry = inputWks.Range("DataEntry")
    Set rngDates = historicalWks.Range("DateSeries")
    
    str = InputBox("The new date must be greater than the first:", "Enter a new date", "yyyy-mm-dd")
    If str = "" Then Exit Sub
    On Error Resume Next
    newDt = 0
    newDt = str
    If newDt = 0 Then
        MsgBox "invalid date format.  use yyyy-mm-dd"
        Exit Sub
    End If
    On Error GoTo 0
    
    If newDt <= rngDates.Cells(1, 1).Value Then
         MsgBox "New Date must be greater than the first in DateSeries"
         Exit Sub
    End If
    On Error Resume Next
    iRow = -1
    'locate the last date that is less than newDt
    iRow = Application.Match(CLng(newDt), rngDates, 1)
    On Error GoTo 0
    If iRow = -1 Then
        MsgBox "Something wrong: could not locate record"
        Exit Sub
    End If
    If newDt = rngDates.Cells(iRow, 1).Value Then
        inputWks.Range("currRec") = iRow
        modViewData.ViewLogCurrent
        Exit Sub
    End If
    
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    
    With inputWks
        rngDataEntry.ClearContents
        .Range("InputAnchor").Value = newDt
        .Range("RecSelected").Value = newDt
        'a new record will be inserted below the selected row
        .Range("currRec").Value = iRow + 1
    End With
    
    With historicalWks
        .Activate
        .Unprotect
        iRow = iRow + dataTblOffset + 1
        If iRow > rngHistorical.Rows.Count Then
            MsgBox "Something wrong: failed to add the record for " & newDt
            Exit Sub
        End If
        rngHistorical.Rows(iRow).Select
        Application.CutCopyMode = False 'got to clear buffer before insert
        Selection.Insert xlShiftDown
        rngDataEntry.Copy
        Selection.PasteSpecial Paste:=xlPasteValues, Transpose:=True
        Application.CutCopyMode = False
        .Protect
    End With
    inputWks.Activate
    inputWks.Range("InputAnchor").Activate
    modViewData.ViewLogCurrent
    Application.EnableEvents = True
    Application.ScreenUpdating = False

End Sub

Sub UpdateLogRecord()

    Dim historicalWks As Worksheet
    Dim inputWks As Worksheet

    Dim lRec As Long
    Dim dKey As Date
    Dim rngDateSeries As Range
    Dim rngHistorical As Range

    Dim myCopy As Range
    Dim myTest As Range
    
    Dim lRsp As Long
    
    Set inputWks = wksDataEntry
    Set historicalWks = wksHistorical
    Set rngDateSeries = historicalWks.Range("DateSeries")
    Set rngHistorical = historicalWks.Range("tblHistorical")
    dKey = inputWks.Range("inputAnchor").Value
    On Error Resume Next
    lRec = -1
    lRec = Application.Match(CLng(dKey), rngDateSeries, 0)
    On Error GoTo 0
    If lRec = -1 Then
        MsgBox "Date " & dKey & " is not in the DateSeries. Click Add first"
        Exit Sub
    End If
    
    Set myCopy = inputWks.Range("DataEntry")

    With inputWks
        Set myTest = myCopy.Offset(0, 2)
        If Application.Count(myTest) > 0 Then
            MsgBox "Please fill in all the cells!"
            Exit Sub
        End If
    End With
    
    Application.ScreenUpdating = False
    With historicalWks
        .Activate
        .Unprotect
        myCopy.Copy
        rngHistorical.Rows(lRec + dataTblOffset).Select
        Selection.PasteSpecial Paste:=xlPasteValues, Transpose:=True
        Application.CutCopyMode = False
        .Protect
    End With
    
    'clear input cells that contain constants
    With inputWks
        .Activate
        If .Range("ShowMsg").Value = "Yes" Then
           MsgBox "Database has been updated."
        End If
    End With
    Application.ScreenUpdating = True
  
End Sub

Sub DeleteLogRecord()
    'do not execute:
    'Exit Sub
    
    Dim historicalWks As Worksheet
    Dim inputWks As Worksheet

    Dim lRec As Long
    Dim lRecRow As Long
    Dim dKey As Date

    Dim myCopy As Range
    
    Set inputWks = wksDataEntry
    Set historicalWks = wksHistorical
        
    'cells to clear after deleting record
    Set myCopy = inputWks.Range("DataEntry")
    lRec = inputWks.Range("currRec").Value
    dKey = inputWks.Range("inputAnchor").Value
    
    If vbYes = MsgBox("Confirm to delete the current record!", _
    vbCritical + vbYesNo, "Delete record") Then
        Application.ScreenUpdating = False
        With historicalWks
            .Activate
            .Unprotect
            On Error Resume Next
            lRecRow = -1
            lRecRow = Application.Match(CLng(dKey), .Range("DateSeries"), 0)
            On Error GoTo 0
            If lRecRow = -1 Then
                MsgBox "The current record is not in the database!"
                Exit Sub
            End If
            With .Range("tblHistorical")
                Application.DisplayAlerts = False
                .Rows(lRecRow + dataTblOffset).EntireRow.Delete xlShiftUp
                Application.DisplayAlerts = True
            End With
            .Protect
        End With
    
        'clear input cells that contain constants
        Application.ScreenUpdating = True
        With inputWks
            .Activate
            .Unprotect
            myCopy.Copy
            .Range("backupAnchor").PasteSpecial xlPasteValues
            Application.CutCopyMode = False
            myCopy.ClearContents
            If lRec >= .Range("LastRec").Value Then
                Call ViewLogLast
            Else
                .Range("currRec") = lRec
                'ViewLogCurrent
            End If
            .Protect
        End With
    End If
End Sub


