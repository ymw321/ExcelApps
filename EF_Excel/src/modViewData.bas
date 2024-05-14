Attribute VB_Name = "modViewData"

Option Explicit

Public Const dataTblOffset As Integer = 0

Sub ViewLogFirst()
    ViewLog1stLast 1
End Sub
Sub ViewLogLast()
    ViewLog1stLast 2
End Sub

Sub ViewLogUp()
    ViewLogNext (-1)
End Sub
Sub ViewLogDown()
    ViewLogNext (1)
End Sub
Sub ViewLogCurrent()
    ViewLogNext (0)
End Sub

Sub ViewLogNext(ByVal step As Integer)
    Dim historyWks As Worksheet
    Dim inputWks As Worksheet
    Dim rngA As Range

    Dim lRec As Long
    Dim dKey As Date
    Dim lRecRow As Long
    Dim lRecCount As Long
    Dim rngNextRec As Range
    Application.EnableEvents = False
    
    Set inputWks = wksDataEntry
    Set historyWks = wksHistorical
    Set rngA = ActiveCell
    
    'assuming a valid record is already showing. o.w., view first or last
    With inputWks
        lRec = .Range("currRec").Value
        dKey = .Range("inputAnchor").Value
    End With
    With historyWks
        lRecCount = .Range("DateSeries").Count
        If Not (lRec >= 1 And lRec <= lRecCount) Then
            MsgBox "Incorrect record position!" & Chr(10) & "Click First or Last to restart"
            Exit Sub
        Else
            If ((lRec = 1 And step = -1) Or (lRec = lRecCount And step = 1)) Then Exit Sub
        End If
        Set rngNextRec = historyWks.Range("tblHistorical").Rows(dataTblOffset + lRec + step)
    End With

    With inputWks
        .Unprotect
        .Range("CurrRec").Value = lRec + step
        rngNextRec.Copy
        .Range("inputAnchor").PasteSpecial Paste:=xlPasteValues, Transpose:=True
        .Range("RecSelected").Value = .Range("inputAnchor").Value
        rngA.Select
        .Protect
    End With
    Application.EnableEvents = True

End Sub

Sub ViewLog1stLast(ByVal pos As Integer)
  
    Dim historyWks As Worksheet
    Dim inputWks As Worksheet
    Dim rngA As Range

    Dim rngRec As Range
    Dim lRec As Long
    Application.EnableEvents = False
    
    Set inputWks = wksDataEntry
    Set historyWks = wksHistorical
    Set rngA = ActiveCell

    With historyWks
        If pos = 1 Then lRec = 1 Else lRec = .Range("DateSeries").Count
        Set rngRec = .Range("tblHistorical").Rows(lRec + dataTblOffset)
        'rngRec.Copy
    End With

    With inputWks
        .Unprotect
        .Range("CurrRec").Value = lRec
        rngRec.Copy
        .Range("inputAnchor").PasteSpecial Paste:=xlPasteValues, Transpose:=True
        .Range("RecSelected").Value = .Range("inputAnchor").Value
        rngA.Select
        .Protect
    End With
    
    Application.EnableEvents = True

End Sub

Sub updateGroupAssignment()
    Dim lstGroups As ListObject, lstAssets As ListObject, rngAnchor As Range
    wksAssets.Activate
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    
    Set lstGroups = wksAssets.ListObjects("tblGroup")
    Set lstAssets = wksAssets.ListObjects("tblAsset")
    Set rngAnchor = wksAssets.Range("GroupAnchor")
        
    'get asset names
    rngAnchor.Offset(1, 0).End(xlDown).ClearContents
    lstAssets.ListColumns("ShortName").DataBodyRange.Copy
    rngAnchor.Offset(1, 0).PasteSpecial Paste:=xlPasteValues, Transpose:=False
    'get group names
    rngAnchor.Offset(0, 1).End(xlToRight).ClearContents
    lstGroups.ListColumns("GroupShortName").DataBodyRange.Copy
    rngAnchor.Offset(0, 1).PasteSpecial Paste:=xlPasteValues, Transpose:=True
    Application.CutCopyMode = False
    
    MsgBox "Group Assignment coordinates updated!" & Chr(10) _
        & "Please review the assignment before continue"
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub

Sub UpdateActiveAssets()

    Const colInclusion As Integer = 5, colId As Integer = 1, colShortName As Integer = 2, colA_L As Integer = 4
    
    Dim rngAssets, rngInclusion, rngActiveAssets, rngAL As Range
    Set rngAssets = wksAssets.ListObjects("tblAsset").DataBodyRange
    Set rngInclusion = wksAssets.ListObjects("tblAsset").ListColumns(colInclusion).Range
    Set rngActiveAssets = wksStatistics.Range("ActiveAssets")
    Dim i As Integer, j As Integer
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    With wksStatistics
        .Activate
        .Unprotect
        rngActiveAssets.Select
        Range(Selection, Selection.End(xlDown)).ClearContents
    
        j = 1
        For i = 1 To rngAssets.Rows.Count
            If rngAssets.Cells(i, colInclusion) = 1 Then
                rngActiveAssets.Cells(j, 1) = rngAssets.Cells(i, colId)
                rngActiveAssets.Cells(j, 2) = rngAssets.Cells(i, colShortName)
                rngActiveAssets.Cells(j, 3) = rngAssets.Cells(i, colA_L)
                j = j + 1
            End If
        Next i
        If j - 1 <> rngActiveAssets.Rows.Count Then
            MsgBox "Something went wrong: inconsistent active assets count!"
        End If
        '.Sort.SortFields.Add Key:=rngActiveAssets.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With rngActiveAssets
            .Sort Key1:=.Columns(3) _
            , Order1:=xlAscending _
            , Key2:=.Columns(1) _
            , Order2:=xlAscending
        End With
        .Protect
    End With
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub

Sub UpdateActiveGroups()
    Const colInclusion As Integer = 3, colId As Integer = 1, colShortName As Integer = 2
    
    Dim rngGroups, rngInclusion, rngActiveGroups
    Set rngGroups = wksAssets.ListObjects("tblGroup").DataBodyRange
    Set rngInclusion = wksAssets.ListObjects("tblGroup").ListColumns(colInclusion).Range
    Set rngActiveGroups = wksStatistics.Range("ActiveGroups")
    Dim i As Integer, j As Integer
    
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With
    
    With wksStatistics
        .Activate
        .Unprotect
        rngActiveGroups.End(xlToRight).ClearContents
    
        j = 1
        For i = 1 To rngGroups.Rows.Count
            If rngGroups.Cells(i, colInclusion) = 1 Then
                rngActiveGroups.Cells(1, j) = rngGroups.Cells(i, colId)
                rngActiveGroups.Cells(2, j) = rngGroups.Cells(i, colShortName)
                j = j + 1
            End If
        Next i
        If j - 1 <> rngActiveGroups.Columns.Count Then
            MsgBox "Something went wrong: inconsistent active groups count!"
        End If
        '.Sort.SortFields.Add Key:=rngActiveAssets.Columns(3) _
            , SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        With rngActiveGroups
            .Sort Key1:=.Rows(1) _
            , Order1:=xlAscending
        End With
        .Protect
    End With
    
    With Application
        .ScreenUpdating = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With
End Sub
