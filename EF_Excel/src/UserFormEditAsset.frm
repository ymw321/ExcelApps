VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEditAsset 
   Caption         =   "Edit an Asset"
   ClientHeight    =   7080
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9900
   OleObjectBlob   =   "UserFormEditAsset.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormEditAsset"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public EnableEvents As Boolean

Private Sub btnApply_Click()
    'update the record with new values
    Dim idx As Integer
    idx = Me.ListBoxAssets.ListIndex + 1
    
    Dim oSh As Worksheet, tblAsset As Range
    Set oSh = Worksheets("Assets")
    'stop button click event
    Me.EnableEvents = False
    Application.EnableEvents = False
    
    oSh.Unprotect
    Set tblAsset = oSh.ListObjects("tblAsset").DataBodyRange
    'update values
        If Me.TextBoxId = "" Then
            Exit Sub
        ElseIf Not (tblAsset(idx, 1) = CInt(Me.TextBoxId)) Then
            MsgBox "Something is wrong: the selected record does not match the data in Sheet"
            Exit Sub
        End If
        'tblAsset(idx, 1) = CInt(Me.TextBoxId)
        tblAsset(idx, 2) = TextBoxShortName
        tblAsset(idx, 3) = TextBoxLongName
        tblAsset(idx, 4) = TextBoxAorL
        tblAsset(idx, 5) = CInt(Me.TextBox1or0)
        tblAsset(idx, 6) = TextBoxSort
    'done
    Me.ListBoxAssets.ListIndex = idx - 1
    Me.ListBoxAssets.RowSource = "tblAsset"
    
    modViewData.UpdateActiveAssets
    modViewData.updateGroupAssignment
    
    MsgBox "Active Asset List has been updated. " & Chr(10) & "Please examine the Statistics tab before save/close this workbook!"
    oSh.Protect
    Me.EnableEvents = True
    Application.EnableEvents = True
End Sub

Private Sub btnNewAsset_Click()
    Dim oSh As Worksheet, tblAsset As ListObject
    Dim rngObj As Range
    Set oSh = Worksheets("Assets")
    'stop button click event
    'Me.EnableEvents = False
    Application.EnableEvents = False
    
    oSh.Unprotect
    Set tblAsset = oSh.ListObjects("tblAsset")
    Set rngObj = tblAsset.ListColumns("AssetId").Range
    Dim maxId As Long, lastRow As Long
    maxId = Application.WorksheetFunction.Max(rngObj)
    lastRow = Application.WorksheetFunction.Count(rngObj)
    Set rngObj = tblAsset.ListRows(lastRow).Range
    If Not (rngObj(1, 2) = "") Then
        Set rngObj = tblAsset.ListRows.Add(lastRow + 1).Range
        rngObj(1, 1) = maxId + 1
        Me.ListBoxAssets.ListIndex = maxId
        Call ClearInput(maxId + 1)
    Else
        Me.ListBoxAssets.ListIndex = maxId - 1
        Call ClearInput(maxId)
    End If
    Me.ListBoxAssets.RowSource = "tblAsset"
    oSh.Protect
    'Me.EnableEvents = True
    Application.EnableEvents = True
End Sub

Private Sub ListBoxAssets_Click() 'DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.EnableEvents = False Then Exit Sub
    'populate the text boxes with selected item
    Me.EnableEvents = False
    With UserFormEditAsset
        .TextBoxId = .ListBoxAssets.List(.ListBoxAssets.ListIndex, 0)
        .TextBoxShortName = .ListBoxAssets.List(.ListBoxAssets.ListIndex, 1)
        .TextBoxLongName = .ListBoxAssets.List(.ListBoxAssets.ListIndex, 2)
        .TextBoxAorL = .ListBoxAssets.List(.ListBoxAssets.ListIndex, 3)
        .TextBox1or0 = .ListBoxAssets.List(.ListBoxAssets.ListIndex, 4)
        .TextBoxSort = .ListBoxAssets.List(.ListBoxAssets.ListIndex, 5)
    End With
    Me.EnableEvents = True
End Sub

Private Sub ClearInput(ByVal id As Integer)
    Me.EnableEvents = False
    With UserFormEditAsset
        .TextBoxId = id
        .TextBoxShortName = ""
        .TextBoxLongName = ""
        .TextBoxAorL = ""
        .TextBox1or0 = ""
        .TextBoxSort = id
    End With
    Me.EnableEvents = True
End Sub

Private Sub TextBox1or0_Change()
    With Me.TextBox1or0
        If .Value = "" Then Exit Sub
        If .Value <> "1" And .Value <> "0" Then
            MsgBox "Only accepts 1 or 0"
            .Value = ""
        End If
    End With
End Sub

Private Sub TextBoxAorL_Change()
    With Me.TextBoxAorL
        If .Value = "" Then Exit Sub
        If UCase(.Value) <> "L" And UCase(.Value) <> "A" Then
            MsgBox "Only accepts L or A"
            .Value = ""
        End If
        .Value = UCase(.Value)
    End With
End Sub

Private Sub UserForm_Initialize()
    EnableEvents = True
End Sub
