VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEditGroup 
   Caption         =   "Edit a Group"
   ClientHeight    =   6630
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9075
   OleObjectBlob   =   "UserFormEditGroup.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormEditGroup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public EnableEvents As Boolean

Private Sub btnApply_Click()
    'update the record with new values
    Dim idx As Integer
    idx = Me.ListBoxGroups.ListIndex + 1
    
    Dim oSh As Worksheet, tblGroup As Range
    Set oSh = wksAssets
    'stop button click event
    Me.EnableEvents = False
    Application.EnableEvents = False
    
    oSh.Unprotect
    Set tblGroup = oSh.ListObjects("tblGroup").DataBodyRange
    'update values
        If Me.TextBoxId = "" Then
            Exit Sub
        ElseIf Not (tblGroup(idx, 1) = CInt(Me.TextBoxId)) Then
            MsgBox "Something is wrong: the selected record does not match the data in Sheet"
            Exit Sub
        End If
        'tblGroup(idx, 1) = CInt(Me.TextBoxId)
        tblGroup(idx, 2) = TextBoxShortName
        tblGroup(idx, 3) = CInt(Me.TextBox1or0)
        tblGroup(idx, 4) = TextBoxLongName
    'done
    Me.ListBoxGroups.ListIndex = idx - 1
    Me.ListBoxGroups.RowSource = "tblGroup"
    modViewData.updateGroupAssignment
    wksAssets.Range("GroupAnchor").Activate
    oSh.Protect
    Me.EnableEvents = True
    Application.EnableEvents = True
End Sub

Private Sub btnNewGroup_Click()
    Dim oSh As Worksheet, tblGroup As ListObject
    Dim rngObj As Range
    Set oSh = wksAssets
    'stop button click event
    'Me.EnableEvents = False
    Application.EnableEvents = False
    
    oSh.Unprotect
    Set tblGroup = oSh.ListObjects("tblGroup")
    Set rngObj = tblGroup.ListColumns("GroupId").Range
    Dim maxId As Long, lastRow As Long
    maxId = Application.WorksheetFunction.Max(rngObj)
    lastRow = Application.WorksheetFunction.Count(rngObj)
    Set rngObj = tblGroup.ListRows(lastRow).Range
    If Not (rngObj(1, 2) = "") Then
        Set rngObj = tblGroup.ListRows.Add(lastRow + 1).Range
        rngObj(1, 1) = maxId + 1
        Me.ListBoxGroups.ListIndex = maxId
        Call ClearInput(maxId + 1)
    Else
        Me.ListBoxGroups.ListIndex = maxId - 1
        Call ClearInput(maxId)
    End If
    Me.ListBoxGroups.RowSource = "tblGroup"
    oSh.Protect
    'Me.EnableEvents = True
    Application.EnableEvents = True
End Sub

Private Sub ListBoxGroups_Click() 'DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    If Me.EnableEvents = False Then Exit Sub
    'populate the text boxes with selected item
    Me.EnableEvents = False
    With UserFormEditGroup
        .TextBoxId = .ListBoxGroups.List(.ListBoxGroups.ListIndex, 0)
        .TextBoxShortName = .ListBoxGroups.List(.ListBoxGroups.ListIndex, 1)
        .TextBox1or0 = .ListBoxGroups.List(.ListBoxGroups.ListIndex, 2)
        .TextBoxLongName = .ListBoxGroups.List(.ListBoxGroups.ListIndex, 3)
    End With
    Me.EnableEvents = True
End Sub

Private Sub ClearInput(ByVal id As Integer)
    Me.EnableEvents = False
    With UserFormEditAsset
        .TextBoxId = id
        .TextBoxShortName = ""
        .TextBoxLongName = ""
        .TextBox1or0 = ""
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

Private Sub UserForm_Initialize()
    EnableEvents = True
End Sub
