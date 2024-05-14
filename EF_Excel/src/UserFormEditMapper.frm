VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserFormEditMapper 
   Caption         =   "Edit Line Item"
   ClientHeight    =   6510
   ClientLeft      =   420
   ClientTop       =   1650
   ClientWidth     =   9600
   OleObjectBlob   =   "UserFormEditMapper.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserFormEditMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Public EnableEvents As Boolean
Private theWorkAnchor As Range
Private newItemSelected As Boolean
Private newMliItemSelected As Boolean


Private Sub ComboBoxDeloitteItems_Click()
    If Me.EnableEvents = False Then Exit Sub
    'find if the Deloitte item has been mapped. If yes, alert and exit.
    'otherwise, start adding Mli items with mapping multipliers
    Me.EnableEvents = False
    With ComboBoxDeloitteItems
        Dim theItemId As Integer
        theItemId = .Value
        'is theItemId in ListBoxMapper? Add and Show it
        With Me.ListBoxMapper
            Dim rng As Range, anchor As Range, lastCell As Range, foundCell As Range
            Set anchor = Range("tblWorksheet[[#Headers],[DeloitteFieldId]]")
            Set lastCell = anchor.End(xlDown)
            If anchor.Cells(2, 1) Is Nothing Or anchor.Cells(2, 1) = "" Then
                Set lastCell = anchor
            End If
            Set rng = Range(anchor, lastCell)
            On Error Resume Next
            Set foundCell = rng.Find(theItemId)
            If foundCell Is Nothing Then
                'insert a new row below at lastCell
                'lastCell.ListObject.ListRows.Add
                Set theWorkAnchor = lastCell.offset(1, 0)
            Else
                'locate the last row of the same deloitte item
                Dim theOffset As Integer
                theOffset = 1
                While foundCell.offset(theOffset, 0) = theItemId
                    theOffset = theOffset + 1
                Wend
                
                'insert a row below at anchor
                'anchor.ListObject.ListRows.Add
                Set theWorkAnchor = foundCell.offset(theOffset, 0)
            End If
            'theWorkAnchor.ListObject.ListRows.Add Position:=theWorkAnchor.Row - anchor.Row, AlwaysInsert:=True
            'Set theWorkAnchor = theWorkAnchor.offset(-1, 0)
            
        End With
        newItemSelected = True
        
    End With
    Me.EnableEvents = True
End Sub

Private Sub ComboBoxMliItems_Click()
    If Me.EnableEvents = False Then Exit Sub
    'find if the Deloitte item has been mapped. If yes, alert and exit.
    'otherwise, start adding Mli items with mapping multipliers
    newMliItemSelected = True
End Sub

Private Sub btnAdd_Click()
    If Not newItemSelected Then 'do nothing if no new item is selected
        MsgBox "Select a new item to add!"
        Exit Sub
    End If
    'add the selected MLI field along with the selected Deloitte field as a new row to the mapper table
    Dim oSh As Worksheet, tblMapper As ListObject, anchor As Range
    Dim rngObj As Range
    Set oSh = Worksheets("Worksheet")
    Set anchor = Range("tblWorksheet[[#Headers],[DeloitteFieldId]]")
    'stop button click event
    'Me.EnableEvents = False
    Application.EnableEvents = False
    'Me.ListBoxMapper.Selected(theWorkAnchor.Row - anchor.Row) = True
    
    Dim mliId As Integer, mliName As String, delId As Integer, delName As String
    ' add a row and set data
    theWorkAnchor.ListObject.ListRows.Add Position:=theWorkAnchor.Row - anchor.Row, AlwaysInsert:=True
    Set theWorkAnchor = theWorkAnchor.offset(-1, 0)
    With Me.ComboBoxDeloitteItems
        theWorkAnchor.Cells(1, 1) = .Value
        theWorkAnchor.Cells(1, 2) = .Column(2)
    End With
    With Me.ComboBoxMliItems
        theWorkAnchor.Cells(1, 3) = .Value
        theWorkAnchor.Cells(1, 4) = .Column(2)
    End With
    theWorkAnchor.Cells(1, 5) = Me.txtMultiplier.Value
    'Me.ListBoxMapper.RowSource = Range(theWorkAnchor.Cells(1, 1), theWorkAnchor.Cells(10, 7)).Address
    'Me.ListBoxMapper.Selected(theWorkAnchor.Row - anchor.Row) = True
    ' refresh now that new rows are added
    Me.ListBoxMapper.RowSource = "tblWorksheet"

    'Me.Show
    newItemSelected = False
    'Set theWorkAnchor = Range("tblWorksheet[[#Headers],[DeloitteFieldId]]")
    'Me.ComboBoxDeloitteItems.Value = ""
    
    Application.EnableEvents = True
End Sub

Sub test()
    'oSh.Unprotect
    Set tblMapper = oSh.ListObjects("tblWorksheet")
    Set rngObj = tblMapper.ListColumns("DeloitteFieldId").Range
    
    maxId = theWorkAnchor
    lastRow = Application.WorksheetFunction.Count(rngObj)
    Set rngObj = tblMapper.ListRows(lastRow).Range
    If Not (rngObj(1, 2) = "") Then
        Set rngObj = tblMapper.ListRows.Add(lastRow + 1).Range
        rngObj(1, 1) = maxId + 1
        Me.ListBoxMapper.ListIndex = 0
        'Call ClearInput(maxId + 1)
    Else
        Me.ListBoxMapper.ListIndex = maxId - 1
        'Call ClearInput(maxId)
    End If
    Me.ListBoxMapper.RowSource = Range(theWorkAnchor.Cells(1, 1), theWorkAnchor.Cells(10, 7))
    'oSh.Protect
    'Me.EnableEvents = True
    Application.EnableEvents = True
End Sub

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
    Me.EnableEvents = True
    Application.EnableEvents = True
    newItemSelected = False
    newMliItemSelected = False
    Set theWorkAnchor = Range("tblWorksheet[[#Headers],[DeloitteFieldId]]")
    With Me.ComboBoxDeloitteItems
        .Value = ""
    End With
    Me.ListBoxMapper.Selected(0) = True
End Sub

