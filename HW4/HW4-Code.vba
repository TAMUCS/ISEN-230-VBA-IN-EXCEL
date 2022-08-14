'''''''''''''''''''''''''' MODULE 1 ''''''''''''''''''''''''''
Option Explicit

Sub Question1()
    frmBigTen.Show
End Sub

Sub Question2()
    frmBigTenQ2.Show
End Sub

'''''''''''''''''''''''''' frmBigTen (Q1) ''''''''''''''''''''''''''
Option Explicit

Public usrState As String

'[CONSTRUCTOR]: User Form Constructor
Private Sub UserForm_Initialize()
    With frmBigTen
        With .lstStates
            .MultiSelect = fmMultiSelectSingle
            .Font.Size = 14
            .RowSource = "States"
            .Selected(1) = True
        End With
    End With
    usrState = "Indiana"
End Sub

'[EVENT]: User selects state
Private Sub lstStates_AfterUpdate()
    usrState = Me.lstStates
End Sub

'[EVENT]: cmdCancel button is pressed
Private Sub cmdCancel_Click()
    MsgBox "You must not live in a Big Ten state."
    Unload Me
End Sub

'[EVENT]: cmdOK button is pressed
Private Sub cmdOK_Click()
    MsgBox "You live in " & usrState & "."
    Unload Me
End Sub

'[EVENT]: ONLY when frmBigTen red "x" button is pressed
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        MsgBox "You must not live in a Big Ten state."
        Unload Me
    End If
End Sub



'''''''''''''''''''''''''' frmBigTenQ2 ''''''''''''''''''''''''''

Option Explicit

Public lastRow As Long

'[CONSTRUCTOR]: User Form Constructor
Private Sub UserForm_Initialize()
    With frmBigTenQ2
        With .lstStatesQ2
            .MultiSelect = fmMultiSelectMulti
            .Font.Size = 14
            .RowSource = "States"
            .Selected(1) = True
        End With
    End With
    lastRow = Range("A1").End(xlDown).Row
End Sub

'[EVENT]: cmdCancel button is pressed
Private Sub cmdCancelQ2_Click()
    MsgBox "You must not live in a Big Ten state."
    Unload Me
End Sub

'[EVENT]: cmdOK button is pressed
Private Sub cmdOKQ2_Click()
    Dim selectedRows As Collection
    Dim states() As String
    ReDim states(0 To lastRow - 2)
    Set selectedRows = GetSelectedRows(Me.lstStatesQ2)
    Dim outStr As String: outStr = "You have lived in "
    Dim i As Long, j As Long, currRow As Long
    
    'Populate Str array w/ all states from spreadsheet
    For i = 0 To lastRow - 2
        states(i) = Range("A2").Offset(i, 0).Value
    Next i

    Dim numSel As Long: numSel = selectedRows.count
    Dim state As String
    
    'Handle base case (NO SELECTION)
    If numSel = 0 Then
        MsgBox "You must not live in a Big Ten state."
        Unload Me
        Exit Sub
    End If
    
    'Loop through selections; concatenate with output str
    For j = 1 To numSel
        state = states(selectedRows(j))
        With lstStatesQ2
            
            'Handle Base Case (1 SELECTION)
            If numSel = 1 Then
                outStr = outStr & state & "."
                Exit For
            End If
            
            'Handle discrete printing w/ final value detection (j &| numSel > 1)
            If j = numSel Then
                outStr = outStr & "and " & state & "."
            Else
                outStr = outStr & state & ", "
            End If
        End With
    Next j
    
    MsgBox outStr
    Unload Me
End Sub

'[EVENT]: ONLY when frmBigTen red "x" button is pressed
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        MsgBox "You must not live in a Big Ten state."
        Unload Me
    End If
End Sub

' Returns a collection of the index of all the selected items
' CREDIT: https://excelmacromastery.com/vba-listbox/
Function GetSelectedRows(lstBox As MSForms.ListBox) As Collection

    ' Create the collection
    Dim coll As New Collection

    ' Read through each item in the listbox
    Dim i As Long
    For i = 0 To lstBox.ListCount - 1

        ' Check if item at position i is selected
        If lstBox.Selected(i) Then
            coll.Add i
        End If
    Next i

    Set GetSelectedRows = coll

End Function

'''''''''''''''''''''''''' Module 1 (Q3) ''''''''''''''''''''''''''
Option Explicit

Public repData As Variant

Sub Question3()
    frmQ3a.Show
End Sub

'''''''''''''''''''''''''' frmQ3a ''''''''''''''''''''''''''
Option Explicit

Public nRows As Long

Private Sub UserForm_Initialize()
     nRows = Range("A3").End(xlDown).Row - 3
End Sub

Private Sub cmdOK_Click()
    Dim lastName As String: lastName = Me.txtLastName.Value
    Dim firstName As String: firstName = Me.txtFirstName.Value
    
    'Check for empty inputs by user
    If Len(firstName) + Len(lastName) = 0 Then
        Exit Sub
    ElseIf lastName = "" Then
        MsgBox "Enter a last name for the rep you want to find."
        Exit Sub
    ElseIf firstName = "" Then
        MsgBox "Enter a first name for the rep you want to find."
        Exit Sub
    End If
    
    repData = searchRep(LCase(lastName), LCase(firstName))
    
    'Check for invalid match
    If repData(0) = "FALSE" Then
        MsgBox "There is no such rep, so no editing can occur", vbOKOnly, "No such rep"
        Exit Sub
    End If
    
    'Show next userform, hide current
    Unload Me
    frmQ3b.Show
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'[EVENT]: Red "x" button pressed on userform
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = 0 Then
        Unload Me
    End If
End Sub

'Looks for matching rep. If found, returns rep data, else returns arry with "FALSE"
Function searchRep(lstName As String, fstName As String) As Variant
    Dim i As Long, j As Integer
    Dim match As Boolean: match = False
    ReDim repData(0 To 7)
    
    With Range("A4")
        'Checks each row for fist & last name match
        For i = 0 To nRows - 1
            If LCase(.Offset(i, 0).Value) = lstName And LCase(.Offset(i, 1).Value) = fstName Then
                match = True
                repData(7) = i
                
                'Fill array with all rep data with match
                For j = 0 To 6
                    repData(j) = .Offset(i, j).Value
                Next j
            End If
        Next i
    End With
    
    If match Then
        searchRep = repData
    Else
        searchRep = Array("FALSE", "FLASE")
    End If
End Function


'''''''''''''''''''''''''' frmQ3b ''''''''''''''''''''''''''

Option Explicit

Private Sub UserForm_Initialize()
    'Text Box Default Values
    Me.txtLastName.Text = repData(0)
    Me.txtFirstName.Text = repData(1)
    Me.txtAge.Text = repData(5)
    Me.txtYrsExp.Text = repData(4)
    
    'Rating options default value
    Select Case repData(6)
        Case "Mediocre": Me.optRatingMediocre.Value = True
        Case "Good": Me.optRatingGood.Value = True
        Case "Outstanding": Me.optRatingOutstanding.Value = True
    End Select
    
    'Gender options default value
    Select Case repData(2)
        Case "Male": Me.optGenderMale.Value = True
        Case "Female": Me.optGenderFemale.Value = True
    End Select
    'Region options default value
    Select Case repData(3)
        Case "East": Me.optRegionEast = True
        Case "Midwest": Me.optRegionMidwest = True
        Case "Northeast": Me.optRegionNortheast = True
        Case "South": Me.optRegionSouth = True
        Case "West": Me.optRegionWest = True
    End Select
End Sub

Private Sub cmdCancelb_Click()
    Unload Me
    frmQ3a.Show
End Sub

Private Sub cmdOkb_Click()
    Dim i As Long
    Dim newData(0 To 6) As String
    Dim rOffSet As String: rOffSet = repData(7)
    Dim Control As Variant
    
    newData(0) = Me.txtLastName.Value
    newData(1) = Me.txtFirstName.Value
    newData(4) = Me.txtYrsExp.Value
    newData(5) = Me.txtAge.Value
    
    'Set new Rating
    Dim rating As Control
    For Each rating In frameRating.Controls
        If TypeOf rating Is MSForms.OptionButton And rating.Value Then
                newData(6) = rating.Caption
        End If
    Next
    
    'Set new gender
    Dim gender As Control
    For Each gender In frameGender.Controls
        If TypeOf gender Is MSForms.OptionButton And gender.Value Then
                newData(2) = gender.Caption
        End If
    Next
    
    'Set new Region
    Dim region As Control
    For Each region In frameRegion.Controls
        If TypeOf region Is MSForms.OptionButton And region.Value Then
                newData(3) = region.Caption
        End If
    Next
    
    'Set new values in spreadsheet
    With Range("A4")
        For i = 0 To 6
            .Offset(rOffSet, i).Value = newData(i)
        Next i
    End With
    
    Unload Me
    frmQ3a.Show
End Sub
