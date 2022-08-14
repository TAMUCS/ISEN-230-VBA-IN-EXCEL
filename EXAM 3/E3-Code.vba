''''''''''''''''''''''''''''''''' MODULE 1 '''''''''''''''''''''''''''''''''
Option Explicit

Public brndNms As Variant
Public selectedBrands As Variant
Public usrName As String

Sub Prob3()
    frmStart.Show
End Sub

Sub prob2()

wsOptimize.Activate 'Problem2's Worksheet

SolverReset

SolverOk SetCell:=Range("E10"), MaxMinVal:=2, ByChange:=Range("B10:E10")

SolverAdd CellRef:=Range("E5"), Relation:=1, FormulaText:=Range("G5")
SolverAdd CellRef:=Range("E6"), Relation:=1, FormulaText:=Range("G6")
SolverAdd CellRef:=Range("E7"), Relation:=3, FormulaText:=Range("G7")
SolverAdd CellRef:=Range("D10"), Relation:=3, FormulaText:=3
SolverAdd CellRef:=Range("B10:D10"), Relation:=4 'This does NOT force ints?

SolverOptions AssumeLinear:=True, AssumeNonNeg:=True

SolverSolve UserFinish:=True

End Sub

''''''''''''''''''''''''''''''''' frmStart '''''''''''''''''''''''''''''''''

Option Explicit

'[TESTED; WORKING]: Extracts first name from string
Function fstName(strTmp As String) As String
    Dim chr As String, name As String, i As Integer
    
    For i = 1 To Len(strTmp)
        chr = Mid(strTmp, i, 1)
        If chr = " " Then
            fstName = Trim(name)
            Exit For
        Else
            name = name & chr
        End If
    Next i
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

'd
Private Sub cmdOK_Click()
    Dim sel As Boolean: sel = False
    Dim i As Integer
    ReDim selectedBrands(0 To 8) As Boolean
    
    selectedBrands(0) = False
    selectedBrands(1) = False
    selectedBrands(2) = False
    selectedBrands(3) = False
    selectedBrands(4) = False
    selectedBrands(5) = False
    selectedBrands(6) = False
    selectedBrands(7) = False
    selectedBrands(8) = False
    
    For i = 0 To 8
        If Me.lstBrands.Selected(i) Then
            sel = True
            selectedBrands(i) = True
        End If
    Next i

    If Not sel Then
        MsgBox "Please make valid selections before you can proceed", vbInformation, "My program"
        Exit Sub
    Else
        Unload Me
        frmSales.Show
    End If
End Sub

'Saves user-entered name as string
Private Sub txtName_AfterUpdate()
    usrName = Me.txtName.Value 'b
End Sub

'This one explains itself p well
Private Sub UserForm_Initialize()
    ReDim brndNms(0 To 8) As String
    Dim i As Integer
    Dim cell As Range, cutName As String
    
    usrName = "N/A"
    
    'Used shortcut to load up items
    brndNms(0) = "Acer"
    brndNms(1) = "Alienware"
    brndNms(2) = "ASUS"
    brndNms(3) = "Dell"
    brndNms(4) = "HP"
    brndNms(5) = "Lenovo"
    brndNms(6) = "LG"
    brndNms(7) = "Samsung"
    brndNms(8) = "MSI"

    'b
    Me.lstBrands.MultiSelect = fmMultiSelectMulti
    For i = 0 To 8
        Me.lstBrands.AddItem brndNms(i)
    Next i
End Sub


''''''''''''''''''''''''''''''''' frmSales '''''''''''''''''''''''''''''''''
Option Explicit

'[TESTED; WORKING]: Extracts first name from string
Function fstName(strTmp As String) As String
    Dim chr As String, name As String, i As Integer
    
    For i = 1 To Len(strTmp)
        chr = Mid(strTmp, i, 1)
        If chr = " " Then
            fstName = Trim(name)
            Exit For
        Else
            name = name & chr
        End If
    Next i
End Function

'(i)
Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    '(ii)
    If Not (optOne Or optMulti) Then
        MsgBox "Please make valid selections before you can proceed", vbInformation, "My Program"
        Exit Sub
    End If
    
    '(iii)
    If optMulti Then
        Unload Me
        MsgBox "Hello " & usrName & "! We're sorry we cannot handle this transaction at this time", vbCritical, "My Program"
        Exit Sub
    End If
    
    Dim usrSelection As String: usrSelection = Me.lstItems.Value
    Dim cell As Range, price As String, retailer As String
    
    For Each cell In Range(Range("A1").Offset(1, 0), Range("A1").Offset(1, 0).End(xlDown))
        If cell.Value = usrSelection Then
            price = cell.Offset(0, 1).Value
            retailer = cell.Offset(0, 2).Value
        End If
    Next cell
    
    Unload Me
    MsgBox "You have chosen " & usrSelection & " which costs $" & price & " and can be purchased from " & retailer, vbInformation, "My Program"
End Sub

Private Sub optMulti_Click()
    Me.lstItems.Visible = True
End Sub

Private Sub optOne_Click()
    Me.lstItems.Visible = True
End Sub

Private Sub UserForm_Initialize()
    Me.lstItems.MultiSelect = fmMultiSelectSingle '(iv)
    Me.lstItems.Visible = False
    Dim cell As Range, i As Integer, tmp As Integer
    
    For Each cell In Range(Range("A1").Offset(1, 0), Range("A1").Offset(1, 0).End(xlDown))
        For i = 0 To 8
            If selectedBrands(i) = True Then
                If brndNms(i) = fstName(cell.Value) Then
                    Me.lstItems.AddItem cell.Value
                End If
            End If
        Next i
    Next cell
End Sub
