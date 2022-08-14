Option Explicit

''''''''''''''''''''''''''''''''' Problem 2 (Module 1) '''''''''''''''''''''''''''''''''
'!!!!!!!!! The titles for all the MsgBox and InputBox must be given as “My Program”


Sub ProductClassifier()
    '1
    wsShopping.Name = "Problem2"
    
    '2
    wsShopping.Activate
    
    '3
    With Range("A1")
        Range(.Offset(1, 0), .Offset(1, 0).End(xlDown)).Sort key1:=Range("A1"), order1:=xlAscending
    End With
    
    '4
    'a
    Dim response As String
    Dim doPass As Boolean: doPass = False
    Dim mbButton
    Dim respInt As Integer
    
    Do
        response = InputBox("What would you like to purchase today? Enter either 1 or 2" & vbNewLine & "1 - Shoes. 2 - Pants", "My Program", vbOKCancel)
    
        'b
        If Not IsNumeric(response) Then
            MsgBox "Invalid Input." & vbNewLine & "The program will now end." & vbNewLine & "Thank you!", vbInformation, "My Program"
            Exit Sub
        End If
    
        'c
        If CInt(response) = 1 Or CInt(response) = 2 Then
            doPass = True
        Else
            mbButton = MsgBox("You must enter either 1 or 2. Would you like to try again?", vbYesNo, "My Program")
            If mbButton = vbNo Then
                MsgBox "Thank you!", vbOKOnly, "My Program"
                Exit Sub
            End If
        End If
        
        'd
        If doPass = True Then
            MsgBox "Thank you! We have recorded your input as " & CInt(response), vbOKOnly & vbInformation, "My Program"
            respInt = CInt(response)
        End If
    Loop Until doPass = True
    
    '5
    Dim strCat As String
    
    If respInt = 1 Then
        strCat = "Shoes"
    Else
        strCat = "Pants"
    End If
    
    '6
    Dim strType As String
    
    If strCat = "Shoes" Then
        strType = InputBox("Enter a product type among these: Boots, Sandals, Sneakers", "My Program")
    Else
        strType = InputBox("Enter a product type among these: Chinos, Denim, Pant, Shorts", "My Program")
    End If
    
    '8
    Dim strA As String: strA = CStr(GetCount(strCat, strType))
    
    '9
    If strA = "-1" Then
        MsgBox "No match found.", vbOKOnly & vbInformation, "My Program"
    Else
        MsgBox "There are " & strA & " units of " & strType & " available in the category of " & strCat, vbOKOnly & vbInformation, "My Program"
    End If
End Sub

'7
'(i)
Function GetCount(category As String, producttype As String) As Integer
    'The If statement I chose to use to match the category is the "non-existant" type. The reason being checking against the category _
    '    for this set of data is redundant & therefore unneccesary _
    'Even if we were asked to optimize our algorithm, it still wouldn't matter. _
    'With or without checking category using an if, this function would still have O(n) time-complexity, making each of the methods run in the _
    '    exact same discreet time using big-O analysis.
    Dim Product() As String
    Dim matches As Integer: matches = 0
    Dim i As Integer, j As Integer
    Dim nRows As Long: nRows = Range("A1").End(xlDown).Row - 1
    
    '(ii)
    With wsShopping.Range("A1")
        For i = 1 To nRows
            ReDim Preserve Product(1 To i) As String
            Product(i) = .Offset(i, 2).Value
        Next i
    End With
    For j = 1 To UBound(Product) - LBound(Product) + 1
        If LCase(Product(j)) = LCase(producttype) Then
            matches = matches + 1
        End If
    Next j
    
    If matches = 0 Then
        GetCount = -1
    Else
        GetCount = matches
    End If
End Function

''''''''''''''''''''''''''''''''' Problem 3 (Module 2)'''''''''''''''''''''''''''''''''
'!!!!!!!!! The titles for all the MsgBox and InputBox must be given as “My Program”
Sub Prob3()
    Dim strClean As String
    Dim password As String
    Dim noSpace As Boolean: noSpace = True
    Dim i As Long
    
    'a & b
    strClean = Trim(InputBox("Please neter a single word that is atleaset 7 characters long", "My Program"))
    
    'c
    For i = 1 To Len(strClean)
        If Mid(strClean, i, 1) = " " Then
            noSpace = False
        End If
    Next i
    
    If Len(strClean) <= 6 Or noSpace = False Then
        MsgBox "Your input does not meet the required criteria." & vbNewLine & "The program will now terminate", _
            vbOKOnly & vbInformation, "My Program"
        Exit Sub
    End If
    
    'd
    password = UCase(Mid(StrReverse(strClean), 6, 2)) & Asc(UCase(Mid(strClean, 2, 1))) & LCase(Mid(strClean, 3, 3))
    
    'e
    Dim msgResp As Integer
    
    msgResp = MsgBox("Would you like to display the password on the worksheet?", vbYesNoCancel, "My Program")

    If msgResp = vbCancel Then
        MsgBox "Thank you!", vbOKOnly, "My Program"
        Exit Sub
    ElseIf msgResp = vbYes Then
        With wsPswd.Range("A1")
            With .Offset(11, 2)
                .Interior.Color = vbGreen
                With .Borders
                    .LineStyle = xlDash
                    .Weight = xlThick
                End With
                With .Font
                    .Bold = True
                    .Underline = True
                End With
                .Value = password
            End With
        End With
    Else
        MsgBox password, vbOKOnly, "My Program"
    End If
End Sub

''''''''''''''''''''''''''''''''' Problem 4 (ThisWorkbook) '''''''''''''''''''''''''''''''''
Public UserNameTime As String
Public userName As String

Private Sub Workbook_Open()
    Dim StartTime As Double
    StartTime = Timer
    userName = InputBox("Please enter your name", "My Program")
    UserNameTime = Format((Timer - StartTime) / 60, "0.00")
End Sub

Private Sub Workbook_BeforeClose(Cancel As Boolean)
    MsgBox "Hello " & userName & "! It took " & UserNameTime & " minutes to obtain the name.", vbExclamation, "My Program"
End Sub


