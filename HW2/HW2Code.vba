Option Explicit

Function ppuRng(qty As Integer) As Integer
    Dim col As String: col = 0
    If qty >= 50 Then col = 1
    If qty >= 100 Then col = 2
    
    ppuRng = col
End Function

Sub Receipt()
    Dim StartTime As Double
    Dim SecondsElapsed As Double
    StartTime = Timer

    '3) Change the codename of the Sheet1 to wsReceipt and set this as the active sheet
    Dim sht As Object
    Set sht = ActiveWorkbook.VBProject.VBComponents(ActiveWorkbook.Sheets(1).CodeName)
    sht.Name = "wsReceipt"
    wsReceipt.Activate
    
    '4) Change the tab name to "Billing"
    wsReceipt.Name = "Billing"
    
    '5) Clear columns I-K (Contents & formatting, if any)
    wsReceipt.Columns("I:K").Clear

    '6) Ask user for name of company w/ atleast 2 words (store as str) _
    '   Dialog Box title must be "Query" _
    '   Remove all leading and trailing spaces _
    '   Find the number of chars in string after removing the _
    '   leading/trailing spaces (store as int)
    Dim strInp As String: strInp = InputBox("Enter comapny name. Must be atleast two words.", "Query")
    Dim strClean As String: strClean = Trim(strInp)
    Dim numChars As Integer: numChars = Len(strClean)
    
    '7) Create a password using the first 3 characters of the reverse of the first word _
    '   in the company name in lower case concatenated with the number of characters you just found _
    '   and display it in the cell B2. Feel free to create new variables if you would like.
    Dim names() As String, password As String
    names = Split(strClean)
    password = Left(LCase(StrReverse(names(0))), 3) & CStr(numChars)
    Range("B2").Value = password
    
    '8) Ask the user to enter the passcode displayed in cell B2 and store this to a string _
    '   variable called strCheck. _
    '   Dialog Box title must be "Query". _
    '   If the password entered by the user exactly matched with the value show in cell B2, then _
    '   do the following parts from 9 to 14, else show a dialog box that says "Wrong passkey entered. _
    '   Program will now terminate." and the program should terminate.
    Dim strCheck As String: strCheck = InputBox("Enter password displayed in cell B2", "Query")
    
    If Not (strCheck = password) Then
        'Program Terminates
        MsgBox "Wrong passkey entered. Program will now terminate", vbCritical & vbOKOnly, "Query"
        Exit Sub
    End If
        
    '9) Insert title "Receipt" in I3 (bold, underlined & centered) using With..End With
    With Range("I3") 
        With .Font
            .Bold = True
            .Underline = True
        End With
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlCenter
        .Value = "Receipt"
    End With
    
    '10) Merge I3:K3. Apply border line style xlDash.
    With Range("I3:K3")
        .Merge
        .Borders.LineStyle = xlDash
    End With
        
    '11) In cell I5, insert the text "Item Name". In cell J5, insert the text "# of Units". _
    '    In cell K5, insert the text "Total Cost". (ALL TEXT BOLDED)
    Dim vals(2) As String, i As Integer
    vals(0) = "Item Name"
    vals(1) = "# of Units"
    vals(2) = "Total Cost"
        
    For i = 9 To 11
        With cells(5, i)
            .Value = vals(i - 9)
            .Font.Bold = True
        End With
    Next i
    
    '12) Starting at I6, one must copy over the names of all items listed in _
    '   column B from the cell B6 onwards using the Copy method. Your line of code must be such that _
    '   if the user enters any new names below B8, that would also get copied over.
    Dim row As Long, src As Range, dst As Range
    row = 6
    Do While True
        Set src = Range("B" & row)
        Set dst = Range("I" & row)
        
        If Not (src.Value = "") Then
            src.Copy dst
            row = row + 1
        Else
            Exit Do
        End If
    Loop
    
    '13) ask user how many Machine [1-3] they would like to purchase (int);
    Dim m1 As Integer: m1 = CInt(InputBox("Enter quanitiy of machine 1 to purchase.", "Query"))
    Range("J6").Value = m1
    Dim m2 As Integer: m2 = CInt(InputBox("Enter quanitiy of machine 2 to purchase.", "Query"))
    Range("J7").Value = m2
    Dim m3 As Integer: m3 = CInt(InputBox("Enter quanitiy of machine 3 to purchase.", "Query"))
    Range("J8").Value = m3
    
    '14)
    Range("K6:K8").NumberFormat = "#0.00"
    Range("K11").NumberFormat = "$#0.00"
    Range("K13").NumberFormat = "$#0.00"
    Range("K12").NumberFormat = "#0.0#%"
    
    Dim numRows As Long: numRows = Range("B6").End(xlDown).row - 5
    Dim dblSubtotal As Double
    Dim dblDiscount As Double
    Dim percentStr As String
    
    With Range("I5")
        With .Offset(numRows + 3, 1)
            .Value = "Subtotal"
            .Font.Bold = True
        End With
        
        With .Offset(numRows + 4, 1)
            .Value = "Discount"
            .Font.Bold = True
        End With
        
        With .Offset(numRows + 5, 1)
            .Value = "Total"
            .Font.Bold = True
        End With
    
        .Offset(1, 2).Value = .Offset(1, (-4 - ppuRng(m1))).Value * .Offset(1, 1).Value
        .Offset(2, 2).Value = .Offset(2, (-4 - ppuRng(m2))).Value * .Offset(2, 1).Value
        .Offset(3, 2).Value = .Offset(3, (-4 - ppuRng(m3))).Value * .Offset(3, 1).Value
        
        dblSubtotal = WorksheetFunction.Sum(Range("K6:K8"))
        
        .Offset(numRows + 3, 2).Value = dblSubtotal
        
        percentStr = .Offset(-3, -4).Value
        dblDiscount = CDbl(Left(percentStr, Len(percentStr)))
        
        If dblSubtotal >= 65000 Then
            .Offset(numRows + 4, 2).Value = dblDiscount
        Else
            .Offset(numRows + 4, 2).Value = 0
        End If
        
        .Offset(numRows + 5, 2).Value = (1 - dblDiscount) * dblSubtotal
    End With
    
    
    SecondsElapsed = Round(Timer - StartTime, 2)
    MsgBox "This code ran successfully in " & SecondsElapsed & " seconds", vbInformation

End Sub