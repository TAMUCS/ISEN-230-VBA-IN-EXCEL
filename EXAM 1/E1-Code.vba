Option Explicit

Sub Prob3()
    Dim diff As Single
    'a
    Sheet3.Activate
    'b
    Sheet3.Name = "Q3"
    'c
    diff = WorksheetFunction.Sum(Range("C2:C5")) - WorksheetFunction.Sum(wsProb4.Range("C2:C5"))
    'd
    Dim strC1 As String, strC2 As String, strCombo As String
    
    strC1 = Sheet3.Range("A1").Value
    strC2 = Sheet3.Range("A2").Value
    strCombo = strC1 & strC2
    
    Sheet3.Range(strCombo).Value = diff
    'e
    
    With Sheet3
        With Range(strCombo)
            .Font.Size = 24
            .Font.Name = "Arial"
            .Font.Italic = True
            .Interior.Color = vbYellow
            .Value = Format(diff, "00.##")
        End With
    End With
End Sub

