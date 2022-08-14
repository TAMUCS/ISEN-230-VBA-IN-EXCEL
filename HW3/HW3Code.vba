Option Explicit

''''''''''''''''''''''''''''''''''''wsResults CODE''''''''''''''''''''''''''''''''''''
Private Sub Worksheet_Activate()
    Call S3ToOL
    Call calculateOL
End Sub

''''''''''''''''''''''''''''''''''''wsData CODE''''''''''''''''''''''''''''''''''''
Private Sub Worksheet_Activate()
    Call S3ToNRS
End Sub


''''''''''''''''''''''''''''''''''''HW QUESTION 4''''''''''''''''''''''''''''''''''''
Global StartTime As Double
Global minElapsed As Double

Sub Auto_Open()
    StartTime = Timer
    MsgBox "[ATTENTION]: OPEN README TAB BEFORE RUNNING ANY CODE", vbCritical, "IMPORTANT"
    MsgBox "[ATTENTION]: HW QUESTION 3 MAY TAKE UP TO 1 MIN TO COMPLETE WHEN SWITCHING TO wsResults TAB", vbInformation, "IMPORTANT"
End Sub

Sub Auto_Close()
    minElapsed = Round((Timer - StartTime) / 60, 6)
    MsgBox "This code ran successfully in " & minElapsed & " Minutes", vbInformation
End Sub

''''''''''''''''''''''''''''''''''''HW QUESTION 1''''''''''''''''''''''''''''''''''''
'Do Until Bottom
Function dub() As String
    Dim inStrDUB As String
    Do
        inStrDUB = InputBox("Enter a product code (Do Until Bottom)")
    Loop Until inStrDUB <> ""
    dub = inStrDUB
End Function
'Do Until Up
Function duu() As String
    Dim inStrDUU As String
    Do Until inStrDUU <> ""
        inStrDUU = InputBox("Enter a product code (Do Until Up)")
    Loop
    duu = inStrDUU
End Function
'Do While Up
Function dwu() As String
    Dim inStrDWU As String
    Do While inStrDWU = ""
        inStrDWU = InputBox("Enter a product code (Do While Up)")
    Loop
    dwu = inStrDWU
End Function
'Do While Up
Function dwb() As String
    Dim inStrDWB As String
    Do
        inStrDWB = InputBox("Enter a product code (Do While Bottom)")
    Loop While inStrDWB = ""
    dwb = inStrDWB
End Function

Sub HWProb1()
    Dim dubStr As String, duuStr As String, dwbStr As String, dwuStr As String
    'Do Until Bottom
    dubStr = dub()
    MsgBox "You Entered: " & dubStr
    'Do Until Up
    duuStr = duu()
    MsgBox "You Entered: " & duuStr
    'Do While Up
    dwuStr = dwu()
    MsgBox "You Entered: " & dwuStr
    'Do While Bottom
    dwbStr = dwb()
    MsgBox "You Entered: " & dwbStr
End Sub

''''''''''''''''''''''''''''''''''''HW QUESTION 2''''''''''''''''''''''''''''''''''''
Function ascSum(ascChr As String) As Integer
    Dim ascCode As String: ascCode = CStr(Asc(ascChr))
    Dim i As Integer, tmpSum As Integer
    
    For i = 1 To Len(ascCode)
        tmpSum = tmpSum + CInt(Mid(ascCode, i, 1))
    Next i
    ascSum = tmpSum
End Function

Sub CrypticKey()
    CrypticData.Activate
    Dim strResp As String
    
    '(i)
    Do Until strResp <> ""
        strResp = InputBox("Please Enter a Message to Encode")
    Loop
    
    '(ii)
    Dim dict As New Scripting.Dictionary
    dict.CompareMode = vbBinaryCompare
    Dim numRows As Long: numRows = Range("A3").End(xlDown).Row - 3
    Dim oldChrs() As String, newChrs() As String
    ReDim oldChrs(1 To numRows)
    ReDim newChrs(1 To numRows)
    Dim i As Long
    
    With Range("A3")
        For i = 1 To numRows
            oldChrs(i) = .Offset(i, 0).Value
            newChrs(i) = .Offset(i, 1).Value
            dict(oldChrs(i)) = newChrs(i)
        Next i
    End With
    
    '(iii)
    Dim j As Long
    Dim key As String: key = ""
    Dim tmp As String: tmp = ""
    Dim catStr As String: catStr = ""
    
    For j = 1 To Len(strResp)
        tmp = CStr(Mid(strResp, j, 1))
        
        If dict.Exists(tmp) Then
            catStr = dict(tmp)
        ElseIf tmp = " " Then
            catStr = "-"
        Else
            catStr = " " & CStr(ascSum(tmp)) & " "
        End If
        
        key = key & catStr
    Next j
    
    Range("G7").Value = "Key:"
    Range("H7").Value = key
    MsgBox "Your encoded message is:" & vbNewLine & key, vbOKOnly, "Encoded"
End Sub
''''''''''''''''''''''''''''''''''''HW QUESTION 3''''''''''''''''''''''''''''''''''''

'Rename Sheet3 to New Results Sheet
Sub S3ToNRS()
    wsResults.Name = "New Results Sheet"
End Sub

'Rename Sheet3 to Overdue List
Sub S3ToOL()
    wsResults.Name = "Overdue List"
End Sub

'Clears Results Tab Prior to Calculations & Population
Sub clrResultTab()
    With wsResults
        .Columns(1).ClearContents
        .Columns(2).ClearContents
        .Range("A1").Value = "Customers with Payments that are Overdue"
        .Range("A3").Value = "Customer ID"
        .Range("B3").Value = "Pending Payment"
    End With
End Sub

'Calculates, populates Sheet3
Sub calculateOL()
    Call clrResultTab
    Dim numRows As Long: numRows = Range("A3").End(xlDown).Row - 3
    Dim dict As New Scripting.Dictionary
    Dim cID As String, Remainder As Integer, maxCust(1 To 2) As Integer
    Dim i As Long
    dict.CompareMode = vbBinaryCompare
    maxCust(1) = 0
    maxCust(2) = 0
    With wsData.Range("A3")
        For i = 1 To numRows
            cID = .Offset(i, 0).Value 'CID
            Remainder = CInt(.Offset(i, 1).Value) - CInt(.Offset(i, 2).Value) 'Remainder
            
            'Over $1500 Check
            If Remainder > 1500 Then
                dict(cID) = Format(CStr(Remainder), "$#,##0")
                
                ' Keep Track of Max Owed Data
                If Remainder > maxCust(2) Then
                    maxCust(1) = CInt(cID)
                    maxCust(2) = Remainder
                End If
            End If
        Next i
    End With
    
    'Populate wsResults with resulting calculations
    Dim key As Variant
    Dim j As Long: j = 0
    
    With wsResults.Range("A3")
        For j = 0 To dict.Count - 1
            .Offset(j + 1, 0).Value = dict.Keys(j)
            .Offset(j + 1, 1).Value = dict.Items(j)
        Next j
    End With
    
    MsgBox "The customer with the highest payment due has the ID: " & CStr(maxCust(1)) & vbNewLine & _
    "with " & Format(CStr(maxCust(2)), "$#,##0") & " in pending payments."
End Sub
