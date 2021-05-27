'Hint and autosolve
Option Base 1

Sub hint()

Dim gameboard_range As Range
Dim counter, dummyme, multiple, power_factor As Integer
Dim gameboard_array() As Variant
Dim best_move As String

Set gameboard_range = Sheet7.Range("E5:H8")

multiple = 2

'Take initial range and store into array
counter = 0

ReDim gameboard_array(16)

For i = 1 To 4
    For j = 1 To 4
        If counter <= 16 Then
            counter = counter + 1
            ReDim Preserve gameboard_array(counter)
        Else
        End If
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_array(counter) = gameboard_range.Cells(i, j).Value
        Else
            gameboard_array(counter) = 0
        End If
    Next j
Next i

'Take all_right move and move right_range
For i = 1 To 4 Step 1
    For j = 4 To 1 Step -1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i, j - 1) = 1 Then
            dummyme = 1
        ElseIf gameboard_range.Cells(i, j - 1) > 1 And j - 1 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 1)
            gameboard_range.Cells(i, j - 1) = ""
        ElseIf gameboard_range.Cells(i, j - 2) = 1 Then
            dummyme = 1
        ElseIf gameboard_range.Cells(i, j - 2) > 1 And j - 2 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 2)
            gameboard_range.Cells(i, j - 2) = ""
         ElseIf gameboard_range.Cells(i, j - 3) = 1 Then
            dummyme = 1
         ElseIf gameboard_range.Cells(i, j - 3) > 1 And j - 3 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 3)
           gameboard_range.Cells(i, j - 3) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next j
Next i

For i = 1 To 4 Step 1
    For j = 4 To 1 Step -1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True And gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 1) Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j) * 2
            gameboard_range.Cells(i, j - 1) = ""
        End If
    Next j
Next i

For i = 1 To 4 Step 1
    For j = 4 To 1 Step -1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i, j - 1) = 1 Then
            dummyme = 1
        ElseIf gameboard_range.Cells(i, j - 1) > 1 And j - 1 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 1)
            gameboard_range.Cells(i, j - 1) = ""
        ElseIf gameboard_range.Cells(i, j - 2) = 1 Then
            dummyme = 1
        ElseIf gameboard_range.Cells(i, j - 2) > 1 And j - 2 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 2)
            gameboard_range.Cells(i, j - 2) = ""
         ElseIf gameboard_range.Cells(i, j - 3) = 1 Then
            dummyme = 1
         ElseIf gameboard_range.Cells(i, j - 3) > 1 And j - 3 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j - 3)
           gameboard_range.Cells(i, j - 3) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next j
Next i

For i = 1 To 11
    power_factor = Application.WorksheetFunction.Power(multiple, i)
            Sheet4.Cells(14, 3).Offset(-i, 0).Value = Application.WorksheetFunction.CountIf(gameboard_range, power_factor)
Next i

'Replace worksheet values with intermediate inputs
counter = 0

For i = 1 To 4
    For j = 1 To 4
        counter = counter + 1
        If gameboard_array(counter) = 0 Then
            gameboard_range.Cells(i, j) = ""
        Else
        gameboard_range.Cells(i, j) = gameboard_array(counter)
        End If
    Next j
Next i

'Take all_left move and move right_range
For i = 1 To 4 Step 1
    For j = 1 To 4 Step 1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i, j + 1) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j + 1)) = True And j + 1 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 1)
            gameboard_range.Cells(i, j + 1) = ""
        ElseIf gameboard_range.Cells(i, j + 2) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j + 2)) = True And j + 2 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 2)
            gameboard_range.Cells(i, j + 2) = ""
        ElseIf gameboard_range.Cells(i, j + 3) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j + 3)) = True And j + 3 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 3)
            gameboard_range.Cells(i, j + 3) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next j
Next i

For i = 1 To 4 Step 1
    For j = 1 To 4 Step 1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True And gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 1) Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j) * 2
            gameboard_range.Cells(i, j + 1) = ""
        End If
    Next j
Next i

For i = 1 To 4 Step 1
    For j = 1 To 4 Step 1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i, j + 1) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j + 1)) = True And j + 1 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 1)
            gameboard_range.Cells(i, j + 1) = ""
        ElseIf gameboard_range.Cells(i, j + 2) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j + 2)) = True And j + 2 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 2)
            gameboard_range.Cells(i, j + 2) = ""
        ElseIf gameboard_range.Cells(i, j + 3) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j + 3)) = True And j + 3 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j + 3)
            gameboard_range.Cells(i, j + 3) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next j
Next i

For i = 1 To 11
    power_factor = Application.WorksheetFunction.Power(multiple, i)
            Sheet4.Cells(14, 4).Offset(-i, 0).Value = Application.WorksheetFunction.CountIf(gameboard_range, power_factor)
Next i

'Replace worksheet values with intermediate inputs
counter = 0

For i = 1 To 4
    For j = 1 To 4
        counter = counter + 1
        If gameboard_array(counter) = 0 Then
            gameboard_range.Cells(i, j) = ""
        Else
        gameboard_range.Cells(i, j) = gameboard_array(counter)
        End If
    Next j
Next i

'Take all_down move and move right_range
For j = 1 To 4 Step 1
    For i = 4 To 1 Step -1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i - 1, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i - 1, j)) = True And i - 1 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 1, j)
            gameboard_range.Cells(i - 1, j) = ""
        ElseIf gameboard_range.Cells(i - 2, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i - 2, j)) = True And i - 2 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 2, j)
            gameboard_range.Cells(i - 2, j) = ""
        ElseIf gameboard_range.Cells(i - 3, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i - 3, j)) = True And i - 3 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 3, j)
            gameboard_range.Cells(i - 3, j) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next i
Next j

For j = 1 To 4 Step 1
    For i = 4 To 1 Step -1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True And gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 1, j) Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j) * 2
            gameboard_range.Cells(i - 1, j) = ""
        End If
    Next i
Next j

For j = 1 To 4 Step 1
    For i = 4 To 1 Step -1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i - 1, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i - 1, j)) = True And i - 1 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 1, j)
            gameboard_range.Cells(i - 1, j) = ""
        ElseIf gameboard_range.Cells(i - 2, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i - 2, j)) = True And i - 2 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 2, j)
            gameboard_range.Cells(i - 2, j) = ""
        ElseIf gameboard_range.Cells(i - 3, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i - 3, j)) = True And i - 3 > 0 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i - 3, j)
            gameboard_range.Cells(i - 3, j) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next i
Next j

For i = 1 To 11
    power_factor = Application.WorksheetFunction.Power(multiple, i)
            Sheet4.Cells(14, 5).Offset(-i, 0).Value = Application.WorksheetFunction.CountIf(gameboard_range, power_factor)
Next i

'Replace worksheet values with intermediate inputs
counter = 0

For i = 1 To 4
    For j = 1 To 4
        counter = counter + 1
        If gameboard_array(counter) = 0 Then
            gameboard_range.Cells(i, j) = ""
        Else
        gameboard_range.Cells(i, j) = gameboard_array(counter)
        End If
    Next j
Next i

'Take all_up move and move right_range
For j = 1 To 4 Step 1
    For i = 1 To 4 Step 1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i + 1, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i + 1, j)) = True And i + 1 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 1, j)
            gameboard_range.Cells(i + 1, j) = ""
        ElseIf gameboard_range.Cells(i + 2, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i + 2, j)) = True And i + 2 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 2, j)
            gameboard_range.Cells(i + 2, j) = ""
        ElseIf gameboard_range.Cells(i + 3, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i + 3, j)) = True And i + 3 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 3, j)
            gameboard_range.Cells(i + 3, j) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next i
Next j

For j = 1 To 4 Step 1
    For i = 1 To 4 Step 1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True And gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 1, j) Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j) * 2
            gameboard_range.Cells(i + 1, j) = ""
        End If
    Next i
Next j

For j = 1 To 4 Step 1
    For i = 1 To 4 Step 1
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i, j)
        ElseIf gameboard_range.Cells(i + 1, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i + 1, j)) = True And i + 1 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 1, j)
            gameboard_range.Cells(i + 1, j) = ""
        ElseIf gameboard_range.Cells(i + 2, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i + 2, j)) = True And i + 2 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 2, j)
            gameboard_range.Cells(i + 2, j) = ""
        ElseIf gameboard_range.Cells(i + 3, j) = 1 Then
            dummyme = 1
        ElseIf Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i + 3, j)) = True And i + 3 < 5 Then
            gameboard_range.Cells(i, j) = gameboard_range.Cells(i + 3, j)
            gameboard_range.Cells(i + 3, j) = ""
        Else
            gameboard_range.Cells(i, j) = ""
        End If
    Next i
Next j

For i = 1 To 11
    power_factor = Application.WorksheetFunction.Power(multiple, i)
            Sheet4.Cells(14, 6).Offset(-i, 0).Value = Application.WorksheetFunction.CountIf(gameboard_range, power_factor)
Next i

'Replace worksheet values with intermediate inputs
counter = 0

For i = 1 To 4
    For j = 1 To 4
        counter = counter + 1
        If gameboard_array(counter) = 0 Then
            gameboard_range.Cells(i, j) = ""
        Else
        gameboard_range.Cells(i, j) = gameboard_array(counter)
        End If
    Next j
Next i

Call FormatNumbers

If Sheet4.Cells(3, 15).Value = 1 Then
    best_move = "right"
ElseIf Sheet4.Cells(3, 15).Value = 2 Then
    best_move = "left"
ElseIf Sheet4.Cells(3, 15).Value = 3 Then
    best_move = "down"
Else
    best_move = "up"
End If

MsgBox ("Try moving " & best_move & "!")

End Sub
