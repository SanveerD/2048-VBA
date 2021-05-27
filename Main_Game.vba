'Main game
Dim numMoves As Integer

Sub countMoves()
Dim ws As Worksheet
Set ws = Sheet7

If ws Is ActiveSheet Then
    numMoves = numMoves + 1
    Sheet7.Range("G10") = numMoves
Else
    End
End If

End Sub

Sub reset()
Sheet7.Activate
Range("E5:H8").Clear

Call randomnum
Call FormatNumbers

If Worksheets("Helper1").Range("C22") = "Yes" Then
    Call Block
End If

Call resetButtons

If Worksheets("Helper1").Range("numFreeTiles") = 0 Then
    Worksheets("Board").Buttons("Button 10").Font.Color = 8421504
    Worksheets("Board").Buttons("Button 10").OnAction = "buttonDisabled"
End If

If Worksheets("Helper1").Range("D19") = 0 Then
    Worksheets("Board").Buttons("Button 12").Font.Color = 8421504
    Worksheets("Board").Buttons("Button 12").OnAction = "buttonDisabled"
End If

numMoves = -1
Call countMoves

Worksheets("Board").Range("N6") = Worksheets("Helper1").Range("numFreeTiles") - freeTileCount
Worksheets("Board").Range("N7") = Worksheets("Helper1").Range("D19") - numUndos

ReDim gameboardarray(16)

End Sub

Sub FormatNumbers()

Sheet7.Range("E5:H8").BorderAround _
 ColorIndex:=1, Weight:=xlThick
With Sheet7.Range("E5:H8")
    .Interior.ColorIndex = 36
    .Font.Size = 24
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Borders(xlInsideVertical).LineStyle = xlContinuous
End With

Dim currCell As Range, firstCell As Range
Set firstCell = Sheet7.Range("E5")

For j = 0 To 3
    For i = 0 To 3
        Set currCell = firstCell.Offset(j, i)
            
            currCell.VerticalAlignment = xlCenter
            currCell.HorizontalAlignment = xlCenter

            If currCell = "" Then
                currCell.Interior.ColorIndex = 36
            ElseIf currCell.Value = 1 Then
                currCell.Interior.ColorIndex = 3
            ElseIf currCell.Value = 2 Then
                currCell.Interior.ColorIndex = 20
            ElseIf currCell.Value = 4 Then
                currCell.Interior.ColorIndex = 44
            ElseIf currCell.Value = 8 Then
                currCell.Interior.ColorIndex = 46
            ElseIf currCell.Value = 16 Then
                currCell.Interior.ColorIndex = 22
            ElseIf currCell.Value = 32 Then
                currCell.Interior.ColorIndex = 27
            ElseIf currCell.Value = 64 Then
                currCell.Interior.ColorIndex = 43
            ElseIf currCell.Value = 128 Then
                currCell.Interior.ColorIndex = 50
            ElseIf currCell.Value = 256 Then
                currCell.Interior.ColorIndex = 42
            ElseIf currCell.Value = 512 Then
                currCell.Interior.ColorIndex = 53
            ElseIf currCell.Value = 1024 Then
                currCell.Interior.ColorIndex = 24
            ElseIf currCell.Value = 2048 Then
                currCell.Interior.ColorIndex = 17
            End If
    
    Next i
Next j

End Sub

Sub randomnum()

Dim Board, tile As Range
Dim boardSize, numEmptySpace, scatterCell As Integer

Set Board = Sheet7.Range("E5:H8")
boardSize = 16
numEmptySpace = boardSize - WorksheetFunction.CountA(Board)

If numEmptySpace = 0 Then
'if no empty spaces, user has lost and goes to lose screen
    Call LoseScreen
    Exit Sub
Else
   scatterCell = WorksheetFunction.RandBetween(1, numEmptySpace)
   
For Each tile In Board
    If tile.Value = "" Then
        scatterCell = scatterCell - 1
            If scatterCell = 0 Then
                    tile.Value = 2
            End If
    End If

Next tile
End If

End Sub

Sub all_right()

Call countMoves
Call undomove

Dim gameboard_range As Range
Dim dummyme As Integer

Set gameboard_range = Sheet7.Range("E5:H8")

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

Dim max As Integer
max = Application.WorksheetFunction.max(gameboard_range)
If max = 2048 Then
    Call WinScreen
    Exit Sub
Else
    Call randomnum
End If
Call FormatNumbers

Sheet4.Range("E19").Value = 1

End Sub

Sub all_left()

Call countMoves
Call undomove

Dim gameboard_range As Range
Dim dummyme As Integer

Set gameboard_range = Sheet7.Range("E5:H8")

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
Dim max As Integer
max = Application.WorksheetFunction.max(gameboard_range)
If max = 2048 Then
    Call WinScreen
    Exit Sub
Else
    Call randomnum
End If
Call FormatNumbers

Sheet4.Range("E19").Value = 1

End Sub

Sub all_down()

Call countMoves
Call undomove

Dim gameboard_range As Range
Dim dummyme As Integer

Set gameboard_range = Sheet7.Range("E5:H8")

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
Dim max As Integer
max = Application.WorksheetFunction.max(gameboard_range)
If max = 2048 Then
    Call WinScreen
    Exit Sub
Else
    Call randomnum
End If
Call FormatNumbers

Sheet4.Range("E19").Value = 1

End Sub

Sub all_up()
Call undomove

Call countMoves

Dim gameboard_range As Range
Dim dummyme As Integer

Set gameboard_range = Sheet7.Range("E5:H8")

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

Dim max As Integer
max = Application.WorksheetFunction.max(gameboard_range)
If max = 2048 Then
    Call WinScreen
    Exit Sub
Else
    Call randomnum
End If
Call FormatNumbers

Sheet4.Range("E19").Value = 1

End Sub

Sub WinScreen()
Sheet2.Activate
ActiveWindow.DisplayGridlines = False

End Sub

Sub LoseScreen()
Sheet3.Activate
ActiveWindow.DisplayGridlines = False
End Sub

