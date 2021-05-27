'Undo Move
Option Base 1

Public gameboardarray() As Variant
Public undoCounter As Integer

Sub undomove()

Dim gameboard_range As Range
Dim counter As Integer

Set gameboard_range = Sheet7.Range("E5:H8")
counter = 0

ReDim gameboardarray(16)

For i = 1 To 4
    For j = 1 To 4
        If counter <= 16 Then
            counter = counter + 1
            ReDim Preserve gameboardarray(counter)
        Else
        End If
        If Application.WorksheetFunction.IsNumber(gameboard_range.Cells(i, j)) = True Then
            gameboardarray(counter) = gameboard_range.Cells(i, j).Value
        Else
            gameboardarray(counter) = 0
        End If
    Next j
Next i

End Sub

Sub undomove2()

Dim gameboard_range As Range
Dim counter As Integer
Dim numUndos As Integer
Dim checkUndos As Integer

numUndos = Worksheets("Helper1").Range("D19")
checkUndos = Worksheets("Helper1").Range("E19")

If checkUndos = 0 Then
    MsgBox ("Please make a move before using the undo function.")
    Exit Sub
End If

If undoCounter < numUndos Then

Set gameboard_range = Sheet7.Range("E5:H8")
counter = 0

For i = 1 To 4
    For j = 1 To 4
        counter = counter + 1
        If gameboardarray(counter) = 0 Then
            gameboard_range.Cells(i, j) = ""
        Else
        gameboard_range.Cells(i, j) = gameboardarray(counter)
        End If
    Next j
Next i

Call FormatNumbers
undoCounter = undoCounter + 1
End If

If undoCounter >= numUndos Then
    Worksheets("Board").Buttons("Button 12").Font.Color = 8421504
    Worksheets("Board").Buttons("Button 12").OnAction = "buttonDisabled"
End If

Worksheets("Board").Range("N7") = Worksheets("Helper1").Range("D19") - undoCounter

Worksheets("Helper1").Range("E19").Value = 0

End Sub
