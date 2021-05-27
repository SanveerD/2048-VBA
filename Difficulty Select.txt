' Difficulty Select

Public numUndos As Integer
Public freeTileCount As Integer

Sub StartGame()

UserForm2.Show

End Sub

Sub freeTile()
Dim currCell As Range
Dim numFreeTiles As Integer

numFreeTiles = Worksheets("Helper1").Range("numFreeTiles")


If freeTileCount < numFreeTiles Then
    Set currCell = ActiveCell
        If InRange(currCell, Range("E5:H10")) Then
            If currCell = "" Then
                MsgBox ("Please choose a valid tile")
                Exit Sub
            Else
                currCell.ClearContents
                currCell.Interior.ColorIndex = 36
                currCell.Font.Color = vbBlack
            End If
        Else
            ' code to handle that the active cell is not within the right range
            MsgBox "You have to choose a tile on the board first"
            Exit Sub
        End If
End If

freeTileCount = freeTileCount + 1

If freeTileCount >= numFreeTiles Then
    Worksheets("Board").Buttons("Button 10").Font.Color = 8421504
    Worksheets("Board").Buttons("Button 10").OnAction = "buttonDisabled"
End If

If Worksheets("Helper1").Range("numFreeTiles") - freeTileCount < 0 Then
    Worksheets("Board").Range("N6") = 0
Else
    Worksheets("Board").Range("N6") = Worksheets("Helper1").Range("numFreeTiles") - freeTileCount
End If

End Sub

Sub Block()

Dim Board, tile As Range
Dim boardSize, numEmptySpace, scatterCell, probability As Integer

Set Board = Sheet7.Range("E5:H8")
boardSize = 16
numEmptySpace = boardSize - WorksheetFunction.CountA(Board)

If numEmptySpace = 0 Then
'if no empty spaces, user has lost
    MsgBox ("Invalid entry; block cannot be inserted!")
    'go to lose screen
    Call LoseScreen
    Exit Sub
Else
   scatterCell = WorksheetFunction.RandBetween(1, numEmptySpace)
   
For Each tile In Board
    If tile.Value = "" Then
        scatterCell = scatterCell - 1
            If scatterCell = 0 Then
                    tile.Value = 1
                    tile.Font.Color = vbRed
                    tile.Interior.Color = vbRed
                    Exit Sub
            End If
    End If

Next tile
End If

End Sub

Function InRange(Range1 As Range, Range2 As Range) As Boolean
    ' returns True if Range1 is within Range2
    InRange = Not (Application.Intersect(Range1, Range2) Is Nothing)
End Function

Sub resetButtons()
    freeTileCount = 0
    Worksheets("Board").Buttons("Button 10").Font.Color = vbBlack
    Worksheets("Board").Buttons("Button 10").OnAction = "freeTile"
    
    undoCounter = 0
    Worksheets("Helper1").Range("numUndos") = 0
    Worksheets("Board").Buttons("Button 12").Font.Color = vbBlack
    Worksheets("Board").Buttons("Button 12").OnAction = "undomove2"
End Sub

Sub buttonDisabled()
    MsgBox ("You used all of these up!")
End Sub
