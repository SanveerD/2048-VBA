'Leaderboard
Public leaderBoardName As String
Public leaderBoardScore As Integer
Public leaderBoardMoves As Integer

Sub Leaderboards()
Sheet1.Activate
ActiveWindow.DisplayGridlines = False
End Sub

Sub PlayAgain()

UserForm1.Show

End Sub

Sub RunLeaderBoard()
Dim currCell As Range

Sheet1.Activate
Set currCell = Range("B100").End(xlUp)

currCell.Offset(1, 0).Value = leaderBoardName
currCell.Offset(1, 1).Value = leaderBoardScore
currCell.Offset(1, 2).Value = leaderBoardMoves
currCell.Offset(1, 3).Value = Worksheets("Helper1").Range("C30")

Range("B6:E100").Sort key1:=Range("C6:C100"), _
order1:=xlDescending, Header:=xlNo

ActiveWorkbook.Save

End Sub

Sub GoToMenu()
Sheet6.Activate
End Sub

Sub clearleaderboard()
Dim answer As Integer

Sheet1.Activate
answer = MsgBox("You will now clear all score history! Would you like to continue?", vbYesNo + vbQuestion, "Clear")
If answer = vbYes Then
    Range("B6:E100").ClearContents
Else
    Exit Sub
End If
End Sub
