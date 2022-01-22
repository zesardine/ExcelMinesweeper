'Define Basic Variables
Dim boardSize As Integer
Dim pauseSearch As Boolean

Dim Mines()
Dim BombsLeft As Integer
Dim BombsTotal As Integer

Dim GamePlaying As Boolean

Sub Generate()
    'Reset Variables
    pauseSearch = False
    GamePlaying = True
    
    'Check if the 4th column next to the board contains a number, if it does, set the total amount of bombs to be generated to that number
    'If it isn't a number, just set it to the default
    If Cells(4, boardSize + 1).Value <> "" And IsNumeric(Cells(4, boardSize + 1).Value) Then
        BombsTotal = CInt(Cells(4, boardSize + 1).Value)
    Else
        BombsTotal = 10
    End If
    
    'Same with the 5th cloumn, except now it sets the boardsize to the entered number
    If Cells(5, boardSize + 1).Value <> "" And IsNumeric(Cells(5, boardSize + 1).Value) Then
        boardSize = CInt(Cells(5, boardSize + 1).Value)
    Else
        boardSize = 10
    End If
    
    'If the amount of bombs goes over a certain limit, it sets the amount of bombs to that limit
    'This limit is dependant on the boardsize, for example a 10x10 board can only have 20 bombs ((10*10)/5)
    'I found somewhere that a fifth of the boardsize squared is a good maximum limit, you can change it if you want though
    If BombsTotal > (boardSize * boardSize) / 5 Then
        BombsTotal = Round((boardSize * boardSize) / 5, 0)
    End If
    
    'BombsLeft is used as the amount of flags remaining
    BombsLeft = BombsTotal
    
    'This will add a 2 cell padding around the board. This is done so that for loops later on are easier as we don't have to check if we're at the edge of the board, you'll see later
    'The actual playing board will really only be of the size that the player defines
    boardSize = boardSize + 2
    
    'Sets the Mines array to a 2d array that is big boardSize x boardSize
    'Important to note is that this array has some padding
    ReDim Mines(1 To boardSize, 1 To boardSize)
    
    'Next line creates a simple vector array in the format (x, y) which we'll use when generating the board
    'i and j are just variable we'll use in the for loop
    Dim position(1 To 2) As Integer, i As Integer, j As Integer
        
    'Sets some basic formatting for the whole sheet, effectively making it blank and making the cells square
    With Sheets(1)
        .Cells.Value = ""
        .Cells.Borders.LineStyle = xlLineStyleNone
        .Cells.Interior.ColorIndex = 0
        .Cells.ColumnWidth = 4
        .Cells.RowHeight = 24
        .Cells.Font.Bold = True
        .Cells.Font.Size = 15
        .Cells.Font.Name = "Arial Black"
        .Cells.Font.Color = RGB(0, 0, 0)
    End With
    
    Sheets(1).Columns.HorizontalAlignment = xlCenter
    
    Cells(1, 1).Select
    
    'Basically puts back the buttons to the left of the board
    Cells(2, boardSize + 1).Value = "R"
    Cells(4, boardSize + 1).Value = BombsTotal
    Cells(5, boardSize + 1).Value = boardSize - 2
    
    'Sets the column width of the column with buttons to a higher value so the text fits
    Columns(boardSize + 1).ColumnWidth = 8
    
    'Goes through all the cells from (2, 2) to (boardSize - 1, boardSize - 1) and makes them a darker color, also adding a border
    With Range(Cells(2, 2), Cells(boardSize - 1, boardSize - 1))
        .Value = ""
        .Borders.LineStyle = xlContinuous
        .Interior.ColorIndex = 15
    End With
    
    'Initially fills in the mines array with 0's
    For i = 1 To boardSize
         For j = 1 To boardSize
            Mines(i, j) = 0
        Next j
    Next i
    
    i = 0
    
    'Next piece of code fills in the whole board with an amount of X's equal to the bombTotal
    Randomize
    Do
        'First we create a random x position where our bomb could go and save that in the x value of the position vector we created earlier
        position(1) = Int(Rnd * (boardSize - 2)) + 2
        'Same with the y position
        position(2) = Int(Rnd * (boardSize - 2)) + 2
        'Check if that slot in the mines array doesn't already have a bomb, if it doesn't, place a bomb there
        If Mines(position(1), position(2)) <> "X" Then
            Mines(position(1), position(2)) = "X"
            i = i + 1
        End If
        'Repeat until the amount of bombs in the mine array is equal to the bomb amount
    Loop Until i = BombsTotal
    
    'Following for loop goes through all the empty slots in the mines array (where there should be numbers)
    'Then it goes through all the surrounding squares and adds one for every bomb that's in them
    'This is where the padding comes in handy. We don't start at position (1, 1) but instead position (2, 2)
    'What this means is that even if we're checking the square to the left and we're in the left most square, it's still not out of bounds of the array
    'Same with the rightmost square since we're only couting up to boardSize - 1
    For i = 2 To boardSize - 1
         For j = 2 To boardSize - 1
            If Mines(i, j) <> "X" Then
                Mines(i, j) = 0
                If Mines(i - 1, j - 1) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i - 1, j) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i - 1, j + 1) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i, j - 1) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i, j + 1) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i + 1, j - 1) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i + 1, j) = "X" Then Mines(i, j) = Mines(i, j) + 1
                If Mines(i + 1, j + 1) = "X" Then Mines(i, j) = Mines(i, j) + 1
            End If
        Next j
    Next i
    
    'Reselecting the top left empty cell
    Cells(1, 1).Select
End Sub

'The following deals with rightclick events
Sub Worksheet_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Cancel = True 'So the right button menu is not displayed.
    If GamePlaying = True Then 'All of the following only works if the game is being played (not won or lost)
        
        'If the cell is already clear, then exit.
        If Target.Interior.ColorIndex = -4142 Then Exit Sub
        
        'If it has a flag, then remove it.
        If Target.Value = "F" Then
            Target.Value = ""
            Target.Interior.ColorIndex = 15
            Target.Font.Color = RGB(0, 0, 0)
            BombsLeft = BombsLeft + 1
        'If it doesnt have it, then place it.
        Else
            Target.Value = "F"
            Target.Interior.ColorIndex = 16
            Target.Font.Color = RGB(240, 0, 0)
            BombsLeft = BombsLeft - 1
        End If
        
        'Since the bombsValue changed, let's update it in the UI
        Cells(4, boardSize + 1).Value = BombsLeft
        
        'If the player has used all the bombs, this following loop compares the mines array with the actual gameboard, if there's flags where there's bombs, it displays the "You win" message
        If BombsLeft = 0 Then
            Dim Correct As Integer
            Correct = 0
            For i = 2 To boardSize - 1
                For j = 2 To boardSize - 1
                    If Mines(i, j) = "X" Then
                        If Cells(i, j) = "F" Then
                            Correct = Correct + 1
                        End If
                    End If
                Next j
            Next i
            If Correct = BombsTotal Then
                GamePlaying = False
                MsgBox "You Win!"
                'Resets the bomb count for if you want to play again
                Cells(4, boardSize + 1).Value = BombsTotal
            End If
        End If
        
        Cells(1, 1).Select 'and keep this cell selected again
    End If
End Sub

'This is a boolean which is set to true if the right button on the mouse is being pressed, this is used for doubleclicks
Function RightButton() As Boolean
    RightButton = (GetAsyncKeyState(vbKeyRButton) And &H8000)
End Function

'This is now for the left clicks, and also for left+right clicks
Sub Worksheet_SelectionChange(ByVal Target As Range)
    Dim Count As Integer
    Dim R1 As Long, R2 As Long
    R1 = Target.Row: R2 = Target.Column 'Just some variables that we'll use
    
    'Following will run the reset function when you click on the R
    If R1 = 2 And R2 = boardSize + 1 Then
        Generate
        Exit Sub
    End If
    
    If GamePlaying = True Then
        'This just makes sure that if you select multiple cells, only the left top one stays selected
        If Target.Rows.Count > 1 Or Target.Columns.Count > 1 Then
            Cells(Target.Row, Target.Column).Select
            Exit Sub
        End If
        
        'Nothing happens if we clicked outside the board
        If R1 > boardSize - 1 Or R2 > boardSize - 1 Then Exit Sub
        If R1 < 2 Or R2 < 2 Then Exit Sub
        
        'Following deals with what happens when a player double clicks
        'Note that the rightbutton variable has to be true for this to work
        If RightButton And Not pauseSearch Then
            'Checks if the clicked square has a number, aka it's not empty or flagged
            If Target.Value <> "" Then
                Count = 0
                'Counts the amount of flags around the clicked square
                If Cells(R1 - 1, R2 - 1).Value = "F" Then Count = Count + 1
                If Cells(R1 - 1, R2).Value = "F" Then Count = Count + 1
                If Cells(R1 - 1, R2 + 1).Value = "F" Then Count = Count + 1
                If Cells(R1, R2 - 1).Value = "F" Then Count = Count + 1
                If Cells(R1, R2 + 1).Value = "F" Then Count = Count + 1
                If Cells(R1 + 1, R2 - 1).Value = "F" Then Count = Count + 1
                If Cells(R1 + 1, R2).Value = "F" Then Count = Count + 1
                If Cells(R1 + 1, R2 + 1).Value = "F" Then Count = Count + 1
                
                'If the amount of flags equals the number in the square, it opens all the surrounding squares
                'This is actually a self calling function, since even the computer selecting a square will call this very function
                'Thus, the game actually automatically clears out empty areas, although this is done lower
                If Count = Mines(R1, R2) Then
                    pauseSearch = True
                    Cells(R1 - 1, R2 - 1).Select
                    Cells(R1 - 1, R2).Select
                    Cells(R1 - 1, R2 + 1).Select
                    Cells(R1, R2 - 1).Select
                    Cells(R1, R2 + 1).Select
                    Cells(R1 + 1, R2 - 1).Select
                    Cells(R1 + 1, R2).Select
                    Cells(R1 + 1, R2 + 1).Select
                    pauseSearch = False
                    If GamePlaying = False Then
                        'GamePlaying = True
                    End If
                End If
            End If
            Exit Sub
        End If
        
        'This exits the function if the square is empty or flagged
        If Target.Value <> "" Or Target.Interior.ColorIndex = -4142 Then
            Cells(1, 1).Select
            Exit Sub
        End If
        
        'If we've gotten to this point, we actually write the value from the Mines array to the board
        Target.Value = Mines(R1, R2)
        Target.Interior.ColorIndex = 0
        Target.Font.ColorIndex = 0
        
        'If that value is a bomb, we set the game playing to false and show the Game Over Message
        If Mines(R1, R2) = "X" Then
            GamePlaying = False
            Target.Interior.ColorIndex = 3
            MsgBox "Game over!"
            Cells(4, boardSize + 1).Value = BombsTotal
            'Following just goes through the mines array and paints them all onto the board so you can see where they were
            For i = 2 To boardSize - 1
                 For j = 2 To boardSize - 1
                    If Mines(i, j) = "X" Then
                        Cells(i, j) = "X"
                        Cells(i, j).Font.ColorIndex = 0
                        Cells(i, j).Interior.ColorIndex = 3
                    End If
                Next j
            Next i
            'GamePlaying = True
        ElseIf Mines(R1, R2) = 0 Then
            'If the square is empty, we have to clear the area, this is done by calling the select function on the surrounding squares like I talked about earlier
            Target.Font.Color = RGB(255, 255, 255)
            Cells(R1 - 1, R2 - 1).Select
            Cells(R1 - 1, R2).Select
            Cells(R1 - 1, R2 + 1).Select
            Cells(R1, R2 - 1).Select
            Cells(R1, R2 + 1).Select
            Cells(R1 + 1, R2 - 1).Select
            Cells(R1 + 1, R2).Select
            Cells(R1 + 1, R2 + 1).Select
        End If
        
        'And finally, if the mines array only contains a number in the square, we display the number on the board, in it's correct color
        If Mines(R1, R2) = 1 Then
            Target.Font.Color = RGB(0, 0, 200)
        ElseIf Mines(R1, R2) = 2 Then
            Target.Font.Color = RGB(0, 200, 0)
        ElseIf Mines(R1, R2) = 3 Then
            Target.Font.Color = RGB(200, 0, 0)
        ElseIf Mines(R1, R2) = 4 Then
            Target.Font.Color = RGB(50, 0, 220)
        ElseIf Mines(R1, R2) = 5 Then
            Target.Font.Color = RGB(100, 0, 20)
        ElseIf Mines(R1, R2) = 6 Then
            Target.Font.Color = RGB(0, 200, 200)
        ElseIf Mines(R1, R2) = 7 Then
            Target.Font.Color = RGB(0, 0, 0)
        ElseIf Mines(R1, R2) = 8 Then
            Target.Font.Color = RGB(100, 100, 100)
        End If
        
        Cells(1, 1).Select 'Lets once again keep this cell selected
    End If
End Sub
