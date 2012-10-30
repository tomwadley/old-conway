Attribute VB_Name = "basMain"
'Conway's game of Life - Tom Wadley


Public gridReal(1 To 10000) As Integer 'Current cycle, calculate from this array
'In the array above, 1 = Alive, 0 = Dead, 2 = Predetor, 3 = Prey
Public arraySize
Public gameSpeed As Integer 'Sets how often the next cycle is calculated and drawn (milisecends, 1000 = 1 second)
Public aliveColour 'Defines the colour for an alive cell (normally black)
Public deadColour 'Defines the colour for a dead cell (normally white)
Public predColour 'These names are misleading
Public preyColour
Public species
Public NumberofCells 'This is the number of cells to be displayed across the top and the side of the screen (same number) multiply by self for total cells on screen
Public TopLeftCellonScreen 'This is the number of the top left cell on the screen (used in drawCycle)

Public rulestayalive1 As Integer
Public rulestayalive2 As Integer
Public rulebecomealive1 As Integer
Public rulebecomealive2 As Integer

Sub main()
    frmOptions.Show
    frmMain.Show
End Sub

Public Sub nextCycle() 'Calculates the next cycle
    ReDim gridnext(1 To arraySize) 'Next cycle, used for calculating next cycle in nextCycle
    Count = 0
    countpred = 0
    countprey = 0
    
    For Y = 1 To Sqr(arraySize)
        For X = 1 To Sqr(arraySize)
            'Looks at each of the 8 cells around the current cell, if alive add 1 to count
            If TestSurroundingCells(Y, X, 1) = 2 Then
                Count = Count + 10 'any number larger that 8
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, 1)) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, 1)) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, 1)) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, -1) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -1)) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -1)) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -1)) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, Sqr(arraySize)) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize))) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize))) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize))) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, -Sqr(arraySize)) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -Sqr(arraySize))) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -Sqr(arraySize))) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -Sqr(arraySize))) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, Sqr(arraySize) + 1) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize) + 1)) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize) + 1)) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize) + 1)) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, Sqr(arraySize) - 1) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize) - 1)) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize) - 1)) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, Sqr(arraySize) - 1)) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, -(Sqr(arraySize) - 1)) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -(Sqr(arraySize) - 1))) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -(Sqr(arraySize) - 1))) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -(Sqr(arraySize) - 1))) = 3 Then countprey = countprey + 1
                End If
            If TestSurroundingCells(Y, X, -(Sqr(arraySize) + 1)) = 2 Then
                Count = Count + 10
                Else
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -(Sqr(arraySize) + 1))) = 1 Then Count = Count + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -(Sqr(arraySize) + 1))) = 2 Then countpred = countpred + 1
                If gridReal(convert((Y), (X)) + TestSurroundingCells(Y, X, -(Sqr(arraySize) + 1))) = 3 Then countprey = countprey + 1
                End If

            
            If species = 1 Then
                'If cell has 2 or 3 alive cells around it, it becomes alive
                If gridReal(convert((Y), (X))) = 1 Then
                    If Count > rulestayalive1 - 1 And Count < rulestayalive2 + 1 Then
                        gridnext(convert((Y), (X))) = 1
                    Else
                        gridnext(convert((Y), (X))) = 0
                        End If
                Else
                        If gridReal(convert((Y), (X))) = 0 Then
                            If Count > rulebecomealive1 - 1 And Count < rulebecomealive2 + 1 Then
                                gridnext(convert((Y), (X))) = 1
                            Else
                                gridnext(convert((Y), (X))) = 0
                                End If
                            End If
                    End If
            ElseIf species = 2 Then
                If gridReal(convert((Y), (X))) = 2 Then
                    If countpred = 2 Or countpred = 3 Then
                        gridnext(convert((Y), (X))) = 2
                    Else
                        gridnext(convert((Y), (X))) = 0
                        End If
                    If countprey > countpred Then
                        gridnext(convert((Y), (X))) = 0
                        End If
                    'If countpred > 3 Or countpred < 1 Then
                    '    gridnext(convert((Y), (X))) = 0
                    '    End If
                    'Select Case countprey - countpred
                    '    Case Is > 0
                    '        gridnext(convert((Y), (X))) = 2
                    '    Case 0
                    '        gridnext(convert((Y), (X))) = 2
                    '    Case Is < 0
                    '        gridnext(convert((Y), (X))) = 0
                    '    End Select
                ElseIf gridReal(convert((Y), (X))) = 3 Then
                    If countprey = 2 Or countprey = 3 Then
                        gridnext(convert((Y), (X))) = 3
                    Else
                        gridnext(convert((Y), (X))) = 0
                        End If
                    If countpred > countprey Then
                        gridnext(convert((Y), (X))) = 0
                        End If
                    'If countprey = 2 Or countprey = 3 Then
                    '    gridnext(convert((Y), (X))) = 3
                    'Else
                    '    gridnext(convert((Y), (X))) = 0
                    '    End If
                    'If countpred > 1 Then
                    '    gridnext(convert((Y), (X))) = 2
                    '    End If
                    
                    'Select Case countpred - countprey
                    '    Case Is > 0
                    '        gridnext(convert((Y), (X))) = 2
                    '    Case 0
                    '        gridnext(convert((Y), (X))) = 0
                    '    Case Is < 0
                    '        gridnext(convert((Y), (X))) = 3
                    '    End Select
                ElseIf gridReal(convert((Y), (X))) = 0 Then
                    gridnext(convert((Y), (X))) = 0
                    If countprey = 3 Then
                        gridnext(convert((Y), (X))) = 3
                        End If
                    If countpred = 3 Then
                        gridnext(convert((Y), (X))) = 2
                        End If
                    If countprey = 3 And countpred = 3 Then
                        gridnext(convert((Y), (X))) = 0
                        End If
                    'If countpred = 3 Then
                    '    gridnext(convert((Y), (X))) = 2
                    'Else
                    '    gridnext(convert((Y), (X))) = 0
                    '    End If
                    'End If
                End If
            End If
                
                
                
   
            Count = 0
            countpred = 0
            countprey = 0
            Next X
        Next Y

        For i = 1 To arraySize
            gridReal(i) = gridnext(i)
            Next i
End Sub

Sub drawCycle() 'Draws the grid from gridReal onto frmMain using NumberofCells and frmMain height and top
    Dim cellSize As Integer 'store the width/height of each cell (same number)
    If frmMain.Height > frmMain.Width Then 'calculates in which direction frmMain is bigger for use when drawing the grid
        cellSizeDecimals = frmMain.Width / NumberofCells
        Else
        cellSizeDecimals = frmMain.Height / NumberofCells
        End If
    'cellSizeDecimals = cellSizeDecimals - 1 'Minus 1 so the grid cant go over the screen. Not strictly nessasary
    cellSize = cellSizeDecimals 'Removes decimal places (cellSize is and must be integer)
    
    Dim UptoLine 'remembers what line we are upto drawing
    Dim LastPositionFormLeft As Integer 'The last left position used for drawing
    Dim LastPositionFormTop As Integer 'The last top position used for drawing
    Dim UptoPositionArray As Integer 'The position in the gridReal array that is to be drawn next
    UptoLine = 1
    LastPositionFormLeft = 0
    LastPositionFromTop = 0
    UptoPositionArray = TopLeftCellonScreen 'The first cell to be draw is the one set to be on the top left of the screen

    For i = 1 To (NumberofCells * NumberofCells)
        Select Case gridReal(UptoPositionArray)
            Case 0
                colour = deadColour
            Case 1
                colour = aliveColour
            Case 2
                colour = predColour
            Case 3
                colour = preyColour
            End Select
        'If gridReal(UptoPositionArray) = 1 Then colour = aliveColour Else colour = deadColour 'determains colour for next cell
        'End If
        frmMain.Line (LastPositionFormLeft, LastPositionFormTop)-(LastPositionFormLeft + cellSize, LastPositionFormTop + cellSize), colour, BF 'draws 1 cell
        z = False
        For e = 1 To NumberofCells
            If i = NumberofCells * e Then
                z = True
                End If
            Next e
        If z = True Then
        'If i = NumberofCells Or i = (NumberofCells * 2) Or i = (NumberofCells * 3) Or i = (NumberofCells * 4) Or i = (NumberofCells * 5) Or i = (NumberofCells * 6) Or i = (NumberofCells * 7) Or i = (NumberofCells * 8) Or i = (NumberofCells * 9) Or i = (NumberofCells * 10) Then 'Sets up the LastPosition and other vars to match what has just been done for the next i
            LastPositionFormLeft = 0
            linesFinished = UptoLine
            LastPositionFormTop = LastPositionFormTop + cellSize
            UptoPositionArray = TopLeftCellonScreen + (UptoLine * (Sqr(arraySize)))
            UptoLine = UptoLine + 1
            Else
            LastPositionFormLeft = LastPositionFormLeft + cellSize
            UptoPositionArray = UptoPositionArray + 1
            End If
        'LastPositionFormTop = cellSize * UptoLine 'Sets up LastPositonLeft to thr correct value
        Next i
End Sub

Public Sub changeCell(mouseX, mouseY, button)
    Dim cellSize As Integer 'store the width/height of each cell (same number)
    If frmMain.Height > frmMain.Width Then 'calculates in which direction frmMain is bigger for use when drawing the grid
        cellSizeDecimals = frmMain.Width / NumberofCells
        Else
        cellSizeDecimals = frmMain.Height / NumberofCells
        End If
    cellSizeDecimals = cellSizeDecimals - 1 'Minus 1 so the grid cant go over the screen. Not strictly nessasary
    cellSize = cellSizeDecimals 'Removes decimal places (cellSize is and must be integer)
    
    Dim UptoLine 'remembers what line we are upto drawing
    Dim LastPositionFormLeft As Integer 'The last left position used for drawing
    Dim LastPositionFormTop As Integer 'The last top position used for drawing
    Dim UptoPositionArray As Integer 'The position in the gridReal array that is to be drawn next
    UptoLine = 1
    LastPositionFormLeft = 0
    LastPositionFromTop = 0
    UptoPositionArray = TopLeftCellonScreen 'The first cell to be draw is the one set to be on the top left of the screen
    Dim validClick As Boolean
    validClick = True

    For i = 1 To NumberofCells
        If mouseX > LastPositionFormLeft And mouseX < (LastPositionFormLeft + cellSize) Then
            Exit For
            End If
        If i = NumberofCells Then
            validClick = False
            Exit For
            End If
        LastPositionFormLeft = LastPositionFormLeft + cellSize
        UptoPositionArray = UptoPositionArray + 1
        Next i
    For e = 1 To NumberofCells
        If mouseY > LastPositionFormTop And mouseY < (LastPositionFormTop + cellSize) Then
            Exit For
            End If
        If e = NumberofCells Then
            validClick = False
            Exit For
            End If
        LastPositionFormTop = LastPositionFormTop + cellSize
        UptoPositionArray = UptoPositionArray + Sqr(arraySize)
        Next e

If species = 2 Then
    If button = 1 Then 'left click
        If gridReal(UptoPositionArray) = 3 Then
            gridReal(UptoPositionArray) = 0
            colour = deadColour
        Else
            gridReal(UptoPositionArray) = 3
            colour = preyColour
        End If
    Else 'right click (2)
        If gridReal(UptoPositionArray) = 2 Then
            gridReal(UptoPositionArray) = 0
            colour = deadColour
        Else
            gridReal(UptoPositionArray) = 2
            colour = predColour
        End If
        End If
Else
    If gridReal(UptoPositionArray) = 1 Then
        gridReal(UptoPositionArray) = 0
    Else
        gridReal(UptoPositionArray) = 1
        End If
    If gridReal(UptoPositionArray) = 1 Then
        colour = aliveColour
    Else
        colour = deadColour 'determains colour for cell
        End If
    End If

    If validClick = True Then
        frmMain.Line (LastPositionFormLeft, LastPositionFormTop + cellSize)-(LastPositionFormLeft + cellSize, LastPositionFormTop), colour, BF 'draws the cell the same as cells are drawn in drawCycle
        End If

End Sub

Function TestSurroundingCells(arrayPositionY, arrayPositionX, test) 'Tests weather or not a certain cell around a cell specified in NextCycle is dead or alive, called from NextCycle
    TestSurroundingCells = test
    Select Case test
        Case 1
            If arrayPositionX + 1 > Sqr(arraySize) Then TestSurroundingCells = 2
        Case -1
            If arrayPositionX - 1 < 1 Then TestSurroundingCells = 2
        Case Sqr(arraySize)
            If arrayPositionY + 1 > Sqr(arraySize) Then TestSurroundingCells = 2
        Case -Sqr(arraySize)
            If arrayPositionY - 1 < 1 Then TestSurroundingCells = 2
        Case Sqr(arraySize) + 1
            If arrayPositionY + 1 > Sqr(arraySize) Or arrayPositionX + 1 > Sqr(arraySize) Then TestSurroundingCells = 2
        Case Sqr(arraySize) - 1
            If arrayPositionY + 1 > Sqr(arraySize) Or arrayPositionX - 1 < 1 Then TestSurroundingCells = 2
        Case -(Sqr(arraySize) + 1)
            If arrayPositionY - 1 < 1 Or arrayPositionX - 1 < 1 Then TestSurroundingCells = 2
        Case -(Sqr(arraySize) - 1)
            If arrayPositionY - 1 < 1 Or arrayPositionX + 1 > Sqr(arraySize) Then TestSurroundingCells = 2
        End Select
End Function

Function convert(Y As Integer, X As Integer) As Integer 'Converts an X and a Y value into 1 number for the array
    convert = (Y - 1) * Sqr(arraySize) + X
End Function

Public Sub changeSpeed(speed)
    If speed > 0 Then
        frmMain.tmrNextCycle.Interval = speed
        gameSpeed = speed
        End If
End Sub

Public Sub changeRulestayalive1(newrule)
    rulestayalive1 = newrule
End Sub
Public Sub changeRulestayalive2(newrule)
    rulestayalive2 = newrule
End Sub
Public Sub changeRulebecomealive1(newrule)
    rulebecomealive1 = newrule
End Sub
Public Sub changeRulebecomealive2(newrule)
    rulebecomealive2 = newrule
End Sub

'Public Sub changeTopCell(topcell)
'    If unconvertx((topcell + (NumberofCells * NumberofCells))) > Sqr(arraySize) Or unconvertx((topcell + (NumberofCells * NumberofCells))) < 1 Or unconverty((topcell + (NumberofCells * NumberofCells))) > Sqr(arraySize) Or unconverty((topcell + (NumberofCells * NumberofCells))) < 1 Then
'    Else
'        TopLeftCellonScreen = topcell
'        End If
'    drawCycle
'End Sub

'Public Sub changeCells(cells)
'    NumberofCells = cells
'    drawCycle
'End Sub

Public Sub changeArraysize(numberofcellsarray)
    If Sqr(arraySize) = numberofcellsarray Then

    Else
        arraySize = numberofcellsarray * numberofcellsarray
        NumberofCells = Sqr(arraySize)
        For i = 1 To 10000
            gridReal(i) = 0
            Next i
        End If
    drawCycle
End Sub

Public Sub changeStartStop()
    If frmMain.tmrNextCycle.Enabled = False Then
        frmMain.tmrNextCycle.Enabled = True
    Else
        frmMain.tmrNextCycle.Enabled = False
        End If
End Sub

Public Sub updateOptions()
    frmOptions.txtarraysize.Text = NumberofCells * NumberofCells
    frmOptions.txtBecomealive1.Text = rulebecomealive1
    frmOptions.txtBecomealive2.Text = rulebecomealive2
    frmOptions.txtStayalive1.Text = rulestayalive1
    frmOptions.txtStayalive2.Text = rulestayalive2
    frmOptions.txtNumberofcellsarray.Text = Sqr(arraySize)
    frmOptions.txtSpeed.Text = gameSpeed
End Sub
