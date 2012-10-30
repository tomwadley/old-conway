VERSION 5.00
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000004&
   Caption         =   "Conway's Game of Life - Tom W"
   ClientHeight    =   6870
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7260
   FillColor       =   &H00FFFFFF&
   LinkTopic       =   "Form1"
   ScaleHeight     =   6870
   ScaleWidth      =   7260
   Begin VB.Timer tmrNextCycle 
      Interval        =   100
      Left            =   0
      Top             =   0
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    Select Case KeyCode
        Case vbKeyReturn
            changeStartStop
        'Case vbKeyLeft
        '    changeTopCell (TopLeftCellonScreen - 1)
        'Case vbKeyRight
        '    changeTopCell (TopLeftCellonScreen + 1)
        'Case vbKeyUp
        '    changeTopCell (TopLeftCellonScreen - Sqr(arraySize))
        'Case vbKeyDown
        '    changeTopCell (TopLeftCellonScreen + Sqr(arraySize))
        'Case vbKeyS
        '    changeCells (NumberofCells - 1)
        'Case vbKeyX
        '    changeCells (NumberofCells + 1)
        Case vbKeyA
            changeSpeed (tmrNextCycle.Interval + 500)
        Case vbKeyZ
            changeSpeed (tmrNextCycle.Interval - 500)
        Case vbKeyO
            frmOptions.Show
        End Select
            
    updateOptions
End Sub


'I have to put colours in
Private Sub Form_Load()
    species = 1
    arraySize = 1600
    gameSpeed = 1000
    aliveColour = &H0&
    deadColour = &HFFFFFF
    predColour = &HC0C0FF
    preyColour = &HFFC0C0
    NumberofCells = 40
    TopLeftCellonScreen = 1
    tmrNextCycle.Enabled = False
    tmrNextCycle.Interval = gameSpeed
    
    rulestayalive1 = 2
    rulestayalive2 = 3
    rulebecomealive1 = 3
    rulebecomealive2 = 3
    updateOptions
End Sub

Private Sub Form_MouseDown(button As Integer, Shift As Integer, X As Single, Y As Single)
    Call changeCell(X, Y, button)
End Sub

Private Sub Form_Resize()
    drawCycle
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub tmrNextCycle_Timer()
    nextCycle
    drawCycle
End Sub
