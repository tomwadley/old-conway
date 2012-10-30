VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   6135
   ClientLeft      =   7440
   ClientTop       =   630
   ClientWidth     =   5625
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6135
   ScaleWidth      =   5625
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdinfo 
      Caption         =   "<-- Hide Info"
      Height          =   375
      Left            =   1560
      TabIndex        =   16
      Top             =   5640
      Width           =   1335
   End
   Begin VB.CommandButton cmdset 
      Caption         =   "Set to Conway's rule"
      Height          =   375
      Left            =   840
      TabIndex        =   15
      Top             =   2760
      Width           =   1575
   End
   Begin VB.TextBox txtarraysize 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Text            =   "1600"
      Top             =   1800
      Width           =   615
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Clear the grid"
      Height          =   375
      Left            =   1560
      TabIndex        =   9
      Top             =   5160
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "2 Species"
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   5640
      Width           =   1335
   End
   Begin VB.TextBox txtNumberofcellsarray 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   7
      Text            =   "40"
      Top             =   960
      Width           =   615
   End
   Begin VB.TextBox txtBecomealive2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Text            =   "3"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtStayalive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   5
      Text            =   "2"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtStayalive2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   4
      Text            =   "3"
      Top             =   3360
      Width           =   615
   End
   Begin VB.TextBox txtBecomealive1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Text            =   "3"
      Top             =   4320
      Width           =   615
   End
   Begin VB.TextBox txtSpeed 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   1
      Text            =   "1000"
      Top             =   120
      Width           =   615
   End
   Begin VB.CommandButton cmdChange 
      Caption         =   "Make Changes"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   5160
      Width           =   1335
   End
   Begin VB.Label lbl3 
      Caption         =   $"frmOptions.frx":0000
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   18
      Top             =   2760
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.Label lbl1 
      Alignment       =   2  'Center
      Caption         =   "Rule:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   2760
      Width           =   735
   End
   Begin VB.Label lbl2 
      Caption         =   $"frmOptions.frx":00B3
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   14
      Top             =   3120
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Size of grid (Whole thing, number above squared)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   1680
      Width           =   2415
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Size of grid (Number of cells allong the top and the side)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   0
      TabIndex        =   10
      Top             =   720
      Width           =   2415
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Wait between frames (1000's of second)"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   2415
   End
   Begin VB.Label Label5 
      Caption         =   $"frmOptions.frx":0168
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   6135
      Left            =   3120
      TabIndex        =   17
      Top             =   0
      Width           =   2415
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChange_Click()
    changeSpeed (txtSpeed.Text)
    'changeTopCell (txtTopcell.Text)
    'changeCells (txtNumberofcells.Text)
    changeArraysize (txtNumberofcellsarray.Text)
    changeRulestayalive1 (txtStayalive1.Text)
    changeRulestayalive2 (txtStayalive2.Text)
    changeRulebecomealive1 (txtBecomealive1.Text)
    changeRulebecomealive2 (txtBecomealive2.Text)
    updateOptions
    frmMain.SetFocus
End Sub

Private Sub cmdinfo_Click()
    If frmOptions.Width = 5715 Then
        frmOptions.Width = 3165
        cmdinfo.Caption = "Show Info -->"
    Else
        frmOptions.Width = 5715
        cmdinfo.Caption = "<-- Hide Info"
        End If
End Sub

Private Sub Command1_Click()
    If species = 1 Then
        species = 2
        lbl1.Visible = False
        lbl2.Visible = False
        cmdset.Visible = False
        txtStayalive1.Visible = False
        txtStayalive2.Visible = False
        txtBecomealive1.Visible = False
        txtBecomealive2.Visible = False
        Command1.Caption = "Classic Conway"
        lbl3.Visible = True
    Else
        species = 1
        lbl1.Visible = True
        lbl2.Visible = True
        cmdset.Visible = True
        txtStayalive1.Visible = True
        txtStayalive2.Visible = True
        txtBecomealive1.Visible = True
        txtBecomealive2.Visible = True
        Command1.Caption = "2 Species"
        lbl3.Visible = False
        End If
    
    For i = 1 To 10000
        gridReal(i) = 0
        Next i
    drawCycle
    frmMain.SetFocus
End Sub

Private Sub Command2_Click()
    For i = 1 To 10000
        gridReal(i) = 0
        Next i
    drawCycle
    frmMain.SetFocus
End Sub

Private Sub Command3_Click()
    txtStayalive1.Text = 2
    txtStayalive2.Text = 3
    txtBecomealive1.Text = 3
    txtBecomealive2.Text = 3
End Sub

Private Sub Form_Load()
'    drawCycle
End Sub
