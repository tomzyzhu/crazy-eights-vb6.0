VERSION 5.00
Begin VB.Form frmTomZhuCrazyEights_Start 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Start Menu"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3015
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrAnimation 
      Interval        =   55
      Left            =   240
      Top             =   2520
   End
   Begin VB.CommandButton cmdStart 
      BackColor       =   &H0000C000&
      Caption         =   "Start"
      Height          =   495
      Left            =   3240
      MaskColor       =   &H0000C000&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H000080FF&
      Caption         =   "Instructions"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0000FFFF&
      Caption         =   "Close Game"
      Height          =   495
      Left            =   1680
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Shape Shape1 
      Height          =   1935
      Left            =   1680
      Top             =   240
      Width           =   1335
   End
   Begin VB.Shape shpImageBorder 
      BorderColor     =   &H00FFFF00&
      Height          =   2175
      Left            =   1560
      Shape           =   4  'Rounded Rectangle
      Top             =   120
      Width           =   1575
   End
   Begin VB.Label lblIntro2 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Eights"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   3240
      TabIndex        =   4
      Top             =   240
      Width           =   1185
   End
   Begin VB.Label lblIntro1 
      AutoSize        =   -1  'True
      BackColor       =   &H00FFFFFF&
      Caption         =   "Crazy"
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   570
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   1125
   End
   Begin VB.Image imgAnimation 
      Height          =   495
      Left            =   1680
      Top             =   240
      Width           =   1215
   End
End
Attribute VB_Name = "frmTomZhuCrazyEights_Start"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    End
End Sub

Private Sub cmdInstructions_Click()
    frmTomZhuCrazyEights_Instructions.Show
End Sub

Private Sub cmdStart_Click()
    frmTomZhuCrazyEights_Game.Show
    Unload Me
End Sub

Private Sub Form_Load()
    intAnimationCounter = 0 'initialize the animation counter
End Sub

Private Sub tmrAnimation_Timer()
    imgAnimation.Picture = LoadPicture(LoadCard(intAnimationCounter)) 'creates animation
    intAnimationCounter = intAnimationCounter + 1 'alternates between cards
    
    If intAnimationCounter = 52 Then 'resets counter to avoid overflow
        intAnimationCounter = 0
    End If
    
End Sub
