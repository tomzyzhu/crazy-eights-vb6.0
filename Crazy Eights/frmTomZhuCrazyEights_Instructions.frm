VERSION 5.00
Begin VB.Form frmTomZhuCrazyEights_Instructions 
   Caption         =   "Instructions"
   ClientHeight    =   10845
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   10845
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H000080FF&
      Caption         =   "Close Instructions"
      Height          =   735
      Left            =   480
      MaskColor       =   &H000080FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9480
      Width           =   1695
   End
   Begin VB.Label lblInput 
      Caption         =   $"frmTomZhuCrazyEights_Instructions.frx":0000
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   2640
      TabIndex        =   4
      Top             =   9600
      Width           =   11295
   End
   Begin VB.Label lblInstructionsTitle 
      AutoSize        =   -1  'True
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   5160
      TabIndex        =   2
      Top             =   0
      Width           =   1860
   End
   Begin VB.Image imgExample 
      Height          =   1980
      Left            =   360
      Picture         =   "frmTomZhuCrazyEights_Instructions.frx":00E5
      Top             =   3120
      Width           =   1350
   End
   Begin VB.Label lblExamples 
      AutoSize        =   -1  'True
      Caption         =   "Examples"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   6000
      TabIndex        =   1
      Top             =   2400
      Width           =   1260
   End
   Begin VB.Label lblInstructions 
      Caption         =   $"frmTomZhuCrazyEights_Instructions.frx":8D67
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   14175
   End
End
Attribute VB_Name = "frmTomZhuCrazyEights_Instructions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    imgExample.Picture = LoadPicture(App.Path + "\resources\Cards\Examples.bmp")
End Sub
