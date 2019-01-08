VERSION 5.00
Begin VB.Form frmTomZhuCrazyEights_Game 
   BackColor       =   &H0000C000&
   Caption         =   "Game"
   ClientHeight    =   9150
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10815
   LinkTopic       =   "Form1"
   ScaleHeight     =   9150
   ScaleWidth      =   10815
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdInstructions 
      BackColor       =   &H000080FF&
      Caption         =   "Instructions"
      Height          =   495
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   60
      Left            =   10080
      Top             =   1680
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H0000FFFF&
      Caption         =   "Close Game"
      Height          =   495
      Left            =   9480
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton cmdPlay 
      BackColor       =   &H000000FF&
      Caption         =   "Play Selected Card"
      Height          =   495
      Left            =   6360
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton cmdDraw 
      BackColor       =   &H000000FF&
      Caption         =   "Draw Card"
      Height          =   495
      Left            =   1920
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   840
      Width           =   1215
   End
   Begin VB.Label lblInfo 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Hint : A black arrows means you can play the card under the arrow, a red arrows mean you've selected the card under the arrow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   10
      Top             =   2400
      Width           =   3855
   End
   Begin VB.Label lblHint3 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Hint : Forgot how to play? Go to the instructions again"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   9
      Top             =   840
      Width           =   1815
   End
   Begin VB.Label lblHint 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Hint : Click on a card to select that card"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   2400
      Width           =   1935
   End
   Begin VB.Label lblHint2 
      Alignment       =   2  'Center
      BackColor       =   &H0000C000&
      Caption         =   "Hint : A black arrows means you can play the card under the arrow, a red arrows mean you've selected the card under the arrow"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6360
      TabIndex        =   7
      Top             =   960
      Width           =   3855
   End
   Begin VB.Label lblInPlay 
      BackColor       =   &H0000C000&
      Caption         =   "In Play"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4920
      TabIndex        =   5
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblDeck 
      BackColor       =   &H0000C000&
      Caption         =   "Deck"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   3480
      TabIndex        =   4
      Top             =   120
      Width           =   855
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   13
      Left            =   9360
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   12
      Left            =   7920
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   11
      Left            =   6480
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   10
      Left            =   5040
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   9
      Left            =   3600
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   8
      Left            =   2160
      Top             =   6000
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   465
      Index           =   7
      Left            =   720
      Picture         =   "frmTomZhuCrazyEights_Game.frx":0000
      Top             =   6000
      Width           =   345
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   6
      Left            =   9360
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   5
      Left            =   7920
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   4
      Left            =   6480
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   3
      Left            =   5040
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   2
      Left            =   3600
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   495
      Index           =   1
      Left            =   2160
      Top             =   3240
      Width           =   1215
   End
   Begin VB.Image imgArrow 
      Height          =   465
      Index           =   0
      Left            =   720
      Picture         =   "frmTomZhuCrazyEights_Game.frx":08FA
      Top             =   3240
      Width           =   345
   End
   Begin VB.Image imgInPlay 
      Height          =   495
      Left            =   4920
      Top             =   840
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   13
      Left            =   8880
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   12
      Left            =   7440
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   11
      Left            =   6000
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   10
      Left            =   4560
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   9
      Left            =   3120
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   8
      Left            =   1680
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   7
      Left            =   240
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Label lblHandOverlay 
      BackColor       =   &H0000C000&
      Caption         =   "Your Cards:"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   2
      Top             =   2760
      Width           =   1695
   End
   Begin VB.Image imgDeck 
      Height          =   2160
      Left            =   3240
      Picture         =   "frmTomZhuCrazyEights_Game.frx":11F4
      Top             =   840
      Width           =   1425
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   6
      Left            =   8880
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   5
      Left            =   7440
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   4
      Left            =   6000
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   3
      Left            =   4560
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   2
      Left            =   3120
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   1
      Left            =   1680
      Top             =   3960
      Width           =   1215
   End
   Begin VB.Image imgPlayerHand 
      Height          =   495
      Index           =   0
      Left            =   240
      Top             =   3960
      Width           =   1215
   End
End
Attribute VB_Name = "frmTomZhuCrazyEights_Game"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Name:Tom Zhu
'Class:ICS2O1
'Date:01/11/16
'Game:Crazy Eights
'Input:A valid card to play
'Process: A computer reply and if either player has won
'Output: Display information to user
'Goal: Get rid of all the cards in users hand before the computer

Public Function CPU() 'The function that allows the computer to interact
    'also used to check if someone has won yet
    If Check <> -1 Then
        
        MsgBox ("The computer plays " & Card(Check))
        If Check \ 4 = 7 Then 'if the computer used a wild card (some 8 value), decide randomly which suit to set it to
            Randomize
            intInPlay = Int(4 * Rnd + 1) * 100
            If intInPlay = 100 Then
                MsgBox ("Since the computer plays a 8, it has chosen that the next suit must be a diamond")
            ElseIf intInPlay = 200 Then
                MsgBox ("Since the computer plays a 8, it has chosen that the next suit must be a clubs")
            ElseIf intInPlay = 300 Then
                MsgBox ("Since the computer plays a 8, it has chosen that the next suit must be a hearts")
            ElseIf intInPlay = 400 Then
                MsgBox ("Since the computer plays a 8, it has chosen that the next suit must be a spades")
            End If
        Else
            intInPlay = Check
            If intInPlay \ 4 = 1 Then 'if the computer plays a 2, the player must draw 2 cards
                PDraw
                PDraw
                MsgBox ("Since the computer played a 2, you get 2 cards")
            End If
        End If
        intCompHand(intCardPlay) = -1 'Deletes value of the card the computer played and also the index of the array he played
        intCardPlay = -1
        Clean 'clean the computer hand array
    Else
        CDraw 'draw a card for the computer if it has no moves
        MsgBox ("The computer did not play anything and has drawn a card")
    End If
    
    If PlayerNum = 0 Then 'determine if anyone has won and closes the program if someone has won
        MsgBox ("Congratulations, you won!!!")
        End
    ElseIf CompNum = 0 Then
        MsgBox ("You have lost, sorry...")
        End
    End If
    
End Function

Private Function Render() 'Resets animations
    'failsafe incase animation function messed up somewhere
    Dim intRenderCounter 'Set a local counter for any use in the function
    
    Clean 'Just in case any arrays are not in the right order
    For intRenderCounter = 0 To 13
        imgArrow(intRenderCounter).Visible = False
    Next intRenderCounter
    For intRenderCounter = 0 To 13
    
        imgArrow(intRenderCounter).Visible = False
        imgArrow(intRenderCounter).Picture = LoadPicture(App.Path + "\resources\Arrows\Black.bmp")
        If intPlayerHand(intRenderCounter) >= 0 And intPlayerHand(intRenderCounter) <= 51 Then
            imgPlayerHand(intRenderCounter).Picture = LoadPicture(LoadCard(intPlayerHand(intRenderCounter)))
        Else
            imgPlayerHand(intRenderCounter).Visible = False
        End If
        
        If intPlayerHand(intRenderCounter) <> -1 Then
            If intInPlay < 51 Then
                If intPlayerHand(intRenderCounter) Mod 4 = intInPlay Mod 4 Or intPlayerHand(intrendercoutner) \ 4 = intInPlay \ 4 Or intPlayerHand(intRenderCounter) \ 4 = 7 Then
                    imgArrow(intRenderCounter).Visible = True
                End If
            Else
                If intPlayerHand(intRenderCounter) Mod 4 = (intInPlay / 100) - 1 Then
                    imgArrow(intRenderCounter).Visible = True
                End If
            End If
        End If
        
    Next intRenderCounter
    
    If DeckNum = 0 Then
        imgDeck.Visible = False
    Else
        imgDeck.Visible = True
    End If
    
    imgInPlay.Picture = LoadPicture(LoadCard(intInPlay))
    
    
End Function

Private Sub cmdClose_Click()
    End 'close program if user wants to exit
End Sub

Private Sub cmdDraw_Click()
    PDraw 'draw a card for the user
    MsgBox ("You have drawn a card and it is now the computer's turn")
    CPU 'computate the computers turn
    Render 'backplanB
End Sub

Private Sub cmdInformation_Click() 'Displays any information the user might want to know
    MsgBox ("The Computer has " & CompNum & " cards and there are " & DeckNum & " cards in the deck")
End Sub

Private Sub cmdInstructions_Click()
    frmTomZhuCrazyEights_Instructions.Show 'loads the instructions form
End Sub

Private Sub cmdPlay_Click() 'Determines if the player can play a certain card
    If intCardPlay > -1 Then
        If intPlayerHand(intCardPlay) \ 4 = 7 Then 'if the card is a 8, the user must select a suit to change it to
            Dim strSuit 'declare local variable to get the suit input from user
            
            Do 'gets suit input
                strSuit = InputBox("Please change the suit of either `Diamonds`, `Clubs`, `Hearts`, or `Spades`")
                strSuit = UCase(strSuit) 'remove case sensitivity
                If strSuit = "DIAMONDS" Or strSuit = "CLUBS" Or strSuit = "HEARTS" Or strSuit = "SPADES" Then
                    Exit Do
                End If
            Loop
            If strSuit = "DIAMONDS" Then
                intInPlay = 100
            ElseIf strSuit = "CLUBS" Then
                intInPlay = 200
            ElseIf strSuit = "HEARTS" Then
                intInPlay = 300
            ElseIf strSuit = "SPADES" Then
                intInPlay = 400
            End If
            
        ElseIf intPlayerHand(intCardPlay) \ 4 = 1 Then 'if the card played is 2, the compter must draw 2
            CDraw 'Draws 2 cards for the computer
            CDraw
            MsgBox ("You have given the computer 2 extra cards since you played a 2")
            intInPlay = intPlayerHand(intCardPlay)
        Else
            intInPlay = intPlayerHand(intCardPlay) 'Sets the value the inplay card if there are no special effects
        End If
        
        MsgBox ("Your card, " + Card(intPlayerHand(intCardPlay)) + ", has been succesfully played")
        intPlayerHand(intCardPlay) = -1 'resets hand values and the card to play value
        intCardPlay = -1
        
        'if the player or the computer have no cards, let the player know who won, and end program
        If PlayerNum = 0 Then 'determine if anyone has won and closes the program if someone has won
            MsgBox ("Congratulations, you won!!!")
            End
        ElseIf CompNum = 0 Then
            MsgBox ("You have lost, sorry...")
            End
        Else
            CPU
        End If
        
    End If
    Render
End Sub

Private Sub Form_Load()
    Const intInitialDraw As Integer = 7 ' Determines how many cards easy player gets at the start of the game
    intAnimationCounter = 0 'initialize animation counter
        
    Shuffle 'generates new shuffled deck
    
    For intCounter = 0 To 13 'resets all arrow values
        imgArrow(intCounter).Picture = LoadPicture(App.Path + "\resources\Arrows\Black.bmp")
        imgArrow(intCounter).Visible = False
    Next intCounter
    
    For intCounter = 0 To 51 'Initializes hands as no values
        intPlayerHand(intCounter) = -1
        intCompHand(intCounter) = -1
    Next intCounter
    
    For intCounter = 1 To intInitialDraw 'Draws both players however many cards as preset in the constant        PDraw
        CDraw
        PDraw
    Next intCounter
    
    Clean 'cleans all arrays
    
    intInPlay = intDeckOrder(0) 'Sets new in-play card by taking the top card of the deck
    intDeckOrder(0) = -1
    
    Clean 'Clean up the deck array after setting the new in-play value
    
    intCardPlay = -1 'sets the selected card to null
    
End Sub

Private Sub imgPlayerHand_Click(Index As Integer) 'Senses if the player has clicked on any of the images to select his next play
    If imgArrow(Index).Visible = True Then  'checks for if the card is valid
        
        For intCounter = 0 To 13 'resets all arrows to black
            imgArrow(intCounter).Picture = LoadPicture(App.Path + "\resources\Arrows\Black.bmp")
        Next intCounter
    
        imgArrow(Index).Picture = LoadPicture(App.Path + "\resources\Arrows\Red.bmp") 'sets selected card to the red arrow
        
        intCardPlay = Index
    End If
End Sub

Private Sub tmrAnimation_Timer() 'animates arrows up and down as well as being render.planA for all images

    imgInPlay.Picture = LoadPicture(LoadCard(intInPlay)) 'updates the inplay card
    
    Clean 'cleans player hand array before cleaning the animation of image boxes
    
    For intAnimationCounter2 = 0 To 13 'all image boxes are invisible
        imgPlayerHand(intAnimationCounter2).Visible = False
    Next intAnimationCounter2
    
    Dim intOverFlowBarrier
    intOverFlowBarrier = PlayerNum
    
    If intOverFlowBarrier > 14 Then 'placeholder so program doesn't crash if there is more than 14 cards in hand
        intOverFlowBarrier = 14
    End If
    
    For intAnimationCounter2 = 0 To intOverFlowBarrier - 1 'updates only the image boxes that have a card to display their card
        
        imgPlayerHand(intAnimationCounter2).Picture = LoadPicture(LoadCard(intPlayerHand(intAnimationCounter2)))
        imgPlayerHand(intAnimationCounter2).Visible = True
    Next intAnimationCounter2
    
    For intAnimationCounter2 = 0 To 13 'loops over all 14 arrows to check if they need to be updated
        If Valid(intPlayerHand(intAnimationCounter2)) = True Then 'checks if arrows need to be visible or invisible
            imgArrow(intAnimationCounter2).Visible = True
        Else
            imgArrow(intAnimationCounter2).Visible = False
        End If
    Next intAnimationCounter2
    
    If intAnimationCounter \ 7 = 0 Then
        For intAnimationCounter2 = 0 To 13 'moves every arrow down
            imgArrow(intAnimationCounter2).Move imgArrow(intAnimationCounter2).Left, imgArrow(intAnimationCounter2).Top + 25
        Next intAnimationCounter2
    Else
        For intAnimationCounter2 = 0 To 13 'moves every arrow up
            imgArrow(intAnimationCounter2).Move imgArrow(intAnimationCounter2).Left, imgArrow(intAnimationCounter2).Top - 25
        Next intAnimationCounter2
    End If
    
    intAnimationCounter = intAnimationCounter + 1 'cycles though moving up and moving down
    
    If intAnimationCounter = 14 Then 'Doesn't allow overflow to happen
        intAnimationCounter = 0
    End If
    
    'Update information for the Player
    lblInfo.Caption = "The Computer has " & CompNum & " cards in their hand, you have " & PlayerNum & " cards in hand, and there are " & DeckNum & " cards in the deck"
End Sub
