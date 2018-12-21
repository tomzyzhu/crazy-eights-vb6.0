Attribute VB_Name = "mdlComputations"
Public intDeckOrder(51) As Integer 'Deck array
Public intPlayerHand(51) As Integer 'Player`s hand array
Public intCompHand(51) As Integer  'Computer`s hand array
Public intRandomSpot As Integer 'Random deck location variable
Public intCounter As Integer 'Counter variable where required
Public intCounter1 As Integer 'Secondary counter when 2 counters are required
Public intCounter2 As Integer 'Third counter when needed
Public intInPlay As Integer 'The In play card Number
Public intAnimationCounter As Integer 'A counter for the animation
Public intAnimationCounter2 As Integer 'See above
Public intCardPlay As Integer 'The index of either the player`s hand array or the computer`s hand array that he/she/it as selected to play

Public Function LoadCard(ByVal CardNumber As Integer) As String 'Takes a card number input and translates it to a file location path
    Dim intSuitNumber As Integer 'Determines the suit of the card number
    Dim intValueNumber As Integer 'Determines the value of the card number
    
    LoadCard = App.Path + "\resources\Cards\" 'Initialize the path to gather card resources
    
    intSuitNumber = CardNumber Mod 4
    intValueNumber = CardNumber \ 4
    
    If intSuitNumber = 0 Then 'Determines Folder Subsection
        LoadCard = LoadCard + "Diamonds\Diamond"
    ElseIf intSuitNumber = 1 Then
        LoadCard = LoadCard + "Clubs\Club"
    ElseIf intSuitNumber = 2 Then
        LoadCard = LoadCard + "Hearts\Heart"
    ElseIf intSuitNumber = 3 Then
        LoadCard = LoadCard + "Spades\Spade"
    End If
    
    If intValueNumber = 0 Then 'Determines file name
        LoadCard = LoadCard + "A"
    ElseIf intValueNumber = 1 Then
        LoadCard = LoadCard + "2"
    ElseIf intValueNumber = 2 Then
        LoadCard = LoadCard + "3"
    ElseIf intValueNumber = 3 Then
        LoadCard = LoadCard + "4"
    ElseIf intValueNumber = 4 Then
        LoadCard = LoadCard + "5"
    ElseIf intValueNumber = 5 Then
        LoadCard = LoadCard + "6"
    ElseIf intValueNumber = 6 Then
        LoadCard = LoadCard + "7"
    ElseIf intValueNumber = 7 Then
        LoadCard = LoadCard + "8"
    ElseIf intValueNumber = 8 Then
        LoadCard = LoadCard + "9"
    ElseIf intValueNumber = 9 Then
        LoadCard = LoadCard + "10"
    ElseIf intValueNumber = 10 Then
        LoadCard = LoadCard + "J"
    ElseIf intValueNumber = 11 Then
        LoadCard = LoadCard + "Q"
    ElseIf intValueNumber = 12 Then
        LoadCard = LoadCard + "K"
    End If
    
    If CardNumber > 51 Then 'Determines special card values set from playing a crazy 8
        LoadCard = App.Path + "\resources\Cards\"
        If CardNumber = 100 Then
            LoadCard = LoadCard + "Diamonds\Diamond"
        ElseIf CardNumber = 200 Then
            LoadCard = LoadCard + "Clubs\Club"
        ElseIf CardNumber = 300 Then
            LoadCard = LoadCard + "Hearts\Heart"
        ElseIf CardNumber = 400 Then
            LoadCard = LoadCard + "Spades\Spade"
        End If
    End If
    
    LoadCard = LoadCard + ".bmp" 'Adds file extension
    
    Exit Function
End Function

Public Function Card(ByVal CardNum As Integer) As String 'Translates a card into a string value that makes sense in english
    'works pretty much identical to LoadCard, except it`s not a path, so no need for expanation
    
    Dim intSuitNumber As Integer
    Dim intValueNumber As Integer
    
    intSuitNumber = CardNum Mod 4
    intValueNumber = CardNum \ 4
    
    If intValueNumber = 0 Then
        Card = Card + "A"
    ElseIf intValueNumber = 1 Then
        Card = Card + "2"
    ElseIf intValueNumber = 2 Then
        Card = Card + "3"
    ElseIf intValueNumber = 3 Then
        Card = Card + "4"
    ElseIf intValueNumber = 4 Then
        Card = Card + "5"
    ElseIf intValueNumber = 5 Then
        Card = Card + "6"
    ElseIf intValueNumber = 6 Then
        Card = Card + "7"
    ElseIf intValueNumber = 7 Then
        Card = Card + "8"
    ElseIf intValueNumber = 8 Then
        Card = Card + "9"
    ElseIf intValueNumber = 9 Then
        Card = Card + "10"
    ElseIf intValueNumber = 10 Then
        Card = Card + "Jack"
    ElseIf intValueNumber = 11 Then
        Card = Card + "Queen"
    ElseIf intValueNumber = 12 Then
        Card = Card + "King"
    End If
    
    Card = Card + " of "
    If intSuitNumber = 0 Then
        Card = Card + "Diamonds"
    ElseIf intSuitNumber = 1 Then
        Card = Card + "Clubs"
    ElseIf intSuitNumber = 2 Then
        Card = Card + "Hearts"
    ElseIf intSuitNumber = 3 Then
        Card = Card + "Spades"
    End If
    
End Function

Public Function Valid(ByVal CardNum As Integer) As Boolean 'determines if a card is able to be played on the in-play card
    If CardNum = -1 Then 'if a card is -1, it is definitely null and invalid
        Valid = False
        Exit Function
    ElseIf CardNum \ 4 = 7 Then 'if a card is a 8, it will always be able to be played
        Valid = True
        Exit Function
    ElseIf CardNum >= 0 And CardNum <= 51 Then 'if the card is not a special condition
        If intInPlay >= 0 And intInPlay <= 51 Then 'if the in-play card is not a crazy eight value
            If CardNum \ 4 = intInPlay \ 4 Or CardNum Mod 4 = intInPlay Mod 4 Then 'determine if valid
                Valid = True
            Else
                Valid = False
            End If
        ElseIf intInPlay >= 100 Then 'if the in-play card is a special crazy eight value
            If CardNum Mod 4 = (intInPlay / 100) - 1 Then 'determine if valid
                Valid = True
            Else
                Valid = False
            End If
        End If
    End If
End Function

Public Function Check() As Integer 'allows the computer to check if it has a valid play
    Dim intChoseCounter As Integer 'a local variable to cycle through a loop to check the computers hand for a valid play
    
    For intChoseCounter = 0 To 51 'Loop determine a random value the computer will input
        If Valid(intCompHand(intChoseCounter)) = True Then
            Check = intCompHand(intChoseCounter)
            intCardPlay = intChoseCounter 'determine the index of the array of the computer hand for later use
            Exit Function
        ElseIf intChoseCounter = 51 Then 'determine if the computer has no card to play, outputs a null value
            Check = -1
            Exit Function
        End If
    Next intChoseCounter
End Function

Public Function PDraw() 'Draws a card for the Player
    Dim intDrawCounter As Integer
    
    If DeckNum = 0 Then 'creates a new shuffled deck if there is no more cards in the deck
        Shuffle
        MsgBox ("The deck has been reset and re-shuffled due no more cards left in the deck")
    End If
    
    For intDrawCounter = 0 To 51 'loop to check where to put the newly drawn card
        If intPlayerHand(intDrawCounter) = -1 Then 'if player hand array is empty as this spot,
            intPlayerHand(intDrawCounter) = intDeckOrder(0) 'give deck value to the empty spot in the player hand array
            intDeckOrder(0) = -1
            Clean
            Exit For 'stop the loop so it does not give the user multiple cards
        End If
    Next intDrawCounter
    
End Function

Public Function CDraw() 'Draws a card for the Computer ie. same thing as above, only with the computer's hand array
    Dim intDrawCounter As Integer
    
    If DeckNum = 0 Then 'creates a new shuffled deck if there is no more cards in the deck
        Shuffle
        MsgBox ("The deck has been reset and re-shuffled due no more cards left in the deck")
    End If
    
    For intDrawCounter = 0 To 51
        If intCompHand(intDrawCounter) = -1 Then
            intCompHand(intDrawCounter) = intDeckOrder(0)
            intDeckOrder(0) = -1
            Clean
            Exit For
        End If
    Next intDrawCounter
    
End Function

Public Function CompNum() As Integer 'counts the number of cards in the computers hand
    Dim intCompNumCounter 'A local counter for a for loop
    Dim intCompNum  'A local variable to store the value of how many cards there are in the computers hand
    
    intCompNum = 0 'initializing local variable that stores the amount of cards the computer has
    
    For intCompNumCounter = 0 To 51 'loop check over the computers hand array to count how many cards are in the computers hand
        If intCompHand(intCompNumCounter) <> -1 Then
            intCompNum = intCompNum + 1
        End If
    Next intCompNumCounter
    
    CompNum = intCompNum 'give the functions output value
End Function

Public Function PlayerNum() 'Same thing as above, except with the player
    Dim intPlayerNumCounter
    Dim intPlayerNum
    
    intPlayerNum = 0
    
    For intPlayerNumCounter = 0 To 51
        If intPlayerHand(intPlayerNumCounter) <> -1 Then
            intPlayerNum = intPlayerNum + 1
        End If
    Next intPlayerNumCounter
    
    PlayerNum = intPlayerNum
End Function

Public Function DeckNum() As Integer 'Counts the amount of cards left in the deck
    Dim intDeckNumCounter 'local counter for a loop
    Dim intDeckNum 'local variable to store the value of the amount of cards left in the deck
    
    intDeckNum = 0 'initializing local variable that stores the amount of cards left in the deck
    
    For intDeckNumCounter = 0 To 51 'loop checks over the deck array to count the number of cards left in the deck
        If intDeckOrder(intDeckNumCounter) <> -1 Then
            intDeckNum = intDeckNum + 1
        End If
    Next intDeckNumCounter
    
    DeckNum = intDeckNum 'give the functions output value
End Function


Public Function Clean() 'Re-organizes all array values
    Dim intCleanCounter As Integer
    Dim intCleanCounter2 As Integer
    
    For intCleanCounter = 0 To 50 'checks all array values for irregular null values that are not at the index of 51
    
        If intPlayerHand(intCleanCounter) = -1 And intPlayerHand(intCleanCounter + 1) <> -1 Then 'cleans the players hand array
            For intCleanCounter2 = intCleanCounter To 50
                intPlayerHand(intCleanCounter2) = intPlayerHand(intCleanCounter2 + 1)
            Next intCleanCounter2
            intPlayerHand(51) = -1
        End If
        
        If intCompHand(intCleanCounter) = -1 And intCompHand(intCleanCounter + 1) <> -1 Then 'cleans the computer hand array
            For intCleanCounter2 = intCleanCounter To 50
                intCompHand(intCleanCounter2) = intCompHand(intCleanCounter2 + 1)
            Next intCleanCounter2
            intCompHand(51) = -1
        End If
        
        If intDeckOrder(intCleanCounter) = -1 And intDeckOrder(intCleanCounter + 1) <> -1 Then 'cleans the deck array
            For intCleanCounter2 = intCleanCounter To 50
                intDeckOrder(intCleanCounter2) = intDeckOrder(intCleanCounter2 + 1)
            Next intCleanCounter2
            intDeckOrder(51) = -1
        End If
        
    Next intCleanCounter
End Function

Public Function Shuffle() 'creates a new deck to draw from

    For intCounter = 0 To 51 'initialize deck for shuffling
        intDeckOrder(intCounter) = -1
    Next intCounter

    For intCounter = 0 To 51
        Randomize
        intRandomSpot = Int((52 - intCounter) * Rnd) 'pick random spot to skip over
        intCounter1 = 0
        Do
            If intRandomSpot = 0 And intDeckOrder(intCounter1) = -1 Then 'determine if to skip over a value
                intDeckOrder(intCounter1) = intCounter                   'when enough values are skipped,
                Exit Do                                                  'leave do loop
            ElseIf Not intDeckOrder(intCounter1) = -1 Then
                intCounter1 = intCounter1 + 1
            ElseIf intRandomSpot > 0 And intDeckOrder(intCounter1) = -1 Then
                intRandomSpot = intRandomSpot - 1
                intCounter1 = intCounter1 + 1
            End If
        Loop
    Next intCounter
End Function
