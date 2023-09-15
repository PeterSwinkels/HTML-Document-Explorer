Attribute VB_Name = "Connect4Module"
'This module contains this program's core procedures.
Option Explicit

'This enumeration lists the disk colors used in the game.
Private Enum DiskColorsE
   DCOutsideField          'Indicates that the referred location is outside the game's playing field.
   DCNone                  'Indicates that there is no disk.
   DCRed                   'Indicates a red disk.
   DCYellow                'Indicates a yellow disk.
End Enum

'This enumeration lists the states the game can be in.
Private Enum GameStatesE
   GSNeitherPlaying        'Indicates that neither player is playing.
   GSRedPlaying            'Indicates that the player with the red disks is playing.
   GSRedWon                'Indicates that the player with the red disks has won.
   GSTied                  'Indicates that no more moves are possible and neither player has won.
   GSYellowPlaying         'Indicates that the player with the yellow disks is playing.
   GSYellowWon             'Indicates that the player with the yellow disks has won.
End Enum

'This structure defines the players' setup.
Private Type PlayersSetupStr
   ComputerColor As DiskColorsE     'The color of the disks the computer player plays with.
   FirstColor As DiskColorsE        'The color of the disks of the player who makes the first move.
   HumanColor As DiskColorsE        'The color of the disks the human player plays with when the computer player is enabled.
End Type

Public Const SLOT_SIZE As Long = 100               'The size of a disk slot in pixels.
Private Const DROP_DELAY As Single = 0.1            'The time in seconds it takes for a disk to drop one row.
Private Const FIRST_COLUMN As Long = 0              'The first column of disks.
Private Const FIRST_ROW As Long = 0                 'The first row of disks.
Private Const LAST_COLUMN As Long = 6               'The last column of disks.
Private Const LAST_ROW As Long = 5                  'The last row of disks.
Private Const NO_COLUMN As Long = -1                'Indicates that no column has been selected.
Private Const NO_KEY As Long = 0                    'Indicates that no key has been pressed.
Private Const WINNING_LENGTH As Long = 4            'The number of disks of the same color that must be in one line to win.

'This procedure manages the active player.
Private Function ActivePlayerColor(Optional NewPlayer As DiskColorsE = DiskColorsE.DCNone, Optional ChangeTurns As Boolean = False, Optional ResetPlayers As Boolean = False) As DiskColorsE
On Error GoTo ErrorTrap
Static CurrentPlayer As DiskColorsE

   If ChangeTurns Then
      Select Case CurrentPlayer
         Case DiskColorsE.DCRed
            CurrentPlayer = DiskColorsE.DCYellow
         Case DiskColorsE.DCYellow
            CurrentPlayer = DiskColorsE.DCRed
      End Select
   ElseIf Not NewPlayer = DiskColorsE.DCNone Then
      CurrentPlayer = NewPlayer
   ElseIf ResetPlayers Then
      CurrentPlayer = DiskColorsE.DCNone
   End If

EndRoutine:
   ActivePlayerColor = CurrentPlayer
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure determines which moves the computer player will make.
Private Sub ComputerMakeMove()
On Error GoTo ErrorTrap
Dim Column As Long
Dim MovesFound() As Boolean
Dim TriggerLength As Long

   MovesFound() = FindMoves(PlayersSetup().ComputerColor, WINNING_LENGTH, AllowHelpingOpponent:=True)
   If Not FoundMoves(MovesFound) Then
      MovesFound() = FindMoves(PlayersSetup().HumanColor, WINNING_LENGTH, AllowHelpingOpponent:=False)
      If Not FoundMoves(MovesFound) Then
         For TriggerLength = WINNING_LENGTH To 0 Step -1
            MovesFound() = FindMoves(PlayersSetup().ComputerColor, TriggerLength, AllowHelpingOpponent:=False)
            If FoundMoves(MovesFound) Then Exit For
         Next TriggerLength
      
         If Not FoundMoves(MovesFound) Then
            For TriggerLength = WINNING_LENGTH To 0 Step -1
               MovesFound() = FindMoves(PlayersSetup().HumanColor, TriggerLength, AllowHelpingOpponent:=True)
               If FoundMoves(MovesFound) Then Exit For
            Next TriggerLength
         End If
      End If
   End If

   If FoundMoves(MovesFound) Then
      Do While DoEvents() > 0
         Column = Int(Rnd() * (LAST_COLUMN + 1)) + FIRST_COLUMN
         If MovesFound(Column) Then
            DropDisk Column, ColorO:=ActivePlayerColor()
            Exit Do
         End If
      Loop
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns the number of disks of the specified color found using the specified position and direction.
Private Function CountDisks(StartColumn As Long, StartRow As Long, ColorO As DiskColorsE, XDirection As Long, YDirection As Long) As Long
On Error GoTo ErrorTrap
Dim CheckCount As Long
Dim Column As Long
Dim DiskCount As Long
Dim Row As Long

   Column = StartColumn
   CheckCount = 0
   DiskCount = 0
   Row = StartRow
   Do While CheckCount < WINNING_LENGTH And DoEvents() > 0
      Select Case Disks(Column, Row)
         Case ColorO
            DiskCount = DiskCount + 1
         Case Not DiskColorsE.DCNone
            Exit Do
      End Select
         
      Column = Column + XDirection
      Row = Row + YDirection
      CheckCount = CheckCount + 1
   Loop

EndRoutine:
   CountDisks = DiskCount
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the disks inside the game's playing field.
Private Function Disks(Optional Column As Long = FIRST_COLUMN, Optional Row As Long = FIRST_ROW, Optional NewDisk As DiskColorsE = DiskColorsE.DCNone, Optional ResetDisks As Boolean = False) As DiskColorsE
On Error GoTo ErrorTrap
Dim Disk As DiskColorsE
Static CurrentDisks(FIRST_COLUMN To LAST_COLUMN, FIRST_ROW To LAST_ROW) As DiskColorsE

   Disk = DiskColorsE.DCOutsideField

   If ResetDisks Then
      Erase CurrentDisks()
      Disk = DiskColorsE.DCNone

      For Column = FIRST_COLUMN To LAST_COLUMN
         For Row = FIRST_ROW To LAST_ROW
            CurrentDisks(Column, Row) = DiskColorsE.DCNone
         Next Row
      Next Column
   Else
      If Column >= FIRST_COLUMN And Column <= LAST_COLUMN Then
         If Row >= FIRST_ROW And Row <= LAST_ROW Then
            If Not NewDisk = DiskColorsE.DCNone Then CurrentDisks(Column, Row) = NewDisk
            Disk = CurrentDisks(Column, Row)
         End If
      End If
   End If

EndRoutine:
   Disks = Disk
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure displays the help.
Private Sub DisplayHelp()
On Error GoTo ErrorTrap
Dim HelpText As String
   
   HelpText = App.Comments & vbCr & vbCr
   HelpText = HelpText & "Keys:" & vbCr
   HelpText = HelpText & "F1 = This help." & vbCr
   HelpText = HelpText & "F2 = No computer player." & vbCr
   HelpText = HelpText & "F3 = Computer plays as red." & vbCr
   HelpText = HelpText & "F4 = Computer plays as yellow." & vbCr
   HelpText = HelpText & "A = Restart game." & vbCr
   HelpText = HelpText & "I = Information." & vbCr
   HelpText = HelpText & "R = Red plays first." & vbCr
   HelpText = HelpText & "Y = Yellow plays first."
   MsgBox HelpText, vbInformation

EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws a disk of the specified color at the specified position.
Private Sub DrawDisk(Column As Long, Row As Long, ColorO As DiskColorsE)
On Error GoTo ErrorTrap
   Select Case ColorO
      Case DiskColorsE.DCNone
         Interface().FillColor = vbCyan
      Case DiskColorsE.DCRed
         Interface().FillColor = vbRed
      Case DiskColorsE.DCYellow
         Interface().FillColor = vbYellow
   End Select

   Interface().FillStyle = vbFSTransparent
   Interface().Line (Column * SLOT_SIZE, Row * SLOT_SIZE)-Step(SLOT_SIZE, SLOT_SIZE), vbBlack, B
   Interface().FillStyle = vbFSSolid
   Interface().Circle ((Column + 0.5) * SLOT_SIZE, (Row + 0.5) * SLOT_SIZE), SLOT_SIZE / 2.5, vbBlack
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure draws the disks inside the game's playing field.
Private Sub DrawDisks()
On Error GoTo ErrorTrap
Dim Column As Long
Dim Row As Long

   Interface().Cls
   For Column = FIRST_COLUMN To LAST_COLUMN
      For Row = FIRST_ROW To LAST_ROW
         DrawDisk Column, Row, ColorO:=Disks(Column, Row)
      Next Row
   Next Column
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure drops a disk of the specified color into the playing field at the specified column.
Private Sub DropDisk(Column As Long, ColorO As DiskColorsE)
On Error GoTo ErrorTrap
Dim DelayStart As Single
Dim Row As Long
Static DiskIsFalling As Boolean

   If Not DiskIsFalling Then
      If Disks(Column, FIRST_ROW) = DiskColorsE.DCNone Then
         DiskIsFalling = True
         Row = FIRST_ROW
         Do While DoEvents() > 0
            DrawDisk Column, Row, ColorO
            
            If Row = LAST_ROW Then
               Exit Do
            Else
               If Not Disks(Column, Row + 1) = DiskColorsE.DCNone Then Exit Do
            End If
            
            DelayStart = Timer()
            Do While (Timer() < DelayStart + DROP_DELAY) And DoEvents() > 0
               If Timer() < DelayStart Then DelayStart = Timer()
            Loop
            
            DrawDisk Column, Row, DiskColorsE.DCNone
            Row = Row + 1
         Loop
         
         Disks Column, Row, NewDisk:=ColorO
         ActivePlayerColor , ChangeTurns:=True
         DiskIsFalling = False
      End If
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procdure determines which columns can be used by the computer player to make a move.
Private Function FindMoves(ColorO As DiskColorsE, TriggerLength As Long, AllowHelpingOpponent As Boolean) As Boolean()
On Error GoTo ErrorTrap
Dim Column As Long
Dim CheckColumn As Long
Dim CheckCount As Long
Dim CheckRow As Long
Dim DiskCount As Long
Dim FoundMove As Long
Dim MovesFound(FIRST_COLUMN To LAST_COLUMN) As Boolean
Dim Row As Long
Dim XDirection As Long
Dim YDirection As Long

   Erase MovesFound()

   For Column = FIRST_COLUMN To LAST_COLUMN
      For Row = FIRST_ROW To LAST_ROW
         For XDirection = -1 To 1
            For YDirection = -1 To 1
               If Not (XDirection = 0 And YDirection = 0) Then
                  CheckColumn = Column
                  CheckRow = Row
                  CheckCount = 0
                  DiskCount = 0
                  FoundMove = NO_COLUMN

                  Do Until CheckCount = TriggerLength
                     Select Case Disks(CheckColumn, CheckRow)
                        Case ColorO
                           DiskCount = DiskCount + 1
                        Case DiskColorsE.DCNone
                           Select Case Disks(CheckColumn, CheckRow + 1)
                              Case DiskColorsE.DCRed, DiskColorsE.DCYellow, DiskColorsE.DCOutsideField
                                 FoundMove = CheckColumn
                           End Select
                        Case Else
                           Exit Do
                     End Select
                  
                     CheckColumn = CheckColumn + XDirection
                     CheckRow = CheckRow + YDirection
                     CheckCount = CheckCount + 1
                  Loop
               
                  If DiskCount = TriggerLength - 1 Then
                     If Not FoundMove = NO_COLUMN Then
                        If AllowHelpingOpponent Then
                           MovesFound(FoundMove) = True
                        Else
                           If Not MoveHelpsOpponent(PlayersSetup().HumanColor, FoundMove) Then MovesFound(FoundMove) = True
                        End If
                     End If
                  End If
               End If
            Next YDirection
         Next XDirection
      Next Row
   Next Column

EndRoutine:
   FindMoves = MovesFound()
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure checks whether the specified
Private Function FoundMoves(MovesFound() As Boolean) As Boolean
On Error GoTo ErrorTrap
Dim Column As Long
Dim Found As Boolean

   Found = False
   For Column = FIRST_COLUMN To LAST_COLUMN
      If MovesFound(Column) Then
         Found = True
         Exit For
      End If
   Next Column

EndRoutine:
   FoundMoves = Found
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns whether any more moves can be made by the players.
Private Function GameDone() As Boolean
On Error GoTo ErrorTrap
Dim Column As Long
Dim Done As Boolean

   Done = True
   For Column = FIRST_COLUMN To LAST_COLUMN
      If Disks(Column, FIRST_ROW) = DiskColorsE.DCNone Then
         Done = False
         Exit For
      End If
   Next Column

EndRoutine:
   GameDone = Done
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns the game's state.
Private Function GetGameState() As GameStatesE
On Error GoTo ErrorTrap
Dim GameState As GameStatesE

   Select Case WinningPlayer()
      Case DiskColorsE.DCNone
         If GameDone() Then
            GameState = GameStatesE.GSTied
         Else
            Select Case ActivePlayerColor()
               Case DiskColorsE.DCNone
                  GameState = GameStatesE.GSNeitherPlaying
               Case DiskColorsE.DCRed
                  GameState = GameStatesE.GSRedPlaying
               Case DiskColorsE.DCYellow
                  GameState = GameStatesE.GSYellowPlaying
            End Select
         End If
      Case DiskColorsE.DCRed
         GameState = GameStatesE.GSRedWon
      Case DiskColorsE.DCYellow
         GameState = GameStatesE.GSYellowWon
   End Select

EndRoutine:
   GetGameState = GameState
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure greys out all the disks inside game's playing field.
Private Sub GreyOutDisks()
On Error GoTo ErrorTrap
Dim Column As Long
Dim Row As Long

   Interface().FillColor = vbBlack
   Interface().FillStyle = vbDiagonalCross
   For Column = FIRST_COLUMN To LAST_COLUMN
      For Row = FIRST_ROW To LAST_ROW
         Interface().Circle ((Column + 0.5) * SLOT_SIZE, (Row + 0.5) * SLOT_SIZE), SLOT_SIZE / 2.5
      Next Row
   Next Column
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure handles any errors that occur.
Public Sub HandleError()
Dim Description As String
Dim ErrorCode As Long

   Description = Err.Description
   ErrorCode = Err.Number

   On Error Resume Next
   MsgBox Description & vbCr & "Error code: " & CStr(ErrorCode), vbExclamation
End Sub

'This procedure waits for a human player to make a move.
Private Sub HumanMakeMove()
On Error GoTo ErrorTrap
Dim Column As Long

   Column = LastColumnSelected()
   If Not Column = NO_COLUMN Then
      DropDisk Column, ColorO:=ActivePlayerColor()
      LastColumnSelected , ResetColumn:=True
   End If
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure initializes the game.
Private Sub InitializeGame()
On Error GoTo ErrorTrap
Static CurrentFirstColor As DiskColorsE

   Randomize
   ActivePlayerColor , , ResetPlayers:=True
   Disks , , , ResetDisks:=True
   DrawDisks
   GreyOutDisks
   LastColumnSelected , ResetColumn:=True
EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure manages the game's interface window.
Private Function Interface(Optional NewInterface As Form = Nothing) As Form
On Error GoTo ErrorTrap
Static CurrentInterface As Form

   If Not NewInterface Is Nothing Then
      Set CurrentInterface = NewInterface
      
      With CurrentInterface
         .BackColor = vbBlue
         .Width = ((Abs(LAST_COLUMN - FIRST_COLUMN) + 1) * SLOT_SIZE) * Screen.TwipsPerPixelX
         .Width = .Width + (.Width - (.ScaleWidth * Screen.TwipsPerPixelX))
         .Height = ((Abs(LAST_ROW - FIRST_ROW) + 1) * SLOT_SIZE) * Screen.TwipsPerPixelY
         .Height = .Height + (.Height - (.ScaleHeight * Screen.TwipsPerPixelY))
         .Left = (Screen.Width / 2) - (.Width / 2)
         .Top = (Screen.Height / 2) - (.Height / 2)
         
         .Caption = vbNullString
      End With
   End If

EndRoutine:
   Set Interface = CurrentInterface
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the most recently selected column.
Public Function LastColumnSelected(Optional NewSelectedColumn As Long = NO_COLUMN, Optional ResetColumn As Boolean = False) As Long
On Error GoTo ErrorTrap
Dim LastSelected As Long
Static SelectedColumn As Long

   LastSelected = NO_COLUMN
   If ResetColumn Then SelectedColumn = NO_COLUMN

   If Not GetGameState() = GameStatesE.GSNeitherPlaying Then
      If NewSelectedColumn = NO_COLUMN Then
         LastSelected = SelectedColumn
         SelectedColumn = NO_COLUMN
      Else
         SelectedColumn = NewSelectedColumn
         LastSelected = SelectedColumn
      End If
   End If

EndRoutine:
   LastColumnSelected = LastSelected
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the most recently made key stroke.
Public Function LastKeyStroke(Optional NewKeyStroke As Long = NO_KEY) As Long
On Error GoTo ErrorTrap
Dim KeyStroke As Long
Static CurrentKeyStroke As Long

   If NewKeyStroke = NO_KEY Then
      KeyStroke = CurrentKeyStroke
      CurrentKeyStroke = NO_KEY
   Else
      CurrentKeyStroke = NewKeyStroke
      KeyStroke = CurrentKeyStroke
   End If

EndRoutine:
   LastKeyStroke = KeyStroke
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure is executed when this program is started.
Private Sub Main()
On Error GoTo ErrorTrap
   ChDrive Left$(App.Path, InStr(App.Path, ":"))
   ChDir App.Path

   Interface NewInterface:=InterfaceWindow
   Interface().Show

   PlayGame
EndRoutine:
   Unload Interface()
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure indicates whether dropping a disk at the specified column can help the specified opponent.
Private Function MoveHelpsOpponent(OpponentColor As DiskColorsE, Column As Long) As Boolean
On Error GoTo ErrorTrap
Dim HelpsOpponent As Boolean
Dim Row As Long
Dim XDirection As Long
Dim YDirection As Long

   HelpsOpponent = False
   Row = 0
   Do Until (Row = LAST_ROW) Or (Not Disks(Column, Row + 1) = DiskColorsE.DCNone) Or (DoEvents() = 0)
      Row = Row + 1
   Loop

   If Row > FIRST_ROW Then
      Row = Row - 1

      For XDirection = -1 To 1
         For YDirection = -1 To 1
            If Not (XDirection = 0 And YDirection = 0) Then
               If CountDisks(Column, Row, OpponentColor, XDirection, YDirection) = WINNING_LENGTH - 1 Then
                  HelpsOpponent = True
               End If
            End If
         Next YDirection
      Next XDirection
   End If

EndRoutine:
   MoveHelpsOpponent = HelpsOpponent
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the players' setup.
Private Function PlayersSetup(Optional NewComputerColor As DiskColorsE = DiskColorsE.DCNone, Optional NewFirstColor As DiskColorsE = DiskColorsE.DCNone, Optional NoComputerPlayer As Boolean = False) As PlayersSetupStr
On Error GoTo ErrorTrap
Static CurrentPlayersSetup As PlayersSetupStr

   With CurrentPlayersSetup
      If NoComputerPlayer Then
         .ComputerColor = DiskColorsE.DCNone
      Else
         If Not NewComputerColor = DiskColorsE.DCNone Then .ComputerColor = NewComputerColor
      End If

      If Not NewFirstColor = DiskColorsE.DCNone Then .FirstColor = NewFirstColor

      Select Case .ComputerColor
         Case DiskColorsE.DCNone
            .HumanColor = DiskColorsE.DCNone
         Case DiskColorsE.DCRed
            .HumanColor = DiskColorsE.DCYellow
         Case DiskColorsE.DCYellow
            .HumanColor = DiskColorsE.DCRed
      End Select
   End With

EndRoutine:
   PlayersSetup = CurrentPlayersSetup
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure manages the game.
Private Sub PlayGame()
On Error GoTo ErrorTrap
Dim KeyStroke As Long

   PlayersSetup NewComputerColor:=DiskColorsE.DCYellow, NewFirstColor:=DiskColorsE.DCRed
   InitializeGame

   Do While DoEvents() > 0
      Interface().Caption = App.Title & " - " & StateText() & " F1 = Help"
      
      Select Case GetGameState()
         Case GameStatesE.GSRedPlaying, GameStatesE.GSYellowPlaying
            If PlayersSetup().ComputerColor = DiskColorsE.DCNone Then
               HumanMakeMove
            Else
               If ActivePlayerColor() = PlayersSetup().ComputerColor Then
                  ComputerMakeMove
               Else
                  HumanMakeMove
               End If
            End If
         Case GameStatesE.GSRedWon, GameStatesE.GSYellowWon, GameStatesE.GSTied
            GreyOutDisks
      End Select

      KeyStroke = LastKeyStroke()
      If Not KeyStroke = NO_KEY Then
         Select Case GetGameState()
            Case GameStatesE.GSNeitherPlaying
               DrawDisks
               ActivePlayerColor NewPlayer:=PlayersSetup().FirstColor
            Case GameStatesE.GSRedWon, GameStatesE.GSYellowWon, GameStatesE.GSTied
               InitializeGame
            Case Else
               Select Case KeyStroke
                  Case vbKeyA
                     InitializeGame
                  Case vbKeyI
                     With App
                        MsgBox .Title & " v" & CStr(.Major) & "." & CStr(.Minor) & CStr(.Revision) & " - by: " & .CompanyName, vbInformation
                     End With
                  Case vbKeyR
                     PlayersSetup , NewFirstColor:=DiskColorsE.DCRed
                     InitializeGame
                  Case vbKeyY
                     PlayersSetup , NewFirstColor:=DiskColorsE.DCYellow
                     InitializeGame
                  Case vbKeyF1
                     DisplayHelp
                  Case vbKeyF2
                     PlayersSetup , , NoComputerPlayer:=True
                     InitializeGame
                  Case vbKeyF3
                     PlayersSetup NewComputerColor:=DiskColorsE.DCRed
                     InitializeGame
                  Case vbKeyF4
                     PlayersSetup NewComputerColor:=DiskColorsE.DCYellow
                     InitializeGame
               End Select
         End Select
      End If
   Loop

EndRoutine:
   Exit Sub

ErrorTrap:
   HandleError
   Resume EndRoutine
End Sub

'This procedure returns a text description of the game's state.
Private Function StateText() As String
On Error GoTo ErrorTrap
Dim Text As String

   Text = vbNullString
   Select Case GetGameState()
      Case GameStatesE.GSNeitherPlaying
         Text = "Inactive. - Press any key."
      Case GameStatesE.GSRedPlaying
         Text = "Red's turn."
      Case GameStatesE.GSYellowPlaying
         Text = "Yellow's turn."
      Case GameStatesE.GSRedWon
         Text = "Red won."
      Case GameStatesE.GSYellowWon
         Text = "Yellow won."
      Case GameStatesE.GSTied
         Text = "Game is tied."
   End Select

EndRoutine:
   StateText = Text
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

'This procedure returns which player has won.
Private Function WinningPlayer() As DiskColorsE
On Error GoTo ErrorTrap
Dim ColorO As DiskColorsE
Dim Column As Long
Dim Row As Long
Dim XDirection As Long
Dim YDirection As Long
Dim Winner As DiskColorsE

   Winner = DiskColorsE.DCNone

   For ColorO = DiskColorsE.DCRed To DiskColorsE.DCYellow
      For Column = FIRST_COLUMN To LAST_COLUMN
         For Row = FIRST_ROW To LAST_ROW
            For XDirection = -1 To 1
               For YDirection = -1 To 1
                  If Not (XDirection = 0 And YDirection = 0) Then
                     If CountDisks(Column, Row, ColorO, XDirection, YDirection) = WINNING_LENGTH Then
                        Winner = ColorO
                     End If
                  End If
               Next YDirection
            Next XDirection
         Next Row
      Next Column
   Next ColorO

EndRoutine:
   WinningPlayer = Winner
   Exit Function

ErrorTrap:
   HandleError
   Resume EndRoutine
End Function

