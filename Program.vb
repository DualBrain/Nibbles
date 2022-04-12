'
'                         Q B a s i c   N i b b l e s
'
'                   Copyright (C) Microsoft Corporation 1990
'
' Nibbles is a game for one or two players.  Navigate your snakes
' around the game board trying to eat up numbers while avoiding
' running into walls or other snakes.  The more numbers you eat up,
' the more points you gain and the longer your snake becomes.
'
' To run this game, press Shift+F5.
'
' To exit QBasic, press Alt, F, X.
'
' To get help on a BASIC keyword, move the cursor to the keyword and press
' F1 or click the right mouse button.
'
' NOTE: Converted to VB for .NET Core by Cory Smith 2019-06-26.  Enjoy!
'
' TODO:
'  
'  - Needs music / game sound.  Currently using BEEP.
'  - Possibly need ESC (with prompt) to abort the game.
'  - Game speed might not be 100% accurate with regards to the original.
'  - Need to test on Linux / Mac.
'  - Need to test with "other" consoles on Windows.

Option Explicit On
Option Strict On
Option Infer On

Imports System.Text
Imports System.Console

Module Nibbles

  'User-defined TYPEs

  Structure SnakeBody
    Public row As Integer
    Public col As Integer
  End Structure

  'This type defines the player's snake

  Structure SnakeType
    Public head As Integer
    Public length As Integer
    Public row As Integer
    Public col As Integer
    Public direction As Integer
    Public lives As Integer
    Public score As Integer
    Public scolor As ConsoleColor
    Public alive As Boolean 'Integer
  End Structure

  'This type is used to represent the playing screen in memory
  'It is used to simulate graphics in text mode, and has some interesting,
  'and slightly advanced methods to increasing the speed of operation.
  'Instead of the normal 80x25 text graphics using "█", we will be
  'using "▄" and "▀" and "█" to mimic an 80x50 pixel screen.
  'Check out sub-programs SET and POINTISTHERE to see how this is implemented
  'feel free to copy these (as well as arenaType and the DIM ARENA stmt and the
  'initialization code in the DrawScreen subprogram) and use them in your own
  'programs
  Structure ArenaType
    Public realRow As Integer 'Maps the 80x50 point into the real 80x25
    Public acolor As ConsoleColor 'Stores the current color of the point
    Public sister As Integer  'Each char has 2 points in it. Sister is
  End Structure               '-1 if sister point is above, +1 if below

  'Constants

  Const MAXSNAKELENGTH = 1000
  Const STARTOVER = 1             ' Parameters to 'Level' SUB
  Const SAMELEVEL = 2
  Const NEXTLEVEL = 3

  'Global Variables

  Private ReadOnly arena(50, 80) As ArenaType
  Private curLevel%
  Private ReadOnly colorTable(10) As ConsoleColor

  Sub Main()

#Region "Prepare Console Window"

    Console.OutputEncoding = Encoding.UTF8

    Dim OrgBufferHeight%, OrgBufferWidth%, OrgWindowHeight%, OrgWindowWidth%

    OrgBufferHeight = Console.BufferHeight
    OrgBufferWidth = Console.BufferWidth
    OrgWindowHeight = Console.WindowHeight
    OrgWindowWidth = Console.WindowWidth

    Resize(80, 26)

    Console.CursorVisible = False

#End Region

    Dim numPlayers%
    Dim speed%
    Dim diff = ""
    Dim monitor = ""

    Randomize(Timer)

    Call ClearKeyLocks()

    Intro()

    GetInputs(numPlayers, speed, diff, monitor)
    Call SetColors(monitor)
    DrawScreen()

    Do
      PlayNibbles(numPlayers, speed, diff$)
    Loop While StillWantsToPlay()

    Call RestoreKeyLocks()

    ForegroundColor = ConsoleColor.White
    BackgroundColor = ConsoleColor.Black
    Clear()

  End Sub

  Private Sub SetColors(monitor$)

    ' snake1     snake2   Walls  Background  Dialogs-Fore  Back

    Dim values As Integer()

    If monitor$ = "M" Then
      values = {15, 7, 7, 0, 15, 0}
    Else
      values = {14, 13, 12, 1, 15, 4}
    End If

    For A = 1 To 6
      colorTable(A) = CType(values(A - 1), ConsoleColor)
    Next

  End Sub

  Private Sub RestoreKeyLocks()
    'DEF SEG = 0                     ' Restore CapLock, NumLock and ScrollLock states
    'POKE(1047, KeyFlags)
    'DEF SEG
  End Sub

  Private Sub ClearKeyLocks()
    'DEF SEG = 0                     ' Turn off CapLock, NumLock and ScrollLock
    'KeyFlags = PEEK(1047)
    'POKE(1047, &H0)
    'DEF SEG
  End Sub

  ' Centers text on given row.
  Sub Center(row%, text$)
    CWrite(text, row - 1, 40 - Len(text) \ 2)
  End Sub

  ' Draws playing field.
  Sub DrawScreen()

    ' Initialize screen.
    ForegroundColor = colorTable(1)
    BackgroundColor = colorTable(4)

    Clear()

    ' Print title & message.
    Center(1, "Nibbles!")
    Center(11, "Initializing Playing Field...")

    ' Initialize arena array.
    For row = 1 To 50
      For col = 1 To 80
        arena(row, col).realRow = (row + 1) \ 2
        arena(row, col).sister = (row Mod 2) * 2 - 1
      Next
    Next

  End Sub

  ' Erases snake to facilitate moving through playing field.
  Sub EraseSnake(snake() As SnakeType, snakeBod(,) As SnakeBody, snakeNum%)

    For c = 0 To 9
      For b = snake(snakeNum).length - c To 0 Step -10
        Dim tail = (snake(snakeNum).head + MAXSNAKELENGTH - b) Mod MAXSNAKELENGTH
        [Set](snakeBod(tail, snakeNum).row, snakeBod(tail, snakeNum).col, colorTable(4))
      Next b
    Next c

  End Sub

  Private Sub CWrite(text$, row%, col%)
    SetCursorPosition(col, row)
    Write(text)
  End Sub

  'Private Sub CWrite(text$, row%, col%, fg As ConsoleColor)
  '  ForegroundColor = fg
  '  SetCursorPosition(col, row)
  '  Write(text)
  'End Sub

  'Private Sub CWrite(text$, row%, col%, fg As ConsoleColor, bg As ConsoleColor)
  '  ForegroundColor = fg
  '  BackgroundColor = bg
  '  SetCursorPosition(col, row)
  '  Write(text)
  'End Sub

  '  Gets player inputs
  Sub GetInputs(ByRef NumPlayers%, ByRef speed%, ByRef diff$, ByRef monitor$)

    ForegroundColor = ConsoleColor.Gray
    BackgroundColor = ConsoleColor.Black

    Clear()

    Dim num$
    Do
      CWrite(Space(34), 4, 46)
      SetCursorPosition(19, 4) : Write("How many players ([1] or 2)? ")
      num = ReadLine()
      If String.IsNullOrEmpty(num) Then
        num = "1"
      End If
    Loop Until Val(num$) = 1 OrElse Val(num$) = 2
    NumPlayers = CInt(Fix(Val(num$)))

    CWrite("Skill level (1 to 100)? ", 7, 20)
    CWrite("[1] = Novice", 8, 21)
    CWrite("90  = Expert", 9, 21)
    CWrite("100 = Twiddle Fingers", 10, 21)
    CWrite("(Computer speed may affect your skill level)", 11, 14)
    Dim gamespeed As String
    Do
      CWrite(Space(35), 7, 44)
      SetCursorPosition(44, 7)
      gamespeed = ReadLine()
      If String.IsNullOrEmpty(gamespeed) Then
        gamespeed = "1"
      End If
    Loop Until Val(gamespeed$) >= 1 AndAlso Val(gamespeed$) <= 100
    speed = CInt(Fix(Val(gamespeed$)))

    speed = (100 - speed) * 2 + 1

    Do
      CWrite(Space(25), 14, 55)
      SetCursorPosition(14, 14)
      Write("Increase game speed during play ([Y] or N)? ")
      diff = ReadLine()
      If String.IsNullOrEmpty(diff) Then
        diff = "Y"
      End If
      diff = UCase(diff)
    Loop Until diff = "Y" OrElse diff = "N"

    Do
      CWrite(Space(34), 16, 45)
      SetCursorPosition(16, 16)
      Write("Monochrome or color monitor (M or [C])? ")
      monitor = ReadLine()
      If String.IsNullOrEmpty(monitor) Then
        monitor = "C"
      End If
      monitor = UCase(monitor)
    Loop Until monitor = "M" OrElse monitor = "C"

  End Sub

  ' Initializes playing field colors.
  Sub InitColors()

    For row = 1 To 50
      For col = 1 To 80
        arena(row, col).acolor = colorTable(4)
      Next col
    Next row

    Clear()

    ' Set (turn on) pixels for screen border.
    For col = 1 To 80
      [Set](3, col, colorTable(3))
      [Set](50, col, colorTable(3))
    Next col

    For row = 4 To 49
      [Set](row, 1, colorTable(3))
      [Set](row, 80, colorTable(3))
    Next row

  End Sub

  ' Displays game introduction.
  Sub Intro()

    ForegroundColor = ConsoleColor.White
    BackgroundColor = ConsoleColor.Black
    Clear()

    Center(4, "Q B a s i c   N i b b l e s")
    ForegroundColor = ConsoleColor.Gray
    Center(6, "Copyright (C) Microsoft Corporation 1990")
    Center(8, "Nibbles is a game for one or two players.  Navigate your snakes")
    Center(9, "around the game board trying to eat up numbers while avoiding")
    Center(10, "running into walls or other snakes.  The more numbers you eat up,")
    Center(11, "the more points you gain and the longer your snake becomes.")
    Center(13, " Game Controls ")
    Center(15, "  General             Player 1               Player 2    ")
    Center(16, "                        (Up)                   (Up)      ")
    Center(17, "P - Pause                ↑                      W       ")
    Center(18, "                     (Left) ←   → (Right)   (Left) A   D (Right)  ")
    Center(19, "                         ↓                      S       ")
    Center(20, "                       (Down)                 (Down)     ")
    Center(24, "Press any key to continue")

    'PLAY("MBT160O1L8CDEDCDL4ECC")

    SparklePause()

  End Sub

  ' Sets game level.
  Sub Level(whatToDo%, sammy() As SnakeType)

    Select Case (whatToDo)

      Case STARTOVER
        curLevel = 1
      Case NEXTLEVEL
        curLevel += 1
    End Select

    ' Initialize Snakes.
    sammy(1).head = 1
    sammy(1).length = 2
    sammy(1).alive = True
    sammy(2).head = 1
    sammy(2).length = 2
    sammy(2).alive = True

    InitColors()

    Select Case curLevel
      Case 1
        sammy(1).row = 25 : sammy(2).row = 25
        sammy(1).col = 50 : sammy(2).col = 30
        sammy(1).direction = 4 : sammy(2).direction = 3


      Case 2
        For i = 20 To 60
          [Set](25, i, colorTable(3))
        Next
        sammy(1).row = 7 : sammy(2).row = 43
        sammy(1).col = 60 : sammy(2).col = 20
        sammy(1).direction = 3 : sammy(2).direction = 4

      Case 3
        For i = 10 To 40
          [Set](i, 20, colorTable(3))
          [Set](i, 60, colorTable(3))
        Next
        sammy(1).row = 25 : sammy(2).row = 25
        sammy(1).col = 50 : sammy(2).col = 30
        sammy(1).direction = 1 : sammy(2).direction = 2

      Case 4
        For i = 4 To 30
          [Set](i, 20, colorTable(3))
          [Set](53 - i, 60, colorTable(3))
        Next
        For i = 2 To 40
          [Set](38, i, colorTable(3))
          [Set](15, 81 - i, colorTable(3))
        Next
        sammy(1).row = 7 : sammy(2).row = 43
        sammy(1).col = 60 : sammy(2).col = 20
        sammy(1).direction = 3 : sammy(2).direction = 4

      Case 5
        For i = 13 To 39
          [Set](i, 21, colorTable(3))
          [Set](i, 59, colorTable(3))
        Next
        For i = 23 To 57
          [Set](11, i, colorTable(3))
          [Set](41, i, colorTable(3))
        Next
        sammy(1).row = 25 : sammy(2).row = 25
        sammy(1).col = 50 : sammy(2).col = 30
        sammy(1).direction = 1 : sammy(2).direction = 2

      Case 6
        For i = 4 To 49
          If i > 30 Or i < 23 Then
            [Set](i, 10, colorTable(3))
            [Set](i, 20, colorTable(3))
            [Set](i, 30, colorTable(3))
            [Set](i, 40, colorTable(3))
            [Set](i, 50, colorTable(3))
            [Set](i, 60, colorTable(3))
            [Set](i, 70, colorTable(3))
          End If
        Next
        sammy(1).row = 7 : sammy(2).row = 43
        sammy(1).col = 65 : sammy(2).col = 15
        sammy(1).direction = 2 : sammy(2).direction = 1

      Case 7
        For i = 4 To 49 Step 2
          [Set](i, 40, colorTable(3))
        Next
        sammy(1).row = 7 : sammy(2).row = 43
        sammy(1).col = 65 : sammy(2).col = 15
        sammy(1).direction = 2 : sammy(2).direction = 1

      Case 8
        For i = 4 To 40
          [Set](i, 10, colorTable(3))
          [Set](53 - i, 20, colorTable(3))
          [Set](i, 30, colorTable(3))
          [Set](53 - i, 40, colorTable(3))
          [Set](i, 50, colorTable(3))
          [Set](53 - i, 60, colorTable(3))
          [Set](i, 70, colorTable(3))
        Next
        sammy(1).row = 7 : sammy(2).row = 43
        sammy(1).col = 65 : sammy(2).col = 15
        sammy(1).direction = 2 : sammy(2).direction = 1

      Case 9
        For i = 6 To 47
          [Set](i, i, colorTable(3))
          [Set](i, i + 28, colorTable(3))
        Next
        sammy(1).row = 40 : sammy(2).row = 15
        sammy(1).col = 75 : sammy(2).col = 5
        sammy(1).direction = 1 : sammy(2).direction = 2

      Case Else
        For i = 4 To 49 Step 2
          [Set](i, 10, colorTable(3))
          [Set](i + 1, 20, colorTable(3))
          [Set](i, 30, colorTable(3))
          [Set](i + 1, 40, colorTable(3))
          [Set](i, 50, colorTable(3))
          [Set](i + 1, 60, colorTable(3))
          [Set](i, 70, colorTable(3))
        Next
        sammy(1).row = 7 : sammy(2).row = 43
        sammy(1).col = 65 : sammy(2).col = 15
        sammy(1).direction = 2 : sammy(2).direction = 1

    End Select
  End Sub

  '  Main routine that controls game play.
  Sub PlayNibbles(numPlayers%, ByRef speed%, diff$)

    ' Initialize Snakes.
    Dim sammyBody(MAXSNAKELENGTH - 1, 2) As SnakeBody
    Dim sammy(2) As SnakeType
    sammy(1).lives = 5
    sammy(1).score = 0
    sammy(1).scolor = colorTable(1)
    sammy(2).lives = 5
    sammy(2).score = 0
    sammy(2).scolor = colorTable(2)

    Level(STARTOVER, sammy)
    'startRow1 = sammy(1).row : startCol1 = sammy(1).col
    'startRow2 = sammy(2).row : startCol2 = sammy(2).col

    Dim curSpeed = speed

    ' Play Nibbles until finished.

    SpacePause("     Level" + Str(curLevel) + ",  Push Space")
    'gameOver = False
    Do

      If numPlayers = 1 Then
        sammy(2).row = 0
      End If

      Dim number = 1          'Current number that snakes are trying to run into
      Dim nonum = True        'nonum = TRUE if a number is not on the screen

      Dim playerDied = False
      PrintScore(numPlayers, sammy(1).score, sammy(2).score, sammy(1).lives, sammy(2).lives)
      'PLAY("T160O1>L20CDEDCDL10ECC")
      Beep()

      Dim numberRow%
      Dim numberCol%

      Do
        'Print number if no number exists
        If nonum = True Then
          Dim sisterRow%
          Do
            numberRow = CInt(Fix(Int(Rnd(1) * 47 + 3)))
            numberCol = CInt(Fix(Int(Rnd(1) * 78 + 2)))
            sisterRow = numberRow + arena(numberRow, numberCol).sister
          Loop Until Not PointIsThere(numberRow, numberCol, colorTable(4)) And Not PointIsThere(sisterRow, numberCol, colorTable(4))
          numberRow = arena(numberRow, numberCol).realRow
          nonum = False
          ForegroundColor = colorTable(1)
          BackgroundColor = colorTable(4)
          CWrite(Right(Str(number), 1), numberRow - 1, numberCol - 1)
          'Dim count = 0
        End If

        ' Delay game.
        Threading.Thread.Sleep(curSpeed)

        ' Get keyboard input & Change direction accordingly.

        If Console.KeyAvailable Then

          Dim ki = Console.ReadKey(True)

          Select Case ki.Key
            Case ConsoleKey.W : If sammy(2).direction <> 2 Then sammy(2).direction = 1
            Case ConsoleKey.S : If sammy(2).direction <> 1 Then sammy(2).direction = 2
            Case ConsoleKey.A : If sammy(2).direction <> 4 Then sammy(2).direction = 3
            Case ConsoleKey.D : If sammy(2).direction <> 3 Then sammy(2).direction = 4
            Case ConsoleKey.UpArrow : If sammy(1).direction <> 2 Then sammy(1).direction = 1
            Case ConsoleKey.DownArrow : If sammy(1).direction <> 1 Then sammy(1).direction = 2
            Case ConsoleKey.LeftArrow : If sammy(1).direction <> 4 Then sammy(1).direction = 3
            Case ConsoleKey.RightArrow : If sammy(1).direction <> 3 Then sammy(1).direction = 4
            Case ConsoleKey.P : SpacePause(" Game Paused ... Push Space  ")
            Case Else
          End Select

        End If

        For A = 1 To numPlayers
          ' Move Snake.
          Select Case sammy(A).direction
            Case 1 : sammy(A).row = sammy(A).row - 1
            Case 2 : sammy(A).row = sammy(A).row + 1
            Case 3 : sammy(A).col = sammy(A).col - 1
            Case 4 : sammy(A).col = sammy(A).col + 1
          End Select

          ' If snake hits number, respond accordingly.
          If numberRow = Int((sammy(A).row + 1) \ 2) And numberCol = sammy(A).col Then
            ' PLAY("MBO0L16>CCCE") 
            Beep()
            If sammy(A).length < (MAXSNAKELENGTH - 30) Then
              sammy(A).length = sammy(A).length + number * 4
            End If
            sammy(A).score = sammy(A).score + number
            PrintScore(numPlayers, sammy(1).score, sammy(2).score, sammy(1).lives, sammy(2).lives)
            number += 1
            If number = 10 Then
              EraseSnake(sammy, sammyBody, 1)
              EraseSnake(sammy, sammyBody, 2)
              CWrite(" ", numberRow - 1, numberCol - 1)
              Level(NEXTLEVEL, sammy)
              PrintScore(numPlayers, sammy(1).score, sammy(2).score, sammy(1).lives, sammy(2).lives)
              SpacePause("     Level" + Str(curLevel) + ",  Push Space")
              If numPlayers = 1 Then sammy(2).row = 0
              number = 1
              If diff$ = "Y" Then speed -= 10 : curSpeed = speed
            End If
            nonum = True
            'If curSpeed < 1 Then curSpeed = 1
            If curSpeed < 50 Then curSpeed = 50
          End If
        Next

        For a = 1 To numPlayers
          ' If player runs into any point, or the head of the other snake, it dies. 
          If PointIsThere(sammy(a).row, sammy(a).col, colorTable(4)) OrElse (sammy(1).row = sammy(2).row And sammy(1).col = sammy(2).col) Then
            'PLAY("MBO0L32EFGEFDC")
            Beep()
            BackgroundColor = colorTable(4)
            CWrite(" ", numberRow - 1, numberCol - 1)

            playerDied = True
            sammy(a).alive = False
            sammy(a).lives = sammy(a).lives - 1
          Else
            ' Otherwise, move the snake, and erase the tail.
            sammy(a).head = (sammy(a).head + 1) Mod MAXSNAKELENGTH
            sammyBody(sammy(a).head, a).row = sammy(a).row
            sammyBody(sammy(a).head, a).col = sammy(a).col
            Dim tail = (sammy(a).head + MAXSNAKELENGTH - sammy(a).length) Mod MAXSNAKELENGTH
            [Set](sammyBody(tail, a).row, sammyBody(tail, a).col, colorTable(4))
            sammyBody(tail, a).row = 0
            [Set](sammy(a).row, sammy(a).col, sammy(a).scolor)
          End If
        Next

      Loop Until playerDied

      ' Reset speed to initial value.
      curSpeed = speed

      For a = 1 To numPlayers

        EraseSnake(sammy, sammyBody, a)

        ' If dead, then erase snake in really cool way.
        If Not sammy(a).alive Then

          ' Update score.
          sammy(a).score = sammy(a).score - 10
          PrintScore(numPlayers, sammy(1).score, sammy(2).score, sammy(1).lives, sammy(2).lives)

          If a = 1 Then
            SpacePause(" Sammy Dies! Push Space! --->")
          Else
            SpacePause(" <---- Jake Dies! Push Space ")
          End If

        End If

      Next

      Level(SAMELEVEL, sammy)
      PrintScore(numPlayers, sammy(1).score, sammy(2).score, sammy(1).lives, sammy(2).lives)

      ' Play next round, until either of snake's lives have run out..

    Loop Until sammy(1).lives = 0 OrElse sammy(2).lives = 0

  End Sub

  ' Checks the global  arena array to see if the boolean flag is set.
  Function PointIsThere(row%, col%, acolor As ConsoleColor) As Boolean
    If row <> 0 Then
      If arena(row, col).acolor <> acolor Then
        Return True
      Else
        Return False
      End If
    Else
      Return False
    End If
  End Function

  ' Prints players scores and number of lives remaining.
  Sub PrintScore(numPlayers%, score1%, score2%, lives1%, lives2%)

    ForegroundColor = ConsoleColor.White
    BackgroundColor = colorTable(4)

    If numPlayers = 2 Then
      CWrite($"{score2:N}  Lives: {lives2}  <--JAKE", 0, 0)
    End If

    CWrite($"SAMMY-->  Lives:  {lives1}     {score1:N}", 0, 48)

  End Sub

  '  Sets row and column on playing field to given color to facilitate moving
  '  of snakes around the field.
  Sub [Set](row%, col%, acolor As ConsoleColor)

    If row <> 0 Then

      ' Assign color to arena.
      arena(row, col).acolor = acolor
      ' Get real row of pixel.
      Dim realRow% = arena(row, col).realRow
      ' Get arena row of sister.
      Dim sisterRow = row + arena(row, col).sister
      ' Determine sister's color.
      Dim sisterColor = arena(sisterRow, col).acolor

      SetCursorPosition(col - 1, realRow - 1)

      If acolor = sisterColor Then
        ' If both points are same Print "█".
        ForegroundColor = acolor
        BackgroundColor = acolor
        Write("█")
      Else
        ' Deduce whether pixel is on top, or bottom.
        If arena(row, col).sister = 1 Then
          ' Since you cannot have bright backgrounds determine best combo to use.
          If acolor > 7 Then
            ForegroundColor = acolor
            BackgroundColor = sisterColor
            Write("▀")
          Else
            ForegroundColor = sisterColor
            BackgroundColor = acolor
            Write("▄")
          End If
        Else
          If acolor > 7 Then
            ForegroundColor = acolor
            BackgroundColor = sisterColor
            Write("▄")
          Else
            ForegroundColor = sisterColor
            BackgroundColor = acolor
            Write("▀")
          End If
        End If
      End If
    End If
  End Sub

  ' Pauses game play and waits for space bar to be pressed before continuing.
  Private Sub SpacePause(text$)

    ' Adjust the size of text to fit the box.
    text = Left(text$ + Space(28), 28)

    ' Draw the box / message.
    ForegroundColor = colorTable(5)
    BackgroundColor = colorTable(6)
    Center(11, "████████████████████████████████")
    Center(12, "█ " + text + " █")
    Center(13, "████████████████████████████████")

    Do While Console.KeyAvailable
      Console.ReadKey(True)
    Loop

    Do
      If Console.KeyAvailable Then
        Dim ki = Console.ReadKey(True)
        Select Case ki.Key
          Case ConsoleKey.Spacebar : Exit Do
          Case Else
        End Select
      End If
    Loop

    ForegroundColor = ConsoleColor.White
    BackgroundColor = colorTable(4)

    ' Restore the screen background.
    For i = 21 To 26
      For j = 24 To 56
        [Set](i, j, arena(i, j).acolor)
      Next
    Next

  End Sub

  ' Creates flashing border for intro screen.
  Private Sub SparklePause()

    ForegroundColor = ConsoleColor.Red
    BackgroundColor = ConsoleColor.Black

    Dim lights = "*    *    *    *    *    *    *    *    *    *    *    *    *    *    *    *    *    "

    ' Clear keyboard buffer.

    Do While Console.KeyAvailable
      Console.ReadKey(True)
    Loop

    Do

      If Console.KeyAvailable Then ' Any key to continue.
        Console.ReadKey(True)
        Exit Do
      End If

      For a = 1 To 5

        ' Print horizontal sparkles.
        CWrite(Mid(lights, a, 80), 0, 0)
        CWrite(Mid(lights, 6 - a, 80), 21, 0)

        ' Print Vertical sparkles.
        For b = 2 To 21
          Dim c = (a + b) Mod 5
          If c = 1 Then
            CWrite("*", b - 1, 79)
            CWrite("*", 23 - b - 1, 0)
          Else
            CWrite(" ", b - 1, 79)
            CWrite(" ", 23 - b - 1, 0)
          End If
        Next

        Threading.Thread.Sleep(100)

      Next

    Loop

  End Sub

  '  Determines if users want to play game again.
  Private Function StillWantsToPlay() As Boolean

    ForegroundColor = colorTable(5)
    BackgroundColor = colorTable(6)
    Center(10, "█████████████████████████████████")
    Center(11, "█       G A M E   O V E R       █")
    Center(12, "█                               █")
    Center(13, "█      Play Again?   (Y/N)      █")
    Center(14, "█████████████████████████████████")

    Do While Console.KeyAvailable
      Console.ReadKey(True)
    Loop

    Dim again As Boolean
    Do
      If Console.KeyAvailable Then
        Dim ki = Console.ReadKey(True)
        Select Case ki.Key
          Case ConsoleKey.Y : again = True : Exit Do
          Case ConsoleKey.N : again = False : Exit Do
          Case Else
        End Select
      End If
    Loop

    ForegroundColor = ConsoleColor.White
    BackgroundColor = colorTable(4)
    Center(10, "                                 ")
    Center(11, "                                 ")
    Center(12, "                                 ")
    Center(13, "                                 ")
    Center(14, "                                 ")

    If Not again Then
      ForegroundColor = ConsoleColor.Gray
      BackgroundColor = ConsoleColor.Black
      Clear()
    End If

    Return again

  End Function

  Public Sub Resize(cols%, rows%)
    If OperatingSystem.IsWindows Then
      ' HACK!!! Resize the console window attempting to remove scroll bars.
      Try
        Console.SetWindowSize(cols%, rows%) ' Set the windows size...
      Catch ex As Exception
        Console.WriteLine("1 - " & ex.ToString)
        Threading.Thread.Sleep(5000)
      End Try
      Try
        Console.SetBufferSize(cols%, rows%) ' Then set the buffer size to the now window size...
      Catch ex As Exception
        Console.WriteLine("2 - " & ex.ToString)
        Threading.Thread.Sleep(5000)
      End Try
      Try
        Console.SetWindowSize(cols%, rows%) ' Then set the window size again so that the scroll bar area is removed.
      Catch ex As Exception
        Console.WriteLine("3 - " & ex.ToString)
        Threading.Thread.Sleep(5000)
      End Try
    End If
  End Sub

End Module