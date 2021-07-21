Imports System

Module Program
    'Variables for board set up
    Dim Rows As Integer
    Dim Columns As Integer
    Dim Blank As Char = "."
    Dim Board(0 To 5, 0 To 6) As Char

    'Variables for game set up
    Dim ThisPlayer As Char 'stores the character of the player whose turn it is
    Dim GameOver As Boolean
    Dim FreeRow As Integer
    Dim WinnerFound As Boolean
    Dim IsFull As Boolean

    'Variables for choosing a column
    Dim ChosenColumn As Integer
    Dim ValidColumnNumber As Boolean

    'Variables to check for a winner
    Dim Winner As String

    Sub BoardSetUp()
        For Rows = 0 To 5
            For Columns = 0 To 6
                Board(Rows, Columns) = Blank
            Next
        Next
    End Sub
    Sub SetUpGame()
        ThisPlayer = "X" 'Means player X always starts
        FreeRow = 5
        GameOver = False
        WinnerFound = False
        IsFull = False
    End Sub
    Sub OutputBoard()
        Console.Write(" 123456")
        Console.WriteLine()
        For Rows = 0 To 5
            Console.Write(Rows + 1)

            For Columns = 0 To 6
                Console.Write(Board(Rows, Columns))
            Next
            Console.WriteLine()
        Next
        Console.WriteLine()
    End Sub

    Sub ChooseColumn()
        Console.WriteLine("Player " & ThisPlayer & "'s turn")
        Console.WriteLine("Which column would you like to add to")
        ChosenColumn = Console.ReadLine

        If ChosenColumn > 7 Or ChosenColumn < 1 Then
            Do
                Console.WriteLine("Please enter a valid column")
                ChosenColumn = Console.ReadLine
            Loop Until ChosenColumn >= 1 And ChosenColumn <= 7


        End If

        ChosenColumn = ChosenColumn - 1 'This is cause the array is 0 to 6
    End Sub

    Sub FindNextSlot() 'This subroutine makes sure the characters are added on top of each other
        If Board(FreeRow, ChosenColumn) <> Blank Then
            Do
                FreeRow = FreeRow - 1
                If FreeRow < 0 Then
                    Do 'This loop ensures people dont try add to a column thats full
                        Console.WriteLine("This column is full. Please choose a valid column")
                        ChooseColumn() 'Reenter a column thats valid
                        FreeRow = 5 'so it starts searching the bottom of the column again
                    Loop Until FreeRow > -1
                End If
            Loop Until Board(FreeRow, ChosenColumn) = Blank
        End If
    End Sub

    Sub StoreMove()
        Board(FreeRow, ChosenColumn) = ThisPlayer
        FreeRow = 5
    End Sub
    Sub SwapTurns()
        If ThisPlayer = "X" Then
            ThisPlayer = "O"
        Else
            ThisPlayer = "X"
        End If
    End Sub
    Sub CheckForWinner()
        Dim RowCheck As Integer = 0
        Dim ColumnCheck As Integer = 0
        Dim CountRow As Integer
        Dim CountColumn As Integer

        'Checking rows
        For CountRow = 0 To 5
            For CountColumn = 0 To 6
                If Board(RowCheck, ColumnCheck) = ThisPlayer And Board(RowCheck, ColumnCheck + 1) = ThisPlayer And Board(RowCheck, ColumnCheck + 2) = ThisPlayer And Board(RowCheck, ColumnCheck + 3) = ThisPlayer And Board(RowCheck, ColumnCheck + 4) = ThisPlayer And Board(RowCheck, ColumnCheck + 5) = ThisPlayer Then
                    WinnerFound = True
                    Winner = ThisPlayer
                End If
                RowCheck = RowCheck + 1

                If RowCheck = 6 Then
                    RowCheck = RowCheck - 1
                End If
            Next
        Next

        'Checking Columns
        If WinnerFound <> True Then
            RowCheck = 0
            For CountColumn = 0 To 6
                For CountRow = 0 To 5
                    If Board(RowCheck, ColumnCheck) = ThisPlayer And Board(RowCheck + 1, ColumnCheck) = ThisPlayer And Board(RowCheck + 2, ColumnCheck) = ThisPlayer And Board(RowCheck + 3, ColumnCheck) = ThisPlayer And Board(RowCheck + 4, ColumnCheck) = ThisPlayer And Board(RowCheck + 5, ColumnCheck) = ThisPlayer Then
                        WinnerFound = True
                        Winner = ThisPlayer
                    End If
                    ColumnCheck = ColumnCheck + 1

                    If RowCheck >= 6 Then
                        RowCheck = 5
                    End If

                    If ColumnCheck >= 7 Then
                        ColumnCheck = 6
                    End If
                Next

            Next
        End If
    End Sub
    Sub CheckFullBoard()
        Dim Inner As Integer
        Dim Outer As Integer

        IsFull = True
        For Outer = 0 To 5
            For Inner = 0 To 6
                If Board(Outer, Inner) = Blank Then
                    IsFull = False
                End If
            Next
        Next
    End Sub

    Sub IsGameOver()
        If IsFull = True Then
            GameOver = True
            Winner = "Draw"
        ElseIf WinnerFound = True Then
            GameOver = True
        End If
    End Sub
    Sub Main(args As String())

        SetUpGame()
        BoardSetUp()
        OutputBoard()

        Do
            ChooseColumn()
            FindNextSlot()
            StoreMove()
            OutputBoard()
            CheckForWinner()
            CheckFullBoard()
            IsGameOver()
            If GameOver = False Then
                SwapTurns()
            End If
        Loop Until GameOver = True
        Console.WriteLine()
        Console.WriteLine("The game is over.")
        Console.WriteLine("The winner is " & Winner)
        Console.ReadKey()
    End Sub
End Module
