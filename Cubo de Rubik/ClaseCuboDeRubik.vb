Option Explicit On
Option Strict On

Imports Cubo_de_Rubik.Utilidades

Public Class ClaseCuboDeRubik
    Public ArrayOfRubiksCube(5) As Integer
    '   Private FixedPairs(11) As Boolean
    '    Private ObjectivePairs(11) As Boolean
    '    Private FixedCorners(11) As Boolean
    '    Private ObjectiveCorners(11) As Boolean
    Public ListOfNeededMovements As String
    Public InitialArray(5) As Integer
    Public ListOfMovements() As Integer 'This is the array on which we will base ourselves

    '    Public ArrayOfRepeatedMovements() As Integer


    Public Sub New(AuxiliarArrayOfRubiksCube() As Integer)
        Dim Counter As Integer
        For Counter = 0 To 5
            ArrayOfRubiksCube(Counter) = AuxiliarArrayOfRubiksCube(Counter)
            InitialArray(Counter) = AuxiliarArrayOfRubiksCube(Counter)
        Next
        ListOfNeededMovements = vbNullString
    End Sub

    'Public Sub New(AuxiliarArrayOfRubiksCube() As Integer, ListOfMovementsSoFar As String)
    '    Dim Counter As Integer
    '    For Counter = 0 To 5
    '        ArrayOfRubiksCube(Counter) = AuxiliarArrayOfRubiksCube(Counter)
    '        InitialArray(Counter) = AuxiliarArrayOfRubiksCube(Counter)
    '    Next
    '    ListOfNeededMovements = ListOfMovementsSoFar
    'End Sub

    Public Sub New()

    End Sub


    Public Function IsTheFaceSolved(FaceNumber As Integer) As Boolean
        Select Case FaceNumber
            Case 0 To 5
                Dim Counter As Integer
                For Counter = 0 To 8
                    If SquareColour(Counter, ArrayOfRubiksCube(FaceNumber)) <> FaceNumber Then Return False
                Next
                Return True
            Case Else
                WeTerminateWithError(22) : Stop
        End Select
    End Function

    Public Sub Rotate90UpFaceClockwise()   'Arriba girar a la izquierda
        AddElementToArray(0, ListOfMovements)
        Dim ArrayOfAddends(5) As Integer
        ArrayOfAddends(0) += CType(6 ^ 0, Integer) * SquareColour(2, ArrayOfRubiksCube(4)) + CType(6 ^ 1, Integer) * SquareColour(5, ArrayOfRubiksCube(4)) +
            CType(6 ^ 2, Integer) * SquareColour(8, ArrayOfRubiksCube(4))
        ArrayOfAddends(0) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(0)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(0)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(0))
        ArrayOfAddends(1) = 0
        ArrayOfAddends(2) = SquareColour(2, ArrayOfRubiksCube(0)) + CType(6 ^ 3, Integer) * SquareColour(1, ArrayOfRubiksCube(0)) +
            CType(6 ^ 6, Integer) * SquareColour(0, ArrayOfRubiksCube(0))
        ArrayOfAddends(2) -= SquareColour(0, ArrayOfRubiksCube(2)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(2)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(2))
        ArrayOfAddends(3) = ColourConfigurationOfTheFaceOnceRotated90Clockwise(ArrayOfRubiksCube(3))
        ArrayOfAddends(3) -= ArrayOfRubiksCube(3)
        ArrayOfAddends(4) = CType(6 ^ 2, Integer) * SquareColour(8, ArrayOfRubiksCube(5)) + CType(6 ^ 5, Integer) * SquareColour(7, ArrayOfRubiksCube(5)) +
            CType(6 ^ 8, Integer) * SquareColour(6, ArrayOfRubiksCube(5))
        ArrayOfAddends(4) -= CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(4)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(4)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(4))
        ArrayOfAddends(5) += CType(6 ^ 6, Integer) * SquareColour(0, ArrayOfRubiksCube(2)) + CType(6 ^ 7, Integer) * SquareColour(3, ArrayOfRubiksCube(2)) +
            CType(6 ^ 8, Integer) * SquareColour(6, ArrayOfRubiksCube(2))
        ArrayOfAddends(5) -= CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(5)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(5)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(5))
        ApplyAddends(ArrayOfAddends)
    End Sub

    Public Sub Rotate90UpFaceClockwise(FrontFace%, DownFace%)
        Select Case DownFace
            Case 1
                Select Case FrontFace
                    Case 0, 2, 4, 5 : Rotate90UpFaceClockwise()
                    Case 1, 3 : WeTerminateWithError(21) : Stop
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 3
                Select Case FrontFace
                    Case 0, 2, 4, 5 : Rotate90DownFaceClockwise()
                    Case 1, 3 : WeTerminateWithError(21) : Stop
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 0, 2, 4, 5
                If FrontFace = DownFace Or FrontFace = BackFace(DownFace) Then WeTerminateWithError(21) : Stop
                Select Case 6 * DownFace + FrontFace
                    Case 0, 5, 2 * 6 + 2, 2 * 6 + 4, 4 * 6 + 2, 4 * 6 + 4, 5 * 6 + 0, 5 * 6 + 5 : WeTerminateWithError(22) : Stop
                    Case 1 To 4 : Rotate90BackFaceClockwise()
                    Case 2 * 6 + 0, 2 * 6 + 1, 2 * 6 + 3, 2 * 6 + 5 : Rotate90RightFaceClockwise()
                    Case 4 * 6 + 0, 4 * 6 + 1, 4 * 6 + 3, 4 * 6 + 5 : Rotate90LeftFaceClockwise()
                    Case 5 * 6 + 1 To 5 * 6 + 4 : Rotate90FrontFaceClockwise()
                    Case Else
                        Select Case 6 * DownFace + FrontFace
                            Case 0 To 35 : WeTerminateWithError(22) : Stop
                            Case Is > 35 : WeTerminateWithError(5) : Stop
                        End Select
                End Select
            Case Else
                WeTerminateWithError(5) : Stop
        End Select
    End Sub

    Public Sub Rotate90UpFaceCounterClockwise() ' Arriba girar a la derecha
        Rotate90UpFaceClockwise() : Rotate90UpFaceClockwise() : Rotate90UpFaceClockwise()
    End Sub

    Public Sub Rotate90UpFaceCounterClockwise(FrontFace%, DownFace%)
        Rotate90UpFaceClockwise(FrontFace, DownFace) : Rotate90UpFaceClockwise(FrontFace, DownFace) : Rotate90UpFaceClockwise(FrontFace, DownFace)
    End Sub

    Public Sub Rotate90DownFaceCounterClockwise()    ' Abajo girar a la izquierda
        AddElementToArray(2, ListOfMovements)
        Dim ArrayOfAddends(5) As Integer
        ArrayOfAddends(0) += CType(6 ^ 6, Integer) * SquareColour(0, ArrayOfRubiksCube(4)) + CType(6 ^ 7, Integer) * SquareColour(3, ArrayOfRubiksCube(4)) +
            CType(6 ^ 8, Integer) * SquareColour(6, ArrayOfRubiksCube(4))
        ArrayOfAddends(0) -= CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(0)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(0)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(0))
        ArrayOfAddends(1) += ColourConfigurationOfTheFaceOnceRotated90CounterClockwise(ArrayOfRubiksCube(1))
        ArrayOfAddends(1) -= ArrayOfRubiksCube(1)
        ArrayOfAddends(2) += CType(6 ^ 2, Integer) * SquareColour(8, ArrayOfRubiksCube(0)) + CType(6 ^ 5, Integer) * SquareColour(7, ArrayOfRubiksCube(0)) +
            CType(6 ^ 8, Integer) * SquareColour(6, ArrayOfRubiksCube(0))
        ArrayOfAddends(2) -= CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(2)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(2)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(2))
        ArrayOfAddends(3) += 0
        ArrayOfAddends(4) += CType(6 ^ 0, Integer) * SquareColour(2, ArrayOfRubiksCube(5)) + CType(6 ^ 3, Integer) * SquareColour(1, ArrayOfRubiksCube(5)) +
            CType(6 ^ 6, Integer) * SquareColour(0, ArrayOfRubiksCube(5))
        ArrayOfAddends(4) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(4)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(4)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(4))
        ArrayOfAddends(5) += CType(6 ^ 0, Integer) * SquareColour(2, ArrayOfRubiksCube(2)) + CType(6 ^ 1, Integer) * SquareColour(5, ArrayOfRubiksCube(2)) +
            CType(6 ^ 2, Integer) * SquareColour(8, ArrayOfRubiksCube(2))
        ArrayOfAddends(5) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(5)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(5)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(5))
        ApplyAddends(ArrayOfAddends)
    End Sub

    Public Sub Rotate90DownFaceCounterClockwise(FrontFace%, DownFace%)
        Select Case DownFace
            Case 1
                Select Case FrontFace
                    Case 0, 2, 4, 5 : Rotate90DownFaceCounterClockwise()
                    Case 1, 3 : WeTerminateWithError(21) : Stop
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 3
                Select Case FrontFace
                    Case 0, 2, 4, 5 : Rotate90UpFaceCounterClockwise()
                    Case 1, 3 : WeTerminateWithError(21) : Stop
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 0, 2, 4, 5
                If DownFace = FrontFace Or BackFace(DownFace) = FrontFace Then WeTerminateWithError(21)
                Select Case 6 * DownFace + FrontFace
                    Case 0, 5 : WeTerminateWithError(22) : Stop
                    Case 1 To 4 : Rotate90FrontFaceCounterClockwise()

                    Case 2 * 6 + 0, 2 * 6 + 1, 2 * 6 + 3, 2 * 6 + 5 : Rotate90LeftFaceCounterClockwise()
                    Case 2 * 6 + 2, 2 * 6 + 4 : WeTerminateWithError(22) : Stop

                    Case 4 * 6 + 0, 4 * 6 + 1, 4 * 6 + 3, 4 * 6 + 5 : Rotate90RightFaceCounterClockwise()
                    Case 4 * 6 + 2, 4 * 6 + 4 : WeTerminateWithError(22) : Stop

                    Case 5 + 6 * 0, 5 * 6 + 5 : WeTerminateWithError(22) : Stop
                    Case 5 * 6 + 1, 5 * 6 + 2, 5 * 6 + 3, 5 * 6 + 4 : Rotate90BackFaceCounterClockwise()

                    Case Else
                        Select Case 6 * DownFace + FrontFace
                            Case 0 To 35 : WeTerminateWithError(22) : Stop
                            Case Else : WeTerminateWithError(5) : Stop
                        End Select
                End Select
            Case Else
                WeTerminateWithError(5) : Stop
        End Select
    End Sub

    Public Sub Rotate90DownFaceClockwise()  ' Abajo girar a la derecha
        Rotate90DownFaceCounterClockwise() : Rotate90DownFaceCounterClockwise() : Rotate90DownFaceCounterClockwise()
    End Sub

    Public Sub Rotate90DownFaceClockwise(FrontFace%, DownFace%)
        Rotate90DownFaceCounterClockwise(FrontFace, DownFace) : Rotate90DownFaceCounterClockwise(FrontFace, DownFace) : Rotate90DownFaceCounterClockwise(FrontFace, DownFace)
    End Sub

    Public Sub Rotate90LeftFaceCounterClockwise()   ' Izquierda girar arriba
        AddElementToArray(4, ListOfMovements)
        Dim ArrayOfAddends(5) As Integer
        ArrayOfAddends(0) += CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(1)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(1)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(1))
        ArrayOfAddends(0) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(0)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(0)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(0))
        ArrayOfAddends(1) += CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(5)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(5)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(5))
        ArrayOfAddends(1) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(1)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(1)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(1))
        ArrayOfAddends(2) += ColourConfigurationOfTheFaceOnceRotated90CounterClockwise(ArrayOfRubiksCube(2))
        ArrayOfAddends(2) -= ArrayOfRubiksCube(2)
        ArrayOfAddends(3) += CType(6 ^ 2, Integer) * SquareColour(6, ArrayOfRubiksCube(0)) + CType(6 ^ 5, Integer) * SquareColour(3, ArrayOfRubiksCube(0)) +
            CType(6 ^ 8, Integer) * SquareColour(0, ArrayOfRubiksCube(0))
        ArrayOfAddends(3) -= CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(3)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(3)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(3))
        ArrayOfAddends(4) = 0
        ArrayOfAddends(5) += CType(6 ^ 0, Integer) * SquareColour(8, ArrayOfRubiksCube(3)) + CType(6 ^ 3, Integer) * SquareColour(5, ArrayOfRubiksCube(3)) +
            CType(6 ^ 6, Integer) * SquareColour(2, ArrayOfRubiksCube(3))
        ArrayOfAddends(5) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(5)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(5)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(5))
        ApplyAddends(ArrayOfAddends)
    End Sub

    Public Sub Rotate90LeftFaceCounterClockwise(FrontFace%, DownFace%)
        Select Case DownFace
            Case 2
                Select Case FrontFace
                    Case 0 : Rotate90UpFaceCounterClockwise()
                    Case 1 : Rotate90FrontFaceCounterClockwise()
                    Case 2, 4 : WeTerminateWithError(22) : Stop
                    Case 3 : Rotate90BackFaceCounterClockwise()
                    Case 5 : Rotate90DownFaceCounterClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 4
                Select Case FrontFace
                    Case 0 : Rotate90DownFaceCounterClockwise()
                    Case 1 : Rotate90BackFaceCounterClockwise()
                    Case 2, 4 : WeTerminateWithError(22) : Stop
                    Case 3 : Rotate90FrontFaceCounterClockwise()
                    Case 5 : Rotate90UpFaceCounterClockwise()
                End Select
            Case 0, 1, 3, 5
                If FrontFace = DownFace Or FrontFace = BackFace(DownFace) Then WeTerminateWithError(21) : Stop
                Select Case 6 * DownFace + FrontFace
                    Case 0, 5 : WeTerminateWithError(22) : Stop
                    Case 1 : Rotate90RightFaceCounterClockwise()
                    Case 2 : Rotate90DownFaceCounterClockwise()
                    Case 3 : Rotate90LeftFaceCounterClockwise()
                    Case 4 : Rotate90UpFaceCounterClockwise()

                    Case 1 * 6 + 0 : Rotate90LeftFaceCounterClockwise()
                    Case 1 * 6 + 1, 1 * 6 + 3 : WeTerminateWithError(22) : Stop
                    Case 1 * 6 + 2 : Rotate90BackFaceCounterClockwise()
                    Case 1 * 6 + 4 : Rotate90FrontFaceCounterClockwise()
                    Case 1 * 6 + 5 : Rotate90RightFaceCounterClockwise()

                    Case 3 * 6 + 0 : Rotate90RightFaceCounterClockwise()
                    Case 3 * 6 + 1, 3 * 6 + 3 : WeTerminateWithError(22) : Stop
                    Case 3 * 6 + 2 : Rotate90FrontFaceCounterClockwise()
                    Case 3 * 6 + 4 : Rotate90BackFaceCounterClockwise()
                    Case 3 * 6 + 5 : Rotate90LeftFaceCounterClockwise()

                    Case 5 * 6 + 0, 5 * 6 + 5 : WeTerminateWithError(22) : Stop
                    Case 5 * 6 + 1 : Rotate90LeftFaceCounterClockwise()
                    Case 5 * 6 + 2 : Rotate90UpFaceCounterClockwise()
                    Case 5 * 6 + 3 : Rotate90RightFaceCounterClockwise()
                    Case 5 * 6 + 4 : Rotate90DownFaceCounterClockwise()
                    Case Else
                        Select Case 6 * DownFace + FrontFace
                            Case 0 To 35 : WeTerminateWithError(22) : Stop
                            Case Is > 35 : WeTerminateWithError(5) : Stop
                        End Select
                End Select
            Case Else
                WeTerminateWithError(5) : Stop
        End Select
    End Sub

    Public Sub Rotate90LeftFaceClockwise()    ' Izquierda girar abajo
        Rotate90LeftFaceCounterClockwise() : Rotate90LeftFaceCounterClockwise() : Rotate90LeftFaceCounterClockwise()
    End Sub

    Public Sub Rotate90LeftFaceClockwise(FrontFace%, DownFace%)
        Rotate90LeftFaceCounterClockwise(FrontFace, DownFace) : Rotate90LeftFaceCounterClockwise(FrontFace, DownFace) : Rotate90LeftFaceCounterClockwise(FrontFace, DownFace)
    End Sub

    Public Sub Rotate90RightFaceClockwise() ' Derecha girar arriba
        AddElementToArray(6, ListOfMovements)
        Dim ArrayOfAddends(5) As Integer
        ArrayOfAddends(0) += CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(1)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(1)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(1))
        ArrayOfAddends(0) -= CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(0)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(0)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(0))
        ArrayOfAddends(1) += CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(5)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(5)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(5))
        ArrayOfAddends(1) -= CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(1)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(1)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(1))
        ArrayOfAddends(2) = 0
        ArrayOfAddends(3) += CType(6 ^ 0, Integer) * SquareColour(8, ArrayOfRubiksCube(0)) + CType(6 ^ 3, Integer) * SquareColour(5, ArrayOfRubiksCube(0)) +
            CType(6 ^ 6, Integer) * SquareColour(2, ArrayOfRubiksCube(0))
        ArrayOfAddends(3) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(3)) + CType(6 ^ 3, Integer) * SquareColour(3, ArrayOfRubiksCube(3)) +
            CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(3))
        ArrayOfAddends(4) += ColourConfigurationOfTheFaceOnceRotated90Clockwise(ArrayOfRubiksCube(4))
        ArrayOfAddends(4) -= ArrayOfRubiksCube(4)
        ArrayOfAddends(5) += CType(6 ^ 2, Integer) * SquareColour(6, ArrayOfRubiksCube(3)) + CType(6 ^ 5, Integer) * SquareColour(3, ArrayOfRubiksCube(3)) +
            CType(6 ^ 8, Integer) * SquareColour(0, ArrayOfRubiksCube(3))
        ArrayOfAddends(5) -= CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(5)) + CType(6 ^ 5, Integer) * SquareColour(5, ArrayOfRubiksCube(5)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(5))
        ApplyAddends(ArrayOfAddends)
    End Sub

    Public Sub Rotate90RightFaceClockwise(FrontFace%, DownFace%)
        Select Case DownFace
            Case 2
                Select Case FrontFace
                    Case 0 : Rotate90DownFaceClockwise()
                    Case 1 : Rotate90BackFaceClockwise()
                    Case 2, 4 : WeTerminateWithError(21) : Stop
                    Case 3 : Rotate90FrontFaceClockwise()
                    Case 5 : Rotate90UpFaceClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 4
                Select Case FrontFace
                    Case 0 : Rotate90UpFaceClockwise()
                    Case 1 : Rotate90FrontFaceClockwise()
                    Case 2, 4 : WeTerminateWithError(21) : Stop
                    Case 3 : Rotate90BackFaceClockwise()
                    Case 5 : Rotate90DownFaceClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 0, 1, 3, 5
                If DownFace = FrontFace Or DownFace = BackFace(FrontFace) Then WeTerminateWithError(21) : Stop
                Select Case 6 * DownFace + FrontFace
                    Case 0, 5 : WeTerminateWithError(22) : Stop
                    Case 1 : Rotate90LeftFaceClockwise()
                    Case 2 : Rotate90UpFaceClockwise()
                    Case 3 : Rotate90RightFaceClockwise()
                    Case 4 : Rotate90DownFaceClockwise()

                    Case 1 * 6 + 0 : Rotate90RightFaceClockwise()
                    Case 1 * 6 + 1, 1 * 6 + 3 : WeTerminateWithError(22) : Stop
                    Case 1 * 6 + 2 : Rotate90FrontFaceClockwise()
                    Case 1 * 6 + 4 : Rotate90BackFaceClockwise()
                    Case 1 * 6 + 5 : Rotate90LeftFaceClockwise()

                    Case 3 * 6 + 0 : Rotate90LeftFaceClockwise()
                    Case 3 * 6 + 1, 3 * 6 + 3 : WeTerminateWithError(22) : Stop
                    Case 3 * 6 + 2 : Rotate90BackFaceClockwise()
                    Case 3 * 6 + 4 : Rotate90FrontFaceClockwise()
                    Case 3 * 6 + 5 : Rotate90RightFaceClockwise()

                    Case 5 * 6 + 0, 5 * 6 + 5 : WeTerminateWithError(22) : Stop
                    Case 5 * 6 + 1 : Rotate90RightFaceClockwise()
                    Case 5 * 6 + 2 : Rotate90DownFaceClockwise()
                    Case 5 * 6 + 3 : Rotate90LeftFaceClockwise()
                    Case 5 * 6 + 4 : Rotate90UpFaceClockwise()

                    Case Else
                        Select Case 6 * DownFace + FrontFace
                            Case 0 To 35 : WeTerminateWithError(22) : Stop
                            Case Else : WeTerminateWithError(5) : Stop
                        End Select
                End Select
        End Select
    End Sub

    Public Sub Rotate90RightFaceCounterClockwise()  ' Derecha girar abajo
        Rotate90RightFaceClockwise() : Rotate90RightFaceClockwise() : Rotate90RightFaceClockwise()
    End Sub

    Public Sub Rotate90RightFaceCounterClockwise(FrontFace%, DownFace%)
        Rotate90RightFaceClockwise(FrontFace, DownFace) : Rotate90RightFaceClockwise(FrontFace, DownFace) : Rotate90RightFaceClockwise(FrontFace, DownFace)
    End Sub

    Public Sub Rotate90FrontFaceCounterClockwise()   ' Alante rotar a la izquierda
        AddElementToArray(8, ListOfMovements)
        Dim ArrayOfAddends(5) As Integer
        ArrayOfAddends(0) += ColourConfigurationOfTheFaceOnceRotated90CounterClockwise(ArrayOfRubiksCube(0))
        ArrayOfAddends(0) -= ArrayOfRubiksCube(0)
        ArrayOfAddends(1) += CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(2)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(2)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(2))
        ArrayOfAddends(1) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(1)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(1)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(1))
        ArrayOfAddends(2) += CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(3)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(3)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(3))
        ArrayOfAddends(2) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(2)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(2)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(2))
        ArrayOfAddends(3) += CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(4)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(4)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(4))
        ArrayOfAddends(3) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(3)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(3)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(3))
        ArrayOfAddends(4) += CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(1)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(1)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(1))
        ArrayOfAddends(4) -= CType(6 ^ 0, Integer) * SquareColour(0, ArrayOfRubiksCube(4)) + CType(6 ^ 1, Integer) * SquareColour(1, ArrayOfRubiksCube(4)) +
            CType(6 ^ 2, Integer) * SquareColour(2, ArrayOfRubiksCube(4))
        ArrayOfAddends(5) = 0
        ApplyAddends(ArrayOfAddends)
    End Sub

    Public Sub Rotate90FrontFaceCounterClockwise(FrontFace%, DownFace%)
        Select Case DownFace
            Case 0, 5
                Select Case FrontFace
                    Case 0, 5 : WeTerminateWithError(21) : Stop
                    Case 1 : Rotate90DownFaceCounterClockwise()
                    Case 2 : Rotate90LeftFaceCounterClockwise()
                    Case 3 : Rotate90UpFaceCounterClockwise()
                    Case 4 : Rotate90RightFaceCounterClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 1, 3
                Select Case FrontFace
                    Case 0 : Rotate90FrontFaceCounterClockwise()
                    Case 1, 3 : WeTerminateWithError(21) : Stop
                    Case 2 : Rotate90LeftFaceCounterClockwise()
                    Case 4 : Rotate90RightFaceCounterClockwise()
                    Case 5 : Rotate90BackFaceCounterClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 2, 4
                Select Case FrontFace
                    Case 0 : Rotate90FrontFaceCounterClockwise()
                    Case 1 : Rotate90DownFaceCounterClockwise()
                    Case 2, 4 : WeTerminateWithError(21) : Stop
                    Case 3 : Rotate90UpFaceCounterClockwise()
                    Case 5 : Rotate90BackFaceCounterClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case Else
                WeTerminateWithError(5) : Stop
        End Select
    End Sub

    Public Sub Rotate90FrontFaceClockwise() ' Alante rotar a la derecha
        Rotate90FrontFaceCounterClockwise() : Rotate90FrontFaceCounterClockwise() : Rotate90FrontFaceCounterClockwise()
    End Sub

    Public Sub Rotate90FrontFaceClockwise(FrontFace%, DownFace%)
        Rotate90FrontFaceCounterClockwise(FrontFace, DownFace) : Rotate90FrontFaceCounterClockwise(FrontFace, DownFace) : Rotate90FrontFaceCounterClockwise(FrontFace, DownFace)
    End Sub

    Public Sub Rotate90BackFaceClockwise()    ' Atrás rotar a la izquierda
        AddElementToArray(10, ListOfMovements)
        Dim ArrayOfAddends(5) As Integer
        ArrayOfAddends(0) = 0
        ArrayOfAddends(1) += CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(2)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(2)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(2))
        ArrayOfAddends(1) -= CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(1)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(1)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(1))
        ArrayOfAddends(2) += CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(3)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(3)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(3))
        ArrayOfAddends(2) -= CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(2)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(2)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(2))
        ArrayOfAddends(3) += CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(4)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(4)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(4))
        ArrayOfAddends(3) -= CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(3)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(3)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(3))
        ArrayOfAddends(4) += CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(1)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(1)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(1))
        ArrayOfAddends(4) -= CType(6 ^ 6, Integer) * SquareColour(6, ArrayOfRubiksCube(4)) + CType(6 ^ 7, Integer) * SquareColour(7, ArrayOfRubiksCube(4)) +
            CType(6 ^ 8, Integer) * SquareColour(8, ArrayOfRubiksCube(4))
        ArrayOfAddends(5) += ColourConfigurationOfTheFaceOnceRotated90Clockwise(ArrayOfRubiksCube(5))
        ArrayOfAddends(5) -= ArrayOfRubiksCube(5)
        ApplyAddends(ArrayOfAddends)
    End Sub

    Public Sub Rotate90BackFaceClockwise(FrontFace%, DownFace%)
        Select Case DownFace
            Case 0, 5
                Select Case FrontFace
                    Case 0, 5 : WeTerminateWithError(21) : Stop
                    Case 1 : Rotate90UpFaceClockwise()
                    Case 2 : Rotate90RightFaceClockwise()
                    Case 3 : Rotate90DownFaceClockwise()
                    Case 4 : Rotate90LeftFaceClockwise()
                    Case Else : WeTerminateWithError(5)
                End Select
            Case 1, 3
                Select Case FrontFace
                    Case 0 : Rotate90BackFaceClockwise()
                    Case 1, 3 : WeTerminateWithError(21) : Stop
                    Case 2 : Rotate90RightFaceClockwise()
                    Case 4 : Rotate90LeftFaceClockwise()
                    Case 5 : Rotate90FrontFaceClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
            Case 2, 4
                Select Case FrontFace
                    Case 0 : Rotate90BackFaceClockwise()
                    Case 1 : Rotate90UpFaceClockwise()
                    Case 2, 4 : WeTerminateWithError(21) : Stop
                    Case 3 : Rotate90DownFaceClockwise()
                    Case 5 : Rotate90FrontFaceClockwise()
                    Case Else : WeTerminateWithError(5) : Stop
                End Select
        End Select
    End Sub

    Public Sub Rotate90BackFaceCounterClockwise()  ' Atrás rotar a la derecha
        Rotate90BackFaceClockwise() : Rotate90BackFaceClockwise() : Rotate90BackFaceClockwise()
    End Sub

    Public Sub Rotate90BackFaceCounterClockwise(FrontFace%, DownFace%)
        Rotate90BackFaceClockwise(FrontFace, DownFace) : Rotate90BackFaceClockwise(FrontFace, DownFace) : Rotate90BackFaceClockwise(FrontFace, DownFace)
    End Sub

    Private Sub ApplyAddends(ArrayOfAddends() As Integer)
        Dim Counter As Integer
        For Counter = 0 To 5
            ArrayOfRubiksCube(Counter) += ArrayOfAddends(Counter)
        Next
    End Sub


    Public Sub ExecuteMovement(MovementNumber As Integer)
        Select Case MovementNumber
            Case 0 : Rotate90UpFaceClockwise()
            Case 1 : Rotate90UpFaceCounterClockwise()
            Case 2 : Rotate90DownFaceCounterClockwise()
            Case 3 : Rotate90DownFaceClockwise()
            Case 4 : Rotate90LeftFaceCounterClockwise()
            Case 5 : Rotate90LeftFaceClockwise()
            Case 6 : Rotate90RightFaceClockwise()
            Case 7 : Rotate90RightFaceCounterClockwise()
            Case 8 : Rotate90FrontFaceCounterClockwise()
            Case 9 : Rotate90FrontFaceClockwise()
            Case 10 : Rotate90BackFaceClockwise()
            Case 11 : Rotate90BackFaceCounterClockwise()
            Case Else : WeTerminateWithError(19) : Stop
        End Select
    End Sub

    'Public Sub ExecuteMovement(MovementNumber%, FrontFace%, DownFace%)
    '    Select Case MovementNumber
    '        Case 0 : Rotate90UpFaceClockwise(FrontFace, DownFace)
    '        Case 1 : Rotate90UpFaceCounterClockwise(FrontFace, DownFace)
    '        Case 2 : Rotate90DownFaceCounterClockwise(FrontFace, DownFace)
    '        Case 3 : Rotate90DownFaceClockwise(FrontFace, DownFace)
    '        Case 4 : Rotate90LeftFaceCounterClockwise(FrontFace, DownFace)
    '        Case 5 : Rotate90LeftFaceClockwise(FrontFace, DownFace)
    '        Case 6 : Rotate90RightFaceClockwise(FrontFace, DownFace)
    '        Case 7 : Rotate90RightFaceCounterClockwise(FrontFace, DownFace)
    '        Case 8 : Rotate90FrontFaceCounterClockwise(FrontFace, DownFace)
    '        Case 9 : Rotate90FrontFaceClockwise(FrontFace, DownFace)
    '        Case 10 : Rotate90BackFaceClockwise(FrontFace, DownFace)
    '        Case 11 : Rotate90BackFaceCounterClockwise(FrontFace, DownFace)
    '        Case Else : WeTerminateWithError(19) : Stop
    '    End Select
    'End Sub



End Class

