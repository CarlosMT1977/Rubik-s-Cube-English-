Option Strict On
Option Explicit On
Imports Cubo_de_Rubik.Utilidades

Public Class SolveTheCubeDeLuxe
    ' The first thing we will do is solving the yellow face
    Private Const YellowColour As Integer = 0
    Private CubeThatWeMustSolve As ClaseCuboDeRubik
    'Private ListOfMovements() As Integer

    Public Sub New(CubeArgument As ClaseCuboDeRubik)
        CubeThatWeMustSolve = New ClaseCuboDeRubik(CloningOfArray(CubeArgument.ArrayOfRubiksCube))
    End Sub

    Public Sub ShowSolutionOfTheCube()
        Dim TotalString As String = FromNumberOfMovementToString(CubeThatWeMustSolve, True)

        CubeThatWeMustSolve.ArrayOfRubiksCube = CloningOfArray(CubeThatWeMustSolve.InitialArray)
        Dim TextOfMessageBox, CaptionOfMessageBox As String
        TextOfMessageBox = "Next, we are going to see the " & CubeThatWeMustSolve.ListOfMovements.GetLength(0) & " needed movements"
        CaptionOfMessageBox = CubeThatWeMustSolve.ListOfMovements.GetLength(0) & " movements"
        MessageBox.Show(TextOfMessageBox, CaptionOfMessageBox)

        Dim LittleString As String
        Do While TotalString <> vbNullString
            If TotalString.IndexOf(vbCrLf & vbCrLf) <> -1 Then
                LittleString = TotalString.Substring(0, TotalString.IndexOf(vbCrLf & vbCrLf))
                TotalString = TotalString.Substring(TotalString.IndexOf(vbCrLf & vbCrLf) + 4)
            ElseIf TotalString <> vbNullString Then
                LittleString = TotalString
                TotalString = vbNullString
            End If
            MessageBox.Show(LittleString, "Rotate 90°")
        Loop
    End Sub

    Private Function IsTheYellowFaceSolvedButTheFirstRowCantBeSolved() As Boolean
        If Not IsTheFaceSolved(YellowColour) Then Return False
        Dim CurrentFace, FollowingFace As Integer
        Dim CurrentColour, FollowingColour As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            CurrentColour = SquareColour(0, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace))
            FollowingColour = SquareColour(2, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace))
            If Not (FollowingColour = CurrentColour Mod 4 + 1) Then Return True
        Next
        Return False
    End Function

    Private Function IsTheWhiteFaceSolvedButTheLastRowCantBeSolved() As Boolean
        ' We presuppose that the yellow face is already solved, as well as the two upper faces in each adjacent face
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then Return False
        If Not IsTheFaceSolved(5) Then Return False
        Dim CurrentFace, FollowingFace As Integer
        Dim CurrentColour, FollowingColour As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            CurrentColour = SquareColour(6, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace))
            FollowingColour = SquareColour(8, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace))
            If Not (FollowingColour = CurrentColour Mod 4 + 1) Then Return True
        Next
        Return False
    End Function

    Private Sub MakePreviousCheckings()
        Dim Counter, CurrentColour, FaceCounter, SquareCounter As Integer
        Dim AppearancesOfEachColour(5) As Integer
        For FaceCounter = 0 To 5
            If Not AreAllTheSquaresOfTheSoughtColour(FaceCounter, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceCounter), 4) Then WeTerminateWithError(67) : Stop : End
        Next
        InitializeArray(AppearancesOfEachColour)
        For FaceCounter = 0 To 5
            For SquareCounter = 0 To 8
                CurrentColour = SquareColour(SquareCounter, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceCounter))
                AppearancesOfEachColour(CurrentColour) += 1
                If AppearancesOfEachColour(CurrentColour) > 9 Then WeTerminateWithError(67) : Stop : End
            Next
        Next

        Dim AparicionesDeCadaPar(5, 5) As Integer

        Dim ColorSuperior, ColorLateralInferior, ColorLateralSuperior, ColorInferior, ColorLateralIzquierdo, ColorLateralDerecho As Integer
        Dim YellowSquare, WhiteSquare As Integer
        Dim CaraLateral As Integer
        For CaraLateral = 1 To 4
            Select Case CaraLateral
                Case 1 : YellowSquare = 7 : WhiteSquare = 1
                Case 2 : YellowSquare = 3 : WhiteSquare = 3
                Case 3 : YellowSquare = 1 : WhiteSquare = 7
                Case 4 : YellowSquare = 5 : WhiteSquare = 5
                Case Else : WeTerminateWithError(22) : Stop
            End Select
            ColorSuperior = SquareColour(YellowSquare, CubeThatWeMustSolve.ArrayOfRubiksCube(0))
            ColorInferior = SquareColour(WhiteSquare, CubeThatWeMustSolve.ArrayOfRubiksCube(5))
            ColorLateralSuperior = SquareColour(1, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraLateral))
            ColorLateralInferior = SquareColour(7, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraLateral))
            ColorLateralDerecho = SquareColour(3, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraLateral))
            ColorLateralIzquierdo = SquareColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraLateral Mod 4 + 1))

            AparicionesDeCadaPar(ColorSuperior, ColorLateralSuperior) += 1
            AparicionesDeCadaPar(ColorLateralSuperior, ColorSuperior) += 1
            AparicionesDeCadaPar(ColorInferior, ColorLateralInferior) += 1
            AparicionesDeCadaPar(ColorLateralInferior, ColorInferior) += 1
            AparicionesDeCadaPar(ColorLateralIzquierdo, ColorLateralDerecho) += 1
            AparicionesDeCadaPar(ColorLateralDerecho, ColorLateralIzquierdo) += 1
        Next
        Dim ColorUno, ColorDos As Integer
        For ColorUno = 0 To 5
            For ColorDos = 0 To 5
                If ColorUno = ColorDos Or ColorUno = BackFace(ColorDos) Then
                    If AparicionesDeCadaPar(ColorUno, ColorDos) <> 0 Then WeTerminateWithError(67) : Stop : End
                Else
                    If AparicionesDeCadaPar(ColorUno, ColorDos) <> 1 Then WeTerminateWithError(67) : Stop : End
                End If
            Next
        Next

        Dim AparicionesDeCadaTrio(5, 5, 5) As Integer
        Dim TrioEsquina(2) As Integer


        Dim CaraPosterior, ColorTres As Integer


        For CaraLateral = 1 To 4
            Select Case CaraLateral
                Case 1 : YellowSquare = 6 : WhiteSquare = 0
                Case 2 : YellowSquare = 0 : WhiteSquare = 6
                Case 3 : YellowSquare = 2 : WhiteSquare = 8
                Case 4 : YellowSquare = 8 : WhiteSquare = 2
                Case Else : WeTerminateWithError(22) : Stop
            End Select
            CaraPosterior = CaraLateral Mod 4 + 1

            ColorUno = SquareColour(YellowSquare, CubeThatWeMustSolve.ArrayOfRubiksCube(0))
            ColorDos = SquareColour(0, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraLateral))
            ColorTres = SquareColour(2, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraPosterior))
            AparicionesDeCadaTrio(ColorUno, ColorDos, ColorTres) += 1
            AparicionesDeCadaTrio(ColorUno, ColorTres, ColorDos) += 1
            AparicionesDeCadaTrio(ColorDos, ColorUno, ColorTres) += 1
            AparicionesDeCadaTrio(ColorDos, ColorTres, ColorUno) += 1
            AparicionesDeCadaTrio(ColorTres, ColorUno, ColorDos) += 1
            AparicionesDeCadaTrio(ColorTres, ColorDos, ColorUno) += 1

            ColorUno = SquareColour(WhiteSquare, CubeThatWeMustSolve.ArrayOfRubiksCube(5))
            ColorDos = SquareColour(6, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraLateral))
            ColorTres = SquareColour(8, CubeThatWeMustSolve.ArrayOfRubiksCube(CaraPosterior))
            AparicionesDeCadaTrio(ColorUno, ColorDos, ColorTres) += 1
            AparicionesDeCadaTrio(ColorUno, ColorTres, ColorDos) += 1
            AparicionesDeCadaTrio(ColorDos, ColorUno, ColorTres) += 1
            AparicionesDeCadaTrio(ColorDos, ColorTres, ColorUno) += 1
            AparicionesDeCadaTrio(ColorTres, ColorUno, ColorDos) += 1
            AparicionesDeCadaTrio(ColorTres, ColorDos, ColorUno) += 1
        Next

        Dim AparicionesAmarillas, AparicionesBlancas As Integer
        Dim ContadorDePruebas As Integer = 0
        For ColorUno = 0 To 5
            For ColorDos = 0 To 5
                For ColorTres = 0 To 5
                    AparicionesAmarillas = 0 : AparicionesBlancas = 0
                    If ColorUno = 0 Then AparicionesAmarillas += 1
                    If ColorDos = 0 Then AparicionesAmarillas += 1
                    If ColorTres = 0 Then AparicionesAmarillas += 1

                    If ColorUno = 5 Then AparicionesBlancas += 1
                    If ColorDos = 5 Then AparicionesBlancas += 1
                    If ColorTres = 5 Then AparicionesBlancas += 1

                    If Not (AparicionesAmarillas = 1 And AparicionesBlancas = 0) And Not (AparicionesAmarillas = 0 And AparicionesBlancas = 1) Then Continue For

                    Dim ColoresLateralesDeLaEsquina(1) As Integer
                    ColoresLateralesDeLaEsquina = {-1, -1}

                    Dim PivoteActual As Integer = 0
                    If ColorUno <> 0 And ColorUno <> 5 Then
                        ColoresLateralesDeLaEsquina(PivoteActual) = ColorUno
                        PivoteActual += 1
                    End If
                    If ColorDos <> 0 And ColorDos <> 5 Then
                        ColoresLateralesDeLaEsquina(PivoteActual) = ColorDos
                        PivoteActual += 1
                    End If
                    If ColorTres <> 0 And ColorTres <> 5 Then
                        ColoresLateralesDeLaEsquina(PivoteActual) = ColorTres
                        PivoteActual += 1
                    End If
                    Select Case Math.Abs(ColoresLateralesDeLaEsquina(0) - ColoresLateralesDeLaEsquina(1))
                        Case 0, 2 : Continue For
                        Case 1, 3
                        Case Else : WeTerminateWithError(67) : Stop : End
                    End Select

                    ContadorDePruebas += 1
                    If AparicionesDeCadaTrio(ColorUno, ColorDos, ColorTres) <> 1 Then WeTerminateWithError(67) : Stop : End
                Next
            Next
        Next
        If ContadorDePruebas <> 48 Then WeTerminateWithError(67) : Stop : End
    End Sub

    Public Sub SolveTheCube()
        MakePreviousCheckings()
        If Not AreTheFourDownCornerPiecesSolved() Then SolveTheFourDownCornerPieces()
        Dim NumberOfRepetitionsOfTheLoop As Integer = 0
        Do While Not IsTheCubeSolved()
            NumberOfRepetitionsOfTheLoop += 1
            If NumberOfRepetitionsOfTheLoop > 10 Then WeTerminateWithError(67) : Stop : End
            If Not AreTheFourDownCornerPiecesSolved() Then WeTerminateWithError(22) : Stop
            Dim NumberOfUnsolvedDownEdgePieces As Integer = HowManyUnsolvedEdgePiecesAreThereDown()
            Dim DownFace As Integer = 0
            Select Case NumberOfUnsolvedDownEdgePieces
                Case 4 : DownFace = 1
                Case 3
                    Dim Counter As Integer
                    For Counter = 1 To 4
                        If AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7) Then
                            DownFace = (Counter + 1) Mod 4 + 1
                            Exit For
                        End If
                    Next
                    If DownFace = 0 Then WeTerminateWithError(67) : Stop : End
                Case Else
                    WeTerminateWithError(22) : Stop
            End Select
            If DownFace < 1 Or DownFace > 4 Then WeTerminateWithError(67) : Stop : End
            Dim LeftFace, RightFace, TopFace As Integer
            Dim WeTurnRight As Boolean
            RightFace = DownFace Mod 4 + 1
            LeftFace = (DownFace + 2) Mod 4 + 1
            TopFace = (DownFace + 1) Mod 4 + 1

            If Not AreAllTheSquaresOfTheSoughtColour(TopFace, CubeThatWeMustSolve.ArrayOfRubiksCube(TopFace), 7) Then
                WeTurnRight = True
            ElseIf AreAllTheSquaresOfTheSoughtColour(RightFace, CubeThatWeMustSolve.ArrayOfRubiksCube(LeftFace), 7) Then
                WeTurnRight = True
            ElseIf AreAllTheSquaresOfTheSoughtColour(DownFace, CubeThatWeMustSolve.ArrayOfRubiksCube(LeftFace), 7) Then
                WeTurnRight = False
            Else
                WeTerminateWithError(67) : Stop : End
            End If

            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            If WeTurnRight Then
                CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
            Else
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(5, DownFace)
            End If
            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
            If WeTurnRight Then
                CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
            Else
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(5, DownFace)
            End If
            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            SolveTheFourDownCornerPieces()
        Loop
        SimplifyArrayOfMovements(CubeThatWeMustSolve.ListOfMovements)
        Clipboard.SetText(FromNumberOfMovementToString(CubeThatWeMustSolve, True))
        MessageBox.Show("It seems to be that the cube is fully solved")
    End Sub

    Private Function IsTheCubeSolved() As Boolean
        Dim FaceCounter, SquareCounter As Integer
        For FaceCounter = 0 To 5
            For SquareCounter = 0 To 8
                If SquareColour(SquareCounter, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceCounter)) <> FaceCounter Then Return False
            Next
        Next
        Return True
    End Function

    Private Function HowManyUnsolvedEdgePiecesAreThereDown() As Integer
        ' Damos por hecho, al entrar, aquí, que todo lo que no sean los cuatro bordes de abajo está todo montado
        If Not AreTheFourDownCornerPiecesSolved() Then WeTerminateWithError(22) : Stop
        Dim Counter As Integer
        Dim Gatherer As Integer = 0
        For Counter = 1 To 4
            If Not AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7) Then Gatherer += 1
        Next
        If Gatherer <> 3 And Gatherer <> 4 Then WeTerminateWithError(67) : Stop : End
        Return Gatherer
    End Function

    Private Function AreTheFourDownCornerPiecesSolved() As Boolean
        ' Aquí en la resolución damos por hecho que está montada la cara amarilla completa, las dos líneas superiores de cada una de las cuatro caras adyacentes, y la cara blanca completa
        ' Here we presuppose that we have already solved the whole yellow face, the two upper rows of each of the four adjacent faces, and the whole white face
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then Return False
        If Not IsTheFaceSolved(5) Then Return False
        Dim Counter As Integer
        For Counter = 1 To 4
            If Not AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 6, 8) Then Return False
        Next
        If Not AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 0, 2, 6, 8) Then WeTerminateWithError(22) : Stop
        Return True
    End Function

    Private Function AreThereTwoAndOnlyTwoDownAdjacentCornersSolved() As Boolean
        ' Aquí en la resolución damos por hecho que está montada la cara amarilla completa, las dos líneas superiores de cada una de las cuatro caras adyacentes, y la cara blanca completa
        ' Here we presuppose that we have already solved the whole yellow face, the two upper rows of each of the four adjacent faces, and the whole white face

        If AreTheFourDownCornerPiecesSolved() Then Return False
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then Return False
        If Not IsTheFaceSolved(5) Then Return False

        Dim CurrentFace, PreviousFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            If Not AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 6, 8) Then Continue For
            If Not AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 8) Then WeTerminateWithError(22) : Stop
            If Not AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(PreviousFace), 6) Then WeTerminateWithError(22) : Stop
            Return True
        Next

        Return False
    End Function

    Private Function AreThereTwoOppositeCornerPiecesDownSolvedAndOnlyTwo() As Boolean
        ' Aquí en la resolución damos por hecho que está montada la cara amarilla completa, las dos líneas superiores de cada una de las cuatro caras adyacentes, y la cara blanca completa
        ' Here we presuppose that we have already solved the whole yellow face, the two upper rows of each of the four adjacent faces, and the whole white face

        If AreTheFourDownCornerPiecesSolved() Then Return False
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then Return False
        If Not IsTheFaceSolved(5) Then Return False

        Dim CurrentFace, PreviousFace, FollowingFace, OppositeFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            OppositeFace = (CurrentFace + 1) Mod 4 + 1
            If Not AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 6) Then Continue For
            If Not AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 8) Then WeTerminateWithError(22) : Stop
            If Not AreAllTheSquaresOfTheSoughtColour(OppositeFace, CubeThatWeMustSolve.ArrayOfRubiksCube(OppositeFace), 6) Then Continue For
            If Not AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(PreviousFace), 8) Then WeTerminateWithError(22) : Stop
            Return True
        Next
        Return False
    End Function

    Private Function AreThereTwoAndOnlyTwoDownCornerPiecesSolved() As Boolean
        ' Aquí en la resolución damos por hecho que está montada la cara amarilla completa, las dos líneas superiores de cada una de las cuatro caras adyacentes, y la cara blanca completa
        ' Here we presuppose that we have already solved the whole yellow face, the two upper rows of each of the four adjacent faces, and the whole white face

        Dim One, Two As Boolean
        One = AreThereTwoAndOnlyTwoDownAdjacentCornersSolved()
        Two = AreThereTwoOppositeCornerPiecesDownSolvedAndOnlyTwo()
        If One And Two Then WeTerminateWithError(65) : Stop
        Return One Xor Two
    End Function

    Public Sub SolveTheFourDownCornerPieces()
        ' Aquí damos por hecho que tenemos que montar también la cara amarilla, las dos líneas superiores de cada una de las cuatro caras adyacentes, y la cara blanca
        ' Here we presuppose that we also have to solve the yellow face, the two upper rows in each of the four adjacent faces, and also the whole white face

        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then SolveTheYellowFacedWithTheFirstAndSecondRow()
        If Not IsTheFaceSolved(5) Then SolveTheDownWhiteFace()
        If IsTheWhiteFaceSolvedButTheLastRowCantBeSolved() Then WeTerminateWithError(67) : Stop : End
        Do While Not AreTheFourDownCornerPiecesSolved()
            Do While Not AreThereTwoAndOnlyTwoDownCornerPiecesSolved() AndAlso Not AreTheFourDownCornerPiecesSolved()
                CubeThatWeMustSolve.Rotate90BackFaceClockwise()
            Loop
            If AreTheFourDownCornerPiecesSolved() Then Exit Do
            Dim DownFace As Integer = 0
            If AreThereTwoAndOnlyTwoDownAdjacentCornersSolved() Then
                Dim CurrentFace, PreviousFace, FollowingFace As Integer
                For CurrentFace = 1 To 4
                    FollowingFace = CurrentFace Mod 4 + 1
                    PreviousFace = (CurrentFace + 2) Mod 4 + 1
                    If Not AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 6, 8) Then Continue For
                    If Not AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 8) Then WeTerminateWithError(22) : Stop
                    If Not AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(PreviousFace), 6) Then WeTerminateWithError(22) : Stop
                    DownFace = (CurrentFace + 1) Mod 4 + 1
                    Exit For
                Next
                If CurrentFace > 4 Then WeTerminateWithError(22) : Stop
            ElseIf AreThereTwoOppositeCornerPiecesDownSolvedAndOnlyTwo() Then
                DownFace = 1
            Else
                WeTerminateWithError(22) : Stop
            End If
            If DownFace = 0 Then WeTerminateWithError(22) : Stop
            If DownFace < 1 Or DownFace > 4 Then WeTerminateWithError(66) : Stop

            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90UpFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90UpFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90UpFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90UpFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
        Loop
    End Sub


    Private Function IsThereWhiteCross() As Boolean
        Return AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1, 3, 5, 7)
    End Function

    Private Function IsThereAWhiteLineWithoutPeak() As Boolean
        Dim BoolAuxOne, BoolAuxTwo As Boolean
        BoolAuxOne = AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1, 7)
        BoolAuxTwo = AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 3, 5)
        Return BoolAuxOne Xor BoolAuxTwo
    End Function

    Private Function IsThereAWhitePeakWithoutLine() As Boolean
        Dim BoolAuxOne, BoolAuxTwo As Boolean
        BoolAuxOne = AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1, 3)
        BoolAuxTwo = AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 5, 7)
        If BoolAuxOne Xor BoolAuxTwo Then Return True
        BoolAuxOne = AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1, 5)
        BoolAuxTwo = AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 3, 7)
        If BoolAuxOne Xor BoolAuxTwo Then Return True
        Return False
    End Function


    Private Sub SolveTheCrossFromTheHorizontalLine(ByVal DownFace As Integer)
        CubeThatWeMustSolve.Rotate90DownFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(5, DownFace)
    End Sub

    Private Sub MakeTheHorizontalLineFromTheNorthEasternPeak(ByVal DownFace As Integer)
        ' Aquí hay que comprobar primero si existe pico nororiental
        ' Here we must first check if the North Eastern Peak exists

        CubeThatWeMustSolve.Rotate90UpFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90UpFaceCounterClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90UpFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
        CubeThatWeMustSolve.Rotate90UpFaceCounterClockwise(5, DownFace)
    End Sub

    Private Sub MakeTheCrossFromNothing(ByVal DownFace As Integer)
        Dim FollowingFace As Integer = DownFace Mod 4 + 1
        SolveTheCrossFromTheHorizontalLine(FollowingFace)
        MakeTheHorizontalLineFromTheNorthEasternPeak(DownFace)
        SolveTheCrossFromTheHorizontalLine(DownFace)
    End Sub


    Private Function IsTheVoidDown() As Boolean
        ' Se refiere a que no está ninguno de los bordes blancos que formarían la cruz, pero las esquinas blancas pueden estar
        ' It means that there is noone of the white edges that would solve the cross, but the white corners can be

        Return NoOneOfTheSquaresIsTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1, 3, 5, 7)
    End Function

    Private Sub SolveTheWhiteCross()
        If IsThereWhiteCross() Then
            MessageBox.Show("The white cross was already solved")
            ' Se supone que aquí no podemos llegar, porque si estamos en este procedimiento es porque no había cruz blanca.
            ' It is supposed that we can't be here, because if we are in this procedure it's bewcause there wasn't any white cross.
            Exit Sub
        ElseIf IsTheVoidDown() Then
            MakeTheCrossFromNothing(1)
        ElseIf IsThereAWhiteLineWithoutPeak() Then
            If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 3, 5) Then
                SolveTheCrossFromTheHorizontalLine(1)
            ElseIf AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1, 7) Then
                SolveTheCrossFromTheHorizontalLine(2)
            Else
                WeTerminateWithError(22) : Stop
            End If
        ElseIf IsThereAWhitePeakWithoutLine() Then
            Dim DownFace As Integer
            If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 3) Then
                    DownFace = 4
                ElseIf AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 5) Then
                    DownFace = 3
                Else
                    WeTerminateWithError(22) : Stop
                End If
            ElseIf AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 7) Then
                If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 3) Then
                    DownFace = 1
                ElseIf AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), 5) Then
                    DownFace = 2
                Else
                    WeTerminateWithError(22) : Stop
                End If
            Else
                WeTerminateWithError(22) : Stop
            End If
            MakeTheHorizontalLineFromTheNorthEasternPeak(DownFace)
            SolveTheCrossFromTheHorizontalLine(DownFace)
        Else
            WeTerminateWithError(67) : Stop : End
        End If

        If Not IsThereWhiteCross() Then WeTerminateWithError(67) : Stop : End


    End Sub

    Public Sub SolveTheDownWhiteFace()
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then SolveTheYellowFacedWithTheFirstAndSecondRow()
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then WeTerminateWithError(64) : Stop
        If Not IsThereWhiteCross() Then SolveTheWhiteCross()
        If Not IsThereWhiteCross() Then WeTerminateWithError(67) : Stop : End

        Dim NumberOfRepetitionsOfTheLoop As Integer = 0
        Do While Not IsTheFaceSolved(5)
            NumberOfRepetitionsOfTheLoop += 1
            If NumberOfRepetitionsOfTheLoop > 10 Then WeTerminateWithError(67) : Stop : End
            If Not IsThereWhiteCross() Or Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then WeTerminateWithError(22) : Stop : End
            Dim DownFace, FacePreviousToTheDownFace, Counter As Integer
            Dim NumberOfWhiteCorners As Integer = 0
            For Counter = 0 To 8 Step 2
                If Counter = 4 Then Continue For
                If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), Counter) Then NumberOfWhiteCorners += 1
            Next
            If Not IsThereWhiteCross() Or Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then WeTerminateWithError(67) : Stop : End
            DownFace = 0
            Select Case NumberOfWhiteCorners
                Case 0
                    For Counter = 1 To 4
                        FacePreviousToTheDownFace = (Counter + 6) Mod 4 + 1
                        If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(FacePreviousToTheDownFace), 6) Then
                            DownFace = Counter
                            Exit For
                        End If
                    Next
                Case 1
                    For Counter = 0 To 8 Step 2
                        If Counter = 4 Then Continue For
                        If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(5), Counter) Then
                            Select Case Counter
                                Case 0 : DownFace = 2
                                Case 2 : DownFace = 1
                                Case 6 : DownFace = 3
                                Case 8 : DownFace = 4
                                Case Else : WeTerminateWithError(67) : Stop : End
                            End Select
                            Exit For
                        End If
                    Next
                Case 2, 3, 4
                    For Counter = 1 To 4
                        If AreAllTheSquaresOfTheSoughtColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 8) Then
                            DownFace = Counter
                            Exit For
                        End If
                    Next
                Case Else
                    WeTerminateWithError(22) : Stop : End
            End Select
            If DownFace = 0 Then WeTerminateWithError(67) : Stop : End
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(5, DownFace)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(5, DownFace)
        Loop
    End Sub

    Private Function IsItPossibleToPlaceDirectlyAnEdgePieceOfASecondRow() As Boolean
        Dim CurrentFace, PreviousFace, FollowingFace, OppositeFace, WhiteSquare As Integer
        For CurrentFace = 1 To 4
            Select Case CurrentFace
                Case 1 : WhiteSquare = 1
                Case 2 : WhiteSquare = 3
                Case 3 : WhiteSquare = 7
                Case 4 : WhiteSquare = 5
                Case Else : WeTerminateWithError(22) : Stop
            End Select
            FollowingFace = CurrentFace Mod 4 + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            OppositeFace = (CurrentFace + 1) Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(OppositeFace, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquare) Then
                If AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 7) Then Return True
                If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 7) Then Return True
            End If
        Next
        Return False
    End Function

    Private Sub PlaceDirectlyAnEdgeOfSecondRow()
        If Not IsItPossibleToPlaceDirectlyAnEdgePieceOfASecondRow() Then WeTerminateWithError(60) : Stop
        Dim CurrentFace, PreviousFace, FollowingFace, OppositeFace, WhiteSquare As Integer
        For CurrentFace = 1 To 4
            Select Case CurrentFace
                Case 1 : WhiteSquare = 1
                Case 2 : WhiteSquare = 3
                Case 3 : WhiteSquare = 7
                Case 4 : WhiteSquare = 5
                Case Else : WeTerminateWithError(22) : Stop
            End Select
            FollowingFace = CurrentFace Mod 4 + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            OppositeFace = (CurrentFace + 1) Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(OppositeFace, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquare) Then
                If AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 7) Then
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                    Exit Sub
                End If
                If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 7) Then
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(22) : Stop
    End Sub

    Private Function IsItPossibleToPlaceINDIRECTLYAnEdgeOfSecondRow() As Boolean
        Dim Result As Boolean = False
        Dim Counter As Integer
        For Counter = 1 To 4
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
            If IsItPossibleToPlaceDirectlyAnEdgePieceOfASecondRow() Then Result = True
        Next
        Return Result
    End Function

    Private Sub PlaceIndirectlyAnEdgeOfTheSecondRow()
        If Not IsItPossibleToPlaceINDIRECTLYAnEdgeOfSecondRow() Then WeTerminateWithError(61) : Stop
        Dim NumberOfMovements, Counter As Integer
        For NumberOfMovements = 4 To 1 Step -1
            For Counter = 1 To NumberOfMovements
                CubeThatWeMustSolve.Rotate90BackFaceClockwise()
            Next
            If IsItPossibleToPlaceDirectlyAnEdgePieceOfASecondRow() Then
                PlaceDirectlyAnEdgeOfSecondRow()
                Exit Sub
            End If
        Next
        WeTerminateWithError(22) : Stop
    End Sub

    Private Function IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() As Boolean
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstRow() Then Return False
        Dim Counter As Integer
        For Counter = 1 To 4
            If Not AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 3, 4, 5) Then Return False
        Next
        Return True
    End Function

    Private Function IsItPossibleToTakeDownAnInvertedEdgePieceFromTheSecondRow() As Boolean
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 3) Then
                If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 5) Then Return True
            End If
        Next
        Return False
    End Function

    Private Sub TakeDownAnInvertedEdgePieceFromTheSecondRow()
        If Not IsItPossibleToTakeDownAnInvertedEdgePieceFromTheSecondRow() Then WeTerminateWithError(62) : Stop
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 3) Then
                If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 5) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(22) : Stop
    End Sub

    Private Sub TakeDownAnEdgePieceFromTheSecondRow()
        If IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then WeTerminateWithError(63) : Stop
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = CurrentFace Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 3) Then
                If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 5) Then Continue For
            End If
            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
            Exit Sub
        Next
        WeTerminateWithError(22) : Stop
    End Sub

    Public Sub SolveTheYellowFacedWithTheFirstAndSecondRow()
        If IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then
            Dim TextOfMessageBox As String = "The yellow face, as well as the two upper rows, are already solved, so there's nothing to solve here"
            Dim CaptionOfMessageBox As String = "The thing is already solved"
            MessageBox.Show(TextOfMessageBox, CaptionOfMessageBox, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        Else
            SolveTheYellowFaceWithTheFirstRow()
        End If
        Do While Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow()
            If IsItPossibleToPlaceDirectlyAnEdgePieceOfASecondRow() Then
                PlaceDirectlyAnEdgeOfSecondRow()
            ElseIf IsItPossibleToPlaceINDIRECTLYAnEdgeOfSecondRow() Then
                PlaceIndirectlyAnEdgeOfTheSecondRow()
            ElseIf IsItPossibleToTakeDownAnInvertedEdgePieceFromTheSecondRow() Then
                TakeDownAnInvertedEdgePieceFromTheSecondRow()
            Else
                TakeDownAnEdgePieceFromTheSecondRow()
            End If
        Loop
        If Not IsTheYellowFaceSolvedAsWellAsTheFirstAndSecondRow() Then WeTerminateWithError(22) : Stop
    End Sub


    Public Sub SolveTheYellowFaceWithTheFirstRow()
        If IsTheYellowFaceSolvedAsWellAsTheFirstRow() Then
            Dim TextOfMessageBox As String = "The yellow face is already solved, as well as the first row, so there's nothing to solve here"
            Dim CaptionOfMessageBox As String = "The yellow face and the first row are already solved"
        End If

        If Not IsTheFaceSolved(YellowColour) Then
            SolveTheYellowFace()
        End If

        If IsTheYellowFaceSolvedButTheFirstRowCantBeSolved() Then WeTerminateWithError(67) : Stop : End

        WeRotateTheYellowFaceUntilWeGetTheOptimalPosition()

        Do While Not AreTheEdgePiecesOfTheFirstRowSolved()
            If IsItPossibleToExchangeFourConsecutiveEdgePieces() Then
                ExchangeFourConsecutiveEdgePieces()
            ElseIf IsItPossibleToExchangeFourDiabolicalEdgePieces() Then
                ExchangeFourDiabolicalEdgePieces()
            ElseIf IsItPossibleToExchangeThreeConsecutiveEdgePieces() Then
                ExchangeThreeConsecutiveEdgePieces()
            ElseIf IsItPossibleToExchangeTwoAdjacentEdgePieces() Then
                ExchangeTwoAdjacentEdgePieces()
            ElseIf IsItPossibleToExchangeTwoOppositeEdgePieces() Then
                ExchangeTwoOppositeEdgePieces()
            Else
                WeTerminateWithError(56) : Stop
            End If
        Loop

        Do While Not AreTheCornerPiecesOfTheFirstRowSolved()
            If IsItPossibleToExchangeFourConsecutiveCornerPieces() Then
                ExchangeFourConsecutiveCornerPieces()
            ElseIf IsItPossibleToExchangeFourDiabolicalCornerPieces() Then
                ExchangeFourDiabolicalCornerPieces()
            ElseIf IsItPossibleToExchangeThreeConsecutiveCornerPieces() Then
                ExchangeThreeConsecutiveCornerPieces()
            ElseIf IsItPossibleToExchangeTwoAdjacentCornerPieces() Then
                ExchangeTwoAdjacentCornerPieces()
            ElseIf IsItPossibleToExchangeTwoOppositeCornerPieces() Then
                ExchangeTwoOppositeCornerPieces()
            Else
                WeTerminateWithError(56) : Stop
            End If
        Loop

    End Sub


    Private Sub WeRotateTheYellowFaceUntilWeGetTheOptimalPosition()
        If Not IsTheFaceSolved(YellowColour) Then WeTerminateWithError(59) : Stop
        Dim Counter As Integer
        Dim BestNumberOfMovementsSoFar As Integer = 9999
        Dim BestCounterSoFar As Integer = 5
        Dim NeededMovements(4) As Integer
        For Counter = 1 To 4
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
            NeededMovements(Counter) = HowManyMovementsMustBeMadeWithTheFirstRow()
            Select Case Counter
                Case 4 : NeededMovements(Counter) += 0
                Case 2 : NeededMovements(Counter) += 2
                Case 1, 3 : NeededMovements(Counter) += 1
                Case Else : WeTerminateWithError(22) : Stop
            End Select
            If NeededMovements(Counter) < BestNumberOfMovementsSoFar Then
                BestNumberOfMovementsSoFar = NeededMovements(Counter)
                BestCounterSoFar = Counter
            End If
        Next
        For Counter = 1 To BestCounterSoFar
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
        Next

    End Sub

    Private Function HowManyMovementsMustBeMadeWithTheFirstRow() As Integer
        If Not IsTheFaceSolved(YellowColour) Then WeTerminateWithError(59) : Stop
        Dim Gatherer As Integer = 0

        If IsItPossibleToExchangeFourConsecutiveEdgePieces() Then
            Gatherer += 29
        ElseIf IsItPossibleToExchangeFourDiabolicalEdgePieces() Then
            Gatherer += 27
        ElseIf IsItPossibleToExchangeFourAdjacentEdgePiecesTwoByTwo() Then
            Gatherer += 34
        ElseIf IsItPossibleToExchangeFourEdgePiecesOppositeTwoByTwo() Then
            Gatherer += 30
        ElseIf IsItPossibleToExchangeThreeConsecutiveEdgePieces() Then
            Gatherer += 22
        ElseIf IsItPossibleToExchangeTwoAdjacentEdgePieces() Then
            Gatherer += 17
        ElseIf IsItPossibleToExchangeTwoOppositeEdgePieces() Then
            Gatherer += 15
        End If


        If IsItPossibleToExchangeFourConsecutiveCornerPieces() Then
            Gatherer += 18
        ElseIf IsItPossibleToExchangeFourDiabolicalCornerPieces() Then
            Gatherer += 18
        ElseIf IsItPossibleToExchangeFourCornerPiecesOppositeTwoByTwo() Then
            Gatherer += 22
        ElseIf IsItPossibleToExchangeFourCornerPiecesAdjacentTwoByTwo() Then
            Gatherer += 18
        ElseIf IsItPossibleToExchangeThreeConsecutiveCornerPieces() Then
            Gatherer += 13
        ElseIf IsItPossibleToExchangeTwoAdjacentCornerPieces() Then
            Gatherer += 9
        ElseIf IsItPossibleToExchangeTwoOppositeCornerPieces() Then
            Gatherer += 11
        End If

        Return Gatherer
    End Function


    Private Function IsTheYellowFaceSolvedAsWellAsTheFirstRow() As Boolean
        If Not IsTheFaceSolved(YellowColour) Then Return False
        Dim Counter As Integer
        For Counter = 1 To 4
            If Not AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), {0, 1, 2}) Then Return False
        Next
        Return True
    End Function

    Private Function AreTheEdgePiecesOfTheFirstRowSolved() As Boolean
        ' Damos por hecho que la cara amarilla ya está montada, por eso sólo miramos los bordes de la primera línea sin comprobar si está montada la cara
        ' We just presuppose that the yellow face is already solved, that's why we just focus on the edge pieces of the first line without checking if the yellow face is solved
        Dim Counter As Integer
        For Counter = 1 To 4
            If Not AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 1) Then Return False
        Next
        Return True
    End Function

    Private Function AreTheCornerPiecesOfTheFirstRowSolved() As Boolean
        ' Damos por hecho que la cara amarilla ya está montada, por eso sólo miramos las esquinas de la primera línea sin comprobar si está montada la cara
        ' We presuppose that the yellow face is already solved, that's why we only focus on the corner pieces of the first row without checking if the yellow face is solved
        Dim Counter As Integer
        For Counter = 1 To 4
            If Not AreAllTheSquaresOfTheSoughtColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), {0, 2}) Then Return False
        Next
        Return True
    End Function


    Private Function IsItPossibleToExchangeFourAdjacentEdgePiecesTwoByTwo() As Boolean
        Dim Gatherer, CurrentFace, FollowingFace As Integer
        Gatherer = 0
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then Gatherer += 1
            End If
        Next
        If Gatherer = 2 Then
            Return True
        ElseIf Gatherer < 2 Then
            Return False
        Else
            WeTerminateWithError(22) : Stop
        End If
    End Function

    Private Function IsItPossibleToExchangeFourEdgePiecesOppositeTwoByTwo() As Boolean
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 2
            FollowingFace = ((CurrentFace + 1) Mod 4) + 1
            If Not AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then Return False
            If Not AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then Return False
        Next
        Return True
    End Function

    Private Function IsItPossibleToExchangeFourCornerPiecesAdjacentTwoByTwo() As Boolean
        Dim Gatherer, CurrentFace, FollowingFace, PreviousFace As Integer
        Gatherer = 0
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 0) Then Gatherer += 1
            End If
        Next
        If Gatherer = 2 Then
            Return True
        ElseIf Gatherer < 4 Then
            Return False
        Else
            WeTerminateWithError(22) : Stop
        End If
    End Function

    Private Function IsItPossibleToExchangeFourCornerPiecesOppositeTwoByTwo() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 2
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1
            If Not AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then Return False
            If Not AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then Return False
            If Not AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then Return False
            If Not AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then Return False
        Next
        Return True
    End Function


    Private Function IsItPossibleToExchangeFourDiabolicalCornerPieces() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = FaceOne Mod 4 + 1
            FaceThree = FaceTwo Mod 4 + 1
            FaceFour = FaceThree Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then Return True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Return False
    End Function


    Private Sub ExchangeFourDiabolicalCornerPieces()
        If Not IsItPossibleToExchangeFourDiabolicalCornerPieces() Then WeTerminateWithError(57) : Stop

        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = FaceOne Mod 4 + 1
            FaceThree = FaceTwo Mod 4 + 1
            FaceFour = FaceThree Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then

                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceOne, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceTwo, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceThree, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceFour, 5)

                                            EmbedDirectlyLowerCorner()
                                            Exit Sub


                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next

        WeTerminateWithError(57) : Stop
    End Sub


    Private Function IsItPossibleToExchangeFourConsecutiveCornerPiecesInASCENDINGOrder() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = FaceOne Mod 4 + 1
            FaceThree = FaceTwo Mod 4 + 1
            FaceFour = FaceThree Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then Return True
                                    End If
                                End If
                            End If
                        End If
                    End If

                End If
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToExchangeFourConsecutiveCornerPiecesInDESCENDINGOrder() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = FaceOne Mod 4 + 1
            FaceThree = FaceTwo Mod 4 + 1
            FaceFour = FaceThree Mod 4 + 1

            If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then Return True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToExchangeFourConsecutiveCornerPieces() As Boolean
        Return IsItPossibleToExchangeFourConsecutiveCornerPiecesInASCENDINGOrder() Xor IsItPossibleToExchangeFourConsecutiveCornerPiecesInDESCENDINGOrder()
    End Function


    Private Sub ExchangeFourConsecutiveCornerPiecesInASCENDINGOrder()
        If Not IsItPossibleToExchangeFourConsecutiveCornerPiecesInASCENDINGOrder() Then WeTerminateWithError(57) : Stop

        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = FaceOne Mod 4 + 1
            FaceThree = FaceTwo Mod 4 + 1
            FaceFour = FaceThree Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then

                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceOne, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceTwo, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceFour, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceFour, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceFour, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceFour, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceFour, 5)

                                            EmbedDirectlyLowerCorner()
                                            Exit Sub

                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If

                End If
            End If
        Next

        WeTerminateWithError(57) : Stop
    End Sub

    Private Sub ExchangeFourConsecutiveCornerPiecesInDESCENDINGOrder()
        If Not IsItPossibleToExchangeFourConsecutiveCornerPiecesInDESCENDINGOrder() Then WeTerminateWithError(57) : Stop

        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = FaceOne Mod 4 + 1
            FaceThree = FaceTwo Mod 4 + 1
            FaceFour = FaceThree Mod 4 + 1

            If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then

                                            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(FaceOne, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceThree, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceOne, 5)

                                            EmbedDirectlyLowerCorner()
                                            Exit Sub

                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If

            End If
        Next

        WeTerminateWithError(57) : Stop
    End Sub

    Private Sub ExchangeFourConsecutiveCornerPieces()
        If IsItPossibleToExchangeFourConsecutiveCornerPiecesInASCENDINGOrder() Then
            ExchangeFourConsecutiveCornerPiecesInASCENDINGOrder()
        ElseIf IsItPossibleToExchangeFourConsecutiveCornerPiecesInDESCENDINGOrder() Then
            ExchangeFourConsecutiveCornerPiecesInDESCENDINGOrder()
        Else
            WeTerminateWithError(57) : Stop
        End If
    End Sub


    Private Function IsItPossibleToExchangeThreeConsecutiveCornerPiecesInASCENDINGOrder() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then Return True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToExchangeThreeConsecutiveCornerPiecesInDESCENDINGOrder() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1

            If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then Return True
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Next
        Return False
    End Function

    Private Function IsItPossibleToExchangeThreeConsecutiveCornerPieces() As Boolean
        Return IsItPossibleToExchangeThreeConsecutiveCornerPiecesInASCENDINGOrder() Xor IsItPossibleToExchangeThreeConsecutiveCornerPiecesInDESCENDINGOrder()
    End Function


    Private Sub ExchangeThreeConsecutiveCornerPiecesInASCENDINGOrder()
        If Not IsItPossibleToExchangeThreeConsecutiveCornerPiecesInASCENDINGOrder() Then WeTerminateWithError(57) : Stop

        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then

                                            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90FrontFaceClockwise(FaceTwo, 5)

                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceThree, 5)

                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceTwo, 5)

                                            Exit Sub

                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
        Next

        WeTerminateWithError(57) : Stop
    End Sub

    Private Sub ExchangeThreeConsecutiveCornerPiecesInDESCENDINGOrder()
        If Not IsItPossibleToExchangeThreeConsecutiveCornerPiecesInDESCENDINGOrder() Then WeTerminateWithError(57) : Stop

        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1

            If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 0) Then
                            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 2) Then
                                If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                                    If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                                        If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 0) Then

                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceTwo, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceTwo, 5)

                                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceThree, 5)
                                            CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceThree, 5)

                                            Exit Sub

                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

        Next
        WeTerminateWithError(57) : Stop
    End Sub

    Private Sub ExchangeThreeConsecutiveCornerPieces()
        If IsItPossibleToExchangeThreeConsecutiveCornerPiecesInASCENDINGOrder() Then
            ExchangeThreeConsecutiveCornerPiecesInASCENDINGOrder()
        ElseIf IsItPossibleToExchangeThreeConsecutiveCornerPiecesInDESCENDINGOrder() Then
            ExchangeThreeConsecutiveCornerPiecesInDESCENDINGOrder()
        Else
            WeTerminateWithError(57) : Stop
        End If
    End Sub


    Private Function IsItPossibleToExchangeTwoAdjacentCornerPieces() As Boolean
        Dim CurrentFace, FollowingFace, PreviousFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 0) Then Return True
            End If
        Next
        Return False
    End Function

    Private Sub ExchangeTwoAdjacentCornerPieces()
        If Not IsItPossibleToExchangeTwoAdjacentCornerPieces() Then WeTerminateWithError(57) : Stop

        Dim CurrentFace, FollowingFace, PreviousFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            PreviousFace = (CurrentFace + 2) Mod 4 + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 2) Then
                If AreAllTheSquaresOfTheSoughtColour(PreviousFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 0) Then
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    Exit Sub
                End If
            End If
        Next

        WeTerminateWithError(57) : Stop
    End Sub

    Private Function IsItPossibleToExchangeTwoOppositeCornerPieces() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Private Sub ExchangeTwoOppositeCornerPieces()
        If Not IsItPossibleToExchangeTwoOppositeCornerPieces() Then WeTerminateWithError(57) : Stop

        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = (FaceTwo Mod 4) + 1
            FaceFour = (FaceThree Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 0) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 2) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 0) Then
                        If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 2) Then
                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90DownFaceClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90BackFaceClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)
                            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceOne, 5)
                            Exit Sub
                        End If
                    End If
                End If
            End If
        Next

        WeTerminateWithError(57) : Stop
    End Sub


    Private Function IsItPossibleToExchangeFourDiabolicalEdgePieces() As Boolean
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = ((FaceTwo + 1) Mod 4) + 1
            FaceFour = (FaceTwo Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 1) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 1) Then
                        If Not AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 1) Then WeTerminateWithError(56) : Stop
                        Return True
                    End If
                End If
            End If
        Next
        Return False
    End Function

    Private Sub ExchangeFourDiabolicalEdgePieces()
        If Not IsItPossibleToExchangeFourDiabolicalEdgePieces() Then WeTerminateWithError(55) : Stop
        Dim FaceOne, FaceTwo, FaceThree, FaceFour As Integer
        For FaceOne = 1 To 4
            FaceTwo = (FaceOne Mod 4) + 1
            FaceThree = ((FaceTwo + 1) Mod 4) + 1
            FaceFour = (FaceTwo Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FaceTwo, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceOne), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceThree, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceTwo), 1) Then
                    If AreAllTheSquaresOfTheSoughtColour(FaceFour, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceThree), 1) Then
                        If Not AreAllTheSquaresOfTheSoughtColour(FaceOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFour), 1) Then WeTerminateWithError(56) : Stop
                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceOne, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceOne, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(FaceOne, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceOne, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceOne, 5)
                        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceOne, 5)

                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceTwo, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceTwo, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(FaceTwo, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceTwo, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceTwo, 5)

                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceThree, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceThree, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(FaceThree, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceThree, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceThree, 5)
                        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(FaceThree, 5)

                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(FaceFour, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(FaceFour, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(FaceFour, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(FaceFour, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(FaceFour, 5)

                        EmbedDirectlyLowerEdge()
                        Exit Sub
                    End If
                End If
            End If
        Next
        WeTerminateWithError(55) : Stop
    End Sub


    Private Function IsItPossibleToExchangeFourConsecutiveEdgePiecesInASCENDINGOrder() As Boolean
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            If NoOneOfTheSquaresIsTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then Return False
        Next
        Return True
    End Function

    Private Function IsItPossibleToExchangeFourConsecutiveEdgePiecesInDESCENDINGOrder() As Boolean
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = ((CurrentFace + 2) Mod 4) + 1
            If NoOneOfTheSquaresIsTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then Return False
        Next
        Return True
    End Function

    Private Function IsItPossibleToExchangeFourConsecutiveEdgePieces() As Boolean
        Return IsItPossibleToExchangeFourConsecutiveEdgePiecesInASCENDINGOrder() Xor IsItPossibleToExchangeFourConsecutiveEdgePiecesInDESCENDINGOrder()
    End Function


    Private Sub ExchangeFourConsecutiveEdgePiecesInASCENDINGOrder()
        If Not IsItPossibleToExchangeFourConsecutiveEdgePiecesInASCENDINGOrder() Then WeTerminateWithError(55) : Stop
        Dim CurrentFace As Integer
        For CurrentFace = 1 To 4
            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)

            If CurrentFace = 1 Then
                CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
            Else
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
            End If

            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)

            If CurrentFace = 1 Then
                CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
            Else
                CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
            End If
        Next
        EmbedDirectlyLowerEdge()
    End Sub

    Private Sub ExchangeFourConsecutiveEdgePiecesInDESCENDINGOrder()
        If Not IsItPossibleToExchangeFourConsecutiveEdgePiecesInDESCENDINGOrder() Then WeTerminateWithError(55) : Stop
        Dim CurrentFace As Integer
        For CurrentFace = 4 To 1 Step -1
            CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)

            If CurrentFace = 4 Then
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
            Else
                CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
            End If

            CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
            CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)

            If CurrentFace = 4 Then
                CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
            Else
                CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
            End If
        Next
        EmbedDirectlyLowerEdge()

    End Sub

    Private Sub ExchangeFourConsecutiveEdgePieces()
        If Not IsItPossibleToExchangeFourConsecutiveEdgePieces() Then WeTerminateWithError(55) : Stop
        If IsItPossibleToExchangeFourConsecutiveEdgePiecesInASCENDINGOrder() Then
            ExchangeFourConsecutiveEdgePiecesInASCENDINGOrder()
        ElseIf IsItPossibleToExchangeFourConsecutiveEdgePiecesInDESCENDINGOrder() Then
            ExchangeFourConsecutiveEdgePiecesInDESCENDINGOrder()
        Else
            WeTerminateWithError(55) : Stop
        End If
    End Sub


    Private Function IsItPossibleToExchangeThreeConsecutiveEdgePiecesInASCENDINGOrder() As Boolean
        Dim CurrentFace, FollowingFace, FaceFollowingTheFollowingOne As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            FaceFollowingTheFollowingOne = (FollowingFace Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFollowingTheFollowingOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                    If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFollowingTheFollowingOne), 1) Then Return True

                End If
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToExchangeThreeConsecutiveEdgePiecesInDESCENDINGOrder() As Boolean
        Dim CurrentFace, FollowingFace, FaceFollowingTheFollowingOne As Integer
        For CurrentFace = 1 To 4
            FollowingFace = ((CurrentFace + 2) Mod 4) + 1
            FaceFollowingTheFollowingOne = ((FollowingFace + 2) Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFollowingTheFollowingOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                    If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFollowingTheFollowingOne), 1) Then Return True
                End If
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToExchangeThreeConsecutiveEdgePieces() As Boolean
        Return IsItPossibleToExchangeThreeConsecutiveEdgePiecesInASCENDINGOrder() Xor IsItPossibleToExchangeThreeConsecutiveEdgePiecesInDESCENDINGOrder()
    End Function

    Private Sub ExchangeThreeConsecutiveEdgePiecesInASCENDINGOrder()
        If Not IsItPossibleToExchangeThreeConsecutiveEdgePiecesInASCENDINGOrder() Then WeTerminateWithError(55) : Stop
        Dim CurrentFace, FollowingFace, FaceFollowingTheFollowingOne As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            FaceFollowingTheFollowingOne = (FollowingFace Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFollowingTheFollowingOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                    If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFollowingTheFollowingOne), 1) Then

                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                        Exit Sub
                    End If
                End If
            End If
        Next
        WeTerminateWithError(55) : Stop
    End Sub

    Private Sub ExchangeThreeConsecutiveEdgePiecesInDESCENDINGOrder()
        If Not IsItPossibleToExchangeThreeConsecutiveEdgePiecesInDESCENDINGOrder() Then WeTerminateWithError(55) : Stop

        Dim CurrentFace, FollowingFace, FaceFollowingTheFollowingOne As Integer
        For CurrentFace = 1 To 4
            FollowingFace = ((CurrentFace + 2) Mod 4) + 1
            FaceFollowingTheFollowingOne = ((FollowingFace + 2) Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FaceFollowingTheFollowingOne, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                    If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceFollowingTheFollowingOne), 1) Then

                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                        CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                        Exit Sub
                    End If
                End If
            End If
        Next

        WeTerminateWithError(55) : Stop
    End Sub

    Private Sub ExchangeThreeConsecutiveEdgePieces()
        If IsItPossibleToExchangeThreeConsecutiveEdgePiecesInASCENDINGOrder() Then
            ExchangeThreeConsecutiveEdgePiecesInASCENDINGOrder()
        ElseIf IsItPossibleToExchangeThreeConsecutiveEdgePiecesInDESCENDINGOrder() Then
            ExchangeThreeConsecutiveEdgePiecesInDESCENDINGOrder()
        Else
            WeTerminateWithError(55) : Stop
        End If
    End Sub


    Private Function IsItPossibleToExchangeTwoOppositeEdgePieces() As Boolean
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = ((CurrentFace + 1) Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then Return True
            End If
        Next
        Return False
    End Function

    Private Sub ExchangeTwoOppositeEdgePieces()
        If Not IsItPossibleToExchangeTwoOppositeEdgePieces() Then WeTerminateWithError(55) : Stop
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = ((CurrentFace + 1) Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    Exit Sub
                End If
            End If
        Next

        WeTerminateWithError(55) : Stop
    End Sub


    Private Function IsItPossibleToExchangeTwoAdjacentEdgePieces() As Boolean
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then Return True
            End If
        Next
        Return False
    End Function

    Private Sub ExchangeTwoAdjacentEdgePieces()
        If Not IsItPossibleToExchangeTwoAdjacentEdgePieces() Then WeTerminateWithError(55) : Stop
        Dim CurrentFace, FollowingFace As Integer
        For CurrentFace = 1 To 4
            FollowingFace = (CurrentFace Mod 4) + 1
            If AreAllTheSquaresOfTheSoughtColour(FollowingFace, CubeThatWeMustSolve.ArrayOfRubiksCube(CurrentFace), 1) Then
                If AreAllTheSquaresOfTheSoughtColour(CurrentFace, CubeThatWeMustSolve.ArrayOfRubiksCube(FollowingFace), 1) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(CurrentFace, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(CurrentFace, 5)
                    Exit Sub
                End If
            End If
        Next

        WeTerminateWithError(55) : Stop
    End Sub


    Public Sub SolveTheYellowFace()
        If IsTheFaceSolved(YellowColour) Then
            Dim TextOfMessageBox As String = "The yellow face is already solved, so we have  nothing to solve here"
            Dim CaptionOfMessageBox As String = "The yellow face is already solved"
            MessageBox.Show(TextOfMessageBox, CaptionOfMessageBox, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End If
        Do While Not IsTheFaceSolved(YellowColour)
            If IsItPossibleToInsertDirectlyAColumnTrio() Then
                InsertDirectlyColumnTrio()
            ElseIf IsItPossibleToInsertINDIRECTLYAColumnTrio() Then
                InsertINDIRECTLYColumnTrio()

            ElseIf IsItPossibleToInsertDirectlyAnUndergrounColumnTrio() Then
                InsertDirectlyUndergroundColumnTrio()
            ElseIf IsItPossibleToInsertINDIRECTLYAnUndergroundColumnTrio() Then
                InsertINDIRECTLYUndergroundColumnTrio()

            ElseIf IsItPossibleToInsertDirectlyAVerticalPair() Then
                InsertDirectlyVerticalPair()
            ElseIf IsItPossibleToInsertINDIRECTLYAVerticalPair() Then
                InsertINDIRECTLYVerticalPair()

            ElseIf IsItPossibleToInsertDirectlyAnUndergroundPair() Then
                InsertDirectlyUndergroundPair()
            ElseIf IsItPossibleToInsertINDIRECTLYAnUndergrounPair() Then
                InsertINDIRECTLYUndergroundPair()

            ElseIf IsItPossibleToInsertDirectlyAnUpperCorner() Then
                InsertDirectlyUpperCorner()
            ElseIf IsItPossibleToInsertDirectlyALateralEdge() Then
                InsertDirectlyLateralEdge()
            ElseIf IsItPossibleToInsertDirectlyALowerCorner() Then
                InsertDirectlyLowerCorner()

            ElseIf IsItPossibleToInsertINDIRECTLYALateralEdge() Then
                InsertINDIRECTLYALateralEdge()
            ElseIf IsItPossibleToInsertINDIRECTLYALowerCorner() Then
                InsertINDIRECTLYLowerCorner()

            ElseIf IsItPossibleToEmbedDirectlyALowerEdgeCornerPair() Then
                InsertDirectlyLowerEdgeCornerPair()
            ElseIf IsItPossibleToEmbedINDIRECTLYALowerEdgeCornerPair() Then
                EmbedINDIRECTLYLowerEdgeCornerPair()

            ElseIf IsItPossibleToEmbedDirectlyAHorizontalPair() Then
                EmbedDirectlyHorizontalPair()
            ElseIf IsItPossibleToEmbedINDIRECTLYAHorizontalPair() Then
                EmbedINDIRECTLYHorizontalPair()

            ElseIf IsItPossibleToEmbedDirectlyALowerCorner() Then
                EmbedDirectlyLowerCorner()
            ElseIf IsItPossibleToEmbedDirectlyALowerEdge() Then
                EmbedDirectlyLowerEdge()
            ElseIf IsItPossibleToEmbedDirectlyALateralEdge() Then
                EmbedDirectlyLateralEdge()

            ElseIf IsItPossibleToEmbedINDIRECTLYALowerCorner() Then
                EmbedINDIRECTLYLowerCorner()
            ElseIf IsItPossibleToEmbedINDIRECTLYALowerEdge() Then
                EmbedINDIRECTLYLowerEdge()
            ElseIf IsItPossibleToEmbedINDIRECTLYALateralEdge() Then
                EmbedINDIRECTLYLateralEdge()

            ElseIf IsItPossibleToPlaceAnyhowAnUpperCorner() Then
                PlaceAnyhowUpperCorner()
            ElseIf IsItPossibleToPlaceAnyhowAnUpperEdge() Then
                PlaceAnyhowUpperEdge()

            ElseIf IsItPossibleToEmbedDirectlyAnUndergroundEdge() Then
                EmbedDirectlyUndergroundEdge()
            ElseIf IsItPossibleToEmbedINDIRECTLYAnUndergroundEdge() Then
                EmbedINDIRECTLYUndergroundEdge()

            ElseIf IsItPossibleToEmbedDirectlyAnUndergroundCorner() Then
                EmbedDirectlyUndergroundCorner()
            ElseIf IsItPossibleToEmbedINDIRECTLYAnUndergroundCorner() Then
                EmbedINDIRECTLYUndergroundCorner()



            Else
                WeTerminateWithError(49) : Stop
            End If
        Loop
    End Sub


    Private Function IsItPossibleToInsertINDIRECTLYAnUndergrounPair() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        Dim Counter As Integer
        Dim Result As Boolean = False
        For Counter = 1 To 4
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
            If IsItPossibleToInsertDirectlyAnUndergroundPair() Then Result = True
        Next
        Return Result
    End Function


    Private Function IsItPossibleToInsertDirectlyAnUndergroundPair() As Boolean
        ' Como da igual ponerlo por la izquierda o por la derecha (porque si te pones enfrente es el contrario), vamos a suponer que va por la izquierda
        ' Since it is the same inserting it through the left or through the right (because if you place yourself opposite, it is the contrary), let's suppose that it goes through the left

        If Not IsThereAboveAnyFreeLine() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(2), WhiteSquares(2) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 3, 6} : WhiteSquares = {0, 3, 6}
                Case 2 : YellowSquares = {0, 1, 2} : WhiteSquares = {6, 7, 8}
                Case 3 : YellowSquares = {2, 5, 8} : WhiteSquares = {2, 5, 8}
                Case 4 : YellowSquares = {6, 7, 8} : WhiteSquares = {0, 1, 2}
            End Select
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                Dim IndexesOfChosenPair(1), ValuesOfChosenPair(1) As Integer

                IndexesOfChosenPair = {0, 1}
                ValuesOfChosenPair = {WhiteSquares(IndexesOfChosenPair(0)), WhiteSquares(IndexesOfChosenPair(1))}
                If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), ValuesOfChosenPair) Then Return True

                IndexesOfChosenPair = {0, 2}
                ValuesOfChosenPair = {WhiteSquares(IndexesOfChosenPair(0)), WhiteSquares(IndexesOfChosenPair(1))}
                If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), ValuesOfChosenPair) Then Return True

                IndexesOfChosenPair = {1, 2}
                ValuesOfChosenPair = {WhiteSquares(IndexesOfChosenPair(0)), WhiteSquares(IndexesOfChosenPair(1))}
                If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), ValuesOfChosenPair) Then Return True
            End If
        Next
        Return False
    End Function

    Private Sub InsertDirectlyUndergroundPair()
        If Not IsItPossibleToInsertDirectlyAnUndergroundPair() Then WeTerminateWithError(53) : Stop
        ' Como da igual ponerlo por la izquierda o por la derecha (porque si te pones enfrente es el contrario), vamos a suponer que va por la izquierda
        ' Since it is the same putting it through the left or through the right (because if you put yourself opposite, it is the contrary), let's suppose that it will go through the left
        Dim Counter As Integer
        Dim YellowSquares(2), WhiteSquares(2) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 3, 6} : WhiteSquares = {0, 3, 6}
                Case 2 : YellowSquares = {0, 1, 2} : WhiteSquares = {6, 7, 8}
                Case 3 : YellowSquares = {2, 5, 8} : WhiteSquares = {2, 5, 8}
                Case 4 : YellowSquares = {6, 7, 8} : WhiteSquares = {0, 1, 2}
            End Select
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                Dim IndexesOfChosenPair(1), ValuesOfChosenPair(1) As Integer

                IndexesOfChosenPair = {0, 1}
                ValuesOfChosenPair = {WhiteSquares(IndexesOfChosenPair(0)), WhiteSquares(IndexesOfChosenPair(1))}
                If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), ValuesOfChosenPair) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If

                IndexesOfChosenPair = {0, 2}
                ValuesOfChosenPair = {WhiteSquares(IndexesOfChosenPair(0)), WhiteSquares(IndexesOfChosenPair(1))}
                If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), ValuesOfChosenPair) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If

                IndexesOfChosenPair = {1, 2}
                ValuesOfChosenPair = {WhiteSquares(IndexesOfChosenPair(0)), WhiteSquares(IndexesOfChosenPair(1))}
                If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), ValuesOfChosenPair) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(53) : Stop
    End Sub

    Private Sub InsertINDIRECTLYUndergroundPair()
        If Not IsItPossibleToInsertINDIRECTLYAnUndergrounPair() Then WeTerminateWithError(54) : Stop
        Do While Not IsItPossibleToInsertDirectlyAnUndergroundPair()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        InsertDirectlyUndergroundPair()
    End Sub


    Private Function IsItPossibleToInsertDirectlyAnUndergrounColumnTrio() As Boolean
        'Como da igual ponerlo por la izquierda o por la derecha (porque si te pones enfrente es el contrario), vamos a suponer que va por la izquierda
        ' Since it is the same putting it through the left or through the right (because if you put yourself opposite, it is the contrary), let's suppose that it will go through the left

        If Not IsThereAboveAnyFreeLine() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(), WhiteSquares() As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 3, 6} : WhiteSquares = {0, 3, 6}
                Case 2 : YellowSquares = {0, 1, 2} : WhiteSquares = {6, 7, 8}
                Case 3 : YellowSquares = {2, 5, 8} : WhiteSquares = {2, 5, 8}
                Case 4 : YellowSquares = {6, 7, 8} : WhiteSquares = {0, 1, 2}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquares) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYAnUndergroundColumnTrio() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        Dim Counter As Integer
        Dim WhiteSquares() As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : WhiteSquares = {0, 3, 6}
                Case 2 : WhiteSquares = {6, 7, 8}
                Case 3 : WhiteSquares = {2, 5, 8}
                Case 4 : WhiteSquares = {0, 1, 2}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquares) Then Return True
        Next
        Return False
    End Function

    Private Sub InsertDirectlyUndergroundColumnTrio()
        If Not IsItPossibleToInsertDirectlyAnUndergrounColumnTrio() Then WeTerminateWithError(51) : Stop
        Dim Counter As Integer
        Dim YellowSquares(), WhiteSquares() As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 3, 6} : WhiteSquares = {0, 3, 6}
                Case 2 : YellowSquares = {0, 1, 2} : WhiteSquares = {6, 7, 8}
                Case 3 : YellowSquares = {2, 5, 8} : WhiteSquares = {2, 5, 8}
                Case 4 : YellowSquares = {6, 7, 8} : WhiteSquares = {0, 1, 2}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquares) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(51) : Stop
    End Sub

    Private Sub InsertINDIRECTLYUndergroundColumnTrio()
        If Not IsItPossibleToInsertINDIRECTLYAnUndergroundColumnTrio() Then WeTerminateWithError(52) : Stop
        Do While Not IsItPossibleToInsertDirectlyAnUndergrounColumnTrio()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        InsertDirectlyUndergroundColumnTrio()
    End Sub


    Private Function IsItPossibleToPlaceAnyhowAnUpperEdge() As Boolean
        Return IsThereBelowAnyUpperEdge()
    End Function

    Private Sub PlaceAnyhowUpperEdge()
        If Not IsItPossibleToPlaceAnyhowAnUpperEdge() Then WeTerminateWithError(58) : Stop
        Dim Counter As Integer
        For Counter = 1 To 4
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 1) Then
                CubeThatWeMustSolve.Rotate90LeftFaceClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90UpFaceClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90FrontFaceClockwise(Counter, 5)
                CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                Exit Sub
            End If
        Next
        WeTerminateWithError(58) : Stop
    End Sub


    Private Function IsItPossibleToPlaceAnyhowLeftUpperCorner() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 0) Then Return True
        Next
        Return False
    End Function

    Private Function IsItPossibleToPlaceAnyhowRightUpperCorner() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 2) Then Return True
        Next
        Return False
    End Function

    Private Function IsItPossibleToPlaceAnyhowAnUpperCorner() As Boolean
        Return IsItPossibleToPlaceAnyhowLeftUpperCorner() Or IsItPossibleToPlaceAnyhowRightUpperCorner()
    End Function

    Private Sub PlaceAnyhowLeftUpperCorner()
        If Not IsItPossibleToPlaceAnyhowLeftUpperCorner() Then WeTerminateWithError(50) : Stop
        Dim Counter As Integer
        For Counter = 1 To 4
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 0) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                Exit Sub
            End If
        Next
        WeTerminateWithError(50) : Stop
    End Sub

    Private Sub PlaceAnyhowRightUpperCorner()
        If Not IsItPossibleToPlaceAnyhowRightUpperCorner() Then WeTerminateWithError(50) : Stop
        Dim Counter As Integer
        For Counter = 1 To 4
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 2) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                Exit Sub
            End If
        Next
        WeTerminateWithError(50) : Stop
    End Sub

    Private Sub PlaceAnyhowUpperCorner()
        If IsItPossibleToPlaceAnyhowLeftUpperCorner() Then
            PlaceAnyhowLeftUpperCorner()
        ElseIf IsItPossibleToPlaceAnyhowRightUpperCorner() Then
            PlaceAnyhowRightUpperCorner()
        Else
            WeTerminateWithError(50) : Stop
        End If
    End Sub


    Private Sub EmbedDirectlyLeftLowerEdgeCornerPair()
        If Not IsItPossibleToEmbedDirectlyLeftLowerEdgeCornerPair() Then WeTerminateWithError(47) : Stop
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 1}
                Case 2 : YellowSquares = {2, 5}
                Case 3 : YellowSquares = {7, 8}
                Case 4 : YellowSquares = {3, 6}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 3, 6) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90UpFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(47) : Stop
    End Sub

    Private Sub EmbedDirectlyRightLowerEdgeCornerPair()
        If Not IsItPossibleToEmbedDirectlyRightLowerEdgeCornerPair() Then WeTerminateWithError(47) : Stop

        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {1, 2}
                Case 2 : YellowSquares = {5, 8}
                Case 3 : YellowSquares = {6, 7}
                Case 4 : YellowSquares = {0, 3}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 5, 8) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90UpFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(47) : Stop
    End Sub

    Private Sub InsertDirectlyLowerEdgeCornerPair()
        If Not IsItPossibleToEmbedDirectlyALowerEdgeCornerPair() Then WeTerminateWithError(47) : Stop
        If IsItPossibleToEmbedDirectlyLeftLowerEdgeCornerPair() Then
            EmbedDirectlyLeftLowerEdgeCornerPair()
        ElseIf IsItPossibleToEmbedDirectlyRightLowerEdgeCornerPair() Then
            EmbedDirectlyRightLowerEdgeCornerPair()
        Else
            WeTerminateWithError(47) : Stop
        End If
    End Sub

    Private Sub EmbedINDIRECTLYLowerEdgeCornerPair()
        If Not IsItPossibleToEmbedINDIRECTLYALowerEdgeCornerPair() Then WeTerminateWithError(48) : Stop

        If (IsThereBelowAnyLeftLowerEdgeCornerPair() AndAlso IsItPossibleToEmbedINDIRECTLYALeftLowerEdgeCornerPair()) Or
                (IsThereBelowAnyRightLowerEdgeCornerPair() AndAlso IsItPossibleToEmbedINDIRECTLYRightLowerEdgeCornerPair()) Then
            Do While Not IsItPossibleToEmbedDirectlyALowerEdgeCornerPair()
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
            Loop
            InsertDirectlyLowerEdgeCornerPair()
        Else
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
            EmbedINDIRECTLYLowerEdgeCornerPair()
        End If
    End Sub


    Private Function IsItPossibleToEmbedDirectlyAnUndergroundCorner() As Boolean
        If Not IsThereUndergroundAnyCorner() Then Return False
        Dim Counter, YellowSquare, WhiteSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 8 : WhiteSquare = 2
                Case 2 : YellowSquare = 6 : WhiteSquare = 0
                Case 3 : YellowSquare = 0 : WhiteSquare = 6
                Case 4 : YellowSquare = 2 : WhiteSquare = 8
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquare) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYAnUndergroundCorner() As Boolean
        Return IsThereUndergroundAnyCorner()
    End Function

    Private Sub EmbedDirectlyUndergroundCorner()
        If Not IsItPossibleToEmbedDirectlyAnUndergroundCorner() Then WeTerminateWithError(44) : Stop
        Dim Counter, YellowSquare, WhiteSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 8 : WhiteSquare = 2
                Case 2 : YellowSquare = 6 : WhiteSquare = 0
                Case 3 : YellowSquare = 0 : WhiteSquare = 6
                Case 4 : YellowSquare = 2 : WhiteSquare = 8
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquare) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90UpFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(44) : Stop
    End Sub

    Private Sub EmbedINDIRECTLYUndergroundCorner()
        If Not IsItPossibleToEmbedINDIRECTLYAnUndergroundCorner() Then WeTerminateWithError(45) : Stop
        Do While Not IsItPossibleToEmbedDirectlyAnUndergroundCorner()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        EmbedDirectlyUndergroundCorner()
    End Sub


    Private Function IsItPossibleToEmbedDirectlyAnUndergroundEdge() As Boolean
        If Not IsThereUndergroundAnyEdge() Then Return False
        Dim Counter, YellowSquare, WhiteSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 7 : WhiteSquare = 1
                Case 2 : YellowSquare = 3 : WhiteSquare = 3
                Case 3 : YellowSquare = 1 : WhiteSquare = 7
                Case 4 : YellowSquare = 5 : WhiteSquare = 5
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquare) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYAnUndergroundEdge() As Boolean
        Return IsThereUndergroundAnyEdge()
    End Function

    Private Sub EmbedDirectlyUndergroundEdge()
        If Not IsItPossibleToEmbedDirectlyAnUndergroundEdge() Then WeTerminateWithError(42) : Stop
        Dim Counter, YellowSquare, WhiteSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 7 : WhiteSquare = 1
                Case 2 : YellowSquare = 3 : WhiteSquare = 3
                Case 3 : YellowSquare = 1 : WhiteSquare = 7
                Case 4 : YellowSquare = 5 : WhiteSquare = 5
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(5), WhiteSquare) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90LeftFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(42) : Stop
    End Sub

    Private Sub EmbedINDIRECTLYUndergroundEdge()
        If Not IsItPossibleToEmbedINDIRECTLYAnUndergroundEdge() Then WeTerminateWithError(43) : Stop
        Do While Not IsItPossibleToEmbedDirectlyAnUndergroundEdge()
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
        Loop
        EmbedDirectlyUndergroundEdge()
    End Sub


    Private Sub EmbedDirectlyCentreLeftPair()
        If Not IsItPossibleToEmbedDirectlyCentreLeftPair() Then WeTerminateWithError(40) : Stop
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 3}
                Case 2 : YellowSquares = {1, 2}
                Case 3 : YellowSquares = {5, 8}
                Case 4 : YellowSquares = {6, 7}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 6, 7) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(40) : Stop
    End Sub

    Private Sub EmbedDirectlyCentreRightPair()
        If Not IsItPossibleToEmbedDirectlyACentreRightPair() Then WeTerminateWithError(40) : Stop
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {2, 5}
                Case 2 : YellowSquares = {7, 8}
                Case 3 : YellowSquares = {3, 6}
                Case 4 : YellowSquares = {0, 1}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7, 8) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(40) : Stop
    End Sub

    Private Sub EmbedDirectlyHorizontalPair()
        If IsItPossibleToEmbedDirectlyCentreLeftPair() Then
            EmbedDirectlyCentreLeftPair()
        ElseIf IsItPossibleToEmbedDirectlyACentreRightPair() Then
            EmbedDirectlyCentreRightPair()
        Else
            WeTerminateWithError(40) : Stop
        End If
    End Sub

    Private Sub EmbedINDIRECTLYCentreLeftPair()
        If Not IsItPossibleToEmbedINDIRECTLYACentreLeftPair() Then WeTerminateWithError(41) : Stop
        Do While Not IsItPossibleToEmbedDirectlyCentreLeftPair()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        EmbedDirectlyCentreLeftPair()
    End Sub

    Private Sub EmbedINDIRECTLYCentreRightPair()
        If Not IsItPossibleToEmbedINDIRECTLYACentreRightPair() Then WeTerminateWithError(41) : Stop
        Do While Not IsItPossibleToEmbedDirectlyACentreRightPair()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        EmbedDirectlyCentreRightPair()
    End Sub

    Private Sub EmbedINDIRECTLYHorizontalPair()
        If IsItPossibleToEmbedDirectlyAHorizontalPair() Then
            EmbedDirectlyHorizontalPair()
        ElseIf IsItPossibleToEmbedINDIRECTLYACentreLeftPair() Then
            EmbedINDIRECTLYCentreLeftPair()
        ElseIf IsItPossibleToEmbedINDIRECTLYACentreRightPair() Then
            EmbedINDIRECTLYCentreRightPair()
        Else
            WeTerminateWithError(41) : Stop
        End If
    End Sub


    Private Function IsItPossibleToEmbedDirectlyCentreLeftPair() As Boolean
        If Not IsThereBelowAnyCentreLeftPair() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 3}
                Case 2 : YellowSquares = {1, 2}
                Case 3 : YellowSquares = {5, 8}
                Case 4 : YellowSquares = {6, 7}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 6, 7) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyACentreRightPair() As Boolean
        If Not IsThereBelowAnyCentreRightPair() Then Return False
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {2, 5}
                Case 2 : YellowSquares = {7, 8}
                Case 3 : YellowSquares = {3, 6}
                Case 4 : YellowSquares = {0, 1}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7, 8) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyAHorizontalPair() As Boolean
        Return IsItPossibleToEmbedDirectlyCentreLeftPair() Or IsItPossibleToEmbedDirectlyACentreRightPair()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYACentreLeftPair() As Boolean
        Return IsThereBelowAnyCentreLeftPair() And IsThereAboveAnyFreeCentreLeftPair()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYACentreRightPair() As Boolean
        Return IsThereBelowAnyCentreRightPair() And IsThereAboveAnyFreeCentreRightPair()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYAHorizontalPair() As Boolean
        Return IsItPossibleToEmbedINDIRECTLYACentreLeftPair() Or IsItPossibleToEmbedINDIRECTLYACentreRightPair()
    End Function


    Private Function IsItPossibleToEmbedDirectlyLeftLowerEdgeCornerPair() As Boolean
        If Not IsThereBelowAnyLeftLowerEdgeCornerPair() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 1}
                Case 2 : YellowSquares = {2, 5}
                Case 3 : YellowSquares = {7, 8}
                Case 4 : YellowSquares = {3, 6}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 3, 6) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyRightLowerEdgeCornerPair() As Boolean
        If Not IsThereBelowAnyRightLowerEdgeCornerPair() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {1, 2}
                Case 2 : YellowSquares = {5, 8}
                Case 3 : YellowSquares = {6, 7}
                Case 4 : YellowSquares = {0, 3}
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 5, 8) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyALowerEdgeCornerPair() As Boolean
        Return IsItPossibleToEmbedDirectlyLeftLowerEdgeCornerPair() Or IsItPossibleToEmbedDirectlyRightLowerEdgeCornerPair()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYALeftLowerEdgeCornerPair() As Boolean
        If Not IsItPossibleBelowAnyLeftLowerEdgeCornerPair() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {0, 1}
                Case 2 : YellowSquares = {2, 5}
                Case 3 : YellowSquares = {7, 8}
                Case 4 : YellowSquares = {3, 6}
            End Select
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYRightLowerEdgeCornerPair() As Boolean
        If Not IsItPossibleAnyRightLowerEdgeCornerPair() Then Return False
        Dim Counter As Integer
        Dim YellowSquares(1) As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquares = {1, 2}
                Case 2 : YellowSquares = {5, 8}
                Case 3 : YellowSquares = {6, 7}
                Case 4 : YellowSquares = {0, 3}
            End Select
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquares) Then Return True
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYALowerEdgeCornerPair() As Boolean
        Return IsItPossibleToEmbedINDIRECTLYALeftLowerEdgeCornerPair() Or IsItPossibleToEmbedINDIRECTLYRightLowerEdgeCornerPair()
    End Function


    Private Function IsItPossibleToEmbedDirectlyALeftEdge() As Boolean
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 1
                Case 2 : YellowSquare = 5
                Case 3 : YellowSquare = 7
                Case 4 : YellowSquare = 3
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 3) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyARightEdge() As Boolean
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 1
                Case 2 : YellowSquare = 5
                Case 3 : YellowSquare = 7
                Case 4 : YellowSquare = 3
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 5) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyALateralEdge() As Boolean
        Return IsItPossibleToEmbedDirectlyALeftEdge() Or IsItPossibleToEmbedDirectlyARightEdge()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYALeftEdge() As Boolean
        Return IsThereBelowAnyLeftEdge()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYARightEdge() As Boolean
        Return IsThereBelowAnyRightEdge()
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYALateralEdge() As Boolean
        Return IsItPossibleToEmbedINDIRECTLYALeftEdge() Or IsItPossibleToEmbedINDIRECTLYARightEdge()
    End Function


    Private Sub EmbedDirectlyLeftEdge()
        If Not IsItPossibleToEmbedDirectlyALeftEdge() Then WeTerminateWithError(38) : Stop
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 1
                Case 2 : YellowSquare = 5
                Case 3 : YellowSquare = 7
                Case 4 : YellowSquare = 3
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 3) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90UpFaceCounterClockwise(Counter, 5)
                    If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube((Counter Mod 4) + 1), 6) Then
                        CubeThatWeMustSolve.Rotate90DownFaceClockwise(Counter, 5)
                    End If
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(38) : Stop
    End Sub

    Private Sub EmbedDirectlyRightEdge()
        If Not IsItPossibleToEmbedDirectlyARightEdge() Then WeTerminateWithError(38) : Stop
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 1
                Case 2 : YellowSquare = 5
                Case 3 : YellowSquare = 7
                Case 4 : YellowSquare = 3
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 5) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90UpFaceClockwise(Counter, 5)
                    If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(((Counter + 2) Mod 4) + 1), 8) Then
                        CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(Counter, 5)
                    End If
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(38) : Stop
    End Sub

    Private Sub EmbedDirectlyLateralEdge()
        If Not IsItPossibleToEmbedDirectlyALateralEdge() Then WeTerminateWithError(38) : Stop
        If IsItPossibleToEmbedDirectlyALeftEdge() Then
            EmbedDirectlyLeftEdge()
        ElseIf IsItPossibleToEmbedDirectlyARightEdge() Then
            EmbedDirectlyRightEdge()
        Else
            WeTerminateWithError(38) : Stop
        End If
    End Sub

    Private Sub EmbedINDIRECTLYLateralEdge()
        If Not IsItPossibleToEmbedINDIRECTLYALateralEdge() Then WeTerminateWithError(39) : Stop
        Do While Not IsItPossibleToEmbedDirectlyALateralEdge()
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
        Loop
        EmbedDirectlyLateralEdge()
    End Sub


    Private Sub EmbedDirectlyLeftLowerCorner()
        If Not IsItPossibleToEmbedDirectlyALeftLowerCorner() Then WeTerminateWithError(35) : Stop
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 0
                Case 2 : YellowSquare = 2
                Case 3 : YellowSquare = 8
                Case 4 : YellowSquare = 6
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 6) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(35) : Stop
    End Sub

    Private Sub EmbedDirectlyRightLowerCorner()
        If Not IsItPossibleToEmbedDirectlyARightLowerCorner() Then WeTerminateWithError(35) : Stop
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 2
                Case 2 : YellowSquare = 8
                Case 3 : YellowSquare = 6
                Case 4 : YellowSquare = 0
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 8) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90DownFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(35) : Stop
    End Sub

    Private Sub EmbedDirectlyLowerCorner()
        If IsItPossibleToEmbedDirectlyALeftLowerCorner() Then
            EmbedDirectlyLeftLowerCorner()
        ElseIf IsItPossibleToEmbedDirectlyARightLowerCorner() Then
            EmbedDirectlyRightLowerCorner()
        Else
            WeTerminateWithError(35) : Stop
        End If
    End Sub

    Private Sub EmbedINDIRECTLYLowerCorner()
        If Not IsItPossibleToEmbedINDIRECTLYALowerCorner() Then WeTerminateWithError(37) : Stop
        Do While Not IsItPossibleToEmbedDirectlyALowerCorner()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        EmbedDirectlyLowerCorner()
    End Sub


    Private Sub EmbedDirectlyLowerEdgeThroughTheLeft()
        If Not IsItPossibleToEmbedDirectlyALowerEdgeThroughTheLeft() Then WeTerminateWithError(36) : Stop
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 3
                Case 2 : YellowSquare = 1
                Case 3 : YellowSquare = 5
                Case 4 : YellowSquare = 7
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(36) : Stop
    End Sub

    Private Sub EmbedDirectlyLowerEdgeThroughTheRight()
        If Not IsItPossibleToEmbedDirectlyALowerEdgeThroughTheRight() Then WeTerminateWithError(36) : Stop
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 5
                Case 2 : YellowSquare = 7
                Case 3 : YellowSquare = 3
                Case 4 : YellowSquare = 1
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90RightFaceClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90BackFaceCounterClockwise(Counter, 5)
                    CubeThatWeMustSolve.Rotate90FrontFaceClockwise(Counter, 5)
                    Exit Sub
                End If
            End If
        Next
        WeTerminateWithError(36) : Stop
    End Sub

    Private Sub EmbedDirectlyLowerEdge()
        If IsItPossibleToEmbedDirectlyALowerEdgeThroughTheLeft() Then
            EmbedDirectlyLowerEdgeThroughTheLeft()
        ElseIf IsItPossibleToEmbedDirectlyALowerEdgeThroughTheRight() Then
            EmbedDirectlyLowerEdgeThroughTheRight()
        Else
            WeTerminateWithError(36) : Stop
        End If
    End Sub

    Private Sub EmbedINDIRECTLYLowerEdge()
        If Not IsItPossibleToEmbedINDIRECTLYALowerEdge() Then WeTerminateWithError(36) : Stop
        Do While Not IsItPossibleToEmbedDirectlyALowerEdge()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        EmbedDirectlyLowerEdge()
    End Sub


    Private Function IsItPossibleToEmbedDirectlyALeftLowerCorner() As Boolean
        If Not IsThereBelowAnyLeftLowerCorner() Then Return False
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 0
                Case 2 : YellowSquare = 2
                Case 3 : YellowSquare = 8
                Case 4 : YellowSquare = 6
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 6) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyARightLowerCorner() As Boolean
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 2
                Case 2 : YellowSquare = 8
                Case 3 : YellowSquare = 6
                Case 4 : YellowSquare = 0
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 8) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyALowerCorner() As Boolean
        Return IsItPossibleToEmbedDirectlyALeftLowerCorner() Or IsItPossibleToEmbedDirectlyARightLowerCorner()
    End Function


    Private Function IsItPossibleToEmbedINDIRECTLYALeftLowerCorner() As Boolean
        If Not IsThereBelowAnyLeftLowerCorner() Then Return False
        Dim Counter As Integer
        For Counter = 0 To 8 Step 2
            If Counter = 4 Then Continue For
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), Counter) Then Return True
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYARightLowerCorner() As Boolean
        If Not IsThereBelowAnyRightLowerCorner() Then Return False
        Dim Counter As Integer
        For Counter = 0 To 8 Step 2
            If Counter = 4 Then Continue For
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), Counter) Then Return True
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedINDIRECTLYALowerCorner() As Boolean
        Return IsItPossibleToEmbedINDIRECTLYALeftLowerCorner() Or IsItPossibleToEmbedINDIRECTLYARightLowerCorner()
    End Function


    Private Function IsItPossibleToEmbedDirectlyALowerEdgeThroughTheLeft() As Boolean
        If Not IsThereBelowAnyLowerEdge() Then Return False
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 3
                Case 2 : YellowSquare = 1
                Case 3 : YellowSquare = 5
                Case 4 : YellowSquare = 7
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyALowerEdgeThroughTheRight() As Boolean
        If Not IsThereBelowAnyLowerEdge() Then Return False
        Dim Counter, YellowSquare As Integer
        For Counter = 1 To 4
            Select Case Counter
                Case 1 : YellowSquare = 5
                Case 2 : YellowSquare = 7
                Case 3 : YellowSquare = 3
                Case 4 : YellowSquare = 1
            End Select
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter), 7) Then
                If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), YellowSquare) Then Return True
            End If
        Next
        Return False
    End Function

    Private Function IsItPossibleToEmbedDirectlyALowerEdge() As Boolean
        Return IsItPossibleToEmbedDirectlyALowerEdgeThroughTheLeft() Or IsItPossibleToEmbedDirectlyALowerEdgeThroughTheRight()
    End Function


    Private Function IsItPossibleToEmbedINDIRECTLYALowerEdge() As Boolean
        Return IsThereBelowAnyLowerEdge()
    End Function


    Private Sub InsertINDIRECTLYVerticalPair()
        If Not IsItPossibleToInsertINDIRECTLYAVerticalPair() Then WeTerminateWithError(32)
        If IsItPossibleToInsertDirectlyAVerticalPair() Then InsertDirectlyVerticalPair() : Exit Sub
        If IsItPossibleToInsertDirectlyALeftUpperCorner() Then
            If IsThereBelowAnyLeftEdge() Then
                Do While Not IsItPossibleToInsertDirectlyAVerticalPair()
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
                Loop
                InsertDirectlyVerticalPair() : Exit Sub
            ElseIf IsThereBelowAnyLeftLowerCorner() Then
                Do While Not IsItPossibleToInsertDirectlyAVerticalPair()
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise()
                Loop
                InsertDirectlyVerticalPair() : Exit Sub
            End If
        End If

        If IsItPossibleToInsertDirectlyAnUpperRightCorner() Then
            If IsThereBelowAnyRightEdge() Then
                Do While Not IsItPossibleToInsertDirectlyAVerticalPair()
                    CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
                Loop
                InsertDirectlyVerticalPair() : Exit Sub
            ElseIf IsThereBelowAnyRightLowerCorner() Then
                Do While Not IsItPossibleToInsertDirectlyAVerticalPair()
                    CubeThatWeMustSolve.Rotate90BackFaceClockwise()
                Loop
                InsertDirectlyVerticalPair() : Exit Sub
            End If
        End If

        If IsThereBelowAnyLowerEdgeCornerPair() Then
            Do While Not IsItPossibleToInsertDirectlyAVerticalPair()
                CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
            Loop
            InsertDirectlyVerticalPair() : Exit Sub
        Else
            CreateLowerEdgeCornerPair()
            InsertINDIRECTLYVerticalPair() : Exit Sub
        End If
        WeTerminateWithError(32) : Stop
    End Sub


    Private Sub CreateLeftLowerEdgeCornerPair()
        If Not IsItPossibleBelowAnyLeftLowerEdgeCornerPair() Then WeTerminateWithError(28) : Stop
        Do While Not IsThereBelowAnyLeftLowerEdgeCornerPair()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
    End Sub

    Private Sub CreateRightLowerEdgeCornerPair()
        If Not IsItPossibleAnyRightLowerEdgeCornerPair() Then WeTerminateWithError(28) : Stop
        Do While Not IsThereBelowAnyRightLowerEdgeCornerPair()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
    End Sub

    Private Sub CreateLowerEdgeCornerPair()
        If Not IsThereBelowAPossibilityOfAnyLowerEdgeCornerPair() Then WeTerminateWithError(28) : Stop
        Do While Not IsThereBelowAnyLowerEdgeCornerPair()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
    End Sub


    Private Sub InsertINDIRECTLYTrioColumnThroughTheLeft()
        If Not IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheLeft() Then WeTerminateWithError(31) : Stop
        If IsItPossibleToInsertDirectlyAColumnTrio() Then InsertDirectlyColumnTrio() : Exit Sub
        CreateLeftLowerEdgeCornerPair()
        Do While Not IsItPossibleToInsertDirectlyAColumnTrio()
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
        Loop
        InsertDirectlyColumnTrio()
    End Sub

    Private Sub InsertINDIRECTLYTrioColumnThroughTheRight()
        If Not IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheRight() Then WeTerminateWithError(31) : Stop
        If IsItPossibleToInsertDirectlyAColumnTrio() Then InsertDirectlyColumnTrio() : Exit Sub
        CreateRightLowerEdgeCornerPair()
        Do While Not IsItPossibleToInsertDirectlyAColumnTrio()
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
        Loop
        InsertDirectlyColumnTrio()
    End Sub

    Private Sub InsertINDIRECTLYColumnTrio()
        If Not IsItPossibleToInsertINDIRECTLYAColumnTrio() Then WeTerminateWithError(31) : Stop
        If IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheLeft() Then
            InsertINDIRECTLYTrioColumnThroughTheLeft()
        ElseIf IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheRight() Then
            InsertINDIRECTLYTrioColumnThroughTheRight()
        Else
            WeTerminateWithError(31) : Stop
        End If
    End Sub


    Private Sub InsertDirectlyColumnTrio()
        If Not IsItPossibleToInsertDirectlyAColumnTrio() Then WeTerminateWithError(23) : Stop
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0, 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2, 5, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(3, 5) : Exit Sub
            End If
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0, 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2, 5, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(4, 5) : Exit Sub
            End If
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2, 5, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0, 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(3, 5) : Exit Sub
            End If
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2, 5, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0, 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(4, 5) : Exit Sub
            End If
        End If
        WeTerminateWithError(23) : Stop
    End Sub

    Private Sub InsertDirectlyVerticalPair()
        If Not IsItPossibleToInsertDirectlyAVerticalPair() Then WeTerminateWithError(24) : Stop
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0, 3) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0, 6) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2, 5) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 5, 8) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0, 3) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0, 6) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2, 5) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 5, 8) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(4, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2, 5) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2, 8) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 5, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0, 3) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0, 6) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2, 5) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2, 8) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 5, 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0, 3) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0, 6) Or
                    AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 3, 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(4, 5) : Exit Sub
            End If
        End If

        WeTerminateWithError(24) : Stop
    End Sub

    Private Sub InsertDirectlyUpperCorner()
        If Not IsItPossibleToInsertDirectlyAnUpperCorner() Then WeTerminateWithError(25) : Stop
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(4, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(4, 5) : Exit Sub
            End If
        End If

        WeTerminateWithError(25) : Stop
    End Sub

    Private Sub InsertDirectlyLateralEdge()
        If Not IsItPossibleToInsertDirectlyALateralEdge() Then WeTerminateWithError(26) : Stop

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 3) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 5) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 3) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 5) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(4, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 5) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 3) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 5) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 3) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(4, 5) : Exit Sub
            End If
        End If

        WeTerminateWithError(26) : Stop
    End Sub

    Private Sub InsertDirectlyLowerCorner()
        If Not IsItPossibleToInsertDirectlyALowerCorner() Then WeTerminateWithError(27) : Stop

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(4, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(1, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(3, 5) : Exit Sub
            End If
        End If

        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 8) Then
                CubeThatWeMustSolve.Rotate90RightFaceClockwise(2, 5) : Exit Sub
            End If
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 6) Then
                CubeThatWeMustSolve.Rotate90LeftFaceCounterClockwise(4, 5) : Exit Sub
            End If
        End If

        WeTerminateWithError(27) : Stop
    End Sub


    Private Sub InsertINDIRECTLYALateralEdge()
        If Not IsItPossibleToInsertINDIRECTLYALateralEdge() Then WeTerminateWithError(33) : Stop
        Do While Not IsItPossibleToInsertDirectlyALateralEdge()
            CubeThatWeMustSolve.Rotate90FrontFaceCounterClockwise()
        Loop
        InsertDirectlyLateralEdge()
    End Sub

    Private Sub InsertINDIRECTLYLowerCorner()
        If Not IsItPossibleToInsertINDIRECTLYALowerCorner() Then WeTerminateWithError(34) : Stop
        Do While Not IsItPossibleToInsertDirectlyALowerCorner()
            CubeThatWeMustSolve.Rotate90BackFaceClockwise()
        Loop
        InsertDirectlyLowerCorner()
    End Sub


    Private Function IsItPossibleToInsertDirectlyAColumnTrio() As Boolean
        If Not (IsThereAboveAnyFreeLine() And IsThereBelowAnyColumnTrio()) Then Return False
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0, 3, 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2, 5, 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0, 3, 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2, 5, 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0, 3, 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2, 5, 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0, 3, 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2, 5, 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheLeft() As Boolean
        Return IsItPossibleToInsertDirectlyALeftUpperCorner() And IsThereBelowAPossibilityOfAnyLeftColumnTrio()
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheRight() As Boolean
        Return IsItPossibleToInsertDirectlyAnUpperRightCorner() And IsThereBelowAPossibilityOfAnyRightColumnTrio()
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYAColumnTrio() As Boolean
        Return IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheLeft() Or IsItPossibleToInsertINDIRECTLYATrioColumnThroughTheRight()
    End Function


    Private Function IsItPossibleToInsertDirectlyAVerticalPair() As Boolean
        If Not (IsThereAboveAnyFreeLine() And (IsThereBelowAnyLowerEdgeCornerPair() Or IsThereBelowAnyUpperEdgeCornerPair() Or IsThereBelowAnyPairOfCorners())) Then Return False
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0, 3) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0, 6) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 3, 6) Then Return True

            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2, 5) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 5, 8) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2, 8) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2, 5) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 5, 8) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2, 8) Then Return True

            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0, 3) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 3, 6) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0, 6) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2, 5) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 5, 8) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2, 8) Then Return True

            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0, 3) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 3, 6) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0, 6) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0, 3) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0, 6) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 3, 6) Then Return True

            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2, 5) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 5, 8) Then Return True
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2, 8) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYAVerticalPair() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If IsThereBelowAPossibilityOfAnyLowerEdgeCornerPair() Then Return True
        If IsItPossibleToInsertDirectlyALeftUpperCorner() Then
            If IsThereBelowAnyLeftLowerCorner() Or IsThereBelowAnyLeftEdge() Then Return True
        End If
        If IsItPossibleToInsertDirectlyAnUpperRightCorner() Then
            If IsThereBelowAnyRightLowerCorner() Or IsThereBelowAnyRightEdge() Then Return True
        End If
        Return False
    End Function


    Private Function IsItPossibleToInsertDirectlyALeftUpperCorner() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 0) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 0) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 0) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 0) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertDirectlyAnUpperRightCorner() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 2) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 2) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 2) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 2) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertDirectlyAnUpperCorner() As Boolean
        Return IsItPossibleToInsertDirectlyALeftUpperCorner() Or IsItPossibleToInsertDirectlyAnUpperRightCorner()
    End Function


    Private Function IsItPossibleToInsertDirectlyALeftLowerCornerPiece() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 6) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertDirectlyARightLowerCornerPiece() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then Return True
        End If
        If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 8) Then
            If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertDirectlyALowerCorner() As Boolean
        Return IsItPossibleToInsertDirectlyALeftLowerCornerPiece() Or IsItPossibleToInsertDirectlyARightLowerCornerPiece()
    End Function


    Private Function IsItPossibleToInsertINDIRECTLYALeftLowerCornerPiece() As Boolean
        Return IsThereBelowAnyLeftLowerCorner() And IsThereAboveAnyFreeLine()
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYARightLowerCornerPiece() As Boolean
        Return IsThereBelowAnyRightLowerCorner() And IsThereAboveAnyFreeLine()
    End Function

    Private Function IsItPossibleToInsertINDIRECTLYALowerCorner() As Boolean
        Return IsItPossibleToInsertINDIRECTLYALeftLowerCornerPiece() Or IsItPossibleToInsertINDIRECTLYARightLowerCornerPiece()
    End Function


    Private Function IsItPossibleToInsertDirectlyALeftEdgePiece() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If Not IsThereBelowAnyLeftEdge() Then Return False
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 3) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 3) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 3) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 3) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertDirectlyARightEdgePiece() As Boolean
        If Not IsThereAboveAnyFreeLine() Then Return False
        If Not IsThereBelowAnyRightEdge() Then Return False
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 3, 6) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(3), 5) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 2, 5, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(1), 5) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 0, 1, 2) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(4), 5) Then Return True
        End If
        If NoOneOfTheSquaresIsTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(0), 6, 7, 8) Then
            If AreAllTheSquaresOfTheSoughtColour(YellowColour, CubeThatWeMustSolve.ArrayOfRubiksCube(2), 5) Then Return True
        End If
        Return False
    End Function

    Private Function IsItPossibleToInsertDirectlyALateralEdge() As Boolean
        Return IsItPossibleToInsertDirectlyALeftEdgePiece() Or IsItPossibleToInsertDirectlyARightEdgePiece()
    End Function


    'Private Function IsItPossibleToInsertINDIRECTLYALeftEdgePiece() As Boolean
    '    Return IsThereAboveAnyFreeLine() And IsThereBelowAnyLeftEdge()
    'End Function

    'Private Function IsItPossibleToInsertINDIRECTLYARightEdgePiece() As Boolean
    '    Return IsThereAboveAnyFreeLine() And IsThereBelowAnyRightEdge()
    'End Function

    Private Function IsItPossibleToInsertINDIRECTLYALateralEdge() As Boolean
        Return IsThereAboveAnyFreeLine() And IsThereBelowAnyLateralEdgePiece()
    End Function


    Private Function AreAllTheSquaresOfTheSoughtColour(SoughtColour As Integer, Configuracion As Integer, ParamArray ListOfSquares() As Integer) As Boolean
        Dim FaceNumber As Integer = SquareColour(4, Configuracion)
        Dim Counter As Integer
        For Counter = 0 To ListOfSquares.GetUpperBound(0)
            If SquareColour(ListOfSquares(Counter), Configuracion) <> SoughtColour Then Return False
        Next
        Return True
    End Function

    Private Function NoOneOfTheSquaresIsTheSoughtColour(SoughtColour As Integer, Configuracion As Integer, ParamArray ListOfSquares() As Integer) As Boolean
        Dim FaceNumber As Integer = SquareColour(4, Configuracion)
        Dim Counter As Integer
        For Counter = 0 To ListOfSquares.GetUpperBound(0)
            If SquareColour(ListOfSquares(Counter), Configuracion) = SoughtColour Then Return False
        Next
        Return True
    End Function


    Private Function IsThereAboveAnyFreeLine() As Boolean
        Dim TrioOfPositionsAuxiliar() As Integer
        TrioOfPositionsAuxiliar = {0, 1, 2}
        If IsTheLineFree(TrioOfPositionsAuxiliar) Then Return True
        TrioOfPositionsAuxiliar = {6, 7, 8}
        If IsTheLineFree(TrioOfPositionsAuxiliar) Then Return True
        TrioOfPositionsAuxiliar = {0, 3, 6}
        If IsTheLineFree(TrioOfPositionsAuxiliar) Then Return True
        TrioOfPositionsAuxiliar = {2, 5, 8}
        If IsTheLineFree(TrioOfPositionsAuxiliar) Then Return True
        Return False
    End Function

    Private Function IsTheLineFree(TrioAuxiliar() As Integer) As Boolean
        Dim Counter As Integer
        For Counter = 0 To 2
            If SquareColour(TrioAuxiliar(Counter), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) = YellowColour Then Return False
        Next
        Return True
    End Function


    Private Function IsThereAboveAnyFreeCentreLeftPair() As Boolean
        ' Lo de "Centro - Izquierda" se refiere a la línea de abajo

        Dim PossiblePairs() As Integer
        PossiblePairs = {1, 2}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        PossiblePairs = {5, 8}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        PossiblePairs = {6, 7}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        PossiblePairs = {0, 3}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        Return False
    End Function

    Private Function IsThereAboveAnyFreeCentreRightPair() As Boolean
        ' Lo de "Centro - Derecha" se refiere a la línea de abajo

        Dim PossiblePairs() As Integer
        PossiblePairs = {7, 8}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        PossiblePairs = {3, 6}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        PossiblePairs = {0, 1}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        PossiblePairs = {2, 5}
        If SquareColour(PossiblePairs(0), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour AndAlso
                SquareColour(PossiblePairs(1), CubeThatWeMustSolve.ArrayOfRubiksCube(YellowColour)) <> YellowColour Then
            Return True
        End If
        Return False
    End Function

    'Private Function IsThereAboveAnyFreePair() As Boolean
    '    Return (IsThereAboveAnyFreeCentreLeftPair() Or IsThereAboveAnyFreeCentreRightPair())
    'End Function


    Private Function IsThereBelowAnyRightLowerCorner() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(8, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyLeftLowerCorner() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(6, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    'Private Function IsThereBelowAnyLowerCorner() As Boolean
    '    Return (IsThereBelowAnyRightLowerCorner() Or IsThereBelowAnyLeftLowerCorner())
    'End Function


    Private Function IsThereBelowAnyRightUpperCorner() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(2, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyLeftUpperCorner() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(0, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    'Private Function IsThereBelowAnyUpperCorner() As Boolean
    '    Return (IsThereBelowAnyRightUpperCorner() Or IsThereBelowAnyLeftUpperCorner())
    'End Function


    Private Function IsThereBelowAnyRightEdge() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyLeftEdge() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(3, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyLateralEdgePiece() As Boolean
        Return (IsThereBelowAnyRightEdge() Or IsThereBelowAnyLeftEdge())
    End Function


    Private Function IsThereBelowAnyLowerEdge() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(7, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyUpperEdge() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(1, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function


    Private Function IsThereBelowAnyLeftLowerEdgeCornerPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(3, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And SquareColour(6, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyRightLowerEdgeCornerPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And SquareColour(8, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyLowerEdgeCornerPair() As Boolean
        Return (IsThereBelowAnyLeftLowerEdgeCornerPair() Or IsThereBelowAnyRightLowerEdgeCornerPair())
    End Function


    Private Function IsItPossibleBelowAnyLeftLowerEdgeCornerPair() As Boolean
        Return (IsThereBelowAnyLeftLowerCorner() And IsThereBelowAnyLeftEdge())
    End Function

    Private Function IsItPossibleAnyRightLowerEdgeCornerPair() As Boolean
        Return (IsThereBelowAnyRightLowerCorner() And IsThereBelowAnyRightEdge())
    End Function

    Private Function IsThereBelowAPossibilityOfAnyLowerEdgeCornerPair() As Boolean
        Return (IsItPossibleBelowAnyLeftLowerEdgeCornerPair() Or IsItPossibleAnyRightLowerEdgeCornerPair())
    End Function


    Private Function IsThereBelowAnyLeftUpperEdgeCornerPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(0, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And SquareColour(3, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyRightUpperEdgeCornerPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(2, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And SquareColour(5, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyUpperEdgeCornerPair() As Boolean
        Return (IsThereBelowAnyLeftUpperEdgeCornerPair() Or IsThereBelowAnyRightUpperEdgeCornerPair())
    End Function


    'Private Function IsThereAPossibilityOfAnyLeftUpperEdgeCornerPair() As Boolean
    '    Return (IsThereBelowAnyLeftEdge() And IsThereBelowAnyLeftUpperCorner())
    'End Function

    'Private Function IsThereBelowAPossibilityOfAnyRightUpperEdgeCornerPair() As Boolean
    '    Return (IsThereBelowAnyRightEdge() And IsThereBelowAnyRightUpperCorner())
    'End Function

    '    Private Function IsThereBelowAPossibilityOfAnyUpperEdgeCornerPair() As Boolean
    '    Return (IsThereAPossibilityOfAnyLeftUpperEdgeCornerPair() Or IsThereBelowAPossibilityOfAnyRightUpperEdgeCornerPair())
    '    End Function


    Private Function IsThereBelowAnyCentreLeftPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(6, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And
                    SquareColour(7, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyCentreRightPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(7, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And
                    SquareColour(8, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then Return True
        Next
        Return False
    End Function

    'Private Function IsThereBelowAnyHorizontalPair() As Boolean
    'Return (IsThereBelowAnyCentreLeftPair() Or IsThereBelowAnyCentreRightPair())
    'End Function


    Private Function IsThereBelowAnyLeftColumnTrio() As Boolean
        Dim SquareCounter, Counter As Integer
        Dim Result As Boolean
        For Counter = 1 To 4
            Result = True
            For SquareCounter = 0 To 6 Step 3
                If SquareColour(SquareCounter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) <> YellowColour Then
                    Result = False
                    Exit For
                End If
            Next
            If Result = True Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyRightColumnTrio() As Boolean
        Dim SquareCounter, Counter As Integer
        Dim Result As Boolean
        For Counter = 1 To 4
            Result = True
            For SquareCounter = 2 To 8 Step 3
                If SquareColour(SquareCounter, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) <> YellowColour Then
                    Result = False
                    Exit For
                End If
            Next
            If Result = True Then Return True
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyColumnTrio() As Boolean
        Return (IsThereBelowAnyRightColumnTrio() Or IsThereBelowAnyLeftColumnTrio())
    End Function


    Private Function IsThereBelowAPossibilityOfAnyLeftColumnTrio() As Boolean
        Return (IsItPossibleBelowAnyLeftLowerEdgeCornerPair() And IsThereBelowAnyLeftUpperCorner())
    End Function

    Private Function IsThereBelowAPossibilityOfAnyRightColumnTrio() As Boolean
        Return (IsItPossibleAnyRightLowerEdgeCornerPair() And IsThereBelowAnyRightUpperCorner())
    End Function

    'Private Function IsThereBelowAPossibilityOfAnyColumnTrio() As Boolean
    '    Return (IsThereBelowAPossibilityOfAnyLeftColumnTrio() Or IsThereBelowAPossibilityOfAnyRightColumnTrio())
    'End Function


    Private Function IsThereBelowAnyLeftCornerPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(0, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And
                    SquareColour(6, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyRightCornerPair() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 4
            If SquareColour(2, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour And
                    SquareColour(8, CubeThatWeMustSolve.ArrayOfRubiksCube(Counter)) = YellowColour Then
                Return True
            End If
        Next
        Return False
    End Function

    Private Function IsThereBelowAnyPairOfCorners() As Boolean
        Return (IsThereBelowAnyLeftCornerPair() Or IsThereBelowAnyRightCornerPair())
    End Function


    'Private Function IsThereBelowAPossibilityOfAnyLeftCornerPair() As Boolean
    '    Return (IsThereBelowAnyLeftLowerCorner() And IsThereBelowAnyLeftUpperCorner())
    'End Function

    'Private Function IsThereBelowAPossibilityOfAnyRightCornerPair() As Boolean
    '    Return (IsThereBelowAnyRightLowerCorner() And IsThereBelowAnyRightUpperCorner())
    'End Function

    'Private Function HayAbajoPosibilidadDeAlgunParDeEsquinas() As Boolean
    '    Return (IsThereBelowAPossibilityOfAnyLeftCornerPair() Or IsThereBelowAPossibilityOfAnyRightCornerPair())
    'End Function


    Private Function IsThereUndergroundAnyCorner() As Boolean
        Dim Counter As Integer
        For Counter = 0 To 8 Step 2
            If Counter = 4 Then Continue For
            If SquareColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(BackFace(YellowColour))) = YellowColour Then Return True
        Next
        Return False
    End Function

    Private Function IsThereUndergroundAnyEdge() As Boolean
        Dim Counter As Integer
        For Counter = 1 To 7 Step 2
            If SquareColour(Counter, CubeThatWeMustSolve.ArrayOfRubiksCube(BackFace(YellowColour))) = YellowColour Then Return True
        Next
        Return False
    End Function


    Public Function IsTheFaceSolved(FaceNumber As Integer) As Boolean
        Select Case FaceNumber
            Case 0 To 5 : Return AreAllTheSquaresOfTheSoughtColour(FaceNumber, CubeThatWeMustSolve.ArrayOfRubiksCube(FaceNumber), {0, 1, 2, 3, 4, 5, 6, 7, 8})
            Case Else
                WeTerminateWithError(22) : Stop : End
        End Select
    End Function



End Class
