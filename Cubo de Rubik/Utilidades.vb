Option Explicit On
Option Strict On


Public Class Utilidades
    'Shared Function ImpresionDeMatriz(Matriz() As Integer, Optional Delimitador As String = " ") As String
    '    Dim Result(Matriz.Length - 1) As String
    '    Dim Counter As Integer
    '    For Counter = 0 To Matriz.GetUpperBound(0)
    '        Result(Counter) = CType(Matriz(Counter), String)
    '    Next
    '    Return Join(Result, Delimitador)
    'End Function

    'Shared Function SonIgualesLasMatrices(MatrizUno() As Integer, MatrizDos() As Integer) As Boolean
    '    If MatrizUno.GetLength(0) <> MatrizDos.GetLength(0) Then Return False
    '    Dim Counter As Integer
    '    For Counter = 0 To MatrizUno.GetUpperBound(0)
    '        If MatrizUno(Counter) <> MatrizDos(Counter) Then Return False
    '    Next
    '    Return True
    'End Function

    Shared Function SquareColour(SquareNumber%, FaceColourConfiguration%) As Integer
        If SquareNumber < 0 Or SquareNumber > 8 Then WeTerminateWithError(1) : Stop
        If FaceColourConfiguration < 0 Or FaceColourConfiguration >= (6 ^ 9) Then WeTerminateWithError(2) : Stop
        Return (FaceColourConfiguration Mod CType((6 ^ (SquareNumber + 1)), Integer)) \ CType(6 ^ SquareNumber, Integer)
    End Function

    Shared Function BackFace(FaceNumber As Integer) As Integer
        Select Case FaceNumber
            Case 0 : Return 5
            Case 1 : Return 3
            Case 2 : Return 4
            Case 3 : Return 1
            Case 4 : Return 2
            Case 5 : Return 0
            Case Else : WeTerminateWithError(5) : Stop
        End Select
    End Function

    Shared Function Maximum(One As Integer, Two As Integer) As Integer
        If One > Two Then Return One Else Return Two
    End Function

    'Shared Sub ProcesarParesDeColores(CasillaUno%, FaceOne%, CasillaDos%, FaceTwo%, ArrayOfRubiksCube() As Integer, ByRef MatrizDeColoresDePrueba() As Integer)
    '    If (CasillaUno < 0 Or CasillaDos < 0 Or CasillaUno > 8 Or CasillaDos > 8) Then WeTerminateWithError(1) : Stop
    '    If FaceOne < 0 Or FaceTwo < 0 Or FaceOne > 5 Or FaceTwo > 5 Then WeTerminateWithError(5) : Stop
    '    If ArrayOfRubiksCube.GetLength(0) <> 6 Then WeTerminateWithError(3) : Stop
    '    MatrizDeColoresDePrueba(0) = SquareColour(CasillaUno, ArrayOfRubiksCube(FaceOne))
    '    MatrizDeColoresDePrueba(1) = SquareColour(CasillaDos, ArrayOfRubiksCube(FaceTwo))
    '    If MatrizDeColoresDePrueba(0) < MatrizDeColoresDePrueba(1) Then
    '        MatrizDeColoresDePrueba(0) += MatrizDeColoresDePrueba(1)
    '        MatrizDeColoresDePrueba(1) = MatrizDeColoresDePrueba(0) - MatrizDeColoresDePrueba(1)
    '        MatrizDeColoresDePrueba(0) -= MatrizDeColoresDePrueba(1)
    '    End If
    'End Sub

    'Function DeNumeroDeTrioAMatrizDeTresColores(NumeroDeTrio As Integer) As Integer()
    '    Select Case NumeroDeTrio
    '        Case 0 : Return {0, 1, 2}
    '        Case 1 : Return {0, 2, 3}
    '        Case 2 : Return {0, 3, 4}
    '        Case 3 : Return {0, 4, 1}
    '        Case 4 : Return {1, 2, 6}
    '        Case 5 : Return {2, 3, 6}
    '        Case 6 : Return {3, 4, 6}
    '        Case 7 : Return {4, 1, 6}
    '        Case Else : WeTerminateWithError(8) : Stop
    '    End Select
    'End Function

    Shared Function FromThreeColourArrayToTrioNumber(ArrayOfThreeColours() As Integer) As Integer
        If ArrayOfThreeColours.GetLength(0) <> 3 Then WeTerminateWithError(9) : Stop
        If ArrayOfThreeColours(0) > ArrayOfThreeColours(1) Then Return FromThreeColourArrayToTrioNumber({ArrayOfThreeColours(1), ArrayOfThreeColours(0), ArrayOfThreeColours(2)})
        If ArrayOfThreeColours(1) > ArrayOfThreeColours(2) Then Return FromThreeColourArrayToTrioNumber({ArrayOfThreeColours(0), ArrayOfThreeColours(2), ArrayOfThreeColours(1)})
        Dim Counter As Integer
        For Counter = 0 To 2
            If ArrayOfThreeColours(Counter) < 0 Or ArrayOfThreeColours(Counter) > 5 Then WeTerminateWithError(10) : Stop
        Next
        If ArrayOfThreeColours(0) = ArrayOfThreeColours(1) Or ArrayOfThreeColours(1) = ArrayOfThreeColours(2) Then WeTerminateWithError(11) : Stop
        If ArrayOfThreeColours(0) = 0 And ArrayOfThreeColours(1) = 1 And ArrayOfThreeColours(2) = 2 Then Return 0
        If ArrayOfThreeColours(0) = 0 And ArrayOfThreeColours(1) = 2 And ArrayOfThreeColours(2) = 3 Then Return 1
        If ArrayOfThreeColours(0) = 0 And ArrayOfThreeColours(1) = 3 And ArrayOfThreeColours(2) = 4 Then Return 2
        If ArrayOfThreeColours(0) = 0 And ArrayOfThreeColours(1) = 1 And ArrayOfThreeColours(2) = 4 Then Return 3
        If ArrayOfThreeColours(0) = 1 And ArrayOfThreeColours(1) = 2 And ArrayOfThreeColours(2) = 5 Then Return 4
        If ArrayOfThreeColours(0) = 2 And ArrayOfThreeColours(1) = 3 And ArrayOfThreeColours(2) = 5 Then Return 5
        If ArrayOfThreeColours(0) = 3 And ArrayOfThreeColours(1) = 4 And ArrayOfThreeColours(2) = 5 Then Return 6
        If ArrayOfThreeColours(0) = 1 And ArrayOfThreeColours(1) = 4 And ArrayOfThreeColours(2) = 5 Then Return 7
        WeTerminateWithError(12) : Stop
    End Function

    Shared Function ColourConfigurationOfTheFaceOnceRotated90Clockwise(ColourConfigurationOfTheFace As Integer) As Integer
        Dim Result As Integer = 0
        Result += CType(6 ^ 0, Integer) * SquareColour(6, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 1, Integer) * SquareColour(3, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 2, Integer) * SquareColour(0, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 5, Integer) * SquareColour(1, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 8, Integer) * SquareColour(2, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 7, Integer) * SquareColour(5, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 6, Integer) * SquareColour(8, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 3, Integer) * SquareColour(7, ColourConfigurationOfTheFace)
        Result += CType(6 ^ 4, Integer) * SquareColour(4, ColourConfigurationOfTheFace)
        Return Result
    End Function

    Shared Function ColourConfigurationOfTheFaceOnceRotated90CounterClockwise(ColourConfigurationOfTheFace As Integer) As Integer
        Return ColourConfigurationOfTheFaceOnceRotated90Clockwise(ColourConfigurationOfTheFaceOnceRotated90Clockwise(ColourConfigurationOfTheFaceOnceRotated90Clockwise(ColourConfigurationOfTheFace)))
    End Function

    'Shared Sub ComprobarSiElCuboEsCorrecto(CubitoDeRubikAuxiliar As ClaseCuboDeRubik)
    '    Dim ArrayOfRubiksCube(5) As Integer
    '    ArrayOfRubiksCube = CloningOfArray(CubitoDeRubikAuxiliar.ArrayOfRubiksCube)
    '    If ArrayOfRubiksCube.GetLength(0) <> 6 Then WeTerminateWithError(3) : Stop : End
    '    Dim Counter, CuentaUno, CuentaDos As Integer

    '    Dim NumeroDeAparicionesDelColor(5) As Integer
    '    Dim FaceCounter, SquareCounter As Integer
    '    For FaceCounter = 0 To 5
    '        For SquareCounter = 0 To 8
    '            NumeroDeAparicionesDelColor(SquareColour(SquareCounter, ArrayOfRubiksCube(FaceCounter))) += 1
    '        Next
    '    Next
    '    For Counter = 0 To 5
    '        If NumeroDeAparicionesDelColor(Counter) <> 9 Then WeTerminateWithError(4) : Stop : End
    '    Next

    '    Dim NumeroDeAparicionesDelParDeColores(5, 5) As Integer
    '    Dim MatrizDeDosColoresDePrueba(1) As Integer
    '    ProcesarParesDeColores(7, 0, 1, 1, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(3, 0, 1, 2, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(1, 0, 1, 3, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(5, 0, 1, 4, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(3, 1, 5, 2, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(3, 2, 5, 3, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(3, 3, 5, 4, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(3, 4, 5, 1, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(1, 5, 7, 1, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(3, 5, 7, 2, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(7, 5, 7, 3, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    ProcesarParesDeColores(5, 5, 7, 4, ArrayOfRubiksCube, MatrizDeDosColoresDePrueba)
    '    NumeroDeAparicionesDelParDeColores(MatrizDeDosColoresDePrueba(0), MatrizDeDosColoresDePrueba(1)) += 1
    '    For CuentaUno = 1 To 5
    '        For CuentaDos = 0 To CuentaUno - 1
    '            If (CuentaUno = 5 And CuentaDos = 0) Or (CuentaUno = 3 And CuentaDos = 1) Or (CuentaUno = 4 And CuentaDos = 2) Then Continue For
    '            If NumeroDeAparicionesDelParDeColores(CuentaUno, CuentaDos) <> 1 Then WeTerminateWithError(6) : Stop
    '        Next
    '    Next
    '    For Counter = 0 To 5
    '        If NumeroDeAparicionesDelParDeColores(Counter, Counter) <> 0 Then WeTerminateWithError(7) : Stop
    '    Next

    '    Dim NumeroDeAparicionesDelTrioDeColores(7) As Integer
    '    Dim ArrayOfThreeColours(2) As Integer
    '    ArrayOfThreeColours = {SquareColour(6, ArrayOfRubiksCube(0)), SquareColour(0, ArrayOfRubiksCube(1)), SquareColour(2, ArrayOfRubiksCube(2))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(0, ArrayOfRubiksCube(0)), SquareColour(0, ArrayOfRubiksCube(2)), SquareColour(2, ArrayOfRubiksCube(3))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(2, ArrayOfRubiksCube(0)), SquareColour(0, ArrayOfRubiksCube(3)), SquareColour(2, ArrayOfRubiksCube(4))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(8, ArrayOfRubiksCube(0)), SquareColour(0, ArrayOfRubiksCube(4)), SquareColour(2, ArrayOfRubiksCube(1))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(0, ArrayOfRubiksCube(5)), SquareColour(6, ArrayOfRubiksCube(1)), SquareColour(8, ArrayOfRubiksCube(2))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(6, ArrayOfRubiksCube(5)), SquareColour(6, ArrayOfRubiksCube(2)), SquareColour(8, ArrayOfRubiksCube(3))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(8, ArrayOfRubiksCube(5)), SquareColour(6, ArrayOfRubiksCube(3)), SquareColour(8, ArrayOfRubiksCube(4))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    ArrayOfThreeColours = {SquareColour(2, ArrayOfRubiksCube(5)), SquareColour(6, ArrayOfRubiksCube(4)), SquareColour(8, ArrayOfRubiksCube(1))}
    '    NumeroDeAparicionesDelTrioDeColores(FromThreeColourArrayToTrioNumber(ArrayOfThreeColours)) += 1
    '    For Counter = 0 To 7
    '        If NumeroDeAparicionesDelTrioDeColores(Counter) <> 1 Then WeTerminateWithError(13) : Stop
    '    Next

    '    MessageBox.Show("Hemos terminado la comprobación", "Comprobación terminada")
    'End Sub

    Shared Function OppositeMovement(MovementNumber As Integer) As Integer
        If MovementNumber Mod 2 = 1 Then Return MovementNumber - 1 Else Return MovementNumber + 1
    End Function

    'Shared Function UltimoMovimientoSegunCadena(CadenaDeMovimientos As String) As Integer
    '    If InStr(CadenaDeMovimientos, ",") = 0 Then Return CType(CadenaDeMovimientos, Integer)
    '    Return CType(CadenaDeMovimientos.Substring(CadenaDeMovimientos.IndexOf(",") + 1), Integer)
    '    WeTerminateWithError(20) : Stop
    'End Function

    Shared Function FromNumberOfMovementToString(MovementNumber As Integer) As String
        Select Case MovementNumber
            Case 0 : Return "Up Face: Clockwise"
            Case 1 : Return "Up Face: Counter-Clockwise"
            Case 2 : Return "Down Face: Counter-Clockwise"
            Case 3 : Return "Down Face: Clockwise"
            Case 4 : Return "Left Face: Counter-Clockwise"
            Case 5 : Return "Left Face: Clockwise"
            Case 6 : Return "Right Face: Clockwise"
            Case 7 : Return "Right Face: Counter-Clockwise"
            Case 8 : Return "Front Face: Counter-Clockwise"
            Case 9 : Return "Front Face: Clockwise"
            Case 10 : Return "Back Face: Clockwise"
            Case 11 : Return "Back Face: Counter-Clockwise"
            Case Else : WeTerminateWithError(19) : Stop
        End Select
    End Function

    'Shared Function FromNumberOfMovementToString(Matriz() As Integer) As String
    '    Dim Result As String = vbNullString
    '    Dim Counter As Integer
    '    For Counter = 0 To Matriz.GetUpperBound(0)
    '        Result &= (Counter + 1) & ") " & FromNumberOfMovementToString(Matriz(Counter)) & vbCrLf
    '        If (Counter + 1) Mod 5 = 0 Then Result &= vbCrLf
    '    Next
    '    Return Result
    'End Function

    Shared Function FromNumberOfMovementToString(RubiksCubeArgument As ClaseCuboDeRubik, Optional WithDividingLines As Boolean = False) As String
        Dim AuxiliarCube As ClaseCuboDeRubik
        If WithDividingLines Then
            AuxiliarCube = New ClaseCuboDeRubik(CloningOfArray(RubiksCubeArgument.InitialArray))
        End If
        Dim Result As String = vbNullString
        Dim Counter As Integer
        For Counter = 0 To RubiksCubeArgument.ListOfMovements.GetUpperBound(0)
            Result &= (Counter + 1) & ") " & FromNumberOfMovementToString(RubiksCubeArgument.ListOfMovements(Counter)) & vbCrLf
            If (Counter + 1) Mod 5 = 0 Then Result &= vbCrLf
            If WithDividingLines Then
                AuxiliarCube.ExecuteMovement(RubiksCubeArgument.ListOfMovements(Counter))
                If AuxiliarCube.IsTheFaceSolved(0) Then
                    Result &= "----------" & vbCrLf
                    If (Counter + 1) Mod 5 = 0 Then Result &= vbCrLf
                End If
            End If
        Next
        Return Result
    End Function

    Shared Sub WeTerminateWithError(NumeroDeError As Integer)
        Dim CaptionOfMessageBox, TextOfMessageBox As String
        Select Case NumeroDeError
            Case 1
                TextOfMessageBox = "The square number must be an integer between 0 and 9, both included"
                CaptionOfMessageBox = "Invalid square number"
            Case 2
                TextOfMessageBox = "The colour configuration number of the face must be an integer between 0 and 6^9-1, both included"
                CaptionOfMessageBox = "Invalid number of colour configuration"
            Case 3
                TextOfMessageBox = "The array of Rubik's Cube must have 6 and only 6 elements"
                CaptionOfMessageBox = "Invalid number of elements in an array"
            Case 4
                TextOfMessageBox = "Each colour must appear in 9 and only 9 squares in all the Rubik's Cube"
                CaptionOfMessageBox = "Invalid number of appearances of a colour"
            Case 5
                TextOfMessageBox = "The face number must be a number between 0 and 5, both included"
                CaptionOfMessageBox = "Invalid face number"
            Case 6
                TextOfMessageBox = "Each colour pair of the list must appear once and only once"
                CaptionOfMessageBox = "Invalid number of appearances of a pair of colours"
            Case 7
                TextOfMessageBox = "It is impossible that two visible squares on a piece have both the same colour"
                CaptionOfMessageBox = "Invalid coincidence of the same colour on two visible squares of a piece"
            Case 8
                TextOfMessageBox = "The trio number must be a number between 0 and 7, both included"
                CaptionOfMessageBox = "Invalid trio number"
            Case 9
                TextOfMessageBox = "You must give as an argument an array of three and only three colours"
                CaptionOfMessageBox = "Invalid number of colours"
            Case 10
                TextOfMessageBox = "The colour must be a number between 0 and 5, both included"
                CaptionOfMessageBox = "Invalid colour number"
            Case 11
                TextOfMessageBox = "On a corner piece, there can't appear the same colour more than once"
                CaptionOfMessageBox = "Unauthorized repetition of a colour on a corner piece"
            Case 12
                TextOfMessageBox = "The colour trio passed as an argument doesn't correspond to any of the 8 corner pieces of the cube"
                CaptionOfMessageBox = "Trío de colores inexistente"
            Case 13
                TextOfMessageBox = "Each of the 7 colour trios must appear once and only once"
                CaptionOfMessageBox = "Colour trio that doesn't appear or that appears more than once"
            Case 14
                TextOfMessageBox = "The corner number must be a number between 0 and 7, both included"
                CaptionOfMessageBox = "Invalid corner number"
            Case 15
                TextOfMessageBox = "The pair number must be a number between 0 and 11, both included"
                CaptionOfMessageBox = "Invalid pair number"
            Case 16
                TextOfMessageBox = "The face number must be a number between 0 and 5, both included"
                CaptionOfMessageBox = "Invalid face number"
            Case 17
                TextOfMessageBox = "When the queue is empty, the beginning and the end point at Nothing; and when the queue _
                    is not empty, noone of the elements points at Nothing. But it is not possible that one element points at Nothing and another one doesn't."
                CaptionOfMessageBox = "It is not possible that one points at Nothing and another one doesn't"
            Case 18
                TextOfMessageBox = "You can't remove anything from the queue, because the queue is empty"
                CaptionOfMessageBox = "You can't remove anything from empty queues"
            Case 19
                TextOfMessageBox = "You must pass a movement number from 0 to 11, both included"
                CaptionOfMessageBox = "Invalid movement number"
            Case 20
                TextOfMessageBox = "Revise that string, because you can't deduct from it which was the last movement"
                CaptionOfMessageBox = "Invalid string of movements"
            Case 21
                TextOfMessageBox = "You can't put a face over itself or other the opposite one"
                CaptionOfMessageBox = "Invalid pair of faces"
            Case 22
                TextOfMessageBox = "We shouldn't be here, because the situation in which we are should have been stopped at any of the previous conditional structures"
                CaptionOfMessageBox = "We shouldn't be here"
            Case 23
                TextOfMessageBox = "It is not possible to make a direct insertion of column trio"
                CaptionOfMessageBox = "Impossible movement"
            Case 24
                TextOfMessageBox = "It is not possible to make a direct insertion of vertical pair"
                CaptionOfMessageBox = "Impossible movement"
            Case 25
                TextOfMessageBox = "It is not possible to make a direct insertion of upper corner"
                CaptionOfMessageBox = "Impossible movement"
            Case 26
                TextOfMessageBox = "It is not possible to make a direct insertion of a lateral edge"
                CaptionOfMessageBox = "Impossible movement"
            Case 27
                TextOfMessageBox = "It is not possible to make a direct insertion of lower corner"
                CaptionOfMessageBox = "Impossible movement"
            Case 28
                TextOfMessageBox = "It is not possible to create the sought pair"
                CaptionOfMessageBox = "Impossible movement"
            Case 29
                TextOfMessageBox = "The face number must be a number between 1 and 4, both included"
                CaptionOfMessageBox = "Impossible movement"
            Case 30
                TextOfMessageBox = "It is not possible to creat the sought trio"
                CaptionOfMessageBox = "Impossible movement"
            Case 31
                TextOfMessageBox = "It is not possible to make an INDIRECT insertion of column trio"
                CaptionOfMessageBox = "Impossible movement"
            Case 32
                TextOfMessageBox = "It is not possible to make an INDIRECT insertion of vertical pair"
                CaptionOfMessageBox = "Impossible movement"
            Case 33
                TextOfMessageBox = "It is not possible to make an INDIRECT insertion of lateral edge"
                CaptionOfMessageBox = "Impossible movement"
            Case 34
                TextOfMessageBox = "It is not possible to make an INDIRECT insertion of lower corner"
                CaptionOfMessageBox = "Impossible movement"
            Case 35
                TextOfMessageBox = "It is not possible to make the direct embedding of lower corner that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 36
                TextOfMessageBox = "It is not possible to make the direct embedding of lower edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 36
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of lower edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 37
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of lower corner that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 38
                TextOfMessageBox = "It is not possible to make the direct embedding of lateral edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 39
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of lateral edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 40
                TextOfMessageBox = "It is not possible to make the direct embedding of horizontal pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 41
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of horizontal pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 42
                TextOfMessageBox = "It is not possible to make the direct embedding of underground edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 43
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of underground edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 44
                TextOfMessageBox = "It is not possible to make the direct embedding of underground corner that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 45
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of underground corner that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 46
                TextOfMessageBox = "It is not possible to make the INDIRECT insertion of vertical pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 47
                TextOfMessageBox = "It is not possible to make the direct embedding of lower edge-corner pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 48
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of lower edge-corner pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 49
                TextOfMessageBox = "This is strange, because technically the face is not still solved, but there isn't either any movement that can be made"
                CaptionOfMessageBox = "We are in a paradox"
            Case 50
                TextOfMessageBox = "It is not possible to anyhow place the upper corner that you want to place"
                CaptionOfMessageBox = "Impossible movement"
            Case 51
                TextOfMessageBox = "It is not possible to make the direct insertion of underground column trio that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 52
                TextOfMessageBox = "It is not possible to make the INDIRECT insertion of underground column trio that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 53
                TextOfMessageBox = "It is not possible to make the direct insertion of underground pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 54
                TextOfMessageBox = "It is not p ossible to make the INDIRECT insertion of underground pair that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 55
                TextOfMessageBox = "It is not possible to make the edge exchange that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 56
                TextOfMessageBox = "Something is going wrong here, because if the condition has been TRUE in the previous cases, it should also be TRUE here"
                CaptionOfMessageBox = "We shouldn't be here"
            Case 57
                TextOfMessageBox = "It is not possible to make the corner exchange that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 58
                TextOfMessageBox = "It is not possible to anyhow place any upper edge"
                CaptionOfMessageBox = "Impossible movement"
            Case 59
                TextOfMessageBox = "Before making anything of this, the yellow face should be solved, but it isn't"
                CaptionOfMessageBox = "We shouldn't be here"
            Case 60
                TextOfMessageBox = "It is not possible to make the direct embedding of a secod-row edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 61
                TextOfMessageBox = "It is not possible to make the INDIRECT embedding of second-row edge that you want to make"
                CaptionOfMessageBox = "Impossible movement"
            Case 62
                TextOfMessageBox = "It is not possible to lower a second-row inverted edge piece"
                CaptionOfMessageBox = "Impossible movement"
            Case 63
                TextOfMessageBox = "How can you lower a second-row edge piece, when the second row is already solved?"
                CaptionOfMessageBox = "Impossible movement"
            Case 64
                TextOfMessageBox = "In order to solve the underground white face, first the yellow face and the two upper rows on each of the four adjacent faces must be solved"
                CaptionOfMessageBox = "Impossible movement"
            Case 65
                TextOfMessageBox = "Those two possibilities are mutually exclusive"
                CaptionOfMessageBox = "Impossible situation"
            Case 66
                TextOfMessageBox = "The down face must be in this case a number between 1 and 4, both included"
                CaptionOfMessageBox = "Invalid value"
            Case 67
                TextOfMessageBox = "The data that you have input are incorrect, that's why we can't solve the cube"
                CaptionOfMessageBox = "Unsolvable cube"
        End Select



        MessageBox.Show(TextOfMessageBox, CaptionOfMessageBox, MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Sub

    Shared Sub WeWarnTheUser(TextOfMessageBox As String, Optional CadenaDeTitulo As String = "¡BE CAREFUL!")
        MessageBox.Show(TextOfMessageBox, CadenaDeTitulo, MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
    End Sub

    Shared Function WhatIsTheColourCode(Colour As Color) As Integer
        Select Case Colour
            Case Color.Yellow : Return 0
            Case Color.Red : Return 1
            Case Color.Blue : Return 2
            Case Color.Orange : Return 3
            Case Color.Green : Return 4
            Case Color.White : Return 5
            Case Else : WeTerminateWithError(22) : Stop : End
        End Select
    End Function

    Shared Function Power(Base%, Exponent%) As Integer
        Dim Result As Integer = 1
        Dim Counter As Integer
        For Counter = 1 To Exponent
            Result *= Base
        Next
        Return Result
    End Function

    Shared Function CloningOfArray(Matriz() As Integer) As Integer()
        Dim Result(Matriz.GetUpperBound(0)) As Integer
        Dim Counter As Integer
        For Counter = 0 To Matriz.GetUpperBound(0)
            Result(Counter) = Matriz(Counter)
        Next
        Return Result
    End Function


    Shared Sub SimplifyArrayOfMovements(ByRef AuxiliarArray() As Integer)
        Dim Counter As Integer
        Do
            For Counter = 1 To AuxiliarArray.GetUpperBound(0)
                If AuxiliarArray(Counter) = OppositeMovement(AuxiliarArray(Counter - 1)) Then
                    RemoveElementFromTheArray(Counter - 1, AuxiliarArray)
                    RemoveElementFromTheArray(Counter - 1, AuxiliarArray)
                    Continue Do
                End If
            Next
            For Counter = 2 To AuxiliarArray.GetUpperBound(0)
                If AuxiliarArray(Counter) = AuxiliarArray(Counter - 1) AndAlso AuxiliarArray(Counter) = AuxiliarArray(Counter - 2) Then
                    AuxiliarArray(Counter - 2) = OppositeMovement(AuxiliarArray(Counter - 2))
                    RemoveElementFromTheArray(Counter - 1, AuxiliarArray)
                    RemoveElementFromTheArray(Counter - 1, AuxiliarArray)
                    Continue Do
                End If
            Next
            For Counter = 0 To AuxiliarArray.GetUpperBound(0) - 2
                Dim AuxiliarIndex, SecondCounter As Integer
                AuxiliarIndex = 0
                For SecondCounter = Counter + 1 To AuxiliarArray.GetUpperBound(0)
                    If AuxiliarArray(SecondCounter) \ 4 <> AuxiliarArray(Counter) \ 4 Then Exit For
                    If AuxiliarArray(SecondCounter) = OppositeMovement(AuxiliarArray(Counter)) Then
                        RemoveElementFromTheArray(SecondCounter, AuxiliarArray)
                        RemoveElementFromTheArray(Counter, AuxiliarArray)
                        Continue Do
                    End If
                    If AuxiliarArray(SecondCounter) = AuxiliarArray(Counter) Then
                        If AuxiliarIndex = 0 Then
                            AuxiliarIndex = SecondCounter
                        Else
                            RemoveElementFromTheArray(SecondCounter, AuxiliarArray)
                            RemoveElementFromTheArray(AuxiliarIndex, AuxiliarArray)
                            RemoveElementFromTheArray(Counter, AuxiliarArray)
                            Continue Do
                        End If
                    End If
                Next
            Next
            Exit Do
        Loop
    End Sub

    'Shared Function ObtenerMatrizDeRepeticionesDeMovimientos(MatrizDeMovimientosAuxiliar() As Integer) As Integer()
    '    Dim Result(11) As Integer
    '    Dim Counter As Integer
    '    For Counter = 0 To MatrizDeMovimientosAuxiliar.GetUpperBound(0)
    '        Result(MatrizDeMovimientosAuxiliar(Counter)) += 1
    '    Next
    '    Return Result
    'End Function

    Shared Sub AddElementToArray(ByVal Element As Integer, ByRef Matriz() As Integer)
        If Matriz Is Nothing Then
            ReDim Matriz(0)
        Else
            ReDim Preserve Matriz(Matriz.GetLength(0))
        End If
        Matriz(Matriz.GetUpperBound(0)) = Element
    End Sub

    Shared Sub RemoveElementFromTheArray(ByVal Indice As Integer, ByRef Matriz() As Integer)
        Dim Counter As Integer
        For Counter = Indice To Matriz.GetUpperBound(0) - 1
            Matriz(Counter) = Matriz(Counter + 1)
        Next
        ReDim Preserve Matriz(Matriz.GetUpperBound(0) - 1)
    End Sub

    'Shared Sub CreateArrayOfRandomMovements(ByRef Matriz() As Integer, NumberOfMovements As Integer, Seed As Long)
    '    Matriz = Nothing
    '    Randomize(Seed)
    '    Dim AuxiliarInteger, Counter As Integer
    '    Dim RealAuxiliar As Single
    '    For Counter = 1 To NumberOfMovements
    '        RealAuxiliar = 12 * Rnd()
    '        RealAuxiliar = Int(RealAuxiliar)
    '        AuxiliarInteger = CType(RealAuxiliar, Integer)
    '        AddElementToArray(AuxiliarInteger, Matriz)
    '    Next
    '    SimplifyArrayOfMovements(Matriz)
    '    Do While Matriz.GetLength(0) < NumberOfMovements
    '        For Counter = Matriz.GetLength(0) + 1 To NumberOfMovements
    '            RealAuxiliar = 12 * Rnd()
    '            RealAuxiliar = Int(RealAuxiliar)
    '            AuxiliarInteger = CType(RealAuxiliar, Integer)
    '            AddElementToArray(AuxiliarInteger, Matriz)
    '        Next
    '        SimplifyArrayOfMovements(Matriz)
    '    Loop
    'End Sub

    'Shared Sub InitializeRubiksCube(AuxiliarCube As ClaseCuboDeRubik)
    '    AuxiliarCube.ArrayOfRubiksCube = {0, 2015539, 4031078, 6046617, 8062156, 10077695}
    'End Sub

    'Shared Sub InitializeRubiksCube(ByRef PointedCube As ClaseCuboDeRubik, NumberOfMovements As Integer, Seed As Long)
    '    Dim AuxiliarCube As ClaseCuboDeRubik = New ClaseCuboDeRubik
    '    InitializeRubiksCube(AuxiliarCube)
    '    Dim ArrayOfMovements() As Integer
    '    CreateArrayOfRandomMovements(ArrayOfMovements, NumberOfMovements, Seed)
    '    MessageBox.Show(FromNumberOfMovementToString(ArrayOfMovements))
    '    Dim Counter As Integer
    '    For Counter = 0 To ArrayOfMovements.GetUpperBound(0)
    '        AuxiliarCube.ExecuteMovement(ArrayOfMovements(Counter))
    '    Next
    '    PointedCube.ArrayOfRubiksCube = CloningOfArray(AuxiliarCube.ArrayOfRubiksCube)
    'End Sub

    Shared Sub InitializeArray(Matriz() As Integer)
        Dim Counter As Integer
        For Counter = 0 To Matriz.GetUpperBound(0)
            Matriz(Counter) = 0
        Next
    End Sub

    'Shared Sub InitializeArray(Matriz(,) As Integer)
    '    Dim One, Cero As Integer
    '    For Cero = 0 To Matriz.GetUpperBound(0)
    '        For One = 0 To Matriz.GetUpperBound(1)
    '            Matriz(Cero, One) = 0
    '        Next
    '    Next
    'End Sub

    'Shared Sub InitializeArray(Matriz(,,) As Integer)
    '    Dim Cero, One, Two As Integer
    '    For Cero = 0 To Matriz.GetUpperBound(0)
    '        For One = 0 To Matriz.GetUpperBound(1)
    '            For Two = 0 To Matriz.GetUpperBound(2)
    '                Matriz(Cero, One, Two) = 0
    '            Next
    '        Next
    '    Next
    'End Sub



End Class


