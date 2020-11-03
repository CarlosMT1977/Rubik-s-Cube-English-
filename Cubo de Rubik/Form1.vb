Option Explicit On
Option Strict On

Imports Cubo_de_Rubik.Utilidades
Imports Cubo_de_Rubik.MontarElCubo

Public Class Form1
    '0: Yellow
    '1: Red
    '2: Blue
    '3: Orange
    '4: Green
    '5: White
    Private InitialRubiksCube, FinalRubiksCube As ClaseCuboDeRubik
    Private Const LeftDistance As Integer = 50
    Private Const TopDistance As Integer = 50
    Private Const MinimumDistance As Integer = 10
    Private Const MaximumDistance As Integer = 20
    Private ArrayOfButtons(53) As Button
    Dim ArrayOfSamples(5) As Button
    Private Const WidthOfTheButton As Integer = 30
    Private Const HeightOfTheButton As Integer = 30
    Private Const WidthOfTheSample As Integer = 55
    Private Const HeightOfTheSample As Integer = 55
    Private DistanceBetweenSamples As Integer
    Private Colours() As Color = {Color.Yellow, Color.Red, Color.Blue, Color.Orange, Color.Green, Color.White}
    Private Const MinimumBorder As Integer = 1
    Private Const MaximumBorder As Integer = 10

    Dim SelectedColour As Integer = -1

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MessageBox.Show("The Front Face is the YELLOW one, and the Down Face is the RED one")
        CreateArrayOfControls()

    End Sub

    Private Sub btnSolveRubiksCube_Click(sender As Object, e As EventArgs) Handles btnSolveRubiksCube.Click

        If Not IsTheArrayOfControlsProperlyFilled() Then
            WeWarnTheUser("You still need to assign a colour somewhere", "Colour assignment needed")
            Exit Sub
        End If

        MessageBox.Show("I see you have filled all the buttons of the form")
        InitialRubiksCube = New ClaseCuboDeRubik
        WeAssignArrayOfControlsToRubiksCube(InitialRubiksCube)




        Dim SolutionOfTheWholeCube As SolveTheCubeDeLuxe = New SolveTheCubeDeLuxe(InitialRubiksCube)
        SolutionOfTheWholeCube.SolveTheCube()
        SolutionOfTheWholeCube.ShowSolutionOfTheCube()


        MessageBox.Show("Check if the whole cube has been already solved")


    End Sub

    Private Sub CreateArrayOfControls()
        Dim Counter, XVariable, YVariable, XFixed, YFixed, XAux, YAux As Integer
        Dim NumberOfFacesOnTheLeft, NumberOfFacesOnTop As Integer
        For Counter = 0 To 53
            ArrayOfButtons(Counter) = New Button
            ArrayOfButtons(Counter).Size = New Size(WidthOfTheButton, HeightOfTheButton)
            ArrayOfButtons(Counter).Text = vbNullString
            Select Case Counter \ 9
                Case 2 : NumberOfFacesOnTheLeft = 0
                Case 0, 1, 3 : NumberOfFacesOnTheLeft = 1
                Case 4 : NumberOfFacesOnTheLeft = 2
                Case 5 : NumberOfFacesOnTheLeft = 3
            End Select
            Select Case Counter \ 9
                Case 3 : NumberOfFacesOnTop = 0
                Case 0, 2, 4, 5 : NumberOfFacesOnTop = 1
                Case 1 : NumberOfFacesOnTop = 2
            End Select
            XFixed = LeftDistance + 3 * NumberOfFacesOnTheLeft * (WidthOfTheButton + MinimumDistance) + NumberOfFacesOnTheLeft * (MaximumDistance - MinimumDistance)
            XVariable = (WidthOfTheButton + MinimumDistance) * (Counter Mod 3)
            YFixed = TopDistance + 3 * NumberOfFacesOnTop * (HeightOfTheButton + MinimumDistance) + NumberOfFacesOnTop * (MaximumDistance - MinimumDistance)
            YVariable = (HeightOfTheButton + MinimumDistance) * ((Counter Mod 9) \ 3)
            ArrayOfButtons(Counter).Location = New Point(XFixed + XVariable, YFixed + YVariable)
            ArrayOfButtons(Counter).Name = CType(Counter, String)
            AddHandler ArrayOfButtons(Counter).Click, AddressOf WhenWeClickOnAButtonOfTheCube
            Me.Controls.Add(ArrayOfButtons(Counter))
            If Counter Mod 9 = 4 Then ArrayOfButtons(Counter).BackColor = Colours(Counter \ 9)
        Next

        Panel1.Size = New Size(6 * WidthOfTheButton + 4 * MinimumDistance + MaximumDistance, 3 * HeightOfTheButton + 2 * MinimumDistance)
        XFixed = ArrayOfButtons(17).Location.X + WidthOfTheButton + MaximumDistance
        YFixed = ArrayOfButtons(8).Location.Y + HeightOfTheButton + MaximumDistance
        Panel1.Location = New Point(XFixed, YFixed)

        DistanceBetweenSamples = CType((6 * WidthOfTheButton + 4 * MinimumDistance + MaximumDistance - 3 * WidthOfTheSample) / 4, Integer)
        For Counter = 0 To 5
            ArrayOfSamples(Counter) = New Button
            ArrayOfSamples(Counter).Size = New Size(WidthOfTheSample, HeightOfTheSample)
            XFixed = DistanceBetweenSamples
            XVariable = (WidthOfTheSample + DistanceBetweenSamples) * (Counter Mod 3)
            YFixed = DistanceBetweenSamples
            YVariable = (HeightOfTheSample + DistanceBetweenSamples) * (Counter \ 3)
            ArrayOfSamples(Counter).Location = New Point(XFixed + XVariable, YFixed + YVariable)
            Panel1.Controls.Add(ArrayOfSamples(Counter))
            ArrayOfSamples(Counter).BackColor = Colours(Counter)
            ArrayOfSamples(Counter).FlatStyle = FlatStyle.Flat
            ArrayOfSamples(Counter).FlatAppearance.BorderColor = Color.Black
            ArrayOfSamples(Counter).FlatAppearance.BorderSize = MinimumBorder
            ArrayOfSamples(Counter).Name = CType(Counter, String)
            AddHandler ArrayOfSamples(Counter).Click, AddressOf WhenWeClickOnAColourSample
        Next
        Panel1.Size = New Size(3 * WidthOfTheSample + 4 * DistanceBetweenSamples, 2 * HeightOfTheSample + 3 * DistanceBetweenSamples)
        XAux = LeftDistance + Maximum(ArrayOfButtons(53).Location.X + WidthOfTheButton, Panel1.Location.X + Panel1.Size.Width)
        YAux = TopDistance + Maximum(ArrayOfButtons(17).Location.Y + HeightOfTheButton, Panel1.Location.Y + Panel1.Size.Height)
        Me.ClientSize = New Size(XAux, YAux)
        XAux = CType((ArrayOfButtons(18).Location.X + ArrayOfButtons(20).Location.X + ArrayOfButtons(20).Size.Width) / 2 - btnSolveRubiksCube.Size.Width / 2, Integer)
        YAux = CType((ArrayOfButtons(27).Location.Y + ArrayOfButtons(33).Location.Y + ArrayOfButtons(33).Size.Height) / 2 - btnSolveRubiksCube.Size.Height / 2, Integer)
        btnSolveRubiksCube.Location = New Point(XAux, YAux)
    End Sub

    Private Sub WhenWeClickOnAButtonOfTheCube(sender As Object, e As EventArgs)
        Dim PointedButton As Button = CType(sender, Button)
        If CType(PointedButton.Name, Integer) Mod 9 = 4 Then
            WeWarnTheUser("You can't modify the colour of the centre squares of each face", "Non-modifiable colour")
        ElseIf SelectedColour = -1 Then
            WeWarnTheUser("In order to assign a colour, you must first select it from the samples", "Select from the samples")
        Else
            PointedButton.BackColor = Colours(SelectedColour)
        End If
    End Sub

    Private Sub WhenWeClickOnAColourSample(sender As Object, e As EventArgs)
        Dim ConnectedSample As Button = CType(sender, Button)
        If SelectedColour <> -1 Then ArrayOfSamples(SelectedColour).FlatAppearance.BorderSize = MinimumBorder
        SelectedColour = CType(ConnectedSample.Name, Integer)
        ArrayOfSamples(SelectedColour).FlatAppearance.BorderSize = MaximumBorder
    End Sub

    Private Function IsTheArrayOfControlsProperlyFilled() As Boolean
        Dim SquareCounter, ColourCounter As Integer
        For SquareCounter = 0 To 53
            For ColourCounter = 0 To 5
                If ArrayOfButtons(SquareCounter).BackColor = Colours(ColourCounter) Then Exit For
            Next
            If ColourCounter = 6 Then Return False
        Next
        Return True
    End Function


    Private Sub WeAssignArrayOfControlsToRubiksCube(ByRef AuxiliarCube As ClaseCuboDeRubik)
        InitializeArray(AuxiliarCube.ArrayOfRubiksCube)
        Dim Counter, BlueSquare, GreenSquare, WhiteSquare As Integer
        For Counter = 0 To 8
            AuxiliarCube.ArrayOfRubiksCube(0) += WhatIsTheColourCode(ArrayOfButtons(Counter).BackColor) * Power(6, Counter)
            AuxiliarCube.ArrayOfRubiksCube(1) += WhatIsTheColourCode(ArrayOfButtons(9 + Counter).BackColor) * Power(6, Counter)
        Next
        For Counter = 0 To 8
            Select Case Counter
                Case 0 : BlueSquare = 6 : GreenSquare = 2 : WhiteSquare = 8
                Case 1 : BlueSquare = 3 : GreenSquare = 5 : WhiteSquare = 7
                Case 2 : BlueSquare = 0 : GreenSquare = 8 : WhiteSquare = 6
                Case 3 : BlueSquare = 7 : GreenSquare = 1 : WhiteSquare = 5
                Case 4 : BlueSquare = 4 : GreenSquare = 4 : WhiteSquare = 4
                Case 5 : BlueSquare = 1 : GreenSquare = 7 : WhiteSquare = 3
                Case 6 : BlueSquare = 8 : GreenSquare = 0 : WhiteSquare = 2
                Case 7 : BlueSquare = 5 : GreenSquare = 3 : WhiteSquare = 1
                Case 8 : BlueSquare = 2 : GreenSquare = 6 : WhiteSquare = 0
                Case Else : WeTerminateWithError(22) : Stop : End
            End Select
            AuxiliarCube.ArrayOfRubiksCube(2) += WhatIsTheColourCode(ArrayOfButtons(18 + Counter).BackColor) * Power(6, BlueSquare)
            AuxiliarCube.ArrayOfRubiksCube(3) += WhatIsTheColourCode(ArrayOfButtons(27 + Counter).BackColor) * Power(6, WhiteSquare)
            AuxiliarCube.ArrayOfRubiksCube(4) += WhatIsTheColourCode(ArrayOfButtons(36 + Counter).BackColor) * Power(6, GreenSquare)
            AuxiliarCube.ArrayOfRubiksCube(5) += WhatIsTheColourCode(ArrayOfButtons(45 + Counter).BackColor) * Power(6, WhiteSquare)
        Next
    End Sub


End Class


