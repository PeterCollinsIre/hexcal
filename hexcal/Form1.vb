Public Class Form1
    Public Structure pointCoordinate
        Public globalNorthing, globalEasting, globalElevation, gridNorthing, gridEasting, gridElevation, afterRolllNorthing, afterRollEasting, afterRollElevation, projectedLength, northingLength, planDistance As Double
        Public isAssigned As Boolean
    End Structure

    Dim leftMast, rightMast, controlPoint1, controlPoint2, aPin, a1Pin,
        a3Pin, bPin, b1Pin, gPin, dPin, fPin, hPin, jPin, rPin, sPin, cPin,
        boomSensor, stickSensor, dogboneSensor As pointCoordinate


    Dim inputcFile As String
    Dim inputcFileParse As String
    Dim linecount As Int16 = 1
    Dim N As String
    Dim c As Int16 = 0
    Dim ctOneCase As Int16 = 0
    Dim ctRot As Double


    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
        TextBox1.Text = OpenFileDialog1.FileName

    End Sub



    Public Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

        Try

            Using MyReader As New Microsoft.VisualBasic.
                        FileIO.TextFieldParser(TextBox1.Text)
                MyReader.TextFieldType = FileIO.FieldType.Delimited
                MyReader.SetDelimiters(",")
                Dim currentRow As String()
                While Not MyReader.EndOfData
                    Try
                        currentRow = MyReader.ReadFields()

                        Dim currentField As String

                        For Each currentField In currentRow



                            If N = 3 Then
                                controlPoint1.globalElevation = currentField
                                N = 0
                            End If

                            If N = 2 Then
                                controlPoint1.globalEasting = currentField
                                N = 3
                            End If

                            If N = 1 Then
                                controlPoint1.globalNorthing = currentField
                                N = 2
                            End If

                            If currentField = "ct1" Or currentField = "CT1" Or currentField = "Ct1" Or currentField = "CTL1" Then
                                controlPoint1.isAssigned = True
                                N = N + 1
                            End If

                            If N = 6 Then
                                controlPoint2.globalElevation = currentField
                                N = 0
                            End If

                            If N = 5 Then
                                controlPoint2.globalEasting = currentField
                                N = 6
                            End If

                            If N = 4 Then
                                controlPoint2.globalNorthing = currentField
                                N = 5
                            End If

                            If currentField = "ct2" Or currentField = "CT2" Or currentField = "Ct2" Or currentField = "CTL2" Then
                                controlPoint2.isAssigned = True
                                N = N + 4
                            End If


                            If N = 9 Then
                                aPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 8 Then
                                aPin.globalEasting = currentField
                                N = 9
                            End If

                            If N = 7 Then
                                aPin.globalNorthing = currentField
                                N = 8
                            End If

                            If currentField = "A" Or currentField = "a" Then
                                aPin.isAssigned = True
                                N = N + 7
                            End If

                            If N = 12 Then
                                bPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 11 Then
                                bPin.globalEasting = currentField
                                N = 12
                            End If

                            If N = 10 Then
                                bPin.globalNorthing = currentField
                                N = 11
                            End If

                            If currentField = "b" Or currentField = "B" Then
                                bPin.isAssigned = True
                                N = N + 10
                            End If

                            If N = 15 Then
                                gPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 14 Then
                                gPin.globalEasting = currentField
                                N = 15
                            End If

                            If N = 13 Then
                                gPin.globalNorthing = currentField
                                N = 14
                            End If

                            If currentField = "g" Or currentField = "G" Then
                                gPin.isAssigned = True
                                N = N + 13
                            End If

                            If N = 18 Then
                                dPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 17 Then
                                dPin.globalEasting = currentField
                                N = 18
                            End If

                            If N = 16 Then
                                dPin.globalNorthing = currentField
                                N = 17
                            End If

                            If currentField = "d" Or currentField = "D" Then
                                dPin.isAssigned = True
                                N = N + 16
                            End If

                            If N = 21 Then
                                fPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 20 Then
                                fPin.globalEasting = currentField
                                N = 21
                            End If

                            If N = 19 Then
                                fPin.globalNorthing = currentField
                                N = 20
                            End If

                            If currentField = "f" Or currentField = "F" Then
                                fPin.isAssigned = True
                                N = N + 19
                            End If

                            If N = 24 Then
                                hPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 23 Then
                                hPin.globalEasting = currentField
                                N = 24
                            End If

                            If N = 22 Then
                                hPin.globalNorthing = currentField
                                N = 23
                            End If

                            If currentField = "h" Or currentField = "H" Then
                                hPin.isAssigned = True
                                N = N + 22
                            End If

                            If N = 27 Then
                                jPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 26 Then
                                jPin.globalEasting = currentField
                                N = 27
                            End If

                            If N = 25 Then
                                jPin.globalNorthing = currentField
                                N = 26
                            End If

                            If currentField = "j" Or currentField = "J" Then
                                jPin.isAssigned = True
                                N = N + 25
                            End If

                            If N = 30 Then
                                b1Pin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 29 Then
                                b1Pin.globalEasting = currentField
                                N = 30
                            End If

                            If N = 28 Then
                                b1Pin.globalNorthing = currentField
                                N = 29
                            End If

                            If currentField = "b1" Or currentField = "B1" Then
                                b1Pin.isAssigned = True
                                N = N + 28
                            End If

                            If N = 33 Then
                                boomSensor.globalElevation = currentField
                                N = 0
                            End If

                            If N = 32 Then
                                boomSensor.globalEasting = currentField
                                N = 33
                            End If

                            If N = 31 Then
                                boomSensor.globalNorthing = currentField
                                N = 32
                            End If

                            If currentField = "sb" Or currentField = "SB" Then
                                boomSensor.isAssigned = True
                                N = N + 31
                            End If

                            If N = 36 Then
                                stickSensor.globalElevation = currentField
                                N = 0
                            End If

                            If N = 35 Then
                                stickSensor.globalEasting = currentField
                                N = 36
                            End If

                            If N = 34 Then
                                stickSensor.globalNorthing = currentField
                                N = 35
                            End If

                            If currentField = "ss" Or currentField = "SS" Then
                                stickSensor.isAssigned = True
                                N = N + 34
                            End If

                            If N = 39 Then
                                dogboneSensor.globalElevation = currentField
                                N = 0
                            End If

                            If N = 38 Then
                                dogboneSensor.globalEasting = currentField
                                N = 39
                            End If

                            If N = 37 Then
                                dogboneSensor.globalNorthing = currentField
                                N = 38
                            End If

                            If currentField = "sd" Or currentField = "SD" Then
                                dogboneSensor.isAssigned = True
                                N = N + 37
                            End If

                            If N = 42 Then
                                a1Pin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 41 Then
                                a1Pin.globalEasting = currentField
                                N = 42
                            End If

                            If N = 40 Then
                                a1Pin.globalNorthing = currentField
                                N = 41
                            End If

                            If currentField = "a1" Or currentField = "A1" Then
                                a1Pin.isAssigned = True
                                N = N + 40
                            End If

                            If N = 45 Then
                                a3Pin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 44 Then
                                a3Pin.globalEasting = currentField
                                N = 45
                            End If

                            If N = 43 Then
                                a3Pin.globalNorthing = currentField
                                N = 44
                            End If

                            If currentField = "a3" Or currentField = "A3" Then
                                a3Pin.isAssigned = True
                                N = N + 43
                            End If

                            If N = 48 Then
                                rPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 47 Then
                                rPin.globalEasting = currentField
                                N = 48
                            End If

                            If N = 46 Then
                                rPin.globalNorthing = currentField
                                N = 47
                            End If

                            If currentField = "r" Or currentField = "R" Then
                                rPin.isAssigned = True
                                N = N + 46
                            End If

                            If N = 51 Then
                                sPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 50 Then
                                sPin.globalEasting = currentField
                                N = 51
                            End If

                            If N = 49 Then
                                sPin.globalNorthing = currentField
                                N = 50
                            End If

                            If currentField = "s" Or currentField = "S" Then
                                sPin.isAssigned = True
                                N = N + 49
                            End If

                            If N = 54 Then
                                cPin.globalElevation = currentField
                                N = 0
                            End If

                            If N = 53 Then
                                cPin.globalEasting = currentField
                                N = 54
                            End If

                            If N = 52 Then
                                cPin.globalNorthing = currentField
                                N = 53
                            End If

                            If currentField = "c" Or currentField = "C" Then
                                cPin.isAssigned = True
                                N = N + 52
                            End If

                            If N = 57 Then
                                leftMast.globalElevation = currentField
                                N = 0
                            End If

                            If N = 56 Then
                                leftMast.globalEasting = currentField
                                N = 57
                            End If

                            If N = 55 Then
                                leftMast.globalNorthing = currentField
                                N = 56
                            End If

                            If currentField = "lm" Or currentField = "LM" Then
                                leftMast.isAssigned = True
                                N = N + 55
                            End If

                            If N = 60 Then
                                rightMast.globalElevation = currentField
                                N = 0
                            End If

                            If N = 59 Then
                                rightMast.globalEasting = currentField
                                N = 60
                            End If

                            If N = 58 Then
                                rightMast.globalNorthing = currentField
                                N = 59
                            End If

                            If currentField = "rm" Or currentField = "RM" Then
                                rightMast.isAssigned = True
                                N = N + 58
                            End If





                        Next
                    Catch ex As Microsoft.VisualBasic.
                  FileIO.MalformedLineException
                        MsgBox("Line " & ex.Message &
        "is not valid and will be skipped.")
                    End Try
                End While
            End Using

        Catch ex As Exception
            MsgBox("No file selected")

        End Try

        'Find the heading of two control points
        ctRot = Math.Atan2(controlPoint2.globalNorthing - controlPoint1.globalNorthing, controlPoint2.globalEasting - controlPoint1.globalEasting)

        'Now we rotate and project the points
        leftMast.projectedLength = EastingDistanceFromControlPointOne(leftMast)
        rightMast.projectedLength = EastingDistanceFromControlPointOne(rightMast)
        aPin.projectedLength = EastingDistanceFromControlPointOne(aPin)
        a1Pin.projectedLength = EastingDistanceFromControlPointOne(a1Pin)
        a3Pin.projectedLength = EastingDistanceFromControlPointOne(a3Pin)
        rPin.projectedLength = EastingDistanceFromControlPointOne(rPin)
        boomSensor.projectedLength = EastingDistanceFromControlPointOne(boomSensor)
        stickSensor.projectedLength = EastingDistanceFromControlPointOne(stickSensor)
        dogboneSensor.projectedLength = EastingDistanceFromControlPointOne(dogboneSensor)
        bPin.projectedLength = EastingDistanceFromControlPointOne(bPin)
        sPin.projectedLength = EastingDistanceFromControlPointOne(sPin)
        cPin.projectedLength = EastingDistanceFromControlPointOne(cPin)
        b1Pin.projectedLength = EastingDistanceFromControlPointOne(b1Pin)
        gPin.projectedLength = EastingDistanceFromControlPointOne(gPin)
        dPin.projectedLength = EastingDistanceFromControlPointOne(dPin)
        fPin.projectedLength = EastingDistanceFromControlPointOne(fPin)
        hPin.projectedLength = EastingDistanceFromControlPointOne(hPin)
        jPin.projectedLength = EastingDistanceFromControlPointOne(jPin)


        ' Solve slope length of points
        Label3.Text = FindSlopeDistance(aPin, bPin)
        Label48.Text = FindSlopeDistance(a3Pin, rPin)
        Label50.Text = FindSlopeDistance(a3Pin, bPin)
        Label47.Text = FindSlopeDistance(aPin, a3Pin)
        Label34.Text = FindSlopeDistance(aPin, boomSensor)
        Label35.Text = FindSlopeDistance(boomSensor, bPin)
        Label52.Text = FindSlopeDistance(bPin, sPin)
        Label56.Text = FindSlopeDistance(cPin, dPin)
        Label54.Text = FindSlopeDistance(rPin, bPin)
        Label5.Text = FindSlopeDistance(bPin, gPin)
        Label7.Text = FindSlopeDistance(dPin, gPin)
        Label9.Text = FindSlopeDistance(dPin, fPin)
        Label11.Text = FindSlopeDistance(hPin, fPin)
        Label24.Text = FindSlopeDistance(gPin, jPin)
        Label22.Text = FindSlopeDistance(hPin, gPin)
        Label58.Text = FindSlopeDistance(sPin, gPin)
        Label60.Text = FindSlopeDistance(cPin, gPin)
        Label30.Text = FindSlopeDistance(bPin, stickSensor)
        Label19.Text = FindSlopeDistance(gPin, fPin)
        Label39.Text = FindSlopeDistance(dPin, dogboneSensor)
        ' Label27.Text = FindSlopeDistance(aPin, b1Pin)
        ' Label26.Text = FindSlopeDistance(b1Pin, bPin)

        ' Find the height differences for calibration values
        Label28.Text = ElevationDifference(bPin, aPin)
        Label15.Text = ElevationDifference(bPin, gPin)
        Label17.Text = ElevationDifference(fPin, dPin)
        ' Label28.Text = Math.Round(b1Pin.globalElevation - aPin.globalElevation, 3)
        ' Label30.Text = Math.Round(bPin.globalElevation - b1Pin.globalElevation, 3)

        ' Find distance of stick sensor from B-G line.
        If stickSensor.isAssigned And bPin.isAssigned And gPin.isAssigned Then
            Label37.Text = Math.Round(Math.Sin(Math.Atan2(bPin.projectedLength - gPin.projectedLength, bPin.globalElevation - gPin.globalElevation) _
                                              - Math.Atan2(stickSensor.projectedLength - gPin.projectedLength, stickSensor.globalElevation - gPin.globalElevation)
                                              ) * FindSlopeDistance(gPin, stickSensor), 3) * -1
        End If

        ' Find distance of bucket sensor from F-D line. 
        If dogboneSensor.isAssigned And fPin.isAssigned And dPin.isAssigned Then
            Label41.Text = Math.Round(Math.Sin(Math.Atan2(dogboneSensor.projectedLength - dPin.projectedLength, dogboneSensor.globalElevation - dPin.globalElevation) -
                                                          Math.Atan2(fPin.projectedLength - dPin.projectedLength, fPin.globalElevation - dPin.globalElevation)) * FindSlopeDistance(dPin, dogboneSensor), 3)
        End If



        ' find if stick is extended or curled
        If bPin.projectedLength >= gPin.projectedLength Then
            Label31.Visible = True
        ElseIf bPin.projectedLength < gPin.projectedLength Then
            Label31.Text = "Extended"
            Label31.Visible = True
        End If

        'solve a-a1 VD and HD
        If a1Pin.isAssigned And aPin.isAssigned Then

            Dim aa1Hd As Double = a1Pin.projectedLength - aPin.projectedLength
            aa1Hd = Math.Round(aa1Hd, 3)
            Label42.Text = aa1Hd


            Dim aa1Vd As Double = aPin.globalElevation - a1Pin.globalElevation
            aa1Vd = Math.Round(aa1Vd, 3)
            Label44.Text = aa1Vd
        End If

        'find left mast position
        If leftMast.isAssigned And aPin.isAssigned And controlPoint1.isAssigned And controlPoint2.isAssigned Then

            ConvertGlobalToGrid(leftMast, NumericUpDown1.Value, NumericUpDown2.Value, ctRot)
            ConvertGlobalToGrid(aPin, NumericUpDown1.Value, NumericUpDown2.Value, ctRot)
            Label63.Text = Math.Round(aPin.afterRollEasting - leftMast.afterRollEasting, 3)
            Label64.Text = Math.Round(leftMast.afterRolllNorthing - aPin.afterRolllNorthing - NumericUpDown3.Value, 3)
            Label65.Text = Math.Round(leftMast.afterRollElevation - aPin.afterRollElevation, 3)
        End If

        If rightMast.isAssigned And aPin.isAssigned And controlPoint1.isAssigned And controlPoint2.isAssigned Then
            ConvertGlobalToGrid(rightMast, NumericUpDown1.Value, NumericUpDown2.Value, ctRot)
            ConvertGlobalToGrid(aPin, NumericUpDown1.Value, NumericUpDown2.Value, ctRot)
            Label68.Text = Math.Round(aPin.afterRollEasting - rightMast.afterRollEasting, 3)
            Label67.Text = Math.Round(aPin.afterRolllNorthing - rightMast.afterRolllNorthing + NumericUpDown3.Value, 3)
            Label66.Text = Math.Round(rightMast.afterRollElevation - aPin.afterRollElevation, 3)
        End If

    End Sub

    Private Function ConvertGlobalToGrid(ByRef point As pointCoordinate, ByRef machineMainfall As Double, ByRef machineRoll As Double, ByRef machineHeading As Double)

        Dim a As Double
        point.gridElevation = point.globalElevation
        a = (Math.Sin((Math.Atan2(point.globalNorthing - controlPoint1.globalNorthing, point.globalEasting - controlPoint1.globalEasting)) - machineHeading)) _
        * Math.Sqrt(Math.Pow(point.globalNorthing - controlPoint1.globalNorthing, 2) + Math.Pow(point.globalEasting - controlPoint1.globalEasting, 2))

        point.gridNorthing = controlPoint1.globalNorthing + a
        point.gridEasting = controlPoint1.globalEasting + point.projectedLength

        'grid norting,easting and elevation are all rotated.

        a = (Math.Cos((Math.Atan2(point.gridElevation - controlPoint1.globalElevation, point.gridNorthing - controlPoint1.globalNorthing)) + machineRoll * Math.PI / 180)) _
            * Math.Sqrt(Math.Pow(point.gridElevation - controlPoint1.globalElevation, 2) + Math.Pow(point.gridNorthing - controlPoint1.globalNorthing, 2))
        point.afterRolllNorthing = controlPoint1.globalNorthing + a

        a = (Math.Sin((Math.Atan2(point.gridElevation - controlPoint1.globalElevation, point.gridNorthing - controlPoint1.globalNorthing)) + machineRoll * Math.PI / 180)) _
        * Math.Sqrt(Math.Pow(point.gridElevation - controlPoint1.globalElevation, 2) + Math.Pow(point.gridNorthing - controlPoint1.globalNorthing, 2))
        point.afterRollElevation = controlPoint1.globalElevation + a

        a = Math.Cos((Math.Atan2(point.afterRollElevation - controlPoint1.globalElevation, point.gridEasting - controlPoint1.globalEasting)) + machineMainfall * Math.PI / 180) _
            * Math.Sqrt(Math.Pow(point.afterRollElevation - controlPoint1.globalElevation, 2) + Math.Pow(point.gridEasting - controlPoint1.globalEasting, 2))
        point.afterRollEasting = controlPoint1.globalEasting + a

        a = (Math.Sin((Math.Atan2(point.afterRollElevation - controlPoint1.globalElevation, point.gridEasting - controlPoint1.globalEasting)) + machineMainfall * Math.PI / 180)) _
        * Math.Sqrt(Math.Pow(point.afterRollElevation - controlPoint1.globalElevation, 2) + Math.Pow(point.gridEasting - controlPoint1.globalEasting, 2))
        point.afterRollElevation = controlPoint1.globalElevation + a


        Return 0

    End Function

    '''<summary>
    '''This function rotates the point around control point 1 by the same angle as control 1 - control 2 to east.
    '''Then it returns the easting difference from control 1 to the point called.
    '''</summary>
    Private Function EastingDistanceFromControlPointOne(ByVal e As pointCoordinate)
        Return (Math.Cos((Math.Atan2(e.globalNorthing - controlPoint1.globalNorthing, e.globalEasting - controlPoint1.globalEasting)) - ctRot)) _
        * Math.Sqrt(Math.Pow(e.globalNorthing - controlPoint1.globalNorthing, 2) + Math.Pow(e.globalEasting - controlPoint1.globalEasting, 2))
    End Function
    Private Function NorthingDistanceFromControlPointOneAfterRotation(ByVal e As pointCoordinate)
        Return (Math.Cos((Math.Atan2(e.gridElevation - controlPoint1.globalElevation, e.gridNorthing - controlPoint1.globalNorthing)) + NumericUpDown2.Value * Math.PI / 180)) _
        * Math.Sqrt(Math.Pow(e.gridElevation - controlPoint1.globalElevation, 2) + Math.Pow(e.gridNorthing - controlPoint1.globalNorthing, 2))
    End Function

    Private Function ElevationDistanceFromControlPointOneAfterRotation(ByVal e As pointCoordinate)
        Return (Math.Sin((Math.Atan2(e.gridElevation - controlPoint1.globalElevation, e.gridNorthing - controlPoint1.globalNorthing)) + NumericUpDown2.Value * Math.PI / 180)) _
        * Math.Sqrt(Math.Pow(e.gridElevation - controlPoint1.globalElevation, 2) + Math.Pow(e.gridNorthing - controlPoint1.globalNorthing, 2))
    End Function

    Private Function NorthingDistanceFromControlPointOne(ByVal e As pointCoordinate)
        Return (Math.Sin((Math.Atan2(e.globalNorthing - controlPoint1.globalNorthing, e.globalEasting - controlPoint1.globalEasting)) - ctRot)) _
        * Math.Sqrt(Math.Pow(e.globalNorthing - controlPoint1.globalNorthing, 2) + Math.Pow(e.globalEasting - controlPoint1.globalEasting, 2))
    End Function




    Private Function FindSlopeDistance(ByVal point1 As pointCoordinate, ByVal point2 As pointCoordinate)
        If point1.isAssigned And point2.isAssigned Then
            Return Math.Round(Math.Sqrt(Math.Pow(point2.projectedLength - point1.projectedLength, 2) +
                                        Math.Pow(point2.globalElevation - point1.globalElevation, 2)), 3)
        Else
            Return "No Data"
        End If
    End Function



    Private Function ElevationDifference(ByVal point1 As pointCoordinate, ByVal point2 As pointCoordinate)
        If point1.isAssigned And point2.isAssigned Then
            Return Math.Round(point1.globalElevation - point2.globalElevation, 3)
        Else
            Return "No Data"
        End If
    End Function

    Private Sub TexttBox1_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub TextBox1_DragDrop(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles TextBox1.DragDrop
        Dim s() As String = e.Data.GetData("FileDrop", False)
        Dim i As Integer
        For i = 0 To s.Length - 1
            TextBox1.Text = (s(i))
        Next i
    End Sub
End Class
