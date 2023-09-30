Sub CalculateFraminghamRiskScore()
    ' Define variables for patient information
    Dim age As Integer
    Dim gender As String
    Dim hdlLevel As Double
    Dim tcLevel As Double
    Dim sbp As Integer
    Dim hasDiabetes As Boolean
    Dim isSmoker As Boolean
    Dim riskPoints As Integer
    Dim riskPercentage As Double
    Dim isTreated As Boolean
    Dim hasFamilyHistoryCVD As Boolean

    ' Assign patient data to variables (replace with actual patient data)
    age = 50
    gender = "M"
    hdlLevel = 1.0
    tcLevel = 5.0
    sbp = 130
    hasDiabetes = False
    isSmoker = True
    isTreated = True ' Set this based on hypertension treatment status
    hasFamilyHistoryCVD = True ' Set this based on family history of premature CVD

    ' Calculate risk points based on age, gender, HDL-C level, and TC level
    If age >= 30 And age <= 34 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 0
            Case "F": riskPoints = riskPoints + 0
        End Select
    ElseIf age >= 35 And age <= 39 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 2
            Case "F": riskPoints = riskPoints + 2
        End Select
    ElseIf age >= 40 And age <= 44 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 5
            Case "F": riskPoints = riskPoints + 4
        End Select
    ElseIf age >= 45 And age <= 49 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 6
            Case "F": riskPoints = riskPoints + 5
        End Select
    ElseIf age >= 50 And age <= 54 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 8
            Case "F": riskPoints = riskPoints + 7
        End Select
    ElseIf age >= 55 And age <= 59 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 10
            Case "F": riskPoints = riskPoints + 8
        End Select
    ElseIf age >= 60 And age <= 64 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 11
            Case "F": riskPoints = riskPoints + 9
        End Select
    ElseIf age >= 65 And age <= 69 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 12
            Case "F": riskPoints = riskPoints + 10
        End Select
    ElseIf age >= 70 And age <= 74 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 14
            Case "F": riskPoints = riskPoints + 11
        End Select
    ElseIf age >= 75 Then
        Select Case gender
            Case "M": riskPoints = riskPoints + 15
            Case "F": riskPoints = riskPoints + 12
        End Select
    End If

    ' Calculate additional risk points based on HDL-C level
    Select Case gender
        Case "M":
            If hdlLevel > 1.6 Then
                riskPoints = riskPoints - 2
            ElseIf hdlLevel >= 1.3 And hdlLevel <= 1.6 Then
                riskPoints = riskPoints - 1
            ElseIf hdlLevel >= 1.0 And hdlLevel < 1.3 Then
                riskPoints = riskPoints + 0
            ElseIf hdlLevel >= 0.9 And hdlLevel <= 1.2 Then
                riskPoints = riskPoints + 1
            ElseIf hdlLevel < 0.9 Then
                riskPoints = riskPoints + 2
            End If
        Case "F":
            If hdlLevel > 1.6 Then
                riskPoints = riskPoints - 2
            ElseIf hdlLevel >= 1.3 And hdlLevel <= 1.6 Then
                riskPoints = riskPoints - 1
            ElseIf hdlLevel >= 1.0 And hdlLevel < 1.3 Then
                riskPoints = riskPoints + 0
            ElseIf hdlLevel >= 0.9 And hdlLevel <= 1.2 Then
                riskPoints = riskPoints + 1
            ElseIf hdlLevel < 0.9 Then
                riskPoints = riskPoints + 2
            End If
    End Select

    ' Calculate additional risk points based on TC level
    Select Case gender
        Case "M":
            If tcLevel < 4.1 Then
                riskPoints = riskPoints + 0
            ElseIf tcLevel >= 4.1 And tcLevel <= 5.2 Then
                riskPoints = riskPoints + 1
            ElseIf tcLevel > 5.2 And tcLevel <= 6.2 Then
                riskPoints = riskPoints + 2
            ElseIf tcLevel > 6.2 And tcLevel <= 7.2 Then
                riskPoints = riskPoints + 3
            ElseIf tcLevel > 7.2 Then
                riskPoints = riskPoints + 4
            End If
        Case "F":
            If tcLevel < 4.1 Then
                riskPoints = riskPoints + 0
            ElseIf tcLevel >= 4.1 And tcLevel <= 5.2 Then
                riskPoints = riskPoints + 1
            ElseIf tcLevel > 5.2 And tcLevel <= 6.2 Then
                riskPoints = riskPoints + 3
            ElseIf tcLevel > 6.2 And tcLevel <= 7.2 Then
                riskPoints = riskPoints + 4
            ElseIf tcLevel > 7.2 Then
                riskPoints = riskPoints + 5
            End If
    End Select

    ' Calculate additional risk points based on SBP and hypertension treatment status
    Select Case gender
        Case "M":
            If sbp < 120 Then
                If isTreated Then
                    riskPoints = riskPoints + 0 ' Treated
                Else
                    riskPoints = riskPoints - 2 ' Not treated
                End If
            ElseIf sbp >= 120 And sbp <= 129 Then
                If isTreated Then
                    riskPoints = riskPoints + 2 ' Treated
                Else
                    riskPoints = riskPoints + 0 ' Not treated
                End If
            ElseIf sbp >= 130 And sbp <= 139 Then
                If isTreated Then
                    riskPoints = riskPoints + 3 ' Treated
                Else
                    riskPoints = riskPoints + 1 ' Not treated
                End If
            ElseIf sbp >= 140 And sbp <= 149 Then
                If isTreated Then
                    riskPoints = riskPoints + 4 ' Treated
                Else
                    riskPoints = riskPoints + 2 ' Not treated
                End If
            ElseIf sbp >= 150 And sbp <= 159 Then
                If isTreated Then
                    riskPoints = riskPoints + 4 ' Treated
                Else
                    riskPoints = riskPoints + 2 ' Not treated
                End If
            ElseIf sbp >= 160 Then
                If isTreated Then
                    riskPoints = riskPoints + 5 ' Treated
                Else
                    riskPoints = riskPoints + 3 ' Not treated
                End If
            End If
        Case "F":
            If sbp < 120 Then
                If isTreated Then
                    riskPoints = riskPoints - 1 ' Treated
                Else
                    riskPoints = riskPoints - 3 ' Not treated
                End If
            ElseIf sbp >= 120 And sbp <= 129 Then
                If isTreated Then
                    riskPoints = riskPoints + 2 ' Treated
                Else
                    riskPoints = riskPoints + 0 ' Not treated
                End If
            ElseIf sbp >= 130 And sbp <= 139 Then
                If isTreated Then
                    riskPoints = riskPoints + 3 ' Treated
                Else
                    riskPoints = riskPoints + 1 ' Not treated
                End If
            ElseIf sbp >= 140 And sbp <= 149 Then
                If isTreated Then
                    riskPoints = riskPoints + 5 ' Treated
                Else
                    riskPoints = riskPoints + 2 ' Not treated
                End If
            ElseIf sbp >= 150 And sbp <= 159 Then
                If isTreated Then
                    riskPoints = riskPoints + 6 ' Treated
                Else
                    riskPoints = riskPoints + 4 ' Not treated
                End If
            ElseIf sbp >= 160 Then
                If isTreated Then
                    riskPoints = riskPoints + 5 ' Treated
                Else
                    riskPoints = riskPoints + 7 ' Not treated
                End If
            End If
    End Select

    ' Calculate additional risk points based on diabetes status
    Select Case gender
        Case "M":
            If hasDiabetes Then
                riskPoints = riskPoints + 3
            End If
        Case "F":
            If hasDiabetes Then
                riskPoints = riskPoints + 4
            End If
    End Select

    ' Calculate additional risk points based on smoking status
    Select Case gender
        Case "M":
            If isSmoker Then
                riskPoints = riskPoints + 4
            End If
        Case "F":
            If isSmoker Then
                riskPoints = riskPoints + 3
            End If
    End Select

    ' Calculate the total risk points
    Dim totalRiskPoints As Integer
    totalRiskPoints = riskPoints

    ' Calculate the risk percentage based on total risk points and gender
    Select Case gender
        Case "M":
            Select Case totalRiskPoints
                Case Is <= 3
                    riskPercentage = 1.0
                Case -2
                    riskPercentage = 1.1
                Case -1
                    riskPercentage = 1.4
                Case 0
                    riskPercentage = 1.6
                Case 1
                    riskPercentage = 1.9
                Case 2
                    riskPercentage = 2.3
                Case 3
                    riskPercentage = 2.8
                Case 4
                    riskPercentage = 3.3
                Case 5
                    riskPercentage = 3.9
                Case 6
                    riskPercentage = 4.7
                Case 7
                    riskPercentage = 5.6
                Case 8
                    riskPercentage = 6.7
                Case 9
                    riskPercentage = 7.9
                Case 10
                    riskPercentage = 9.4
                Case 11
                    riskPercentage = 11.2
                Case 12
                    riskPercentage = 13.3
                Case 13
                    riskPercentage = 15.6
                Case 14
                    riskPercentage = 18.4
                Case 15
                    riskPercentage = 21.6
                Case 16
                    riskPercentage = 25.3
                Case 17
                    riskPercentage = 29.3
                Case Is >= 18
                    riskPercentage = 30.0
            End Select
        Case "F":
            Select Case totalRiskPoints
                Case Is <= 3
                    riskPercentage = 1.0
                Case -2
                    riskPercentage = 1.0
                Case -1
                    riskPercentage = 1.2
                Case 0
                    riskPercentage = 1.5
                Case 1
                    riskPercentage = 1.7
                Case 2
                    riskPercentage = 2.0
                Case 3
                    riskPercentage = 2.4
                Case 4
                    riskPercentage = 2.8
                Case 5
                    riskPercentage = 3.3
                Case 6
                    riskPercentage = 3.9
                Case 7
                    riskPercentage = 4.5
                Case 8
                    riskPercentage = 5.3
                Case 9
                    riskPercentage = 6.3
                Case 10
                    riskPercentage = 7.3
                Case 11
                    riskPercentage = 8.6
                Case 12
                    riskPercentage = 10.0
                Case 13
                    riskPercentage = 11.7
                Case 14
                    riskPercentage = 13.7
                Case 15
                    riskPercentage = 15.9
                Case 16
                    riskPercentage = 18.5
                Case 17
                    riskPercentage = 21.5
                Case 18
                    riskPercentage = 24.8
                Case 19
                    riskPercentage = 27.5
                Case Is >= 20
                    riskPercentage = 30.0
            End Select
    End Select

    ' Check if the patient has a family history of premature CVD
    If hasFamilyHistoryCVD Then
        ' If true, multiply the riskPercentage by 2
        riskPercentage = riskPercentage * 2
    End If

    ' Determine Framingham Risk Score based on riskPercentage
    Dim riskScore As String

    If riskPercentage < 10.0 Then
        riskScore = "Low"
    ElseIf riskPercentage >= 10.0 And riskPercentage < 20.0 Then
        riskScore = "Intermediate"
    Else
        riskScore = "High"
    End If

    ' Display the calculated Framingham Risk Score
    MsgBox "Framingham Risk Score: " & riskScore, vbInformation, "Risk Assessment"
End Sub
