' Calculate age
dobString = Patient.DOB

If Not IsEmpty(dobString) Then
    dob = CDate(dobString)
    age = DateDiff("yyyy", dob, Now)
Else
    Profile.MsgBox("Date of Birth data not found")
End If

Profile.MsgBox("Age: " & age & ", Type: " & TypeName(age))

' Retrieve gender
gender = Patient.Sex

If IsEmpty(gender) Then
    Profile.MsgBox("Gender data not found")
End If

Profile.MsgBox("Gender: " & gender & ", Type: " & TypeName(gender))

' Get blood pressure:
Set aPatient = Patient
aConceptCode = "z..2W" ' Use the concept code for blood pressure
aTermsetCode = "IH"
aDateFrom = #01/01/2000#
aDateTo = Now ' Set the end date to the current date for the most up-to-date blood pressure
aUseCrossRefs = False

Set aLatestHRI = aPatient.GetLatestHRI(aTermsetCode, aConceptCode, aDateFrom, aDateTo, aUseCrossRefs)

If aLatestHRI Is Nothing Then
    Profile.MsgBox("There is no latest blood pressure data available.")
Else
    ' Retrieve the blood pressure as a string
    bloodPressureStr = aLatestHRI.AsString

    ' Extract only the systolic blood pressure (the part before the '/')
    systolicPressure = Split(bloodPressureStr, "/")(0)

    ' Remove the "Blood Pressure" text
    systolicPressure = Replace(systolicPressure, "Blood Pressure: ", "")

    ' Convert systolicPressure to an integer
    systolicPressure = CInt(systolicPressure)

    Profile.MsgBox("Systolic Pressure: " & systolicPressure & ", Type: " & TypeName(systolicPressure))
End If

' Get total cholesterol:
aDateTo = Now
aDateFrom = DateAdd("yyyy", -30, aDateTo)
aDateTo = DateAdd("yyyy", 1, aDateTo)

Set aResult = Patient.GetLatestHRI("FHAM", "14647-2", aDateFrom, aDateTo, 1)

If Not aResult Is Nothing Then
    Set aResultContent = aResult.Content
    aResultContentValue = aResultContent.AsString

    ' Extract only the numerical value from the string
    ' Assuming the value is the first part of the string
    cholesterolValue = CDbl(Split(aResultContentValue, " ")(0))

    Profile.MsgBox("Total Cholesterol: " & cholesterolValue & ", Type: " & TypeName(cholesterolValue))
Else
    Profile.MsgBox("No Total Cholesterol Result Found")
End If

' Get HDL cholesterol:
Set aResult = Patient.GetLatestHRI("FHAM", "14646-4", aDateFrom, aDateTo, 1)
If Not aResult Is Nothing Then
    Set aResultContent = aResult.Content
    aResultContentValue = aResultContent.AsString

    ' Extract only the numerical value from the string
    ' Assuming the value is the first part of the string
    HDLValue = CDbl(Split(aResultContentValue, " ")(0))

    Profile.MsgBox("HDL: " & HDLValue & ", Type: " & TypeName(HDLValue))
Else
    Profile.MsgBox("No HDL Result Found")
End If

' Get smoking status:
smoker = Patient.SmokerType

If smoker = 2 Then
smoker = True
Else
smoker = False
End If

Profile.MsgBox("Smoker Type: " & smoker & ", Type: " & TypeName(smoker))

' Get diabetes status:
Set aProblemList = Patient.ProblemList
Set aCategories = aProblemList.Categories
Dim hasDiabetesFound ' Declare hasDiabetesFound here and initialize it to False

For Each aCategory In aCategories
    If aCategory.Description = "Diagnosis" Then
        Dim aProblems ' Declare aProblems here
        Set aProblems = aCategory.Problems ' Populate aProblems here
        For Each aProblem In aProblems
            If Left(aProblem.DxCode, 3) = "250" And aProblem.Status = 1 Then
                hasDiabetesFound = True
                Exit For ' Exit the loop as soon as diabetes is found
            End If
        Next
        Exit For ' Exit the outer loop once the "Diagnosis" category is found
    End If
Next

' Set HasDiabetes based on whether diabetes was found or not
If hasDiabetesFound Then
    HasDiabetes = 1
Else
    HasDiabetes = 0
End If

Profile.MsgBox("Has Diabetes: " & HasDiabetes & ", Type: " & TypeName(HasDiabetes))

' Ask about hypertension treatment
Dim hypertensionTreatment
hypertensionTreatment = InputBox("Is the patient being treated for hypertension?" & vbCrLf & "Please enter 'Yes/No' or 'Y/N'", "Hypertension Treatment")

hypertensionTreatment = LCase(hypertensionTreatment)
If hypertensionTreatment = "yes" or hypertensionTreatment = "y" Then
hypertensionTreatment = True
Else
hypertensionTreatment = False
End If

Profile.MsgBox "Hypertension Treatment: " & hypertensionTreatment & ", Type: " & TypeName(hypertensionTreatment)



' Ask about family history of CVD
Dim familyHistoryCVD
familyHistoryCVD = InputBox("Does the patient have a positive history of premature CVD in a first-degree relative (<55 for men and <65 for women)" & vbCrLf & "Please enter 'Yes/No' or 'Y/N'", "Family History of CVD")
familyHistoryCVD = LCase(familyhistoryCVD)

If familyhistoryCVD = "yes" or familyHistoryCVD = "y" Then
familyHistoryCVD = True
Else
familyHistoryCVD = False
End If

Profile.MsgBox "Family History of CVD: " & familyHistoryCVD & ", Type: " & TypeName(familyHistoryCVD)
 
' Calculate Framingham Risk Score
' Add your Framingham Risk Score calculation logic here
' You'll need to use the variables age, gender, systolicPressure, cholesterolValue, HDLValue, smoker, HasDiabetes, 
' hypertensionTreatment, and familyHistoryCVD to calculate the risk score.

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
            If HDLValue > 1.6 Then
                riskPoints = riskPoints - 2
            ElseIf HDLValue >= 1.3 And HDLValue <= 1.6 Then
                riskPoints = riskPoints - 1
            ElseIf HDLValue >= 1.0 And HDLValue < 1.3 Then
                riskPoints = riskPoints + 0
            ElseIf HDLValue >= 0.9 And HDLValue <= 1.2 Then
                riskPoints = riskPoints + 1
            ElseIf HDLValue < 0.9 Then
                riskPoints = riskPoints + 2
            End If
        Case "F":
            If HDLValue > 1.6 Then
                riskPoints = riskPoints - 2
            ElseIf HDLValue >= 1.3 And HDLValue <= 1.6 Then
                riskPoints = riskPoints - 1
            ElseIf HDLValue >= 1.0 And HDLValue < 1.3 Then
                riskPoints = riskPoints + 0
            ElseIf HDLValue >= 0.9 And HDLValue <= 1.2 Then
                riskPoints = riskPoints + 1
            ElseIf HDLValue < 0.9 Then
                riskPoints = riskPoints + 2
            End If
    End Select

    ' Calculate additional risk points based on TC level
    Select Case gender
        Case "M":
            If cholesterolValue < 4.1 Then
                riskPoints = riskPoints + 0
            ElseIf cholesterolValue >= 4.1 And cholesterolValue <= 5.2 Then
                riskPoints = riskPoints + 1
            ElseIf cholesterolValue > 5.2 And cholesterolValue <= 6.2 Then
                riskPoints = riskPoints + 2
            ElseIf cholesterolValue > 6.2 And cholesterolValue <= 7.2 Then
                riskPoints = riskPoints + 3
            ElseIf cholesterolValue > 7.2 Then
                riskPoints = riskPoints + 4
            End If
        Case "F":
            If cholesterolValue < 4.1 Then
                riskPoints = riskPoints + 0
            ElseIf cholesterolValue >= 4.1 And cholesterolValue <= 5.2 Then
                riskPoints = riskPoints + 1
            ElseIf cholesterolValue > 5.2 And cholesterolValue <= 6.2 Then
                riskPoints = riskPoints + 3
            ElseIf cholesterolValue > 6.2 And cholesterolValue <= 7.2 Then
                riskPoints = riskPoints + 4
            ElseIf cholesterolValue > 7.2 Then
                riskPoints = riskPoints + 5
            End If
    End Select
    
       ' Calculate additional risk points based on systolicPressure and hypertension treatment status
    Select Case gender
        Case "M":
            If systolicPressure < 120 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 0 ' Treated
                Else
                    riskPoints = riskPoints - 2 ' Not treated
                End If
            ElseIf systolicPressure >= 120 And systolicPressure <= 129 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 2 ' Treated
                Else
                    riskPoints = riskPoints + 0 ' Not treated
                End If
            ElseIf systolicPressure >= 130 And systolicPressure <= 139 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 3 ' Treated
                Else
                    riskPoints = riskPoints + 1 ' Not treated
                End If
            ElseIf systolicPressure >= 140 And systolicPressure <= 149 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 4 ' Treated
                Else
                    riskPoints = riskPoints + 2 ' Not treated
                End If
            ElseIf systolicPressure >= 150 And systolicPressure <= 159 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 4 ' Treated
                Else
                    riskPoints = riskPoints + 2 ' Not treated
                End If
            ElseIf systolicPressure >= 160 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 5 ' Treated
                Else
                    riskPoints = riskPoints + 3 ' Not treated
                End If
            End If
        Case "F":
            If systolicPressure < 120 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints - 1 ' Treated
                Else
                    riskPoints = riskPoints - 3 ' Not treated
                End If
            ElseIf systolicPressure >= 120 And systolicPressure <= 129 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 2 ' Treated
                Else
                    riskPoints = riskPoints + 0 ' Not treated
                End If
            ElseIf systolicPressure >= 130 And systolicPressure <= 139 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 3 ' Treated
                Else
                    riskPoints = riskPoints + 1 ' Not treated
                End If
            ElseIf systolicPressure >= 140 And systolicPressure <= 149 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 5 ' Treated
                Else
                    riskPoints = riskPoints + 2 ' Not treated
                End If
            ElseIf systolicPressure >= 150 And systolicPressure <= 159 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 6 ' Treated
                Else
                    riskPoints = riskPoints + 4 ' Not treated
                End If
            ElseIf systolicPressure >= 160 Then
                If hypertensionTreatment Then
                    riskPoints = riskPoints + 5 ' Treated
                Else
                    riskPoints = riskPoints + 7 ' Not treated
                End If
            End If
    End Select
    
        ' Calculate additional risk points based on diabetes status
    Select Case gender
        Case "M":
            If HasDiabetes Then
                riskPoints = riskPoints + 3
            End If
        Case "F":
            If HasDiabetes Then
                riskPoints = riskPoints + 4
            End If
    End Select
    
        ' Calculate additional risk points based on smoking status
    Select Case gender
        Case "M":
            If smoker Then
                riskPoints = riskPoints + 4
            End If
        Case "F":
            If smoker Then
                riskPoints = riskPoints + 3
            End If
    End Select
 
' Calculate the total risk points
Dim totalRiskPoints
totalRiskPoints = riskPoints

' Calculate the risk percentage based on total risk points and gender
If gender = "M" Then
    If totalRiskPoints <= 3 Then
        riskPercentage = 1.0
    ElseIf totalRiskPoints = -2 Then
        riskPercentage = 1.1
    ElseIf totalRiskPoints = -1 Then
        riskPercentage = 1.4
    ElseIf totalRiskPoints = 0 Then
        riskPercentage = 1.6
    ElseIf totalRiskPoints = 1 Then
        riskPercentage = 1.9
    ElseIf totalRiskPoints = 2 Then
        riskPercentage = 2.3
    ElseIf totalRiskPoints = 3 Then
        riskPercentage = 2.8
    ElseIf totalRiskPoints = 4 Then
        riskPercentage = 3.3
    ElseIf totalRiskPoints = 5 Then
        riskPercentage = 3.9
    ElseIf totalRiskPoints = 6 Then
        riskPercentage = 4.7
    ElseIf totalRiskPoints = 7 Then
        riskPercentage = 5.6
    ElseIf totalRiskPoints = 8 Then
        riskPercentage = 6.7
    ElseIf totalRiskPoints = 9 Then
        riskPercentage = 7.9
    ElseIf totalRiskPoints = 10 Then
        riskPercentage = 9.4
    ElseIf totalRiskPoints = 11 Then
        riskPercentage = 11.2
    ElseIf totalRiskPoints = 12 Then
        riskPercentage = 13.3
    ElseIf totalRiskPoints = 13 Then
        riskPercentage = 15.6
    ElseIf totalRiskPoints = 14 Then
        riskPercentage = 18.4
    ElseIf totalRiskPoints = 15 Then
        riskPercentage = 21.6
    ElseIf totalRiskPoints = 16 Then
        riskPercentage = 25.3
    ElseIf totalRiskPoints = 17 Then
        riskPercentage = 29.3
    ElseIf totalRiskPoints >= 18 Then
        riskPercentage = 30.0
    End If
ElseIf gender = "F" Then
    If totalRiskPoints <= 3 Then
        riskPercentage = 1.0
    ElseIf totalRiskPoints = -2 Then
        riskPercentage = 1.0
    ElseIf totalRiskPoints = -1 Then
        riskPercentage = 1.2
    ElseIf totalRiskPoints = 0 Then
        riskPercentage = 1.5
    ElseIf totalRiskPoints = 1 Then
        riskPercentage = 1.7
    ElseIf totalRiskPoints = 2 Then
        riskPercentage = 2.0
    ElseIf totalRiskPoints = 3 Then
        riskPercentage = 2.4
    ElseIf totalRiskPoints = 4 Then
        riskPercentage = 2.8
    ElseIf totalRiskPoints = 5 Then
        riskPercentage = 3.3
    ElseIf totalRiskPoints = 6 Then
        riskPercentage = 3.9
    ElseIf totalRiskPoints = 7 Then
        riskPercentage = 4.5
    ElseIf totalRiskPoints = 8 Then
        riskPercentage = 5.3
    ElseIf totalRiskPoints = 9 Then
        riskPercentage = 6.3
    ElseIf totalRiskPoints = 10 Then
        riskPercentage = 7.3
    ElseIf totalRiskPoints = 11 Then
        riskPercentage = 8.6
    ElseIf totalRiskPoints = 12 Then
        riskPercentage = 10.0
    ElseIf totalRiskPoints = 13 Then
        riskPercentage = 11.7
    ElseIf totalRiskPoints = 14 Then
        riskPercentage = 13.7
    ElseIf totalRiskPoints = 15 Then
        riskPercentage = 15.9
    ElseIf totalRiskPoints = 16 Then
        riskPercentage = 18.5
    ElseIf totalRiskPoints = 17 Then
        riskPercentage = 21.5
    ElseIf totalRiskPoints = 18 Then
        riskPercentage = 24.8
    ElseIf totalRiskPoints = 19 Then
        riskPercentage = 27.5
    ElseIf totalRiskPoints >= 20 Then
        riskPercentage = 30.0
    End If
End If

' Check if the patient has a family history of premature CVD
If familyHistoryCVD Then
    ' If true, multiply the riskPercentage by 2
    riskPercentage = riskPercentage * 2
End If

' Determine Framingham Risk Score based on riskPercentage
If riskPercentage < 10.0 Then
    riskScore = "Low"
ElseIf riskPercentage >= 10.0 And riskPercentage < 20.0 Then
    riskScore = "Intermediate"
Else
    riskScore = "High"
End If

' Display the calculated Framingham Risk Score
Profile.MsgBox "Framingham Risk Score: " & riskScore, vbInformation, "Risk Assessment"