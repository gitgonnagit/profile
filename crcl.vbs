' Calculate age
dobString = Patient.DOB
If Not IsEmpty(dobString) Then
    dob = CDate(dobString)
    age = DateDiff("yyyy", dob, Now)
Else
    age = "Date of Birth data not found"
End If

' Retrieve gender
gender = Patient.Sex
If IsEmpty(gender) Then
    gender = "Gender data not found"
End If

'GET PATIENT WEIGHT
Dim aPatient
Dim aTermsetCode
Dim aConceptCode
Dim aDateFrom
Dim aDateTo
Dim aUseCrossRefs
Dim aLatestHRI

Set aPatient = Patient
aConceptCode = "z..2T"
aTermsetCode = "IH"
aDateFrom = #01/01/2000#
aDateTo = Now ' Set the end date to the current date for the most up-to-date weight
aUseCrossRefs = False

Set aLatestHRI = aPatient.GetLatestHRI(aTermsetCode, aConceptCode, aDateFrom, aDateTo, aUseCrossRefs)

If aLatestHRI Is Nothing Then
    Variable("PrintVariable").Value = "There is no latest weight data available."
Else
    ' Retrieve the weight as a string
    weightStr = aLatestHRI.AsString
    
    ' Extract just the numerical value (assuming it's the first part of the string)
    weightValue = CDbl(Split(weightStr, " ")(1)) ' Convert to numeric
End If


' Retrieve creatinine and extract numerical value
aDateTo = Now
aDateFrom = DateAdd("yyyy", -30, aDateTo)
aDateTo = DateAdd("yyyy", 1, aDateTo)

Set aResult = Patient.GetLatestHRI("IH", "z.SZz", aDateFrom, aDateTo, 1)
If Not aResult Is Nothing Then
    Set aResultContent = aResult.Content
    aResultContentValue = aResultContent.AsString
    theResult3 = CDbl(Split(aResultContentValue, " ")(0)) ' Convert to numeric
Else
    theResult3 = "No Creatinine Result Found"
End If

' Debug statements
'Profile.MsgBox("Age: " & age & ", Type: " & TypeName(age))
'Profile.MsgBox("Weight: " & weightValue & ", Type: " & TypeName(weightValue))
'Profile.MsgBox("Creatinine: " & theResult3 & ", Type: " & TypeName(theResult3))
'Profile.MsgBox("Gender: " & gender)

' Calculate creatinine clearance based on the metric formula
If IsNumeric(age) And IsNumeric(weightValue) And IsNumeric(theResult3) Then
    If gender = "M" Then
        creatinineClearance = (1.2 * (140 - age) * CDbl(weightValue)) / (CDbl(theResult3))
    ElseIf gender = "F" Then
        creatinineClearance = (1.2 * (140 - age) * CDbl(weightValue)) / (CDbl(theResult3)) * 0.85
    Else
        creatinineClearance = "Unknown Gender"
    End If
    
    ' Convert creatinine clearance to a string with one decimal place
    creatinineClearanceStr = CStr(FormatNumber(creatinineClearance, 1))
Else
    creatinineClearanceStr = "Unable to Calculate"
End If

' Store the calculated creatinine clearance as a string in PrintVariable
Variable("PrintVariable").Value = creatinineClearanceStr



