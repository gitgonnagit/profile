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
hypertensionTreatment = InputBox("Is the patient being treated for hypertension?" & vbCrLf & "Please enter 'Yes' or 'No'", "Hypertension Treatment")

Profile.MsgBox "Hypertension Treatment: " & hypertensionTreatment & ", Type: " & TypeName(hypertensionTreatment)

' Ask about family history of CVD
Dim familyHistoryCVD
familyHistoryCVD = InputBox("Does the patient have a family history of cardiovascular disease (CVD)?" & vbCrLf & "Please enter 'Yes' or 'No'", "Family History of CVD")

Profile.MsgBox "Family History of CVD: " & familyHistoryCVD & ", Type: " & TypeName(familyHistoryCVD)


' Calculate Framingham Risk Score
' Add your Framingham Risk Score calculation logic here
' You'll need to use the variables age, gender, systolicPressure, cholesterolValue, HDLValue, smoker, HasDiabetes, 
' hypertensionTreatment, and familyHistoryCVD to calculate the risk score.

' Display the calculated Framingham Risk Score
' Profile.MsgBox("Framingham Risk Score: " & YourCalculatedRiskScore)