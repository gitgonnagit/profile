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

'Get blood pressure: // returns SBP as an integer 

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

    ' Get the variable type 
    systolicPressureType = TypeName(systolicPressure) 

    ' Display the systolic blood pressure (as an integer) and its variable type in a pop-up box 
    Profile.MsgBox("Latest Systolic Pressure: " & systolicPressure & ", Type: " & systolicPressureType) 
End If 

'Get total cholesterol – returns value as a double: 

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

    ' Get the variable type of cholesterolValue 
    cholesterolType = TypeName(cholesterolValue) 

    ' Display the extracted Total Cholesterol value and its variable type in a pop-up box 
    Profile.MsgBox("Total Cholesterol: " & cholesterolValue & vbNewLine & "Variable Type: " & cholesterolType) 
Else 
    Profile.MsgBox("No Total Cholesterol Result Found") 
End If 

'Get HDL cholesterol – returns value as a double: 
aDateTo = Now 
aDateFrom = DateAdd("yyyy", -30, aDateTo) 
aDateTo = DateAdd("yyyy", 1, aDateTo) 

Set aResult = Patient.GetLatestHRI("FHAM", "14646-4", aDateFrom, aDateTo, 1) 
If Not aResult Is Nothing Then 
    Set aResultContent = aResult.Content 
    aResultContentValue = aResultContent.AsString 

    ' Extract only the numerical value from the string 
    ' Assuming the value is the first part of the string 
    cholesterolValue = CDbl(Split(aResultContentValue, " ")(0)) 

    ' Get the variable type of cholesterolValue 
    cholesterolType = TypeName(cholesterolValue) 

    ' Display the extracted Total Cholesterol value and its variable type in a pop-up box 
    Profile.MsgBox("HDL: " & cholesterolValue & vbNewLine & "Variable Type: " & cholesterolType) 
Else 
    Profile.MsgBox("No HDL Result Found") 
End If 

 'Get smoking status – returns smoking type as a LONG: 
' Retrieve smoking Type  
smoker = Patient.SmokerType  

' Get the variable type  
    smokerType = TypeName(smoker) 

If IsEmpty(smoker) Then  
    gender = "Smoking data not found"  
Else 
    Profile.MsgBox (smoker & vbNewLine & "Variable Type: " & smokerType)    
End If  

'spstUnknown	0	  
'spstNonSmoker	1	  
'spstSmoker	2	  
'spstExSmoker	3	  
'spstPassiveSmoker	4 


'Get diabetes status: 

'Get patient problem list 
Set aProblemList = Patient.ProblemList 

'Get list of problem list categories 
Set aCategories = aProblemList.Categories 

'Loop through all categories to look under "Diagnosis" 
For Each aCategory In aCategories 
    ' Check if the category description is "Diagnosis" 
    If aCategory.Description = "Diagnosis" Then 
        aMessage = "Active Problems with Codes Starting with 250:" & vbNewLine & vbNewLine 
        Set aProblems = aCategory.Problems 
        For Each aProblem In aProblems 

            ' Check if the problem code starts with "706" and is active (Status = 1) 
            If Left(aProblem.DxCode, 3) = "250" And aProblem.Status = 1 Then 
                aMessage = aMessage & "Code: " & aProblem.DxCode & "; Status: " & aProblem.Status & vbNewLine 
            End If 
        Next 
        Exit For ' Exit the loop once the "Diagnosis" category is found 
    End If 
Next 

Profile.MsgBox aMessage 