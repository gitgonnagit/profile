Set aPatient = Patient
Set aFilter = Profile.CreatePatientDocumentFilter
aFilter.PatientId = aPatient.ID
aFilter.IncludeScanDoc = True
Set aDocuments = Profile.LoadPatientDocuments(aFilter)

' Initialize lists to store matching documents and dates for colonoscopies
Dim matchingColonoscopyDescriptions
Set matchingColonoscopyDescriptions = CreateObject("System.Collections.ArrayList")
Set matchingDates = CreateObject("System.Collections.ArrayList")

' Initialize lists to store matching documents and dates for operative reports
Dim matchingOperativeDescriptions
Set matchingOperativeDescriptions = CreateObject("System.Collections.ArrayList")
Dim matchingOperativeDates
Set matchingOperativeDates = CreateObject("System.Collections.ArrayList")

For i = 0 To aDocuments.Count - 1
    Set aDocument = aDocuments.Items(i)
    Dim descriptionLower
    descriptionLower = LCase(aDocument.Description)
    
    If InStr(descriptionLower, "colonoscopy") > 0 Then
        ' Store the matching description and date separately for colonoscopies
        matchingColonoscopyDescriptions.Add(aDocument.Description)
        matchingDates.Add(aDocument.Date)
    End If
    
    If InStr(descriptionLower, "operative report") > 0 Then
        ' Store the matching description and date separately for operative reports
        matchingOperativeDescriptions.Add(aDocument.Description)
        matchingOperativeDates.Add(aDocument.Date)
    End If
Next

' Find the most recent colonoscopy document
Dim maxDate
maxDate = DateValue("01/01/1900") ' Initialize with a very old date
For i = 0 To matchingDates.Count - 1
    If matchingDates(i) > maxDate Then
        maxDate = matchingDates(i)
        mostRecentColonoscopy = matchingColonoscopyDescriptions(i) & " - " & FormatDateTime(maxDate, vbShortDate)
    End If
Next

' Display a message if no colonoscopy documents are found
If mostRecentColonoscopy = "" Then
    Profile.Variable("PrintVariable").Value = "No colonoscopy documents found"
Else
    ' Set the value of PrintVariable to the most recent colonoscopy document
    Profile.Variable("PrintVariable").Value = mostRecentColonoscopy
End If

' Build a message for all matching operative reports in chronological order
Dim resultMessage
resultMessage = "Matching Operative Reports and Colonoscopy Reports:" & vbCrLf & vbCrLf

' Sort matching operative reports by date in chronological order
Dim sortedMatchingOperativeReports
Set sortedMatchingOperativeReports = CreateObject("System.Collections.ArrayList")

For i = 0 To matchingOperativeDescriptions.Count - 1
    ' Pair each operative report description with its date
    sortedMatchingOperativeReports.Add(matchingOperativeDescriptions(i) & " - " & FormatDateTime(matchingOperativeDates(i), vbShortDate))
Next

' Sort the list by date
sortedMatchingOperativeReports.Sort
sortedMatchingOperativeReports.Reverse


For i = 0 To sortedMatchingOperativeReports.Count - 1
    ' Add the matching operative report to the result message
    resultMessage = resultMessage & sortedMatchingOperativeReports(i) & vbCrLf
Next

' Display the pop-up message with all matching reports
Profile.MsgBox resultMessage
