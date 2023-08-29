Attribute VB_Name = "GenerateUniqueNumber"
Function GenerateUniqueNumber() As String
    Dim currentDateTime As Double
    Dim dayValue As Integer
    Dim hourValue As Integer
    Dim minuteValue As Integer
    Dim secondValue As Integer

    currentDateTime = Now
    dayValue = day(currentDateTime)
    hourValue = hour(currentDateTime)
    minuteValue = minute(currentDateTime)
    secondValue = second(currentDateTime)

    GenerateUniqueNumber = CStr(9) & Format(dayValue, "00") & Format(hourValue, "00") & Format(minuteValue, "00") & Format(secondValue, "00")
End Function

