<%
Private Function internationalDateTime(dtmDate)

	Dim strYear
	Dim strMonth
	Dim strDay
	Dim strHour
	Dim strMinute
	Dim strSecound

	strYear = Year(dtmDate)
	strMonth = Month(dtmDate)
	strDay = Day(dtmDate)
	strHour = Hour(dtmDate)
	strMinute = Minute(dtmDate)
	strSecound = Second(dtmDate)

	'Place 0 infront of minutes under 10
	If strMonth < 10 then strMonth = "0" & strMonth
	If strDay < 10 then strDay = "0" & strDay
	If strHour < 10 then strHour = "0" & strHour
	If strMinute < 10 then strMinute = "0" & strMinute
	If strSecound < 10 then strSecound = "0" & strSecound

	'This function returns the ISO internation date and time formats:- yyyy-mm-dd hh:mm:ss
	'Dashes prevent systems that use periods etc. from crashing
	internationalDateTime = strYear & "-" & strMonth & "-" & strDay & " " & strHour & ":" & strMinute& ":" & strSecound
End Function
%>
