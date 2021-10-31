<%
'Function to chop down the length of a string and add '...'
Private Function TrimString(strInputString, intStringLength)

	Dim lngTrimLength

	'Trim the string down
	strInputString = Trim(strInputString) & " "

	'If the length of the text is longer than the max then cut it and place '...' at the end
	If CLng(Len(strInputString)) > intStringLength Then
		
		'Get the part in the string to trim it from
		lngTrimLength = InStr(intStringLength, strInputString, " ", vbTextCompare)
		
		'If lngTrimLength = 0 then set it to the default passed to the function (Error handling, should never be used)
		If lngTrimLength = 0 Then lngTrimLength = intStringLength
		
		'Trim the number of characters down to the required amount, but try not to chop words in half
		strInputString = Mid(strInputString, 1, lngTrimLength)

		'Make sure the user hasn't entered a long line of text with no break (most words won't be over 30 chars
		If CLng(Len(strInputString)) => intStringLength + 30 Then
			strInputString = Mid(Trim(strInputString), 1, intStringLength)
		End If

		'Place '...' at the end
		 strInputString = Trim(strInputString) & "..."
	End If

	'Return string
	TrimString = strInputString
End Function
%>
