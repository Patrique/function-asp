<%
Private Function browserDetect()

	Dim strUserAgent	'Holds info on the users browser
	Dim strMSIEversion

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	'MSIE
	If InStr(1, strUserAgent, "MSIE", 1) AND InStr(1, strUserAgent, "Opera", 1) = 0 Then
		
		'Get the MSIE version
		strMSIEversion = Trim(Mid(strUserAgent, inStr(1, strUserAgent, "MSIE", 1)+5, 2))
		
		'Remove any dots from verion number to prevent problesm with locales that use commas
		strMSIEversion = Replace(strMSIEversion, ".", "")
		
		
		'Check that we are dealing with a numeric number
		If isNumeric(strMSIEversion) Then
			'MSIE 6 or below
			If  CInt(strMSIEversion) <= 6 Then
				browserDetect = "MSIE6-"
			Else
				browserDetect = "MSIE"
			End If
		Else
			browserDetect = "MSIE"
		End If
	
	
	'Trident is IE's new redering engine from IE 8 onward - now function like Gecko
	ElseIf InStr(1, strUserAgent, "Trident", 1) Then
		
		browserDetect = "MSIE"

	'Gekco
	ElseIf inStr(1, strUserAgent, "Gecko", 1) Then
		browserDetect = "Gecko"

	'Opera
	ElseIf inStr(1, strUserAgent, "Opera", 1) Then
		browserDetect = "opera"
		
	'Others
	Else
		browserDetect = "na"
	End If

End Function
%>
