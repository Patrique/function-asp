<%
Private Function mobileBrowser()

	Dim strUserAgent	'Holds info on the users browser

	'Get the users HTTP user agent (web browser)
	strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

	Select Case True
		'Mobile Plateforms
		Case InStr(1, strUserAgent, "Smartphone", 1) > 0, _
			 inStr(1, strUserAgent, "mobile", 1) > 0, _
			 inStr(1, strUserAgent, "portable", 1) > 0, _
			 inStr(1, strUserAgent, "Android", 1) > 0, _
			 inStr(1, strUserAgent, "iPad", 1) > 0, _
			 inStr(1, strUserAgent, "iPod", 1) > 0, _
			 inStr(1, strUserAgent, "iPhone", 1) > 0, _
			 inStr(1, strUserAgent, "Windows CE", 1) > 0, _
			 inStr(1, strUserAgent, "WAP", 1) > 0, _
			 inStr(1, strUserAgent, "Windows Phone", 1) > 0
			 
			 mobileBrowser = true
			 
		'Mobile manufactures	 
		Case inStr(1, strUserAgent, "Blackberry", 1) > 0, _
			 inStr(1, strUserAgent, "Samsung", 1) > 0, _
			 inStr(1, strUserAgent, "Nokia", 1) > 0, _
			 inStr(1, strUserAgent, "Palm", 1) > 0, _
			 inStr(1, strUserAgent, "RIM", 1) > 0, _
			 inStr(1, strUserAgent, "LG", 1) > 0, _
			 inStr(1, strUserAgent, "alcatel", 1) > 0, _
			 inStr(1, strUserAgent, "ericsson", 1) > 0, _
			 inStr(1, strUserAgent, "nokia", 1) > 0, _
			 inStr(1, strUserAgent, "panasonic", 1) > 0, _
			 inStr(1, strUserAgent, "sanyo", 1) > 0, _
			 inStr(1, strUserAgent, "philips", 1) > 0
			 
			 mobileBrowser = true
		
		'Mobile Browsers
		Case InStr(1, strUserAgent, "Opera Mini", 1) > 0, _
			inStr(1, strUserAgent, "Mobile Safari", 1) 
			
			mobileBrowser = true
			
		'Mobile Search Bots
		Case InStr(1, strUserAgent, "Googlebot-Mobile", 1) > 0, _
			inStr(1, strUserAgent, "YahooSeeker/M1A1-R2D2", 1) 	
			
			mobileBrowser = true
			
		'Windows Mobile Edge
		Case InStr(1, strUserAgent, "Windows NT 10", 1) > 0 AND inStr(1, strUserAgent, "ARM", 1) > 0
			
			mobileBrowser = true
			
		Case Else
			mobileBrowser = false
			'mobileBrowser = true  'for testing
	End Select

End Function
%>
