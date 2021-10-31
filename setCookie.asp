<%
'Functions and subs for handling cookies

'Set Cookie
Sub setCookie(strCookieName, strCookieKey, strValue, blnStore)
    	'Write Cookie
	If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & strCookieName).Domain = strCookieDomain
	Response.Cookies(strCookiePrefix & strCookieName).Path = strCookiePath
	Response.Cookies(strCookiePrefix & strCookieName)(strCookieKey) = strValue
	
	If blnStore Then
		Response.Cookies(strCookiePrefix & strCookieName).Expires = DateAdd("yyyy", 1, Now())
	End If
End Sub


'Get Cookie
Function getCookie(strCookieName, strCookieKey)  
	'Read in the cookie
	getCookie = Request.Cookies(strCookiePrefix & strCookieName)(strCookieKey)
End Function


'Clear Cookie
Sub clearCookie()  
	'Clear the cookie
	
	If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & "sID").Domain
	Response.Cookies(strCookiePrefix & "sID").Path = strCookiePath
	Response.Cookies(strCookiePrefix & "sID") = ""
	
	If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & "sLID").Domain
	Response.Cookies(strCookiePrefix & "sLID").Path = strCookiePath
	Response.Cookies(strCookiePrefix & "sLID") = ""
	
	If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & "lVisit").Domain
	Response.Cookies(strCookiePrefix & "lVisit").Path = strCookiePath
	Response.Cookies(strCookiePrefix & "lVisit") = ""
	
	If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & "fID").Domain
	Response.Cookies(strCookiePrefix & "fID").Path = strCookiePath
	Response.Cookies(strCookiePrefix & "fID") = ""
	
	If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & "MobileView").Domain
	Response.Cookies(strCookiePrefix & "MobileView") = ""
	Response.Cookies(strCookiePrefix & "MobileView").Path = strCookiePath
	
	'This one stops user voting in polls so doesn't really want to be cleared
	'If strCookieDomain <> "" Then Response.Cookies(strCookiePrefix & "pID").Domain
	'Response.Cookies(strCookiePrefix & "pID").Path = strCookiePath
	'Response.Cookies(strCookiePrefix & "pID") = ""
	
End Sub
%>
