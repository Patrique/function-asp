<%
Private Function getIP()

	Dim strIPAddr

	'If they are not going through a proxy get the IP address
	If Request.ServerVariables("HTTP_X_FORWARDED_FOR") = "" OR InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), "unknown") > 0 Then

		strIPAddr = Request.ServerVariables("REMOTE_ADDR")

	'If they are going through multiple proxy servers only get the fisrt IP address in the list (,)
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",") > 0 Then

		strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ",")-1)

	'If they are going through multiple proxy servers only get the fisrt IP address in the list (;)
	ElseIf InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";") > 0 Then

		strIPAddr = Mid(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), 1, InStr(Request.ServerVariables("HTTP_X_FORWARDED_FOR"), ";")-1)

	'Get the browsers IP address not the proxy servers IP
	Else
		strIPAddr = Request.ServerVariables("HTTP_X_FORWARDED_FOR")
	End If

	'Remove all tags in IP string
	strIPAddr =  removeAllTags(strIPAddr)


	'Place the IP address back into the returning function
	getIP = Trim(Mid(strIPAddr, 1, 50))
End Function
%>
