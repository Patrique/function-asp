<%
'isBool checks if a Boolean value
Private Function isBool(strExpression) 

	'Convert to lower case string (less work to do later)
	strExpression = CStr(LCase(strExpression))
	
	'See if value is a booleon or not
	Select Case strExpression
		Case "true", "false", "1", "0", "-1"
			isBool = True
		Case Else
			isBool = False
	End Select

End Function
%>
