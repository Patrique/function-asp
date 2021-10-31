<%
'CBool Handling True and False
Private Function BoolC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to a Booleon
	BoolC = CBool(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("Forum Expression handling error; The data being converted is not a valid Boolean Subtype.", "BoolC()", "functions_common.asp")
	
	'Disable error trapping
	On Error goto 0

End Function
%>
