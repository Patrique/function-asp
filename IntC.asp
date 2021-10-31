<%
'CInt Handling Integers to 32,768
Private Function IntC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Integer Number
	IntC = CInt(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("Forum Number handling error; The data being converted is not within range; -32,768 to 32,768.", "IntC()", "functions_common.asp")
	
	'Disable error trapping
	On Error goto 0

End Function
%>
