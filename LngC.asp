<%
'CLng Handling Integers to 2,147,483,648
Private Function LngC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to an Lnog Number
	LngC = CLng(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("Forum Number handling error; The data being converted is not within range; -2,147,483,648 to 2,147,483,648.", "LngC()", "functions_common.asp")
	
	'Disable error trapping
	On Error goto 0

End Function
%>
