<%
'CDbl Handling Floating Point Numbers
Private Function DblC(strExpression) 

	'Set error trapping
	On Error Resume Next
	
	'Converts the string data to a Double Foloatinmg Point number
	DblC = CDbl(strExpression)
	
	'If an error has occurred write an error to the page
	If Err.Number <> 0 Then	Call errorMsg("Forum Number handling error; The data being converted is not a valid Floating Point Number.", "DblC()", "functions_common.asp")
	
	'Disable error trapping
	On Error goto 0

End Function
%>
