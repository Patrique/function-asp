<%
'require password complexity
Private Function passwordComplexity(ByRef strPassword, ByRef intMinPasswordLength)

	Dim objRegExp	'Holds regulare expresions object
	
	'Remove tabs and spaces
	strPassword = Replace(strPassword, " ", "")
	strPassword = Replace(strPassword, VbTab, "")

	'Create regular experssions object
	Set objRegExp = New RegExp

	'Tell the regular experssions object to look for 1 number, 1 lower case, 1 upper case, and 1 Non-Alphanumeric Symbol
	With objRegExp
		.Pattern = "^.*(?=.{" & intMinPasswordLength & ",})(?=.*\d)(?=.*[a-z])(?=.*[A-Z])(?=.*[\W]).*$"
		.IgnoreCase = False
		.Global = True
	End With
	
	'See if password is up to the job
	passwordComplexity = objRegExp.Test(strPassword)
	
	Set objRegExp = nothing


End Function
%>
