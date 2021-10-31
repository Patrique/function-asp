<%
'for URL rewrite search engine friendly page titles
Private Function SeoUrlTitle(ByVal strInputEntry, strPrefix)

	Dim intLoopCounter
	Dim objRegExp

	If blnSeoTitleQueryStrings = False Then Exit Function
	
	'Swap to lower case
	strInputEntry = strInputEntry
	
	'Remove any HTML encoding
	strInputEntry = decodeString(strInputEntry)

	strInputEntry = Replace(strInputEntry, "_", " ", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, ".", " ", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "/", " ", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "+", " ", 1, -1, 1)
	
	
	'Create regular experssions object
	Set objRegExp = New RegExp

	'Tell the regular experssions object to look for tags <xxxx>
	With objRegExp
		.Pattern = "[^\w\d\s]"  'same as [^a-zA-Z0-9 ]
		.IgnoreCase = True
		.Global = True
	End With

	'Strip HTML
	strInputEntry = objRegExp.Replace(strInputEntry, "")

	'Distroy regular experssions object
	Set objRegExp = nothing
	
	
	'Trim the final result
	strInputEntry = Trim(strInputEntry)
	
	'Replace spaces with hyphans
	strInputEntry = Replace(strInputEntry, " ", "-", 1, -1, 1)
	
	'Replace double hyphens
	strInputEntry = Replace(strInputEntry, "---", "-", 1, -1, 1)
	strInputEntry = Replace(strInputEntry, "--", "-", 1, -1, 1)
	
	
	'Return result (if any)
	If strInputEntry = "" Then 
		SeoUrlTitle = ""
	Else
		SeoUrlTitle = LCase(strPrefix & strInputEntry)
	End If

End Function
%>
