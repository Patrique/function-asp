<%
'ucfirst(string)	Capitaliza a primeira letra, como no PHP
Function ucFirst(strWord)
		strWord = trim( strWord & "" )
		if len( strWord ) > 0 then
			ucFirst = uCase(left(strWord,1)) & lcase( right( strWord, len( strWord ) - 1 ) )
		else
			ucFirst = strWord
		end if
	End Function
%>
