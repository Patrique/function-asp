<%
'ucWords(string)	Capitaliza a primeira letra de todas as palavras, como no PHP
Function ucWords( strWords )
		strWords = trim( strWords & "" )
		if len( strWords ) > 0 then
			dim arWords, i, strFormatted
			arWords = split( strWords, " " )
			for i = 0 to uBound( arWords )
				strFormatted = strFormatted & " " & ucFirst( arWords( i ) )
			next
			ucWords = trim(strFormatted & "")
		else
			ucWords = strWords
		end if
	End Function
%>
