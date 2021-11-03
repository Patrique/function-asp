<%
Function StripASPCodeRegExp(ByVal strString)
    Dim loRegExp    '--- Regular Expression Object (Requires vbScript 5.1 or higher)
    Dim boolErr
    '--- If no content was received, exit the function
    if VarType(strString) = vbNull Then Exit function
    if strString = "" Then Exit function
    On Error Resume Next
    '--- Create Regular Expression object
    Set loRegExp = New RegExp
    '--- Error handling
    If Err Then boolErr = True End If
    On Error GoTo 0
    If boolErr then 
        Err.Raise 5108, "StripASPCodeRegExp Function", "This function uses " & _
            "the RegExp object and requires the VBScript " & _
            "Scripting Engine Version 5.1 or higher."
        StripASPCodeRegExp = Null
        Exit Function
    End If
    With loRegExp
        '--- Keep finding links after the first one.                
        .Global = True
        '--- Ignore upper/lower case
        .IgnoreCase = True
        '--- Look for Script pattern
        .Pattern = "<%[\s\S]*?" & Chr(37) & ">"
    End With
    strString = loRegExp.Replace(strString, "!-ASP-REMOVED-!")
    '--- Release regular expression object
    Set loRegExp = Nothing
    StripASPCodeRegExp = strString
End Function
%>

<%
Dim a
a = StripASPCodeRegExp(strString)
%>
