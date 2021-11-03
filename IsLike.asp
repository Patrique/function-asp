<%
 '--- IsLike returns True to indicate a pattern match.
response.write IsLike("Livio Siri", "[A-Z]\D\D\D\D\s[A-Z]\D\D\D")

 '--- IsLike returns False to indicate no pattern match.
response.write IsLike("Livio Siri", "B[^i]")
%>

<%
Private Function IsLike(byVal String, byVal Pattern)
    Dim a, b, boolErr
    On Error Resume Next
    Set a = New RegExp
    If Err Then
        boolErr = True
    End If
    On Error GoTo 0
    If boolErr then 
        Err.Raise 5108, "IsLike Function", "This function uses " & _
            "the RegExp object and requires the VBScript " & _
            "Scripting Engine Version 5.1 or higher."
        IsLike = Null
        Exit Function
    End If
    a.Pattern = pattern
    a.IgnoreCase = false
    a.Global = true
    Set b = a.Execute(String)
    if b.Count > 0 then 
        IsLike = True 
    else 
        IsLike = False
    end if
    Set b = nothing
    Set a = Nothing
End Function
%>
