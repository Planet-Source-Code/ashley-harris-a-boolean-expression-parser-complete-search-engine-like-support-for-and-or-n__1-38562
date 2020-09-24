Attribute VB_Name = "SearchMod"
'search function that allows for the following conventions:
' x and y : returns true when both strings x and y are in the string searched
' x or y  : returns true when either x or y, or both, are in the string searched
' not x   : returns true when x is not in the document
' x xor y : returns true when x or y is in the document, but false if they both are
' x and (y or z) : just what you'd expect
' x and y or z   : same ('or' is parsed before 'and')
' x or y or z    : matches any of the x, y, or z
' (x and y) or z : just what you'd expect
'etc. (you can have nested brackets etc)
'
'copyright Ashley Harris (ashley___harris@hotmail.com)
'You can use this module freely in your own apps/asp scripts.
'so long as I get some credit somewhere, and, that I know about it!
'
'    I mean, you can distribute this in any app you want, and make
'    obscene amounts of money from it. I just, would like to know
'    about it!
'
'Also, if you use this, I require a vote on PSC, a preaty good deal
'if you ask me. (I hide out in the perl, vb, javascript, and asp sections of the site)
'
'Also, if you make an improvement on this script, please let me know what you did
'
'currently, on a 300Mhz Machine, it can do one 'check' in about .25 miliseconds.
'depending on the querystring.
'
'Get in touch with me, if you want to (I don't bite):
'Ashley Harris
'Email:  Ashley___harris@ hotmail.com
'MSN:    Ashley___harris@ hotmail.com
'ICQ   : 153577070
'AIM:    Ashley000Harris
'Y!M   : a_s_h_l_e_y_h_a_r_r_i_s
'
'also, check out my webserver(s) on PSC (in vb section), my mailserver on PSC (vb), and
'my beginers guide to perl on PSC.

Public Function checkstring(ByVal query As String, test As String) As Boolean
    'query is something like "Ashley and Harris and not Maxwell"
    'test is the string in which to search, ie "Ashley Maxwell Harris"
    ' (this example would not match, as maxwell was present)

    If query = "" Then Exit Function
    'query is a string you'd pass to a search engine like altavista etc. ie:
    ''animal AND (bear OR "Baby fox")'
    Dim s As Long, f As Long, a As Long, b As Long
    Dim i() As String, j() As String
    
    On Error GoTo 0

    s = InStrRev(query, "(")
morebrackets:
    If s > 0 Then
        f = InStr(s, query, ")")
        
        If f > 0 Then
            query = Mid(query, 1, s - 1) & UCase(CBool(checkstring(Mid(query, s + 1, f - s - 1), test))) & Mid(query, f + 1)
        End If
    End If
    
    s = InStrRev(query, "(")
    If s > 0 And f > 0 Then GoTo morebrackets
    
    'remove any invalid characters
    query = Replace(query, "?", "")
    query = Replace(query, "  ", " ")
    query = Replace(query, vbCr, "")
    query = Replace(query, vbLf, "")
    query = Replace(query, vbTab, "")
    query = Replace(query, "  ", " ")
    query = Replace(query, ".", "")
    query = Replace(query, "(", "")
    query = Replace(query, ")", "")
    i = Split(query, " ")
    
    'ok, this is weird, say this split resulted in:
    ' 'Bear' 'OR' '"Baby' ' fox"'
    
    'turn it into: 'bear' 'OR' 'Baby Fox'

    
    If InStr(1, query, """") Then
    
        For a = UBound(i) To 0 Step -1
            If Right(i(a), 1) = """" And Left(i(a), 1) <> """" Then
                i(a - 1) = i(a - 1) & " " & i(a)
                i(a) = ""
            End If
            If Right(i(a), 1) = """" And Left(i(a), 1) = """" And Len(i(a)) > 1 Then
                i(a) = "!" & Mid(i(a), 2, Len(i(a)) - 2)
            End If
        Next a
        
        'remove the nulls
        b = 0
        For a = 0 To UBound(i)
            If i(a) <> "" Then
                ReDim Preserve j(b) As String
                j(b) = i(a)
                b = b + 1
            End If
        Next a
    
        i = j
    End If
    'now, do the actual searching for each string
    For a = 0 To UBound(i)
        Select Case UCase(i(a))
        Case "AND"
            i(a) = "AND"
        Case "OR"
            i(a) = "OR"
        Case "NOT"
            i(a) = "NOT"
        Case "XOR"
            i(a) = "XOR"
        Case "TRUE"
            i(a) = "true"
        Case "FALSE"
            i(a) = "false"
        Case Else
            If Left(i(a), 1) = "!" Then i(a) = Mid(i(a), 2)
            i(a) = LCase(CBool(InStr(1, test, i(a), vbTextCompare)))
        End Select
    Next a
    
    'we now have a boolean expression on our hands:
    'false AND true XOR false
    query = Join(i, " ")
    
    'this is cheating, I know, but, It's quicker then parsing out the
    'proper way, so...
keepgoing:

    If InStr(1, query, "NOT") = 0 Then GoTo door
    query = Replace(query, "NOT false", "true")
    query = Replace(query, "NOT true", "false")
    
door:
    If InStr(1, query, "OR") = 0 Then GoTo doxor
    query = Replace(query, "false OR false", "false")
    query = Replace(query, "true OR false", "true")
    query = Replace(query, "false OR true", "true")
    query = Replace(query, "true OR true", "true")
    
doxor:
    If InStr(1, query, "XOR") = 0 Then GoTo doand
    query = Replace(query, "false XOR false", "false")
    query = Replace(query, "true XOR false", "true")
    query = Replace(query, "false XOR true", "true")
    query = Replace(query, "true XOR true", "false")
    
doand:
    If InStr(1, query, "AND") = "0" Then GoTo endbit
    query = Replace(query, "false AND false", "false")
    query = Replace(query, "true AND false", "false")
    query = Replace(query, "false AND true", "false")
    query = Replace(query, "true AND true", "true")
    
endbit:
    If InStr(1, query, " ") Then
        query = Replace(query, "true true", "true")
        query = Replace(query, "false true", "true")
        query = Replace(query, "true false", "true")
        query = Replace(query, "false false", "false")
        
        
        If InStr(1, query, " ") Then query = query & " false": GoTo keepgoing
    End If
    
    checkstring = CBool(query)
End Function
