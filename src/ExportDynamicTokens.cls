Option Explicit

' Class Module (Name) = ExportDynamicTokens. No document/UI/registry access.
' Dynamic operators are read longest first: /*an, /*a, /*n, then plain /*.
Public Function TokenKind(ByVal source As String, ByVal at As Long) As String
    If Mid$(source, at, 4) = "/*an" Then
        TokenKind = "an"
    ElseIf Mid$(source, at, 3) = "/*a" Then
        TokenKind = "a"
    ElseIf Mid$(source, at, 3) = "/*n" Then
        TokenKind = "n"
    End If
End Function

Public Function PlainMarkerPosition(ByVal source As String) As Long
    Dim at As Long, depth As Long, ch As String
    For at = 1 To Len(source)
        ch = Mid$(source, at, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then depth = depth - 1
        If depth = 0 And Mid$(source, at, 2) = "/*" Then
            If Len(TokenKind(source, at)) = 0 Then
                PlainMarkerPosition = at
                Exit Function
            End If
        End If
    Next at
End Function

Public Function HasDynamic(ByVal source As String) As Boolean
    Dim at As Long, depth As Long, ch As String
    For at = 1 To Len(source)
        ch = Mid$(source, at, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then depth = depth - 1
        If depth = 0 And Len(TokenKind(source, at)) > 0 Then HasDynamic = True: Exit Function
    Next at
End Function

' Expand only output values, never replacement targets. Also validates all seeds/caps.
Public Function Expand(ByVal source As String, ByVal offset As Long, ByRef limit As Long) As String
    Dim at As Long, finish As Long, start As Long, depth As Long
    Dim kind As String, seed As String, ch As String, cap As Long, result As String
    limit = 0
    at = 1
    Do While at <= Len(source)
        ch = Mid$(source, at, 1)
        If ch = "{" Then depth = depth + 1
        If ch = "}" Then depth = depth - 1
        kind = vbNullString
        If depth = 0 Then kind = TokenKind(source, at)
        If Len(kind) = 0 Then
            result = result & ch
            at = at + 1
        Else
            start = Len(result)
            Do While start > 0
                If Not IsAlphaNumeric(Mid$(result, start, 1)) Then Exit Do
                start = start - 1
            Loop
            seed = Mid$(result, start + 1)
            If Len(seed) = 0 Then Fail "Token " & "/*" & kind & " memerlukan nilai awal."
            finish = at + 2 + Len(kind)
            cap = 0
            If Mid$(source, finish, 1) = "-" Then
                at = finish + 1
                finish = at
                Do While Mid$(source, finish, 1) Like "[0-9]"
                    finish = finish + 1
                Loop
                cap = PositiveLong(Mid$(source, at, finish - at), "Batas token")
                If limit <> 0 And limit <> cap Then Fail "Batas token dalam satu blok tidak konsisten."
                limit = cap
            End If
            If IsAlphaNumeric(Mid$(source, finish, 1)) Then Fail "Akhir token dinamis tidak valid: " & source
            result = Left$(result, start) & IncrementSeed(seed, kind, offset)
            at = finish
        End If
    Loop
    Expand = result
End Function

Public Function PositiveLong(ByVal source As String, ByVal label As String) As Long
    Dim at As Long
    If Len(source) = 0 Or Len(source) > 10 Then Fail label & " harus integer positif."
    For at = 1 To Len(source)
        If Not Mid$(source, at, 1) Like "[0-9]" Then Fail label & " harus integer positif."
    Next at
    If CDbl(source) < 1 Or CDbl(source) > 2147483647# Then Fail label & " di luar batas integer positif."
    PositiveLong = CLng(source)
End Function

Private Function IncrementSeed(ByVal seed As String, ByVal kind As String, ByVal offset As Long) As String
    Dim at As Long, prefix As String, suffix As String
    Dim digit As Long, carry As Long, letterBase As Long, index As Long
    at = Len(seed)
    If kind = "a" Then
        Do While at > 0
            If Not Mid$(seed, at, 1) Like "[A-Za-z]" Then Exit Do
            at = at - 1
        Loop
        suffix = Mid$(seed, at + 1)
        If Len(suffix) = 0 Then Fail "/*a memerlukan akhiran huruf: " & seed
        If suffix <> UCase$(suffix) And suffix <> LCase$(suffix) Then Fail "Gunakan akhiran huruf seluruhnya besar atau kecil: " & seed
        prefix = Left$(seed, at)
        letterBase = 65
        If suffix = LCase$(suffix) Then letterBase = 97
        carry = offset
        index = Len(suffix)
        Do While carry > 0
            If index > 0 Then
                digit = Asc(Mid$(suffix, index, 1)) - letterBase + carry
                Mid$(suffix, index, 1) = Chr$(letterBase + (digit Mod 26))
                carry = digit \ 26
                index = index - 1
            Else
                digit = carry - 1
                suffix = Chr$(letterBase + (digit Mod 26)) & suffix
                carry = digit \ 26
            End If
        Loop
        IncrementSeed = prefix & suffix
    Else
        Do While at > 0
            If Not Mid$(seed, at, 1) Like "[0-9]" Then Exit Do
            at = at - 1
        Loop
        suffix = Mid$(seed, at + 1)
        If Len(suffix) = 0 Then Fail "/*" & kind & " memerlukan akhiran angka: " & seed
        If kind = "n" And at <> 0 Then Fail "/*n hanya menerima nilai awal angka: " & seed
        ' Increment digit-by-digit: preserve padding and avoid Long/Double precision loss.
        carry = offset
        For index = Len(suffix) To 1 Step -1
            digit = Asc(Mid$(suffix, index, 1)) - 48 + (carry Mod 10)
            carry = carry \ 10
            If digit > 9 Then digit = digit - 10: carry = carry + 1
            Mid$(suffix, index, 1) = Chr$(48 + digit)
        Next index
        If carry > 0 Then suffix = CStr(carry) & suffix
        IncrementSeed = Left$(seed, at) & suffix
    End If
End Function

' All matches are replaced in the original source, without rematching generated text.
Public Function ReplaceTarget(ByVal source As String, ByVal target As String, ByVal value As String) As String
    Dim expression As Object, matches As Object, match As Object
    Dim pattern As String, at As Long, kind As String, literal As String
    Dim beforeValue As String, afterValue As String, replacement As String, result As String
    Dim cursor As Long, wildcardCount As Long, prefixPattern As String
    Dim captures As Object, capture As Object
    Dim special As Boolean
    special = HasDynamic(target) Or InStr(target, "&*") > 0 Or InStr(value, "{~}") > 0
    If Not special Then
        ReplaceTarget = Replace$(source, target, value, 1, -1, vbTextCompare)
        Exit Function
    End If
    at = 1
    Do While at <= Len(target)
        kind = TokenKind(target, at)
        If Len(kind) > 0 Then
            pattern = pattern & LiteralPattern(literal)
            literal = vbNullString
            wildcardCount = wildcardCount + 1
            If wildcardCount = 1 Then prefixPattern = pattern
            Select Case kind
                Case "n": pattern = pattern & "\b[0-9]+(?![A-Za-z0-9])"
                Case "a": pattern = pattern & "\b[A-Za-z]+(?![A-Za-z0-9])"
                Case "an": pattern = pattern & "\b[A-Za-z0-9]+(?![A-Za-z0-9])"
            End Select
            at = at + 2 + Len(kind)
        Else
            literal = literal & Mid$(target, at, 1)
            at = at + 1
        End If
    Loop
    pattern = pattern & LiteralPattern(literal)
    If wildcardCount > 1 And InStr(value, "{~}") > 0 Then Fail "{~} memerlukan tepat satu wildcard pada target."
    Set expression = CreateObject("VBScript.RegExp")
    expression.Pattern = pattern
    expression.Global = True
    expression.IgnoreCase = True
    Set matches = expression.Execute(source)
    If matches.Count = 0 Then Fail "Target nama tidak ditemukan: " & target
    cursor = 1
    For Each match In matches
        replacement = value
        If InStr(value, "{~}") > 0 Then
            If wildcardCount = 1 Then
                ' Find the prefix and wildcard together to respect the full target context.
                expression.Pattern = "^(" & prefixPattern & ")(" & Mid$(pattern, Len(prefixPattern) + 1, Len(pattern) - Len(prefixPattern) - Len(LiteralPattern(literal))) & ")"
                Set captures = expression.Execute(match.Value)
                Set capture = captures(0)
                beforeValue = Trim$(CStr(capture.SubMatches(0)))
                afterValue = Trim$(Mid$(match.Value, capture.Length + 1))
            Else
                expression.Pattern = "[0-9]+"
                Set captures = expression.Execute(match.Value)
                If captures.Count = 0 Then Fail "{~} memerlukan angka atau satu wildcard dalam target."
                Set capture = captures(0)
                beforeValue = Trim$(Left$(match.Value, capture.FirstIndex))
                afterValue = Trim$(Mid$(match.Value, capture.FirstIndex + capture.Length + 1))
            End If
            replacement = PreserveValue(value, beforeValue, afterValue)
        End If
        result = result & Mid$(source, cursor, match.FirstIndex + 1 - cursor) & replacement
        cursor = match.FirstIndex + match.Length + 1
    Next match
    ReplaceTarget = result & Mid$(source, cursor)
End Function

Private Function PreserveValue(ByVal value As String, ByVal beforeValue As String, ByVal afterValue As String) As String
    Dim at As Long, result As String, hasValue As Boolean, ch As String
    at = 1
    Do While at <= Len(value)
        If Mid$(value, at, 3) = "{~}" Then
            If hasValue Then result = result & afterValue Else result = result & beforeValue
            at = at + 3
        Else
            ch = Mid$(value, at, 1)
            result = result & ch
            If ch <> " " Then hasValue = True
            at = at + 1
        End If
    Loop
    PreserveValue = result
End Function

Private Function LiteralPattern(ByVal source As String) As String
    Dim alternatives As Variant, alternative As Variant, result As String
    Dim leading As String, trailing As String
    If InStr(source, "&*") = 0 Then
        LiteralPattern = EscapePattern(source)
        Exit Function
    End If
    ' Literal alternatives share the adjacent wildcard and its surrounding spaces.
    leading = Left$(source, Len(source) - Len(LTrim$(source)))
    trailing = Right$(source, Len(source) - Len(RTrim$(source)))
    alternatives = Split(Trim$(source), "&*")
    For Each alternative In alternatives
        If Len(Trim$(CStr(alternative))) = 0 Then Fail "Alternatif target &* kosong."
        If Len(result) > 0 Then result = result & "|"
        result = result & EscapePattern(Trim$(CStr(alternative)))
    Next alternative
    LiteralPattern = EscapePattern(leading) & "(?:" & result & ")" & EscapePattern(trailing)
End Function

Private Function EscapePattern(ByVal source As String) As String
    Dim at As Long, ch As String
    For at = 1 To Len(source)
        ch = Mid$(source, at, 1)
        If InStr(1, "\.^$|?*+()[]{}", ch, vbBinaryCompare) > 0 Then EscapePattern = EscapePattern & "\"
        EscapePattern = EscapePattern & ch
    Next at
End Function

Private Function IsAlphaNumeric(ByVal ch As String) As Boolean
    IsAlphaNumeric = (Len(ch) = 1 And ch Like "[A-Za-z0-9]")
End Function

Private Sub Fail(ByVal message As String)
    Err.Raise 5, "ExportDynamicTokens", message
End Sub
