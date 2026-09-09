Option Explicit

Public Function ParseReplaceTemplate(ByVal templateText As String, _
                                     ByRef replaceFind As String, _
                                     ByRef replaceValue As String) As Boolean
    Dim markerPos As Long

    markerPos = InStr(1, templateText, "/*", vbTextCompare)
    If markerPos = 0 Then
        ParseReplaceTemplate = False
        Exit Function
    End If

    replaceFind = Trim$(Left$(templateText, markerPos - 1))
    replaceValue = Trim$(Mid$(templateText, markerPos + 2))
    ParseReplaceTemplate = True
End Function

Public Function ParseSequenceTemplate(ByVal templateText As String, _
                                      ByRef replaceFind As String, _
                                      ByRef sequenceText As String) As Boolean
    Dim openPos As Long
    Dim closePos As Long

    openPos = InStr(1, templateText, "(/*", vbTextCompare)
    If openPos = 0 Then
        ParseSequenceTemplate = False
        Exit Function
    End If

    closePos = InStr(openPos + 3, templateText, "):", vbTextCompare)
    If closePos = 0 Then
        ParseSequenceTemplate = False
        Exit Function
    End If

    replaceFind = Trim$(Mid$(templateText, openPos + 3, closePos - (openPos + 3)))
    sequenceText = Trim$(Mid$(templateText, closePos + 2))
    ParseSequenceTemplate = True
End Function

Public Function BuildExportFileName(ByVal baseName As String, _
                                    ByVal templateText As String, _
                                    ByVal pageIndex As Long, _
                                    ByVal pageNumber As Long, _
                                    Optional ByVal syncGroupIndex As Long = -1) As String
    Dim templateValue As String
    Dim replaceFind As String
    Dim replaceValue As String
    Dim commandText As String
    Dim standaloneValues() As String
    Dim parts() As String
    Dim partIndex As Long
    Dim result As String
    Dim affixValues() As String
    Dim isPrefix As Boolean
    Dim usesSync As Boolean
    Dim valueIndex As Long

    templateValue = Trim$(templateText)
    If Len(templateValue) = 0 Then
        BuildExportFileName = baseName
        Exit Function
    End If

    parts = SplitSyncNameValues(templateValue, "|")
    If UBound(parts) > 0 Then
        result = baseName
        ' Bentuk nama utama dulu agar template langsung tidak menghapus affix.
        For partIndex = 0 To UBound(parts)
            If Len(parts(partIndex)) = 0 Then Err.Raise 5, "ExportTemplateParser", "Bagian nama di antara | tidak boleh kosong."
            If Not TryParseAffix(parts(partIndex), affixValues, isPrefix, usesSync) Then
                result = BuildExportFileName(result, parts(partIndex), pageIndex, pageNumber, syncGroupIndex)
            End If
        Next partIndex
        ' Affix diterapkan setelah nama utama selesai, mengikuti urutan antar-affix.
        For partIndex = 0 To UBound(parts)
            If TryParseAffix(parts(partIndex), affixValues, isPrefix, usesSync) Then
                result = BuildExportFileName(result, parts(partIndex), pageIndex, pageNumber, syncGroupIndex)
            End If
        Next partIndex
        BuildExportFileName = result
        Exit Function
    End If

    If TryParseAffix(templateValue, affixValues, isPrefix, usesSync) Then
        valueIndex = pageIndex
        If usesSync And syncGroupIndex >= 0 Then valueIndex = syncGroupIndex
        If UBound(affixValues) = 0 Then valueIndex = 0
        If valueIndex < 0 Or valueIndex > UBound(affixValues) Then
            Err.Raise 5, "ExportTemplateParser", "Nilai prefix/suffix tidak tersedia untuk output atau grup sync ini."
        End If
        If isPrefix Then
            BuildExportFileName = affixValues(valueIndex) & baseName
        Else
            BuildExportFileName = baseName & affixValues(valueIndex)
        End If
        Exit Function
    End If

    If ParseSequenceTemplate(templateValue, replaceFind, commandText) Then
        replaceValue = BuildCommandReplacement(replaceFind, commandText, pageIndex, pageNumber, syncGroupIndex)
        BuildExportFileName = Replace$(baseName, RemoveReplacementBraces(replaceFind), replaceValue, 1, -1, vbTextCompare)
        Exit Function
    End If

    If ParseReplaceTemplate(templateValue, replaceFind, replaceValue) Then
        BuildExportFileName = Replace$(baseName, RemoveReplacementBraces(replaceFind), RemoveReplacementBraces(replaceValue), 1, -1, vbTextCompare)
        Exit Function
    End If

    If TryParseExecutionValues(templateValue, standaloneValues) Then
        If pageIndex < 0 Or pageIndex > UBound(standaloneValues) Then
            Err.Raise 5, "ExportTemplateParser", "Nilai blok nama tidak tersedia untuk output ini."
        End If
        BuildExportFileName = standaloneValues(pageIndex)
    ElseIf TryParseStandaloneValues(templateValue, standaloneValues) Then
        BuildExportFileName = BuildStructuredReplacement(RemoveReplacementBraces(standaloneValues(0)), _
            standaloneValues, pageIndex, pageNumber)
    Else
        BuildExportFileName = templateValue
    End If
End Function

Public Function ValidateNameTemplate(ByVal templateText As String, _
                                     ByVal pageCount As Long, _
                                     Optional ByVal syncGroupCount As Long = 0) As Boolean
    Dim replaceFind As String
    Dim commandText As String
    Dim commandValues() As String
    Dim valueIndex As Long
    Dim expectedCount As Long
    Dim parts() As String
    Dim partIndex As Long
    Dim isPrefix As Boolean
    Dim usesSync As Boolean

    If Len(Trim$(templateText)) = 0 Then
        ValidateNameTemplate = True
        Exit Function
    End If

    parts = SplitSyncNameValues(templateText, "|")
    If UBound(parts) > 0 Then
        For partIndex = 0 To UBound(parts)
            If Len(parts(partIndex)) = 0 Then Exit Function
            If Not ValidateNameTemplate(parts(partIndex), pageCount, syncGroupCount) Then Exit Function
        Next partIndex
        ValidateNameTemplate = True
        Exit Function
    End If

    If TryParseAffix(templateText, commandValues, isPrefix, usesSync) Then
        expectedCount = pageCount
        If usesSync And syncGroupCount > 0 Then expectedCount = syncGroupCount
        ValidateNameTemplate = (UBound(commandValues) = 0 Or UBound(commandValues) + 1 = expectedCount)
        Exit Function
    End If

    If ParseSequenceTemplate(templateText, replaceFind, commandText) Then
        If HasNameSyncDelimiter(commandText) Then
            commandValues = SplitSyncNameValues(commandText)
            expectedCount = syncGroupCount
            If expectedCount < 1 Then expectedCount = pageCount
        Else
            commandValues = SplitNameValues(commandText)
            expectedCount = pageCount
        End If

        If UBound(commandValues) = 0 Then
            ValidateNameTemplate = True
            Exit Function
        End If

        If UsesBracedReplacement(commandText) Then
            For valueIndex = 1 To UBound(commandValues)
                If Not IsSequenceValueToken(Trim$(commandValues(valueIndex))) Then
                    ValidateNameTemplate = False
                    Exit Function
                End If
            Next valueIndex
        End If

        ValidateNameTemplate = (UBound(commandValues) + 1 = expectedCount)
        Exit Function
    End If

    If TryParseExecutionValues(templateText, commandValues) Then
        ValidateNameTemplate = (UBound(commandValues) + 1 = pageCount)
    ElseIf TryParseStandaloneValues(templateText, commandValues) Then
        ValidateNameTemplate = (UBound(commandValues) = 0 Or UBound(commandValues) + 1 = pageCount)
    Else
        ValidateNameTemplate = True
    End If
End Function

Private Function TryParseAffix(ByVal templateText As String, ByRef values() As String, _
                               ByRef isPrefix As Boolean, ByRef usesSync As Boolean) As Boolean
    Dim firstValue As String
    Dim hasPrefix As Boolean
    Dim hasSuffix As Boolean
    Dim literalUnderscores As String
    Dim position As Long
    Dim valueIndex As Long
    Dim markerIndex As Long

    ' Gabungkan replacement dan affix melalui |; __ di dalam replacement tetap teks.
    If InStr(1, templateText, "/*", vbBinaryCompare) > 0 Then Exit Function
    usesSync = HasNameSyncDelimiter(templateText)
    If usesSync Then
        values = SplitSyncNameValues(templateText)
    Else
        values = SplitNameValues(templateText)
    End If
    firstValue = values(0)
    hasPrefix = (Right$(firstValue, 2) = "__")
    hasSuffix = (Left$(firstValue, 2) = "__")
    ' Prefix bisa ditulis A__, B, C maupun A, B, C__.
    If Not hasPrefix And Not hasSuffix And UBound(values) > 0 Then
        markerIndex = UBound(values)
        hasPrefix = (Right$(values(markerIndex), 2) = "__")
    End If
    If Not hasPrefix And Not hasSuffix Then Exit Function
    If hasPrefix And hasSuffix Then
        Err.Raise 5, "ExportTemplateParser", "Pisahkan prefix dan suffix menjadi dua bagian dengan |."
    End If

    isPrefix = hasPrefix
    If isPrefix Then
        firstValue = Left$(values(markerIndex), Len(values(markerIndex)) - 2)
        ' A___, B -> A_, B_: dua underscore terakhir adalah marker.
        position = Len(firstValue)
        Do While position > 0
            If Mid$(firstValue, position, 1) <> "_" Then Exit Do
            literalUnderscores = literalUnderscores & "_"
            position = position - 1
        Loop
    Else
        firstValue = Mid$(firstValue, 3)
        ' ___A, B -> _A, _B: dua underscore pertama adalah marker.
        position = 1
        Do While position <= Len(firstValue)
            If Mid$(firstValue, position, 1) <> "_" Then Exit Do
            literalUnderscores = literalUnderscores & "_"
            position = position + 1
        Loop
    End If
    values(markerIndex) = RemoveReplacementBraces(firstValue)
    For valueIndex = 0 To UBound(values)
        If Len(Trim$(RemoveReplacementBraces(values(valueIndex)))) = 0 Then Err.Raise 5, "ExportTemplateParser", "Nilai prefix/suffix tidak boleh kosong."
        If valueIndex <> markerIndex Then
            If isPrefix Then
                values(valueIndex) = RemoveReplacementBraces(values(valueIndex)) & literalUnderscores
            Else
                values(valueIndex) = literalUnderscores & RemoveReplacementBraces(values(valueIndex))
            End If
        End If
    Next valueIndex
    TryParseAffix = True
End Function

Private Function BuildCommandReplacement(ByVal targetText As String, _
                                         ByVal commandText As String, _
                                         ByVal outputIndex As Long, _
                                         ByVal pageNumber As Long, _
                                         ByVal syncGroupIndex As Long) As String
    Dim commandValues() As String
    Dim commandValue As String
    Dim valueIndex As Long

    If HasNameSyncDelimiter(commandText) Then
        commandValues = SplitSyncNameValues(commandText)
        valueIndex = syncGroupIndex
        If valueIndex < 0 Then valueIndex = outputIndex
    Else
        commandValues = SplitNameValues(commandText)
        valueIndex = outputIndex
    End If

    If UsesBracedReplacement(commandText) Then
        BuildCommandReplacement = BuildStructuredReplacement(targetText, commandValues, valueIndex, pageNumber)
        Exit Function
    End If

    If valueIndex >= 0 And valueIndex <= UBound(commandValues) Then
        commandValue = commandValues(valueIndex)
    Else
        commandValue = CStr(pageNumber)
    End If

    BuildCommandReplacement = commandValue
End Function

Private Function BuildStructuredReplacement(ByVal targetText As String, _
                                            ByRef commandValues() As String, _
                                            ByVal outputIndex As Long, _
                                            ByVal pageNumber As Long) As String
    Dim baseCommand As String
    Dim numberStart As Long
    Dim numberEnd As Long
    Dim targetNumberStart As Long
    Dim targetNumberEnd As Long
    Dim replacementNumber As String
    Dim beforeNumber As String
    Dim afterNumber As String

    baseCommand = commandValues(0)

    If Not FindFirstNumberRangeOutsideBraces(targetText, targetNumberStart, targetNumberEnd) And _
       Not FindFirstSequenceValueRangeOutsideBraces(baseCommand, numberStart, numberEnd) Then
        If outputIndex >= 0 And outputIndex <= UBound(commandValues) Then
            BuildStructuredReplacement = RemoveReplacementBraces(commandValues(outputIndex))
        Else
            BuildStructuredReplacement = RemoveReplacementBraces(baseCommand)
        End If
        Exit Function
    End If

    If Not FindFirstSequenceValueRangeOutsideBraces(baseCommand, numberStart, numberEnd) Then
        BuildStructuredReplacement = RemoveReplacementBraces(baseCommand)
        Exit Function
    End If

    beforeNumber = Left$(baseCommand, numberStart - 1)
    afterNumber = Mid$(baseCommand, numberEnd + 1)
    replacementNumber = Mid$(baseCommand, numberStart, numberEnd - numberStart + 1)

    If outputIndex > 0 And outputIndex <= UBound(commandValues) Then
        replacementNumber = commandValues(outputIndex)
    End If

    BuildStructuredReplacement = _
        RemoveReplacementBraces(beforeNumber) & _
        replacementNumber & _
        RemoveReplacementBraces(afterNumber)
End Function

Private Function FindFirstSequenceValueRangeOutsideBraces(ByVal sourceText As String, _
                                                         ByRef rangeStart As Long, ByRef rangeEnd As Long) As Boolean
    Dim position As Long
    Dim braceDepth As Long
    Dim currentChar As String

    ' Pertahankan prioritas angka pada template lama, tetapi A1/1A harus satu nilai.
    If FindFirstNumberRangeOutsideBraces(sourceText, rangeStart, rangeEnd) Then
        Do While rangeStart > 1
            If Not IsAlphaNumericChar(Mid$(sourceText, rangeStart - 1, 1)) Then Exit Do
            rangeStart = rangeStart - 1
        Loop
        Do While rangeEnd < Len(sourceText)
            If Not IsAlphaNumericChar(Mid$(sourceText, rangeEnd + 1, 1)) Then Exit Do
            rangeEnd = rangeEnd + 1
        Loop
        FindFirstSequenceValueRangeOutsideBraces = True
        Exit Function
    End If

    For position = 1 To Len(sourceText)
        currentChar = Mid$(sourceText, position, 1)
        If currentChar = "{" Then
            braceDepth = braceDepth + 1
        ElseIf currentChar = "}" Then
            If braceDepth > 0 Then braceDepth = braceDepth - 1
        ElseIf braceDepth = 0 And IsAlphaNumericChar(currentChar) Then
            rangeStart = position
            rangeEnd = position
            Do While rangeEnd < Len(sourceText)
                If Not IsAlphaNumericChar(Mid$(sourceText, rangeEnd + 1, 1)) Then Exit Do
                rangeEnd = rangeEnd + 1
            Loop
            FindFirstSequenceValueRangeOutsideBraces = True
            Exit Function
        End If
    Next position
End Function

Private Function IsAlphaNumericChar(ByVal value As String) As Boolean
    IsAlphaNumericChar = (value Like "[A-Za-z0-9]")
End Function

Private Function IsSequenceValueToken(ByVal valueText As String) As Boolean
    Dim position As Long
    valueText = Trim$(valueText)
    If IsNumberToken(valueText) Then
        IsSequenceValueToken = True
        Exit Function
    End If
    If Len(valueText) = 0 Then Exit Function
    For position = 1 To Len(valueText)
        If Not IsAlphaNumericChar(Mid$(valueText, position, 1)) Then Exit Function
    Next position
    IsSequenceValueToken = True
End Function

' Blok nama langsung dijalankan berurutan; setiap blok mempunyai pola sendiri.
' Jumlah nilai menentukan batas blok, terlepas dari pembagian sync txbPage.
Private Function TryParseExecutionValues(ByVal templateText As String, ByRef values() As String) As Boolean
    Dim blocks() As String
    Dim blockValues() As String
    Dim blockIndex As Long
    Dim valueIndex As Long
    Dim valueCount As Long
    Dim outputName As String

    If InStr(1, templateText, "/*", vbBinaryCompare) > 0 Then Exit Function
    If Not HasNameSyncDelimiter(templateText) Then Exit Function
    If Not UsesBracedReplacement(templateText) Then Exit Function

    blocks = SplitSyncNameValues(templateText)
    For blockIndex = 0 To UBound(blocks)
        If Not TryParseStandaloneValues(blocks(blockIndex), blockValues) Then
            Err.Raise 5, "ExportTemplateParser", "Setiap blok nama harus memiliki pola dan nilai, misalnya {A}1,2; {B}1,2."
        End If
        For valueIndex = 0 To UBound(blockValues)
            outputName = BuildStructuredReplacement(RemoveReplacementBraces(blockValues(0)), _
                blockValues, valueIndex, 0)
            AddStringValue values, valueCount, outputName, True
        Next valueIndex
    Next blockIndex
    TryParseExecutionValues = True
End Function

' Tanpa /*: {NAMA}_A,B -> NAMA_A, NAMA_B; spasi setelah koma tetap literal.
' Prefix/suffix dari nilai pertama dipakai ulang oleh BuildStructuredReplacement.
Private Function TryParseStandaloneValues(ByVal templateText As String, ByRef values() As String) As Boolean
    Dim valueIndex As Long
    Dim rangeStart As Long
    Dim rangeEnd As Long

    If InStr(1, templateText, "/*", vbBinaryCompare) > 0 Then Exit Function
    If Not UsesBracedReplacement(templateText) Then Exit Function
    values = SplitNameValues(templateText)
    If Not FindFirstSequenceValueRangeOutsideBraces(values(0), rangeStart, rangeEnd) Then Exit Function
    For valueIndex = 1 To UBound(values)
        If Not IsSequenceValueToken(values(valueIndex)) Then
            Err.Raise 5, "ExportTemplateParser", "Nilai nama harus berupa huruf, angka, atau gabungannya: " & values(valueIndex)
        End If
    Next valueIndex
    TryParseStandaloneValues = True
End Function

Private Function FindFirstNumberRangeOutsideBraces(ByVal sourceText As String, ByRef rangeStart As Long, ByRef rangeEnd As Long) As Boolean
    Dim position As Long
    Dim braceDepth As Long
    Dim currentChar As String

    For position = 1 To Len(sourceText)
        currentChar = Mid$(sourceText, position, 1)
        If currentChar = "{" Then
            braceDepth = braceDepth + 1
        ElseIf currentChar = "}" Then
            If braceDepth > 0 Then braceDepth = braceDepth - 1
        ElseIf braceDepth = 0 And currentChar Like "[0-9]" Then
            rangeStart = position
            rangeEnd = position

            Do While rangeEnd + 1 <= Len(sourceText) And IsNumberContinuation(sourceText, rangeEnd + 1)
                rangeEnd = rangeEnd + 1
            Loop

            FindFirstNumberRangeOutsideBraces = True
            Exit Function
        End If
    Next position
End Function

Private Function SplitNameValues(ByVal commandText As String) As String()
    Dim values() As String
    Dim valueCount As Long
    Dim position As Long
    Dim braceDepth As Long
    Dim tokenStart As Long
    Dim currentChar As String

    ReDim values(0 To 0)
    tokenStart = 1

    For position = 1 To Len(commandText)
        currentChar = Mid$(commandText, position, 1)
        Select Case currentChar
            Case "{"
                braceDepth = braceDepth + 1
            Case "}"
                If braceDepth > 0 Then braceDepth = braceDepth - 1
            Case ","
                If braceDepth = 0 Then
                    AddStringValue values, valueCount, Mid$(commandText, tokenStart, position - tokenStart), True
                    tokenStart = position + 1
                End If
        End Select
    Next position

    AddStringValue values, valueCount, Mid$(commandText, tokenStart), True
    SplitNameValues = values
End Function

Private Function SplitSyncNameValues(ByVal commandText As String, _
                                     Optional ByVal delimiter As String = ";") As String()
    Dim values() As String
    Dim valueCount As Long
    Dim position As Long
    Dim braceDepth As Long
    Dim tokenStart As Long
    Dim currentChar As String

    ReDim values(0 To 0)
    tokenStart = 1

    For position = 1 To Len(commandText)
        currentChar = Mid$(commandText, position, 1)
        Select Case currentChar
            Case "{"
                braceDepth = braceDepth + 1
            Case "}"
                If braceDepth > 0 Then braceDepth = braceDepth - 1
            Case delimiter
                If braceDepth = 0 Then
                    AddStringValue values, valueCount, Mid$(commandText, tokenStart, position - tokenStart)
                    tokenStart = position + 1
                End If
        End Select
    Next position

    AddStringValue values, valueCount, Mid$(commandText, tokenStart)
    SplitSyncNameValues = values
End Function

Private Function HasNameSyncDelimiter(ByVal commandText As String) As Boolean
    Dim position As Long
    Dim braceDepth As Long
    Dim currentChar As String

    For position = 1 To Len(commandText)
        currentChar = Mid$(commandText, position, 1)
        If currentChar = "{" Then
            braceDepth = braceDepth + 1
        ElseIf currentChar = "}" Then
            If braceDepth > 0 Then braceDepth = braceDepth - 1
        ElseIf currentChar = ";" And braceDepth = 0 Then
            HasNameSyncDelimiter = True
            Exit Function
        End If
    Next position
End Function

Private Sub AddStringValue(ByRef values() As String, ByRef valueCount As Long, ByVal valueText As String, _
                           Optional ByVal preserveSpaces As Boolean = False)
    If valueCount = 0 Then
        ReDim values(0 To 0)
    Else
        ReDim Preserve values(0 To valueCount)
    End If

    If preserveSpaces Then
        values(valueCount) = valueText
    Else
        values(valueCount) = Trim$(valueText)
    End If
    valueCount = valueCount + 1
End Sub

Private Function UsesBracedReplacement(ByVal commandText As String) As Boolean
    UsesBracedReplacement = (InStr(1, commandText, "{", vbBinaryCompare) > 0 Or _
                             InStr(1, commandText, "}", vbBinaryCompare) > 0)
End Function

Private Function RemoveReplacementBraces(ByVal valueText As String) As String
    RemoveReplacementBraces = Replace$(Replace$(valueText, "{", vbNullString), "}", vbNullString)
End Function

Private Function FindFirstNumberRange(ByVal sourceText As String, ByRef rangeStart As Long, ByRef rangeEnd As Long) As Boolean
    Dim position As Long

    For position = 1 To Len(sourceText)
        If Mid$(sourceText, position, 1) Like "[0-9]" Then
            rangeStart = position
            rangeEnd = position

            Do While rangeEnd + 1 <= Len(sourceText) And IsNumberContinuation(sourceText, rangeEnd + 1)
                rangeEnd = rangeEnd + 1
            Loop

            FindFirstNumberRange = True
            Exit Function
        End If
    Next position
End Function

Private Function IsNumberContinuation(ByVal sourceText As String, ByVal position As Long) As Boolean
    Dim currentChar As String
    Dim previousChar As String
    Dim nextChar As String

    currentChar = Mid$(sourceText, position, 1)
    If currentChar Like "[0-9]" Then
        IsNumberContinuation = True
        Exit Function
    End If

    If currentChar = "." Or currentChar = "," Or currentChar = "/" Then
        If position > 1 And position < Len(sourceText) Then
            previousChar = Mid$(sourceText, position - 1, 1)
            nextChar = Mid$(sourceText, position + 1, 1)
            IsNumberContinuation = (previousChar Like "[0-9]" And nextChar Like "[0-9]")
        End If
    End If
End Function

Private Function IsNumberToken(ByVal valueText As String) As Boolean
    Dim position As Long
    Dim currentChar As String
    Dim hasDigit As Boolean

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then Exit Function

    For position = 1 To Len(valueText)
        currentChar = Mid$(valueText, position, 1)
        If currentChar Like "[0-9]" Then
            hasDigit = True
        ElseIf currentChar = "." Or currentChar = "," Or currentChar = "/" Then
            If position = 1 Or position = Len(valueText) Then Exit Function
            If Not Mid$(valueText, position - 1, 1) Like "[0-9]" Then Exit Function
            If Not Mid$(valueText, position + 1, 1) Like "[0-9]" Then Exit Function
        Else
            Exit Function
        End If
    Next position

    IsNumberToken = hasDigit
End Function

Public Function BuildUniqueExportName(ByRef generatedNames() As String, _
                                       ByVal outputCount As Long, _
                                       ByVal exportIndex As Long) As String
    Dim baseOutputName As String
    Dim currentNameIndex As Long
    Dim duplicateCount As Long
    Dim duplicateSequence As Long

    baseOutputName = generatedNames(exportIndex - 1)

    For currentNameIndex = 0 To outputCount - 1
        If LCase$(generatedNames(currentNameIndex)) = LCase$(baseOutputName) Then
            duplicateCount = duplicateCount + 1
        End If
    Next currentNameIndex

    If duplicateCount <= 1 Then
        BuildUniqueExportName = baseOutputName
        Exit Function
    End If

    For currentNameIndex = 0 To exportIndex - 1
        If LCase$(generatedNames(currentNameIndex)) = LCase$(baseOutputName) Then
            duplicateSequence = duplicateSequence + 1
        End If
    Next currentNameIndex

    BuildUniqueExportName = baseOutputName & " (" & duplicateSequence & ")"
End Function

Public Function RemoveDocumentExtension(ByVal fileName As String) As String

    Dim dotIndex As Long

    dotIndex = InStrRev(fileName, ".")
    If dotIndex > 0 Then
        RemoveDocumentExtension = Left$(fileName, dotIndex - 1)
    Else
        RemoveDocumentExtension = fileName
    End If

End Function
