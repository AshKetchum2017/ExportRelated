Option Explicit

Private pSyncGroupIndexes As Collection
Private pCurrentSyncGroupIndex As Long
Private pSyncGroupCount As Long
Private pPageCount As Long
Private pShortfall As Boolean

Public Property Get Shortfall() As Boolean
    Shortfall = pShortfall
End Property

Public Function ResolvePageGroups(ByVal doc As Document, ByVal pageText As String, ByRef pageGroups As Collection) As Boolean
    Dim inputText As String

    Set pageGroups = New Collection
    Set pSyncGroupIndexes = New Collection
    pCurrentSyncGroupIndex = 1
    pSyncGroupCount = 1
    pPageCount = doc.Pages.Count
    pShortfall = False

    inputText = Trim$(pageText)

    If Len(inputText) = 0 Then
        AddSinglePageGroup pageGroups, doc.ActivePage.Index
        ResolvePageGroups = True
        Exit Function
    End If

    If HasPageSyncDelimiter(inputText) Then
        ParseSyncPageText inputText, pageGroups
    ElseIf HasPageGroupingToken(inputText) Then
        pCurrentSyncGroupIndex = 1
        pSyncGroupCount = 1
        ParsePageText inputText, pageGroups
    Else
        pCurrentSyncGroupIndex = 1
        pSyncGroupCount = 1
        AddGroupedPages pageGroups, inputText
    End If

    ResolvePageGroups = (pageGroups.Count > 0)
End Function

Public Function SyncGroupIndex(ByVal outputIndex As Long) As Long
    If pSyncGroupIndexes Is Nothing Then
        SyncGroupIndex = outputIndex
    ElseIf outputIndex < 1 Or outputIndex > pSyncGroupIndexes.Count Then
        SyncGroupIndex = outputIndex
    Else
        SyncGroupIndex = CLng(pSyncGroupIndexes(outputIndex))
    End If
End Function

Public Function SyncGroupCount() As Long
    If pSyncGroupCount < 1 Then
        SyncGroupCount = 1
    Else
        SyncGroupCount = pSyncGroupCount
    End If
End Function

Public Function PageGroupToRangeText(ByVal pageGroup As Variant) As String
    Dim pageIndex As Long
    Dim result As String

    For pageIndex = LBound(pageGroup) To UBound(pageGroup)
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(pageGroup(pageIndex))
    Next pageIndex

    PageGroupToRangeText = result
End Function

Public Function FirstPageInGroup(ByVal pageGroup As Variant) As Long
    FirstPageInGroup = CLng(pageGroup(LBound(pageGroup)))
End Function

' Hanya untuk tampilan; urutan dan duplikat tetap dipertahankan.
' Dua angka berurutan tetap memakai koma, tiga atau lebih menjadi range.
Public Function DisplayNumberList(ByVal values As Variant) As String
    Dim position As Long
    Dim lastPosition As Long
    Dim nextPosition As Long
    Dim result As String
    position = LBound(values)
    lastPosition = UBound(values)
    Do While position <= lastPosition
        nextPosition = position
        Do While nextPosition < lastPosition
            If CLng(values(nextPosition + 1)) <> CLng(values(nextPosition)) + 1 Then Exit Do
            nextPosition = nextPosition + 1
        Loop
        If Len(result) > 0 Then result = result & ","
        result = result & CStr(values(position))
        If nextPosition - position >= 2 Then
            result = result & "-" & CStr(values(nextPosition))
        ElseIf nextPosition > position Then
            result = result & "," & CStr(values(nextPosition))
        End If
        position = nextPosition + 1
    Loop
    DisplayNumberList = result
End Function

Public Function PageGroupCount(ByVal pageGroup As Variant) As Long
    PageGroupCount = UBound(pageGroup) - LBound(pageGroup) + 1
End Function

Private Sub ParseSyncPageText(ByVal inputText As String, ByRef pageGroups As Collection)
    Dim position As Long
    Dim segmentText As String

    position = 1
    pCurrentSyncGroupIndex = 0

    Do While position <= Len(inputText)
        segmentText = Trim$(ReadTopLevelSyncSegment(inputText, position))
        If Len(segmentText) = 0 Then
            Err.Raise 5, , "Token sync txbPage kosong."
        End If

        pCurrentSyncGroupIndex = pCurrentSyncGroupIndex + 1
        ParsePageSegment segmentText, pageGroups
    Loop

    pSyncGroupCount = pCurrentSyncGroupIndex
End Sub

Private Sub ParsePageSegment(ByVal segmentText As String, ByRef pageGroups As Collection)
    If HasPageGroupingToken(segmentText) Then
        ParsePageText segmentText, pageGroups
    Else
        AddGroupedPages pageGroups, segmentText
    End If
End Sub

Private Sub ParsePageText(ByVal inputText As String, ByRef pageGroups As Collection)
    Dim position As Long
    Dim tokenText As String
    Dim closePosition As Long

    position = 1
    Do While position <= Len(inputText)
        SkipSpaces inputText, position
        If position > Len(inputText) Then Exit Do

        If Mid$(inputText, position, 1) = "(" Then
            closePosition = FindMatchingToken(inputText, position, "(", ")")
            tokenText = Mid$(inputText, position + 1, closePosition - position - 1)
            AddGroupedPages pageGroups, tokenText
            position = closePosition + 1
            ConsumeGroupDelimiter inputText, position
        ElseIf Mid$(inputText, position, 1) = "{" Then
            closePosition = FindMatchingToken(inputText, position, "{", "}")
            tokenText = Mid$(inputText, position + 1, closePosition - position - 1)
            AddSplitGroupPages pageGroups, tokenText
            position = closePosition + 1
            ConsumeGroupDelimiter inputText, position
        Else
            tokenText = ReadTopLevelToken(inputText, position)
            AddUngroupedPages pageGroups, tokenText
        End If
    Loop
End Sub

Private Sub AddSplitGroupPages(ByRef pageGroups As Collection, ByVal groupText As String)
    Dim position As Long
    Dim tokenText As String
    Dim closePosition As Long

    groupText = Trim$(groupText)
    If Len(groupText) = 0 Then
        Err.Raise 5, , "Token grup txbPage kosong."
    End If

    position = 1
    Do While position <= Len(groupText)
        SkipSpaces groupText, position
        If position > Len(groupText) Then Exit Do

        If Mid$(groupText, position, 1) = "(" Then
            closePosition = FindMatchingToken(groupText, position, "(", ")")
            tokenText = Mid$(groupText, position + 1, closePosition - position - 1)
            AddGroupedPages pageGroups, tokenText
            position = closePosition + 1
            ConsumeGroupDelimiter groupText, position
        ElseIf Mid$(groupText, position, 1) = "{" Then
            closePosition = FindMatchingToken(groupText, position, "{", "}")
            tokenText = Mid$(groupText, position + 1, closePosition - position - 1)
            AddSplitGroupPages pageGroups, tokenText
            position = closePosition + 1
            ConsumeGroupDelimiter groupText, position
        Else
            tokenText = ReadTopLevelToken(groupText, position)
            AddSeparatePages pageGroups, tokenText
        End If
    Loop
End Sub

Private Sub AddGroupedPages(ByRef pageGroups As Collection, ByVal groupText As String)
    Dim pages As New Collection
    Dim pageValues() As Long
    Dim pageIndex As Long

    AddPagesToCollection pages, groupText
    If pages.Count = 0 Then
        If InStr(groupText, "/*n") > 0 Then Exit Sub
        Err.Raise 5, , "Token grup txbPage kosong."
    End If

    ReDim pageValues(0 To pages.Count - 1)
    For pageIndex = 1 To pages.Count
        pageValues(pageIndex - 1) = CLng(pages(pageIndex))
    Next pageIndex

    pageGroups.Add pageValues
    pSyncGroupIndexes.Add pCurrentSyncGroupIndex
End Sub

Private Sub AddSeparatePages(ByRef pageGroups As Collection, ByVal tokenText As String)
    Dim pages As New Collection
    Dim pageIndex As Long

    AddPagesToCollection pages, tokenText
    If pages.Count = 0 Then
        If InStr(tokenText, "/*n") > 0 Then Exit Sub
        Err.Raise 5, , "Token txbPage kosong."
    End If

    For pageIndex = 1 To pages.Count
        AddSinglePageGroup pageGroups, CLng(pages(pageIndex))
    Next pageIndex
End Sub

Private Sub AddUngroupedPages(ByRef pageGroups As Collection, ByVal tokenText As String)
    Dim pages As New Collection
    Dim pageValues() As Long
    Dim pageIndex As Long

    AddPagesToCollection pages, tokenText
    If pages.Count = 0 Then
        If InStr(tokenText, "/*n") > 0 Then Exit Sub
        Err.Raise 5, , "Token txbPage kosong."
    End If

    For pageIndex = 1 To pages.Count
        ReDim Preserve pageValues(0 To pageIndex - 1)
        pageValues(pageIndex - 1) = CLng(pages(pageIndex))
    Next pageIndex

    pageGroups.Add pageValues
    pSyncGroupIndexes.Add pCurrentSyncGroupIndex
End Sub

Private Function HasPageGroupingToken(ByVal inputText As String) As Boolean
    HasPageGroupingToken = (InStr(1, inputText, "(", vbBinaryCompare) > 0 Or _
                            InStr(1, inputText, ")", vbBinaryCompare) > 0 Or _
                            InStr(1, inputText, "{", vbBinaryCompare) > 0 Or _
                            InStr(1, inputText, "}", vbBinaryCompare) > 0)
End Function

Private Function HasPageSyncDelimiter(ByVal inputText As String) As Boolean
    HasPageSyncDelimiter = ContainsTopLevelDelimiter(inputText, ";")
End Function

Private Sub AddSinglePageGroup(ByRef pageGroups As Collection, ByVal pageNumber As Long)
    Dim pageValues() As Long

    ReDim pageValues(0 To 0)
    pageValues(0) = pageNumber
    pageGroups.Add pageValues
    pSyncGroupIndexes.Add pCurrentSyncGroupIndex
End Sub

Private Sub AddPagesToCollection(ByRef pages As Collection, ByVal tokenText As String)
    Dim rawValues() As String
    Dim rawIndex As Long
    Dim itemText As String

    tokenText = Trim$(tokenText)
    If Len(tokenText) = 0 Then Exit Sub

    rawValues = Split(tokenText, ",")
    For rawIndex = 0 To UBound(rawValues)
        itemText = Trim$(rawValues(rawIndex))
        AddPageToken pages, itemText
    Next rawIndex
End Sub

Private Function ContainsTopLevelDelimiter(ByVal inputText As String, ByVal delimiterText As String) As Boolean
    Dim position As Long
    Dim currentChar As String
    Dim parenDepth As Long
    Dim braceDepth As Long

    For position = 1 To Len(inputText)
        currentChar = Mid$(inputText, position, 1)
        Select Case currentChar
            Case "("
                parenDepth = parenDepth + 1
            Case ")"
                If parenDepth > 0 Then parenDepth = parenDepth - 1
            Case "{"
                braceDepth = braceDepth + 1
            Case "}"
                If braceDepth > 0 Then braceDepth = braceDepth - 1
            Case delimiterText
                If parenDepth = 0 And braceDepth = 0 Then
                    ContainsTopLevelDelimiter = True
                    Exit Function
                End If
        End Select
    Next position
End Function

Private Function ReadTopLevelSyncSegment(ByVal inputText As String, ByRef position As Long) As String
    Dim tokenStart As Long
    Dim currentChar As String
    Dim parenDepth As Long
    Dim braceDepth As Long

    tokenStart = position
    Do While position <= Len(inputText)
        currentChar = Mid$(inputText, position, 1)

        Select Case currentChar
            Case "("
                parenDepth = parenDepth + 1
            Case ")"
                If parenDepth = 0 Then
                    Err.Raise 5, , "Token grup txbPage tidak valid: " & inputText
                End If
                parenDepth = parenDepth - 1
            Case "{"
                braceDepth = braceDepth + 1
            Case "}"
                If braceDepth = 0 Then
                    Err.Raise 5, , "Token grup txbPage tidak valid: " & inputText
                End If
                braceDepth = braceDepth - 1
            Case ";"
                If parenDepth = 0 And braceDepth = 0 Then
                    ReadTopLevelSyncSegment = Mid$(inputText, tokenStart, position - tokenStart)
                    position = position + 1
                    Exit Function
                End If
        End Select

        position = position + 1
    Loop

    If parenDepth <> 0 Or braceDepth <> 0 Then
        Err.Raise 5, , "Token grup txbPage tidak lengkap: " & inputText
    End If

    ReadTopLevelSyncSegment = Mid$(inputText, tokenStart)
End Function

Private Sub AddPageToken(ByRef pages As Collection, ByVal tokenText As String)
    Dim dashPosition As Long
    Dim startPage As Long
    Dim endPage As Long
    Dim pageNumber As Long
    Dim marker As Long, countText As String, requested As Long
    Dim tokens As New ExportDynamicTokens

    If Len(tokenText) = 0 Then
        Err.Raise 5, , "Format txbPage tidak valid."
    End If

    marker = InStr(1, tokenText, "/*", vbBinaryCompare)
    If marker > 0 Then
        If Mid$(tokenText, marker, 3) <> "/*n" Then Err.Raise 5, "ExportPageParser", "txbPage hanya menerima /*n."
        startPage = tokens.PositiveLong(Left$(tokenText, marker - 1), "Page awal")
        countText = Mid$(tokenText, marker + 3)
        endPage = pPageCount
        If Len(countText) > 0 Then
            If Left$(countText, 1) <> "-" Then Err.Raise 5, "ExportPageParser", "Token page dinamis tidak valid: " & tokenText
            requested = tokens.PositiveLong(Mid$(countText, 2), "Batas page")
            If CDbl(startPage) + requested - 1 > pPageCount Then
                pShortfall = True
            Else
                endPage = startPage + requested - 1
            End If
        End If
        If startPage > pPageCount Then
            pShortfall = True
            Exit Sub
        End If
        For pageNumber = startPage To endPage
            pages.Add pageNumber
        Next pageNumber
        Exit Sub
    End If

    dashPosition = InStr(1, tokenText, "-", vbBinaryCompare)
    If dashPosition > 0 Then
        If InStr(dashPosition + 1, tokenText, "-", vbBinaryCompare) > 0 Then
            Err.Raise 5, , "Range txbPage tidak valid: " & tokenText
        End If

        startPage = ParsePositiveLong(Left$(tokenText, dashPosition - 1), tokenText)
        endPage = ParsePositiveLong(Mid$(tokenText, dashPosition + 1), tokenText)
        If endPage < startPage Then
            Err.Raise 5, , "Range txbPage harus dari kecil ke besar: " & tokenText
        End If
        If endPage > pPageCount Then Err.Raise 5, "ExportPageParser", "Page " & CStr(endPage) & " tidak tersedia. Perbaiki range eksplisit txbPage."

        For pageNumber = startPage To endPage
            pages.Add pageNumber
        Next pageNumber
    Else
        pageNumber = ParsePositiveLong(tokenText, tokenText)
        If pageNumber > pPageCount Then Err.Raise 5, "ExportPageParser", "Page " & CStr(pageNumber) & " tidak tersedia. Perbaiki txbPage."
        pages.Add pageNumber
    End If
End Sub

Private Function ParsePositiveLong(ByVal valueText As String, ByVal sourceText As String) As Long
    valueText = Trim$(valueText)
    If Not IsPositiveIntegerText(valueText) Then
        Err.Raise 5, , "Page harus berupa angka: " & sourceText
    End If

    ParsePositiveLong = CLng(valueText)
    If ParsePositiveLong < 1 Then
        Err.Raise 5, , "Nomor page harus lebih besar dari 0: " & sourceText
    End If
End Function

Private Function IsPositiveIntegerText(ByVal valueText As String) As Boolean
    Dim position As Long

    If Len(valueText) = 0 Then Exit Function

    For position = 1 To Len(valueText)
        If Not Mid$(valueText, position, 1) Like "[0-9]" Then
            Exit Function
        End If
    Next position

    IsPositiveIntegerText = True
End Function

Private Sub SkipSpaces(ByVal inputText As String, ByRef position As Long)
    Do While position <= Len(inputText) And Mid$(inputText, position, 1) = " "
        position = position + 1
    Loop
End Sub

Private Function FindMatchingToken(ByVal inputText As String, _
                                   ByVal openPosition As Long, _
                                   ByVal openToken As String, _
                                   ByVal closeToken As String) As Long
    Dim position As Long
    Dim depth As Long
    Dim currentChar As String

    For position = openPosition To Len(inputText)
        currentChar = Mid$(inputText, position, 1)
        If currentChar = openToken Then
            depth = depth + 1
        ElseIf currentChar = closeToken Then
            depth = depth - 1
            If depth = 0 Then
                FindMatchingToken = position
                Exit Function
            End If
        End If
    Next position

    Err.Raise 5, , "Token grup txbPage tidak lengkap: " & inputText
End Function

Private Sub ConsumeGroupDelimiter(ByVal inputText As String, ByRef position As Long)
    SkipSpaces inputText, position
    If position <= Len(inputText) Then
        If Mid$(inputText, position, 1) <> "," Then
            Err.Raise 5, , "Delimiter txbPage tidak valid: " & inputText
        End If
        position = position + 1
    End If
End Sub

Private Function ReadTopLevelToken(ByVal inputText As String, ByRef position As Long) As String
    Dim tokenStart As Long
    Dim currentChar As String
    Dim parenDepth As Long
    Dim braceDepth As Long

    tokenStart = position
    Do While position <= Len(inputText)
        currentChar = Mid$(inputText, position, 1)

        Select Case currentChar
            Case "("
                parenDepth = parenDepth + 1
            Case ")"
                If parenDepth = 0 Then
                    Err.Raise 5, , "Token grup txbPage tidak valid: " & inputText
                End If
                parenDepth = parenDepth - 1
            Case "{"
                braceDepth = braceDepth + 1
            Case "}"
                If braceDepth = 0 Then
                    Err.Raise 5, , "Token grup txbPage tidak valid: " & inputText
                End If
                braceDepth = braceDepth - 1
            Case ","
                If parenDepth = 0 And braceDepth = 0 Then
                    ReadTopLevelToken = Mid$(inputText, tokenStart, position - tokenStart)
                    position = position + 1
                    Exit Function
                End If
        End Select

        position = position + 1
    Loop

    If parenDepth <> 0 Or braceDepth <> 0 Then
        Err.Raise 5, , "Token grup txbPage tidak lengkap: " & inputText
    End If

    ReadTopLevelToken = Mid$(inputText, tokenStart)
End Function
