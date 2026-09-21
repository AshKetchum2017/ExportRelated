Option Explicit

' Class Module (Name) = ExportTokenPlan.
' One preflight plan shared by queue validation, queue export and single export.
' Original input remains intact in settings, so MacroRunner can reuse it on other documents.
Private pGroups As Collection
Private pNames As Collection
Private pWarning As String
Private pParser As ExportTemplateParser
Private pTokens As ExportDynamicTokens

Public Property Get Groups() As Collection
    Set Groups = pGroups
End Property

Public Property Get WarningDescription() As String
    WarningDescription = pWarning
End Property

Public Function NameAt(ByVal index As Long) As String
    NameAt = CStr(pNames(index))
End Function

Public Sub Prepare(ByVal doc As Document, ByVal pageText As String, ByVal nameText As String)
    Dim pages As New ExportPageParser, sourceGroups As Collection
    Dim parts() As String, blocks() As String, body As String, target As String, mode As Long
    Dim baseName As String, part As Long, block As Long, blockCount As Long
    Dim hasDynamic As Boolean, cap As Long, partCap As Long, dummy As String
    Dim first As Long, last As Long, available As Long, count As Long, index As Long
    Dim sourceIndex As Long, localIndex As Long, group As Variant, resolved As String
    Dim explicitSync As Boolean, shortage As Boolean, pageTotal As Long
    Dim values() As String, originalBlocks() As String
    Dim originalMode As Long, originalTarget As String
    Set pParser = New ExportTemplateParser
    Set pTokens = New ExportDynamicTokens
    Set pGroups = New Collection
    Set pNames = New Collection
    pWarning = vbNullString
    If Not pages.ResolvePageGroups(doc, pageText, sourceGroups) Then Fail "Format page tidak valid atau kosong."
    shortage = pages.Shortfall
    baseName = pParser.RemoveDocumentExtension(doc.Name)
    If Len(baseName) = 0 Then baseName = "Export"
    parts = pParser.SplitSyncNameValues(nameText, "|")
    blockCount = 1
    For part = 0 To UBound(parts)
        body = PartBody(parts(part), target, mode)
        If mode = 0 And InStr(body, "{~}") > 0 Then Fail "{~} hanya dipakai pada nilai replacement dengan target."
        If pTokens.HasDynamic(body) Then hasDynamic = True
        blocks = pParser.SplitSyncNameValues(body)
        If UBound(blocks) + 1 > blockCount Then blockCount = UBound(blocks) + 1
    Next part

    If Not hasDynamic Then
        If Not pParser.ValidateNameTemplate(nameText, sourceGroups.Count, pages.SyncGroupCount) Then Fail "Jumlah value nama tidak sesuai dengan grup page."
        For index = 1 To sourceGroups.Count
            group = sourceGroups(index)
            pGroups.Add group
            pNames.Add pParser.BuildExportFileName(baseName, nameText, index - 1, pages.FirstPageInGroup(group), pages.SyncGroupIndex(index) - 1)
        Next index
    Else
        explicitSync = (pages.SyncGroupCount > 1 And blockCount > 1)
        If explicitSync And pages.SyncGroupCount <> blockCount Then Fail "Jumlah blok nama dinamis harus sesuai dengan grup sync txbPage."
        sourceIndex = 1
        For block = 0 To blockCount - 1
            cap = 0
            ' Validate every block even when earlier blocks consume all available pages.
            For part = 0 To UBound(parts)
                body = SelectBody(parts(part), block, blockCount, target, mode)
                dummy = pTokens.Expand(body, 0, partCap)
                values = pParser.SplitNameValues(body)
                If pTokens.HasDynamic(body) And UBound(values) > 0 Then Fail "Pisahkan deret dinamis dan daftar koma menjadi blok ;."
                If UBound(values) > 0 Then partCap = UBound(values) + 1
                ' Existing direct-name blocks can define their length using explicit values.
                originalBlocks = pParser.SplitSyncNameValues(PartBody(parts(part), originalTarget, originalMode))
                If mode = 0 And UBound(originalBlocks) > 0 And Not pTokens.HasDynamic(body) Then
                    If InStr(1, body, "__", vbBinaryCompare) = 0 Then partCap = UBound(values) + 1
                End If
                If partCap > 0 Then
                    If cap > 0 And cap <> partCap Then Fail "Batas eksplisit nama/affix dalam blok " & CStr(block + 1) & " tidak konsisten."
                    cap = partCap
                End If
            Next part
            If cap = 0 And block < blockCount - 1 And Not explicitSync Then Fail "Blok nama sebelum blok terakhir memerlukan batas -n atau grup sync page eksplisit."

            first = sourceIndex
            last = sourceGroups.Count
            If explicitSync Then
                first = 0: last = 0
                For index = 1 To sourceGroups.Count
                    If pages.SyncGroupIndex(index) = block + 1 Then
                        If first = 0 Then first = index
                        last = index
                    End If
                Next index
            End If
            available = last - first + 1
            If first = 0 Or available < 0 Then available = 0
            count = available
            If cap > 0 Then
                If cap < count Then count = cap
                If cap > available Then shortage = True
            End If
            For part = 0 To UBound(parts)
                body = SelectBody(parts(part), block, blockCount, target, mode)
                dummy = pTokens.Expand(body, 0, partCap)
                resolved = AssemblePart(dummy, target, mode)
                If count > 0 Then
                    If Not pParser.ValidateNameTemplate(resolved, count, count) Then Fail "Jumlah value nama/affix tidak sesuai dengan output blok " & CStr(block + 1) & "."
                End If
            Next part
            For localIndex = 0 To count - 1
                group = sourceGroups(first + localIndex)
                resolved = vbNullString
                For part = 0 To UBound(parts)
                    body = SelectBody(parts(part), block, blockCount, target, mode)
                    body = pTokens.Expand(body, localIndex, partCap)
                    If part > 0 Then resolved = resolved & " | "
                    resolved = resolved & AssemblePart(body, target, mode)
                Next part
                pGroups.Add group
                pNames.Add pParser.BuildExportFileName(baseName, resolved, localIndex, pages.FirstPageInGroup(group), localIndex)
            Next localIndex
            sourceIndex = first + count
        Next block
    End If
    If pGroups.Count = 0 Then Fail "Tidak ada output yang dapat dibuat dari page dan nama ini."
    If shortage Then
        For index = 1 To pGroups.Count
            group = pGroups(index)
            pageTotal = pageTotal + pages.PageGroupCount(group)
        Next index
        pWarning = "Page tidak cukup untuk batas token dinamis. Output hanya menghasilkan " & _
            CStr(pGroups.Count) & " file (" & CStr(pageTotal) & " page dalam seluruh output)."
    End If
End Sub

Private Function PartBody(ByVal part As String, ByRef target As String, ByRef mode As Long) As String
    Dim body As String
    mode = 0: target = vbNullString
    If pParser.ParseSequenceTemplate(part, target, body) Then
        mode = 1
    ElseIf pParser.ParseReplaceTemplate(part, target, body) Then
        mode = 2
    Else
        body = part
    End If
    PartBody = body
End Function

Private Function SelectBody(ByVal part As String, ByVal block As Long, ByVal blockCount As Long, _
                            ByRef target As String, ByRef mode As Long) As String
    Dim blocks() As String
    Dim affixValues() As String, isPrefix As Boolean, usesSync As Boolean
    blocks = pParser.SplitSyncNameValues(PartBody(part, target, mode))
    If pParser.TryParseAffix(part, affixValues, isPrefix, usesSync, True) Then
        If usesSync Then
            If UBound(affixValues) + 1 <> blockCount Then Fail "Jumlah blok affix tidak sesuai."
            If isPrefix Then mode = 3 Else mode = 4
            SelectBody = affixValues(block)
            Exit Function
        End If
    End If
    If UBound(blocks) = 0 Then
        SelectBody = blocks(0)
    Else
        If UBound(blocks) + 1 <> blockCount Then Fail "Jumlah blok antarbagian nama/affix tidak sesuai."
        SelectBody = blocks(block)
    End If
    If Len(SelectBody) = 0 Then Fail "Bagian atau blok nama kosong."
End Function

Private Function AssemblePart(ByVal body As String, ByVal target As String, ByVal mode As Long) As String
    Select Case mode
        Case 1: AssemblePart = "(/*" & target & "): " & body
        Case 2: AssemblePart = target & " /* " & body
        Case 3: AssemblePart = body & "__"
        Case 4: AssemblePart = "__" & body
        Case Else: AssemblePart = body
    End Select
End Function

Private Sub Fail(ByVal message As String)
    Err.Raise 5, "ExportTokenPlan", message
End Sub
