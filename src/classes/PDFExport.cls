Option Explicit

Private Const REG_APP_NAME As String = "RinCorelMacros"
Private Const REG_SECTION_NAME As String = "ExportRelatedMacro"
Private Const REG_PDF_PRESET_KEY As String = "PDFPreset"
Private Const REG_PDF_COMPATIBILITY_KEY As String = "PDFCompatibility"
Private Const DEFAULT_PDF_PRESET As String = "Used by Rin's Macro"
Private Const DEFAULT_PDF_COMPATIBILITY As String = "Acrobat 9.0"

Private pLastErrorNumber As Long
Private pLastErrorDescription As String

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = pLastErrorNumber
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = pLastErrorDescription
End Property

Public Function PublishDocument(ByVal doc As Document, ByVal outputPath As String, _
                                Optional ByVal pageRangeText As String = vbNullString, _
                                Optional ByVal queueItem As ExportSettingItem = Nothing) As Boolean
    Dim pdfSettings As PDFVBASettings
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error Resume Next
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString

    Set pdfSettings = doc.PDFSettings
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear

    If errorNumber <> 0 Or pdfSettings Is Nothing Then
        If errorNumber = 0 Then
            errorNumber = 91
            errorDescription = "PDFSettings tidak tersedia."
        End If

        pLastErrorNumber = errorNumber
        pLastErrorDescription = errorDescription
        PublishDocument = False
        Exit Function
    End If

    If queueItem Is Nothing Then
        If Not TryUsePDFPreset(pdfSettings) Then Exit Function
        TryApplyPDFCompatibility pdfSettings
    Else
        If Not ValidateQueueSettings(queueItem.PDFPreset, queueItem.PDFCompatibility) Then Exit Function
        If Not TryUsePDFPreset(pdfSettings, queueItem.PDFPreset, True) Then Exit Function
        TryApplyPDFCompatibility pdfSettings, queueItem.PDFCompatibility, True
        If pLastErrorNumber <> 0 Then Exit Function
    End If

    If IsMultiPageRange(pageRangeText) Then
        pdfSettings.PublishRange = pdfPageRange
        pdfSettings.PageRange = pageRangeText
    Else
        pdfSettings.PublishRange = pdfCurrentPage
    End If
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear

    If errorNumber <> 0 Then
        pLastErrorNumber = errorNumber
        pLastErrorDescription = errorDescription
        PublishDocument = False
        Exit Function
    End If

    doc.PublishToPDF outputPath
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear

    pLastErrorNumber = errorNumber
    pLastErrorDescription = errorDescription
    PublishDocument = (errorNumber = 0)
End Function

Private Function IsMultiPageRange(ByVal pageRangeText As String) As Boolean
    pageRangeText = Trim$(pageRangeText)
    IsMultiPageRange = (InStr(1, pageRangeText, ",", vbBinaryCompare) > 0 Or _
                        InStr(1, pageRangeText, "-", vbBinaryCompare) > 0)
End Function

Private Function TryUsePDFPreset(ByVal pdfSettings As PDFVBASettings, _
                                 Optional ByVal presetOverride As String = vbNullString, _
                                 Optional ByVal requireExisting As Boolean = False) As Boolean
    Dim presetName As String
    Dim presetLoaded As Boolean
    Dim presetSaved As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error Resume Next

    If pdfSettings Is Nothing Then
        TryUsePDFPreset = False
        Exit Function
    End If

    If requireExisting Then
        presetName = presetOverride
    Else
        presetName = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_PDF_PRESET_KEY, DEFAULT_PDF_PRESET)
    End If
    presetName = Trim$(presetName)
    If Len(presetName) = 0 Then
        presetName = DEFAULT_PDF_PRESET
    End If

    presetLoaded = pdfSettings.Load(presetName)
    If Err.Number <> 0 Then
        errorNumber = Err.Number
        errorDescription = Err.Description
        Err.Clear
        presetLoaded = False
    End If

    If Not presetLoaded Then
        presetLoaded = TryLoadNormalizedPDFPreset(pdfSettings, presetName, errorNumber, errorDescription)
    End If

    If presetLoaded Then
        TryUsePDFPreset = True
        Exit Function
    End If

    If requireExisting Or StrComp(presetName, DEFAULT_PDF_PRESET, vbTextCompare) <> 0 Then
        pLastErrorNumber = errorNumber
        If pLastErrorNumber = 0 Then pLastErrorNumber = 5
        pLastErrorDescription = "Preset PDF tidak ditemukan atau tidak bisa dimuat: " & presetName
        If Len(errorDescription) > 0 Then
            pLastErrorDescription = pLastErrorDescription & vbCrLf & errorDescription
        End If
        TryUsePDFPreset = False
        Exit Function
    End If

    presetSaved = pdfSettings.Save(DEFAULT_PDF_PRESET)
    If Err.Number <> 0 Then
        Err.Clear
        TryUsePDFPreset = True
        Exit Function
    End If

    If presetSaved Then
        pdfSettings.Load DEFAULT_PDF_PRESET
        Err.Clear
    End If

    TryUsePDFPreset = True
End Function

Private Function TryLoadNormalizedPDFPreset(ByVal pdfSettings As PDFVBASettings, _
                                            ByVal presetName As String, _
                                            ByRef errorNumber As Long, _
                                            ByRef errorDescription As String) As Boolean
    Dim normalizedPresetName As String

    normalizedPresetName = NormalizePDFPresetLoadName(presetName)
    If StrComp(normalizedPresetName, presetName, vbTextCompare) = 0 Then
        TryLoadNormalizedPDFPreset = False
        Exit Function
    End If

    On Error Resume Next
    TryLoadNormalizedPDFPreset = pdfSettings.Load(normalizedPresetName)
    If Err.Number <> 0 Then
        errorNumber = Err.Number
        errorDescription = Err.Description
        Err.Clear
        TryLoadNormalizedPDFPreset = False
    End If
End Function

Private Function NormalizePDFPresetLoadName(ByVal presetName As String) As String
    Select Case UCase$(Trim$(presetName))
        Case "ARCHIVING (CMYK)"
            NormalizePDFPresetLoadName = "CorelDefaultID_35"
        Case "ARCHIVING (RGB)"
            NormalizePDFPresetLoadName = Trim$(presetName)
        Case "CURRENT PROOF SETTINGS"
            NormalizePDFPresetLoadName = "CorelDefaultID_20"
        Case "DOCUMENT DISTRIBUTION"
            NormalizePDFPresetLoadName = "CorelDefaultID_34"
        Case "EDITING"
            NormalizePDFPresetLoadName = "CorelDefaultID_36"
        Case "PDF/X-1A:2001"
            NormalizePDFPresetLoadName = "CorelDefaultID_84"
        Case "PDF/X-3:2002"
            NormalizePDFPresetLoadName = "CorelDefaultID_85"
        Case "PDF/X-4:2010 (CMYK)"
            NormalizePDFPresetLoadName = "CorelDefaultID_86"
        Case "PREPRESS"
            NormalizePDFPresetLoadName = "CorelDefaultID_17"
        Case "WEB"
            NormalizePDFPresetLoadName = "CorelDefaultID_33"
        Case Else
            NormalizePDFPresetLoadName = Trim$(presetName)
    End Select
End Function

Private Sub TryApplyPDFCompatibility(ByVal pdfSettings As PDFVBASettings, _
                                     Optional ByVal compatibilityOverride As String = vbNullString, _
                                     Optional ByVal strict As Boolean = False)
    Dim compatibilityName As String
    Dim pdfVersionValue As Long

    On Error Resume Next

    If strict Then
        compatibilityName = compatibilityOverride
    Else
        compatibilityName = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_PDF_COMPATIBILITY_KEY, DEFAULT_PDF_COMPATIBILITY)
    End If
    If Not ResolvePDFVersionValue(compatibilityName, pdfVersionValue) Then
        Exit Sub
    End If

    CallByName pdfSettings, "PDFVersion", VbLet, pdfVersionValue
    If Err.Number <> 0 Then
        Err.Clear
        CallByName pdfSettings, "pdfVersion", VbLet, pdfVersionValue
        If strict And Err.Number <> 0 Then
            pLastErrorNumber = Err.Number
            pLastErrorDescription = "Mengatur compatibility PDF item: " & Err.Description
        End If
        Err.Clear
    End If
End Sub

Public Function ValidateQueueSettings(ByVal presetName As String, ByVal compatibilityName As String) As Boolean
    Dim versionValue As Long
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    If Len(Trim$(presetName)) = 0 Then
        pLastErrorNumber = 5
        pLastErrorDescription = "Preset PDF item kosong. Buka Modify lalu Save."
    ElseIf Not ResolvePDFVersionValue(compatibilityName, versionValue) Then
        pLastErrorNumber = 5
        pLastErrorDescription = "Compatibility PDF item tidak valid: " & compatibilityName
    Else
        ValidateQueueSettings = True
    End If
End Function

Private Function ResolvePDFVersionValue(ByVal compatibilityName As String, ByRef pdfVersionValue As Long) As Boolean
    Select Case UCase$(Trim$(compatibilityName))
        Case "ACROBAT DC"
            pdfVersionValue = 8
        Case "ACROBAT 9.0", "ACROBAT 9"
            pdfVersionValue = 9
        Case "PDF/X-1A:2021", "PDF/X-1A", "PDFX1A"
            pdfVersionValue = 4
        Case "PDF/X-3:2002", "PDF/X-3", "PDFX3"
            pdfVersionValue = 5
        Case "PDF/A-1B", "PDFA1B"
            pdfVersionValue = 7
        Case "PDF/X-4:2010", "PDF/X-4", "PDFX4"
            pdfVersionValue = 11
        Case Else
            ResolvePDFVersionValue = False
            Exit Function
    End Select

    ResolvePDFVersionValue = True
End Function
