Option Explicit

' Code-behind untuk UserForm dengan (Name) = PDFSettings.

Private Const REG_APP_NAME As String = "RinCorelMacros"
Private Const REG_SECTION_NAME As String = "ExportRelatedMacro"
Private Const REG_PDF_PRESET_KEY As String = "PDFPreset"
Private Const REG_PDF_PRESET_DISPLAY_KEY As String = "PDFPresetDisplay"
Private Const REG_PDF_COMPATIBILITY_KEY As String = "PDFCompatibility"
Private Const DEFAULT_PDF_PRESET As String = "Used by Rin's Macro"
Private Const DEFAULT_PDF_COMPATIBILITY As String = "Acrobat 9.0"
Private pQueueItem As ExportSettingItem

Public Sub BeginQueueEdit(ByVal item As ExportSettingItem)
    SelectPDFPresetValue cmbPDFPresets, item.PDFPreset
    SelectComboValue cmbCompatibility, item.PDFCompatibility
    Set pQueueItem = item
End Sub

Private Sub UserForm_Initialize()

    PopulatePDFPresets
    PopulateCompatibility
    LoadSavedPDFSettings

    cmdAddPreset.Enabled = False
    cmdRemovePreset.Enabled = False

End Sub

Private Sub cmbCompatibility_Change()

End Sub

Private Sub cmbPDFPresets_Change()

End Sub

Private Sub cmdAddPreset_Click()

    MsgBox "Tambah preset PDF belum tersedia.", vbInformation, "Export Related"

End Sub

Private Sub cmdClose_Click()

    Unload Me
    
End Sub

Private Sub cmdRemovePreset_Click()

    MsgBox "Hapus preset PDF belum tersedia.", vbInformation, "Export Related"

End Sub

Private Sub cmdSave_Click()

    Dim selectedPreset As String
    Dim selectedCompatibility As String

    selectedPreset = Trim$(SelectedPDFPresetLoadName())
    selectedCompatibility = Trim$(cmbCompatibility.Text)

    If Len(selectedPreset) = 0 Then
        MsgBox "Pilih preset PDF terlebih dahulu.", vbExclamation, "Export Related"
        Exit Sub
    End If

    If Len(selectedCompatibility) = 0 Then
        MsgBox "Pilih compatibility PDF terlebih dahulu.", vbExclamation, "Export Related"
        Exit Sub
    End If

    If pQueueItem Is Nothing Then
        SaveSetting REG_APP_NAME, REG_SECTION_NAME, REG_PDF_PRESET_KEY, selectedPreset
        SaveSetting REG_APP_NAME, REG_SECTION_NAME, REG_PDF_PRESET_DISPLAY_KEY, Trim$(cmbPDFPresets.Text)
        SaveSetting REG_APP_NAME, REG_SECTION_NAME, REG_PDF_COMPATIBILITY_KEY, selectedCompatibility
        MsgBox "Pengaturan PDF berhasil disimpan.", vbInformation, "Export Related"
    Else
        pQueueItem.PDFPreset = selectedPreset
        pQueueItem.PDFPresetDisplay = Trim$(cmbPDFPresets.Text)
        pQueueItem.PDFCompatibility = selectedCompatibility
    End If
    Unload Me

End Sub

Private Sub PopulatePDFPresets()

    Dim savedPreset As String

    cmbPDFPresets.Clear
    cmbPDFPresets.ColumnCount = 2
    cmbPDFPresets.BoundColumn = 1
    cmbPDFPresets.ColumnWidths = "120 pt;0 pt"

    AddBuiltInPDFPresetNames cmbPDFPresets
    AddCorelPDFPresetNames cmbPDFPresets
    savedPreset = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_PDF_PRESET_KEY, vbNullString)
    If Len(Trim$(savedPreset)) > 0 Then
        AddPDFPresetValue cmbPDFPresets, PresetDisplayName(savedPreset), savedPreset
    End If

    SortPDFPresetValues cmbPDFPresets

    If cmbPDFPresets.ListCount = 0 Then
        AddPDFPresetValue cmbPDFPresets, DEFAULT_PDF_PRESET, DEFAULT_PDF_PRESET
    End If

    If cmbPDFPresets.ListCount > 0 Then
        cmbPDFPresets.ListIndex = 0
    End If

End Sub

Private Sub PopulateCompatibility()

    cmbCompatibility.Clear
    cmbCompatibility.AddItem "Acrobat DC"
    cmbCompatibility.AddItem "Acrobat 9.0"
    cmbCompatibility.AddItem "PDF/X-1a:2021"
    cmbCompatibility.AddItem "PDF/X-3:2002"
    cmbCompatibility.AddItem "PDF/A-1b"
    cmbCompatibility.AddItem "PDF/X-4:2010"
    cmbCompatibility.ListIndex = 1

End Sub

Private Sub LoadSavedPDFSettings()

    Dim savedPreset As String
    Dim savedCompatibility As String

    savedPreset = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_PDF_PRESET_KEY, DEFAULT_PDF_PRESET)
    savedCompatibility = GetSetting(REG_APP_NAME, REG_SECTION_NAME, REG_PDF_COMPATIBILITY_KEY, DEFAULT_PDF_COMPATIBILITY)

    SelectPDFPresetValue cmbPDFPresets, savedPreset
    SelectComboValue cmbCompatibility, savedCompatibility

End Sub

Private Sub AddBuiltInPDFPresetNames(ByVal cmb As Object)

    AddPDFPresetValue cmb, "Archiving (CMYK)", "Archiving (CMYK)"
    AddPDFPresetValue cmb, "Archiving (RGB)", "Archiving (RGB)"
    AddPDFPresetValue cmb, "Current Proof Settings", "Current Proof Settings"
    AddPDFPresetValue cmb, "Document Distribution", "Document Distribution"
    AddPDFPresetValue cmb, "Editing", "Editing"
    AddPDFPresetValue cmb, "PDF/X-1a:2001", "PDF/X-1a:2001"
    AddPDFPresetValue cmb, "PDF/X-3:2002", "PDF/X-3:2002"
    AddPDFPresetValue cmb, "PDF/X-4:2010 (CMYK)", "PDF/X-4:2010 (CMYK)"
    AddPDFPresetValue cmb, "Prepress", "Prepress"
    AddPDFPresetValue cmb, "Web", "Web"

End Sub

Private Sub AddCorelPDFPresetNames(ByVal cmb As Object)

    Dim configPath As String
    Dim currentStyle As String
    Dim fileNumber As Integer
    Dim lineText As String
    Dim presetName As String

    configPath = ResolveCorelPDFConfigPath()
    If Len(configPath) = 0 Then
        Exit Sub
    End If

    On Error GoTo ReadFailed
    fileNumber = FreeFile
    Open configPath For Input As #fileNumber
    Do While Not EOF(fileNumber)
        Line Input #fileNumber, lineText
        lineText = Trim$(lineText)

        If Left$(lineText, 6) = "Style=" Then
            currentStyle = Trim$(Mid$(lineText, 7))
            If Not IsBuiltInPresetKey(currentStyle) Then
                AddPDFPresetValue cmb, currentStyle, currentStyle
            End If
        ElseIf Left$(lineText, 6) = "[Style" And Right$(lineText, 1) = "]" Then
            presetName = Mid$(lineText, 7, Len(lineText) - 7)
            If Not IsBuiltInPresetKey(presetName) Then
                AddPDFPresetValue cmb, presetName, presetName
            End If
        End If
    Loop

ReadDone:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
    Exit Sub

ReadFailed:
    Resume ReadDone
End Sub

Private Function ResolveCorelPDFConfigPath() As String

    Dim corelRoot As String
    Dim folderName As String
    Dim configPath As String
    Dim fso As Object

    corelRoot = Environ$("APPDATA") & "\Corel\"
    Set fso = CreateObject("Scripting.FileSystemObject")

    folderName = Dir$(corelRoot & "CorelDRAW Graphics Suite *", vbDirectory)

    Do While Len(folderName) > 0
        If folderName <> "." And folderName <> ".." Then
            configPath = corelRoot & folderName & "\Config\corelpdf.ini"
            If fso.FileExists(configPath) Then
                ResolveCorelPDFConfigPath = configPath
                Exit Function
            End If
        End If

        folderName = Dir$()
    Loop
End Function

Private Function SelectedPDFPresetLoadName() As String

    If cmbPDFPresets.ListIndex >= 0 And cmbPDFPresets.ColumnCount > 1 Then
        SelectedPDFPresetLoadName = Trim$(cmbPDFPresets.List(cmbPDFPresets.ListIndex, 1))
    End If

    If Len(SelectedPDFPresetLoadName) = 0 Then
        SelectedPDFPresetLoadName = Trim$(cmbPDFPresets.Text)
    End If
End Function

Private Sub AddPDFPresetValue(ByVal cmb As Object, ByVal displayName As String, ByVal loadName As String)

    Dim itemIndex As Long

    displayName = Trim$(displayName)
    loadName = Trim$(loadName)
    If Len(displayName) = 0 Or Len(loadName) = 0 Then
        Exit Sub
    End If

    For itemIndex = 0 To cmb.ListCount - 1
        If StrComp(cmb.List(itemIndex, 1), loadName, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next itemIndex

    cmb.AddItem displayName
    cmb.List(cmb.ListCount - 1, 1) = loadName

End Sub

Private Sub SelectPDFPresetValue(ByVal cmb As Object, ByVal loadName As String)

    Dim itemIndex As Long
    Dim displayName As String

    loadName = Trim$(loadName)
    If Len(loadName) = 0 Then
        Exit Sub
    End If

    displayName = PresetDisplayName(loadName)

    For itemIndex = 0 To cmb.ListCount - 1
        If StrComp(cmb.List(itemIndex, 1), loadName, vbTextCompare) = 0 Or _
           StrComp(cmb.List(itemIndex, 0), loadName, vbTextCompare) = 0 Or _
           StrComp(cmb.List(itemIndex, 0), displayName, vbTextCompare) = 0 Then
            cmb.ListIndex = itemIndex
            Exit Sub
        End If
    Next itemIndex

    AddPDFPresetValue cmb, displayName, loadName
    cmb.ListIndex = cmb.ListCount - 1

End Sub

Private Function IsBuiltInPresetKey(ByVal presetName As String) As Boolean
    IsBuiltInPresetKey = (Left$(UCase$(Trim$(presetName)), 15) = "CORELDEFAULTID_")
End Function

Private Function PresetDisplayName(ByVal loadName As String) As String

    Select Case UCase$(Trim$(loadName))
        Case "CORELDEFAULTID_35"
            PresetDisplayName = "Archiving (CMYK)"
        Case "CORELDEFAULTID_33"
            PresetDisplayName = "Web"
        Case "CORELDEFAULTID_34"
            PresetDisplayName = "Document Distribution"
        Case "CORELDEFAULTID_36"
            PresetDisplayName = "Editing"
        Case "CORELDEFAULTID_84"
            PresetDisplayName = "PDF/X-1a:2001"
        Case "CORELDEFAULTID_85"
            PresetDisplayName = "PDF/X-3:2002"
        Case "CORELDEFAULTID_86"
            PresetDisplayName = "PDF/X-4:2010 (CMYK)"
        Case "CORELDEFAULTID_18"
            PresetDisplayName = "Editing"
        Case "CORELDEFAULTID_19"
            PresetDisplayName = "Prepress"
        Case "CORELDEFAULTID_20"
            PresetDisplayName = "Current Proof Settings"
        Case "CORELDEFAULTID_17"
            PresetDisplayName = "Prepress"
        Case Else
            PresetDisplayName = Trim$(loadName)
    End Select
End Function

Private Sub SortPDFPresetValues(ByVal cmb As Object)

    Dim outerIndex As Long
    Dim innerIndex As Long
    Dim displayValue As String
    Dim loadValue As String

    For outerIndex = 0 To cmb.ListCount - 2
        For innerIndex = outerIndex + 1 To cmb.ListCount - 1
            If StrComp(cmb.List(outerIndex, 0), cmb.List(innerIndex, 0), vbTextCompare) > 0 Then
                displayValue = cmb.List(outerIndex, 0)
                loadValue = cmb.List(outerIndex, 1)

                cmb.List(outerIndex, 0) = cmb.List(innerIndex, 0)
                cmb.List(outerIndex, 1) = cmb.List(innerIndex, 1)

                cmb.List(innerIndex, 0) = displayValue
                cmb.List(innerIndex, 1) = loadValue
            End If
        Next innerIndex
    Next outerIndex

End Sub

Private Sub AddComboValue(ByVal cmb As Object, ByVal valueText As String)

    Dim itemIndex As Long

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then
        Exit Sub
    End If

    For itemIndex = 0 To cmb.ListCount - 1
        If StrComp(cmb.List(itemIndex), valueText, vbTextCompare) = 0 Then
            Exit Sub
        End If
    Next itemIndex

    cmb.AddItem valueText

End Sub

Private Sub SelectComboValue(ByVal cmb As Object, ByVal valueText As String)

    Dim itemIndex As Long

    valueText = Trim$(valueText)
    If Len(valueText) = 0 Then
        Exit Sub
    End If

    For itemIndex = 0 To cmb.ListCount - 1
        If StrComp(cmb.List(itemIndex), valueText, vbTextCompare) = 0 Then
            cmb.ListIndex = itemIndex
            Exit Sub
        End If
    Next itemIndex

    cmb.AddItem valueText
    cmb.ListIndex = cmb.ListCount - 1

End Sub
