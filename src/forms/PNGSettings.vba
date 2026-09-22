Option Explicit

' Code-behind untuk UserForm dengan (Name) = PNGSettings.
' Save menyimpan snapshot opsi PNG untuk ekspor berikutnya.
Private pPresets As Collection
Private pSelectedPreset As PNGPreset
Private pApplyingPreset As Boolean
Private pChoosingMatte As Boolean
Private pMatteRed As Long
Private pMatteGreen As Long
Private pMatteBlue As Long
Private pLastValidResolution As Double
Private pQueueItem As ExportSettingItem

Public Sub BeginQueueEdit(ByVal item As ExportSettingItem)
    Dim reader As PNGPresetReader
    Dim saved As PNGPreset
    Set reader = New PNGPresetReader
    Set saved = reader.ParseXML(item.PNGSettingsXML)
    If saved Is Nothing Then Err.Raise 5, "PNGSettings", "Snapshot item bukan preset PNG."
    RestoreBitmapSettingsSourcePath saved, item.PNGSettingsXML, "PNGSettings"
    ApplySavedSnapshot saved
    Set pQueueItem = item
End Sub

Private Sub UserForm_Initialize()

    Dim defaults As PNGPreset

    Me.Caption = "PNG Settings"
    pApplyingPreset = True
    Set pPresets = New Collection

    cmbPreset.Clear
    cmbPreset.Style = fmStyleDropDownList
    cmbPreset.Enabled = False

    PopulateBitmapColorModes cmbColorMode, True, True, False, BITMAP_DEFAULT_COLOR_MODE
    PopulateBitmapResolutions cmbResolution, BITMAP_DEFAULT_RESOLUTION

    ' MSForms tidak menggambar swatch per item. Gunakan bidang warna yang diklik.
    ConfigureBitmapMatteControl cmbBgColor

    cmdSave.Enabled = True
    Set defaults = New PNGPreset
    ApplyPreset defaults
    pApplyingPreset = False
    LoadPresetList
    LoadSavedSettings

End Sub

Private Sub LoadSavedSettings()
    Dim store As PNGSettingsStore
    Dim saved As PNGPreset
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo LoadFailed
    Set store = New PNGSettingsStore
    Set saved = store.LoadSettings()
    If saved Is Nothing Then Exit Sub
    ApplySavedSnapshot saved
    Exit Sub

LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pApplyingPreset = False
    MsgBox "Gagal membaca pengaturan PNG tersimpan (" & CStr(errorNumber) & "): " & errorDescription & vbCrLf & _
        "Atur opsi lalu klik Save untuk menyimpan ulang.", vbExclamation, "PNG Settings"
End Sub

Private Sub ApplySavedSnapshot(ByVal saved As PNGPreset)
    Dim preset As PNGPreset
    Dim presetIndex As Long

    pApplyingPreset = True
    cmbPreset.ListIndex = -1
    For presetIndex = 1 To pPresets.Count
        Set preset = pPresets.Item(presetIndex)
        If Len(saved.SourcePath) > 0 Then
            If StrComp(preset.SourcePath, saved.SourcePath, vbTextCompare) = 0 Then
                cmbPreset.ListIndex = presetIndex - 1
                Exit For
            End If
        End If
    Next presetIndex
    ' Snapshot tersimpan mengungguli XML sumber yang mungkin sudah berubah/hilang.
    ApplyPreset saved
    pApplyingPreset = False
End Sub

Private Sub chkTransparency_Click()
    If pApplyingPreset Then Exit Sub
    UpdateMatteControl
End Sub

Private Sub LoadPresetList()
    Dim reader As PNGPresetReader
    Dim presetIndex As Long
    Dim warningIndex As Long
    Dim warningText As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo LoadFailed
    Set reader = New PNGPresetReader
    Set pPresets = reader.LoadPresets()
    pApplyingPreset = True
    cmbPreset.Clear
    For presetIndex = 1 To pPresets.Count
        cmbPreset.AddItem ExportPresetDisplayName(pPresets, presetIndex)
    Next presetIndex
    cmbPreset.Enabled = (pPresets.Count > 0)
    pApplyingPreset = False

    If pPresets.Count > 0 Then cmbPreset.ListIndex = 0
    For warningIndex = 1 To reader.Warnings.Count
        If warningIndex > 5 Then
            warningText = warningText & vbCrLf & "Dan " & CStr(reader.Warnings.Count - 5) & " peringatan lainnya."
            Exit For
        End If
        warningText = warningText & vbCrLf & CStr(reader.Warnings.Item(warningIndex))
    Next warningIndex
    If pPresets.Count = 0 Then warningText = "Tidak ada preset PNG yang dapat dimuat. Pengaturan manual tetap tersedia." & vbCrLf & warningText
    If Len(warningText) > 0 Then MsgBox Trim$(warningText), vbExclamation, "PNG Settings"
    Exit Sub

LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pApplyingPreset = False
    cmbPreset.Enabled = False
    MsgBox "Gagal memuat daftar preset PNG (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "PNG Settings"
End Sub

Private Sub cmbPreset_Change()
    Dim preset As PNGPreset

    If pApplyingPreset Then Exit Sub
    If cmbPreset.ListIndex < 0 Then Exit Sub
    On Error GoTo ApplyFailed
    Set preset = pPresets.Item(cmbPreset.ListIndex + 1)
    ApplyPreset preset
    Exit Sub

ApplyFailed:
    MsgBox "Gagal menerapkan preset PNG (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "PNG Settings"
End Sub

Private Sub ApplyPreset(ByVal preset As PNGPreset)
    Dim wasApplying As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    wasApplying = pApplyingPreset
    pApplyingPreset = True
    On Error GoTo ApplyFailed
    ApplyCommonBitmapPreset Me, preset, pLastValidResolution

    chkTransparency.Value = preset.Transparency
    chkInterlaced.Value = preset.Interlaced
    pMatteRed = preset.MatteRed
    pMatteGreen = preset.MatteGreen
    pMatteBlue = preset.MatteBlue
    UpdateMatteControl
    cmbPreset.ControlTipText = preset.SourcePath
    ' Hanya menyimpan referensi sumber; perubahan kontrol tidak memutasi preset ini.
    Set pSelectedPreset = preset
    pApplyingPreset = wasApplying
    Exit Sub

ApplyFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pApplyingPreset = wasApplying
    Err.Raise errorNumber, "PNGSettings.ApplyPreset", errorDescription
End Sub

Private Sub UpdateMatteControl()
    UpdateBitmapMatteControl cmbBgColor, Not CBool(chkTransparency.Value), pMatteRed, pMatteGreen, pMatteBlue
End Sub

Private Sub cmbBgColor_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then ChooseMatteColor
End Sub

Private Sub cmbBgColor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyF4 Then
        KeyCode = 0
        ChooseMatteColor
    End If
End Sub

Private Sub ChooseMatteColor()
    Dim operation As String
    Dim errorNumber As Long
    Dim errorDescription As String

    If pApplyingPreset Or pChoosingMatte Then Exit Sub
    If CBool(chkTransparency.Value) Then Exit Sub
    pChoosingMatte = True
    On Error GoTo ColorFailed
    If ChooseBitmapMatteColor(pMatteRed, pMatteGreen, pMatteBlue, operation) Then UpdateMatteControl
    pChoosingMatte = False
    Exit Sub

ColorFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pChoosingMatte = False
    MsgBox "Gagal memilih warna matte pada " & operation & " (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "PNG Settings"
End Sub

Private Sub cmbResolution_AfterUpdate()
    Dim dpi As Double

    If pApplyingPreset Then Exit Sub
    If TryBitmapResolution(cmbResolution.Text, dpi) Then
        pLastValidResolution = dpi
    Else
        MsgBox "Resolusi harus bilangan bulat 1 sampai 1200 dpi.", vbExclamation, "PNG Settings"
        cmbResolution.Value = CStr(pLastValidResolution)
    End If
End Sub

' Nilai kontrol mengungguli SourceXML asli tanpa mengubah file preset Corel.
Public Function ReadCurrentSettings() As PNGPreset
    Dim settings As PNGPreset
    Dim dpi As Double

    If Not TryBitmapResolution(cmbResolution.Text, dpi) Then
        Err.Raise 5, "PNGSettings", "Resolusi harus bilangan bulat 1 sampai 1200 dpi."
    End If
    If cmbColorMode.ListIndex < 0 Then Err.Raise 5, "PNGSettings", "Pilih color mode PNG."
    Set settings = New PNGPreset
    If Not pSelectedPreset Is Nothing Then
        settings.DisplayName = pSelectedPreset.DisplayName
        settings.SourcePath = pSelectedPreset.SourcePath
        settings.SourceXML = pSelectedPreset.SourceXML
    End If
    settings.ColorMode = CStr(cmbColorMode.List(cmbColorMode.ListIndex, 1))
    settings.Transparency = CBool(chkTransparency.Value)
    settings.Antialiased = CBool(chkAntialiased.Value)
    settings.EmbedColorProfile = CBool(chkEmbedColorProfile.Value)
    settings.Interlaced = CBool(chkInterlaced.Value)
    settings.CropToPage = CBool(chkCropToPageOnExport.Value)
    settings.Resolution = dpi
    settings.MatteRed = pMatteRed
    settings.MatteGreen = pMatteGreen
    settings.MatteBlue = pMatteBlue
    Set ReadCurrentSettings = settings
End Function

Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdSave_Click()
    Dim store As PNGSettingsStore
    Dim settings As PNGPreset
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo SaveFailed
    Set settings = ReadCurrentSettings()
    Set store = New PNGSettingsStore
    If pQueueItem Is Nothing Then
        store.SaveSettings settings
    Else
        pQueueItem.PNGSettingsXML = store.SerializeSettings(settings)
    End If
    Unload Me
    Exit Sub

SaveFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MsgBox "Gagal menyimpan pengaturan PNG (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "PNG Settings"
End Sub
