Option Explicit

' Code-behind untuk UserForm dengan (Name) = JPEGSettings.
' Save menyimpan snapshot opsi JPEG untuk ekspor berikutnya.
Private pPresets As Collection
Private pSelectedPreset As JPEGPreset
Private pApplyingPreset As Boolean
Private pChoosingMatte As Boolean
Private pMatteRed As Long
Private pMatteGreen As Long
Private pMatteBlue As Long
Private pLastValidResolution As Double
Private pQueueItem As ExportSettingItem

Public Sub BeginQueueEdit(ByVal item As ExportSettingItem)
    Dim reader As JPEGPresetReader
    Dim saved As JPEGPreset
    Set reader = New JPEGPresetReader
    Set saved = reader.ParseXML(item.JPEGSettingsXML)
    If saved Is Nothing Then Err.Raise 5, "JPEGSettings", "Snapshot item bukan preset JPEG."
    RestoreBitmapSettingsSourcePath saved, item.JPEGSettingsXML, "JPEGSettings"
    ApplySavedSnapshot saved
    Set pQueueItem = item
End Sub

Private Sub UserForm_Initialize()
    Dim defaults As JPEGPreset
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo InitializeFailed
    operation = "Menyiapkan form dan Collection preset"
    Me.Caption = "JPEG Settings"
    pApplyingPreset = True
    Set pPresets = New Collection

    operation = "Menyiapkan cmbPreset"
    cmbPreset.Clear
    cmbPreset.Style = fmStyleDropDownList
    cmbPreset.Enabled = False

    operation = "PopulateBitmapColorModes / cmbColorMode"
    PopulateBitmapColorModes cmbColorMode, False, False, True, BITMAP_DEFAULT_COLOR_MODE
    operation = "PopulateSubFormats / cmbSubFormat"
    PopulateSubFormats
    operation = "PopulatePercentages / cmbQuality"
    PopulatePercentages cmbQuality, 80
    operation = "PopulatePercentages / cmbBlur"
    PopulatePercentages cmbBlur, 0
    operation = "PopulateBitmapResolutions / cmbResolution"
    PopulateBitmapResolutions cmbResolution, BITMAP_DEFAULT_RESOLUTION

    operation = "ConfigureBitmapMatteControl / cmbBgColor"
    ConfigureBitmapMatteControl cmbBgColor

    operation = "Membuat default JPEGPreset"
    Set defaults = New JPEGPreset
    operation = "ApplyPreset / default JPEG"
    ApplyPreset defaults
    pApplyingPreset = False
    operation = "LoadPresetList"
    LoadPresetList
    operation = "LoadSavedSettings"
    LoadSavedSettings
    Exit Sub

InitializeFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    pApplyingPreset = False
    ' Tampilkan error asli sebelum COM meneruskannya melalui UserForms.Add.
    MsgBox "Operasi: " & operation & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & _
        "Description: " & errorDescription, vbExclamation, "JPEGSettings - Initialize"
    Err.Raise errorNumber, "JPEGSettings.UserForm_Initialize", _
        operation & vbCrLf & "Source asli: " & errorSource & vbCrLf & errorDescription
End Sub

Private Sub LoadSavedSettings()
    Dim store As JPEGSettingsStore
    Dim saved As JPEGPreset
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo LoadFailed
    Set store = New JPEGSettingsStore
    Set saved = store.LoadSettings()
    If saved Is Nothing Then Exit Sub
    ApplySavedSnapshot saved
    Exit Sub

LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pApplyingPreset = False
    MsgBox "Gagal membaca pengaturan JPEG tersimpan (" & CStr(errorNumber) & "): " & errorDescription & vbCrLf & _
        "Atur opsi lalu klik Save untuk menyimpan ulang.", vbExclamation, "JPEG Settings"
End Sub

Private Sub ApplySavedSnapshot(ByVal saved As JPEGPreset)
    Dim preset As JPEGPreset
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

Private Sub PopulateSubFormats()
    cmbSubFormat.Clear
    cmbSubFormat.Style = fmStyleDropDownList
    cmbSubFormat.ColumnCount = 2
    cmbSubFormat.ColumnWidths = "120 pt;0 pt"
    cmbSubFormat.BoundColumn = 1
    AddKeyedComboItem cmbSubFormat, "Standard (4:2:2)", "standard"
    AddKeyedComboItem cmbSubFormat, "Optional (4:4:4)", "optional"
    cmbSubFormat.ListIndex = 0
End Sub

Private Sub LoadPresetList()
    Dim reader As JPEGPresetReader
    Dim presetIndex As Long
    Dim warningIndex As Long
    Dim warningText As String
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo LoadFailed
    Set reader = New JPEGPresetReader
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
        If Len(warningText) > 0 Then warningText = warningText & vbCrLf
        warningText = warningText & CStr(reader.Warnings.Item(warningIndex))
    Next warningIndex
    If Len(warningText) > 0 Then MsgBox warningText, vbExclamation, "JPEG Preset"
    Exit Sub

LoadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pApplyingPreset = False
    cmbPreset.Enabled = False
    MsgBox "Gagal membaca preset JPEG (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "JPEG Settings"
End Sub

Private Sub cmbPreset_Change()
    Dim preset As JPEGPreset

    If pApplyingPreset Then Exit Sub
    If cmbPreset.ListIndex < 0 Then Exit Sub
    On Error GoTo ApplyFailed
    Set preset = pPresets.Item(cmbPreset.ListIndex + 1)
    ApplyPreset preset
    Exit Sub

ApplyFailed:
    MsgBox "Gagal menerapkan preset JPEG (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "JPEG Settings"
End Sub

Private Sub ApplyPreset(ByVal preset As JPEGPreset)
    Dim wasApplying As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    wasApplying = pApplyingPreset
    pApplyingPreset = True
    On Error GoTo ApplyFailed
    Set pSelectedPreset = preset

    ApplyCommonBitmapPreset Me, preset, pLastValidResolution
    If Not SelectKeyedComboItem(cmbSubFormat, preset.SubFormat) Then Err.Raise 5, "JPEGSettings", "Subformat preset tidak tersedia."

    cmbQuality.ListIndex = preset.Quality
    cmbBlur.ListIndex = preset.Blur
    chkOptimize.Value = preset.Optimize
    chkProgressive.Value = preset.Progressive
    optDocColorSetting.Value = Not preset.UseColorProof
    optColorProofSetting.Value = preset.UseColorProof
    ApplyMatteDisplay preset
    cmbPreset.ControlTipText = preset.SourcePath
    pApplyingPreset = wasApplying
    Exit Sub

ApplyFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pApplyingPreset = wasApplying
    Err.Raise errorNumber, "JPEGSettings.ApplyPreset", errorDescription
End Sub

Private Sub ApplyMatteDisplay(ByVal preset As JPEGPreset)
    pMatteRed = preset.MatteRed
    pMatteGreen = preset.MatteGreen
    pMatteBlue = preset.MatteBlue
    UpdateBitmapMatteControl cmbBgColor, True, pMatteRed, pMatteGreen, pMatteBlue
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
    pChoosingMatte = True
    On Error GoTo ColorFailed
    If ChooseBitmapMatteColor(pMatteRed, pMatteGreen, pMatteBlue, operation) Then
        UpdateBitmapMatteControl cmbBgColor, True, pMatteRed, pMatteGreen, pMatteBlue
    End If
    pChoosingMatte = False
    Exit Sub

ColorFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    pChoosingMatte = False
    MsgBox "Gagal memilih warna matte pada " & operation & " (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "JPEG Settings"
End Sub

Private Sub cmbResolution_AfterUpdate()
    Dim dpi As Double

    If pApplyingPreset Then Exit Sub
    If TryBitmapResolution(cmbResolution.Text, dpi) Then
        pLastValidResolution = dpi
    Else
        MsgBox "Resolusi harus bilangan bulat 1 sampai 1200 dpi.", vbExclamation, "JPEG Settings"
        cmbResolution.Value = CStr(pLastValidResolution)
    End If
End Sub

' Nilai kontrol mengungguli SourceXML asli tanpa mengubah file preset Corel.
Public Function ReadCurrentSettings() As JPEGPreset
    Dim settings As JPEGPreset
    Dim dpi As Double

    If Not TryBitmapResolution(cmbResolution.Text, dpi) Then
        Err.Raise 5, "JPEGSettings", "Resolusi harus bilangan bulat 1 sampai 1200 dpi."
    End If
    If cmbColorMode.ListIndex < 0 Then Err.Raise 5, "JPEGSettings", "Pilih color mode JPEG."
    If cmbSubFormat.ListIndex < 0 Then Err.Raise 5, "JPEGSettings", "Pilih subformat JPEG."
    If cmbQuality.ListIndex < 0 Then Err.Raise 5, "JPEGSettings", "Pilih quality JPEG."
    If cmbBlur.ListIndex < 0 Then Err.Raise 5, "JPEGSettings", "Pilih blur JPEG."
    If Not CBool(optDocColorSetting.Value) And Not CBool(optColorProofSetting.Value) Then
        Err.Raise 5, "JPEGSettings", "Pilih document atau color proof settings."
    End If

    Set settings = New JPEGPreset
    If Not pSelectedPreset Is Nothing Then
        settings.DisplayName = pSelectedPreset.DisplayName
        settings.SourcePath = pSelectedPreset.SourcePath
        settings.SourceXML = pSelectedPreset.SourceXML
    End If
    settings.ColorMode = CStr(cmbColorMode.List(cmbColorMode.ListIndex, 1))
    settings.Antialiased = CBool(chkAntialiased.Value)
    settings.EmbedColorProfile = CBool(chkEmbedColorProfile.Value)
    settings.CropToPage = CBool(chkCropToPageOnExport.Value)
    settings.Resolution = dpi
    settings.SubFormat = CStr(cmbSubFormat.List(cmbSubFormat.ListIndex, 1))
    settings.Quality = cmbQuality.ListIndex
    settings.Blur = cmbBlur.ListIndex
    settings.UseColorProof = CBool(optColorProofSetting.Value)
    settings.Optimize = CBool(chkOptimize.Value)
    settings.Progressive = CBool(chkProgressive.Value)
    settings.MatteRed = pMatteRed
    settings.MatteGreen = pMatteGreen
    settings.MatteBlue = pMatteBlue
    Set ReadCurrentSettings = settings
End Function

Private Sub cmdClose_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    Dim store As JPEGSettingsStore
    Dim settings As JPEGPreset
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo SaveFailed
    Set settings = ReadCurrentSettings()
    Set store = New JPEGSettingsStore
    If pQueueItem Is Nothing Then
        store.SaveSettings settings
    Else
        pQueueItem.JPEGSettingsXML = store.SerializeSettings(settings)
    End If
    Unload Me
    Exit Sub

SaveFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    MsgBox "Gagal menyimpan pengaturan JPEG (" & CStr(errorNumber) & "): " & errorDescription, vbExclamation, "JPEG Settings"
End Sub
