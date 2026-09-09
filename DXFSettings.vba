Option Explicit

' Code-behind UserForm (Name) = DXFSettings. Kontrol sesuai DXFSettings.md.
Private pQueueItem As ExportSettingItem
Private pSettings As DXFExportSettings
Private pApplying As Boolean
Private pChoosingColor As Boolean

Private Sub UserForm_Initialize()
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String
    On Error GoTo InitializeFailed
    pApplying = True
    Me.Caption = "DXF Settings"
    operation = "Populate ComboBox DXF"
    PopulateKeyedChoices cmbAutoCADVersion, Array( _
        "AutoCAD 2018-2023 (DWG/DXF R2018)", "AutoCAD 2013-2017 (DWG/DXF R2013)", _
        "AutoCAD 2010-2012 (DWG/DXF R2010)", "AutoCAD 2007-2009 (DWG/DXF R2007)", _
        "AutoCAD 2004-2006 (DWG/DXF 2004)", "AutoCAD 2000-2002 (DWG/DXF 2000)", _
        "AutoCAD R14 (DWG/DXF R14)", "AutoCAD R13 (DWG/DXF R13)", _
        "AutoCAD R11 (DWG/DXF R11)", "AutoCAD R10 (DWG/DXF R10)", _
        "AutoCAD R9 (DWG/DXF R9)", "AutoCAD R2.6 (DWG/DXF R2.60)", _
        "AutoCAD R2.5 (DWG/DXF R2.50)"), Array(23, 18, 15, 12, 9, 1, 2, 3, 4, 5, 6, 7, 8)
    ' Urutan tampilan Miles/Feet berbeda dari enum DxfUnits.
    PopulateKeyedChoices cmbAutoCADUnits, Array( _
        "Inches", "Miles", "Feet", "Milimeters", "Centimeters", "Meters", "Kilometers", _
        "Microinches", "Mils", "Yards", "Angstroms", "Nanometers", "Microns", _
        "Decimeters", "Hectometers", "Gigameters", "Astronomical Units", "Light Years", "Parsecs"), _
        Array(0, 2, 1, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18)
    PopulateKeyedChoices cmbAutoCADBitmapAs, Array("JPEG Files", "GIF Files", "PNG Files", "BMP Files"), Array(0, 1, 2, 3)
    PopulateKeyedChoices cmbAutoCADCurvesAs, Array("Polycurves", "Splines"), Array("polycurves", "splines")
    Set pSettings = New DXFExportSettings
    PopulatePercentages cmbAutoCADTolerance, pSettings.Tolerance
    operation = "Menyiapkan warna dan grup OptionButton"
    ConfigureBitmapMatteControl cmbAutoCADColor
    optTextAsCurves.GroupName = "DXFText"
    optTextAsText.GroupName = "DXFText"
    optAutoCADColor.GroupName = "DXFFill"
    optAutoCADUnfilled.GroupName = "DXFFill"
    cmbAutoCADCurvesAs.ControlTipText = "Tersimpan; penerapan pilihan kurva ke ekspor belum tersedia."
    cmbAutoCADTolerance.ControlTipText = "Tersimpan; penerapan tolerance ke ekspor belum tersedia."
    operation = "Menerapkan default DXF"
    ApplySettings pSettings
    operation = "Membaca pengaturan DXF tersimpan"
    LoadSavedSettings
    Exit Sub
InitializeFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    pApplying = False
    MsgBox "Operasi: " & operation & vbCrLf & "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & "Description: " & errorDescription, vbExclamation, "DXFSettings - Initialize"
    Err.Raise errorNumber, "DXFSettings.UserForm_Initialize", operation & ": " & errorDescription
End Sub

Private Sub LoadSavedSettings()
    Dim store As DXFSettingsStore
    Dim saved As DXFExportSettings
    On Error GoTo LoadFailed
    Set store = New DXFSettingsStore
    Set saved = store.LoadSettings()
    If Not saved Is Nothing Then ApplySettings saved
    Exit Sub
LoadFailed:
    pApplying = False
    MsgBox "Gagal membaca pengaturan DXF tersimpan (" & CStr(Err.Number) & "): " & Err.Description & vbCrLf & _
        "Atur opsi lalu klik Save untuk menyimpan ulang.", vbExclamation, "DXF Settings"
End Sub

Public Sub BeginQueueEdit(ByVal item As ExportSettingItem)
    ApplySettings item.ReadDXFSnapshot()
    Set pQueueItem = item
End Sub

Private Sub ApplySettings(ByVal settings As DXFExportSettings)
    Dim store As DXFSettingsStore
    Set store = New DXFSettingsStore
    store.ValidateSettings settings
    pApplying = True
    Set pSettings = settings
    If Not SelectKeyedComboItem(cmbAutoCADVersion, CStr(settings.AutoCADVersion)) Then Err.Raise 5, "DXFSettings", "Versi AutoCAD tidak tersedia."
    If Not SelectKeyedComboItem(cmbAutoCADUnits, CStr(settings.AutoCADUnits)) Then Err.Raise 5, "DXFSettings", "Unit AutoCAD tidak tersedia."
    If Not SelectKeyedComboItem(cmbAutoCADBitmapAs, CStr(settings.BitmapType)) Then Err.Raise 5, "DXFSettings", "Format bitmap tidak tersedia."
    If Not SelectKeyedComboItem(cmbAutoCADCurvesAs, settings.CurvesAs) Then Err.Raise 5, "DXFSettings", "Pilihan kurva tidak tersedia."
    cmbAutoCADTolerance.ListIndex = settings.Tolerance
    optTextAsCurves.Value = settings.TextAsCurves
    optTextAsText.Value = Not settings.TextAsCurves
    optAutoCADColor.Value = settings.FillUnmapped
    optAutoCADUnfilled.Value = Not settings.FillUnmapped
    pApplying = False
    SyncEnabledControls
End Sub

Private Sub SyncEnabledControls()
    If pApplying Then Exit Sub
    cmbAutoCADTolerance.Enabled = (cmbAutoCADCurvesAs.ListIndex = 0)
    UpdateBitmapMatteControl cmbAutoCADColor, CBool(optAutoCADColor.Value), _
        pSettings.FillRed, pSettings.FillGreen, pSettings.FillBlue
    cmbAutoCADColor.ControlTipText = "Unmapped fill RGB (" & CStr(pSettings.FillRed) & ", " & _
        CStr(pSettings.FillGreen) & ", " & CStr(pSettings.FillBlue) & "). Klik untuk memilih warna."
End Sub

Private Sub cmbAutoCADCurvesAs_Change()
    SyncEnabledControls
End Sub

Private Sub optAutoCADColor_Click()
    SyncEnabledControls
End Sub

Private Sub optAutoCADUnfilled_Click()
    SyncEnabledControls
End Sub

Private Sub cmbAutoCADColor_MouseDown(ByVal Button As Integer, ByVal Shift As Integer, ByVal X As Single, ByVal Y As Single)
    If Button = 1 Then ChooseFillColor
End Sub

Private Sub cmbAutoCADColor_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeySpace Or KeyCode = vbKeyReturn Or KeyCode = vbKeyF4 Then
        KeyCode = 0
        ChooseFillColor
    End If
End Sub

Private Sub ChooseFillColor()
    Dim operation As String
    Dim red As Long
    Dim green As Long
    Dim blue As Long
    If pApplying Or pChoosingColor Then Exit Sub
    If Not cmbAutoCADColor.Enabled Then Exit Sub
    On Error GoTo ColorFailed
    pChoosingColor = True
    red = pSettings.FillRed
    green = pSettings.FillGreen
    blue = pSettings.FillBlue
    If ChooseBitmapMatteColor(red, green, blue, operation) Then
        pSettings.FillRed = red
        pSettings.FillGreen = green
        pSettings.FillBlue = blue
        SyncEnabledControls
    End If
    pChoosingColor = False
    Exit Sub
ColorFailed:
    pChoosingColor = False
    MsgBox "Gagal memilih warna pada " & operation & " (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "DXF Settings"
End Sub

Public Function ReadCurrentSettings() As DXFExportSettings
    Dim result As DXFExportSettings
    Dim store As DXFSettingsStore
    If cmbAutoCADVersion.ListIndex < 0 Or cmbAutoCADUnits.ListIndex < 0 Or _
       cmbAutoCADBitmapAs.ListIndex < 0 Or cmbAutoCADCurvesAs.ListIndex < 0 Or cmbAutoCADTolerance.ListIndex < 0 Then
        Err.Raise 5, "DXFSettings", "Lengkapi seluruh pilihan DXF."
    End If
    If CBool(optTextAsCurves.Value) = CBool(optTextAsText.Value) Then Err.Raise 5, "DXFSettings", "Pilih satu mode text DXF."
    If CBool(optAutoCADColor.Value) = CBool(optAutoCADUnfilled.Value) Then Err.Raise 5, "DXFSettings", "Pilih satu mode unmapped fills."
    Set result = New DXFExportSettings
    result.DisplayName = "Custom"
    result.AutoCADVersion = CLng(cmbAutoCADVersion.List(cmbAutoCADVersion.ListIndex, 1))
    result.AutoCADUnits = CLng(cmbAutoCADUnits.List(cmbAutoCADUnits.ListIndex, 1))
    result.BitmapType = CLng(cmbAutoCADBitmapAs.List(cmbAutoCADBitmapAs.ListIndex, 1))
    result.CurvesAs = CStr(cmbAutoCADCurvesAs.List(cmbAutoCADCurvesAs.ListIndex, 1))
    result.Tolerance = cmbAutoCADTolerance.ListIndex
    result.TextAsCurves = CBool(optTextAsCurves.Value)
    result.FillUnmapped = CBool(optAutoCADColor.Value)
    result.FillRed = pSettings.FillRed
    result.FillGreen = pSettings.FillGreen
    result.FillBlue = pSettings.FillBlue
    Set store = New DXFSettingsStore
    store.ValidateSettings result
    Set ReadCurrentSettings = result
End Function

Private Sub cmdSave_Click()
    Dim store As DXFSettingsStore
    Dim settings As DXFExportSettings
    On Error GoTo SaveFailed
    Set settings = ReadCurrentSettings()
    Set store = New DXFSettingsStore
    If pQueueItem Is Nothing Then
        store.SaveSettings settings
    Else
        pQueueItem.DXFSettingsXML = store.SerializeSettings(settings)
    End If
    Unload Me
    Exit Sub
SaveFailed:
    MsgBox "Gagal menyimpan pengaturan DXF (" & CStr(Err.Number) & "): " & Err.Description, vbExclamation, "DXF Settings"
End Sub

Private Sub cmdClose_Click()
    Unload Me
End Sub
