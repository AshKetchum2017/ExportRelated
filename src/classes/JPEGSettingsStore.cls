Option Explicit

' Class Module (Name): JPEGSettingsStore.
Private Const REG_JPEG_SETTINGS_KEY As String = "JPEGSettingsXML"

Public Function LoadSettings() As JPEGPreset
    Dim savedXML As String
    Dim reader As JPEGPresetReader
    Dim settings As JPEGPreset

    savedXML = LoadBitmapSettingsXML(REG_JPEG_SETTINGS_KEY)
    If Len(savedXML) = 0 Then Exit Function
    Set reader = New JPEGPresetReader
    Set settings = reader.ParseXML(savedXML)
    If settings Is Nothing Then Err.Raise 5, "JPEGSettingsStore", "Pengaturan tersimpan bukan preset JPEG."
    ValidateSettings settings
    RestoreBitmapSettingsSourcePath settings, savedXML, "JPEGSettingsStore"
    Set LoadSettings = settings
End Function

Public Sub SaveSettings(ByVal settings As JPEGPreset)
    Dim snapshotXML As String

    snapshotXML = SerializeSettings(settings)
    ' Satu nilai registry agar kegagalan simpan tidak meninggalkan opsi campuran.
    SaveBitmapSettingsXML REG_JPEG_SETTINGS_KEY, snapshotXML
End Sub

Public Sub ValidateSettings(ByVal settings As JPEGPreset)
    ValidateCommonBitmapSettings settings, "JPEGSettingsStore"
    Select Case settings.ColorMode
        Case "grayscale", "rgb", "cmyk"
        Case Else: Err.Raise 5, "JPEGSettingsStore", "Color mode JPEG tidak valid."
    End Select
    Select Case settings.SubFormat
        Case "standard", "optional"
        Case Else: Err.Raise 5, "JPEGSettingsStore", "Subformat JPEG tidak valid."
    End Select
    If settings.Quality < 0 Or settings.Quality > 100 Then
        Err.Raise 5, "JPEGSettingsStore", "Quality JPEG harus 0 sampai 100%."
    End If
    If settings.Blur < 0 Or settings.Blur > 100 Then
        Err.Raise 5, "JPEGSettingsStore", "Blur JPEG harus 0 sampai 100%."
    End If
End Sub

Public Function SerializeSettings(ByVal settings As JPEGPreset) As String
    Dim xml As Object
    Dim root As Object
    Dim child As Object
    Dim reader As JPEGPresetReader
    Dim checked As JPEGPreset

    ValidateSettings settings
    Set xml = CreateBitmapSettingsSnapshot(settings, "jpg", "JPEGSettingsStore")
    Set root = xml.documentElement
    root.setAttribute "subformat", settings.SubFormat
    root.setAttribute "qualityamount", CStr(settings.Quality)
    root.setAttribute "bluramount", CStr(settings.Blur)
    root.setAttribute "optimize", LCase$(CStr(settings.Optimize))
    root.setAttribute "progressive", LCase$(CStr(settings.Progressive))
    Set child = EnsureExportPresetChild(xml, root, "colorWorkflow")
    If settings.UseColorProof Then
        child.setAttribute "colorSettings", "Proof"
    Else
        child.setAttribute "colorSettings", "Document"
    End If

    Set reader = New JPEGPresetReader
    Set checked = reader.ParseXML(root.xml)
    ValidateSettings checked
    SerializeSettings = root.xml
End Function
