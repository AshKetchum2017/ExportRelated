Option Explicit

Private Const REG_PNG_SETTINGS_KEY As String = "PNGSettingsXML"

Public Function LoadSettings() As PNGPreset
    Dim savedXML As String
    Dim reader As PNGPresetReader
    Dim settings As PNGPreset

    savedXML = LoadBitmapSettingsXML(REG_PNG_SETTINGS_KEY)
    If Len(savedXML) = 0 Then Exit Function
    Set reader = New PNGPresetReader
    Set settings = reader.ParseXML(savedXML)
    If settings Is Nothing Then Err.Raise 5, "PNGSettingsStore", "Pengaturan tersimpan bukan preset PNG."
    ValidateSettings settings
    RestoreBitmapSettingsSourcePath settings, savedXML, "PNGSettingsStore"
    Set LoadSettings = settings
End Function

Public Sub SaveSettings(ByVal settings As PNGPreset)
    Dim snapshotXML As String
    snapshotXML = SerializeSettings(settings)
    ' Satu nilai registry agar kegagalan simpan tidak meninggalkan opsi campuran.
    SaveBitmapSettingsXML REG_PNG_SETTINGS_KEY, snapshotXML
End Sub

Public Sub ValidateSettings(ByVal settings As PNGPreset)
    ValidateCommonBitmapSettings settings, "PNGSettingsStore"
    Select Case settings.ColorMode
        Case "bw", "grayscale", "paletted", "rgb"
        Case Else: Err.Raise 5, "PNGSettingsStore", "Color mode PNG tidak valid."
    End Select
End Sub

Public Function SerializeSettings(ByVal settings As PNGPreset) As String
    Dim xml As Object
    Dim root As Object
    Dim reader As PNGPresetReader
    Dim checked As PNGPreset

    ValidateSettings settings
    Set xml = CreateBitmapSettingsSnapshot(settings, "png", "PNGSettingsStore")
    Set root = xml.documentElement
    root.setAttribute "transparency", LCase$(CStr(settings.Transparency))
    root.setAttribute "interlaced", LCase$(CStr(settings.Interlaced))

    Set reader = New PNGPresetReader
    Set checked = reader.ParseXML(root.xml)
    ValidateSettings checked
    SerializeSettings = root.xml
End Function
