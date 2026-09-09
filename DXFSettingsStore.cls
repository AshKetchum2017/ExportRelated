Option Explicit

' Class Module (Name): DXFSettingsStore.
Private Const REG_DXF_SETTINGS_KEY As String = "DXFSettingsXML"

Public Function LoadSettings() As DXFExportSettings
    Dim savedXML As String
    savedXML = LoadBitmapSettingsXML(REG_DXF_SETTINGS_KEY)
    If Len(savedXML) = 0 Then Exit Function
    Set LoadSettings = ParseXML(savedXML)
End Function

Public Sub SaveSettings(ByVal settings As DXFExportSettings)
    SaveBitmapSettingsXML REG_DXF_SETTINGS_KEY, SerializeSettings(settings)
End Sub

Public Sub ValidateSettings(ByVal settings As DXFExportSettings)
    If settings Is Nothing Then Err.Raise 5, "DXFSettingsStore", "Pengaturan DXF kosong."
    Select Case settings.AutoCADVersion
        Case 23, 18, 15, 12, 9, 1, 2, 3, 4, 5, 6, 7, 8
        Case Else: Err.Raise 5, "DXFSettingsStore", "Versi AutoCAD tidak valid."
    End Select
    If settings.AutoCADUnits < 0 Or settings.AutoCADUnits > 18 Then Err.Raise 5, "DXFSettingsStore", "Unit AutoCAD tidak valid."
    If settings.BitmapType < 0 Or settings.BitmapType > 3 Then Err.Raise 5, "DXFSettingsStore", "Format bitmap DXF tidak valid."
    Select Case settings.CurvesAs
        Case "polycurves", "splines"
        Case Else: Err.Raise 5, "DXFSettingsStore", "Pilihan kurva DXF tidak valid."
    End Select
    If settings.Tolerance < 0 Or settings.Tolerance > 100 Then Err.Raise 5, "DXFSettingsStore", "Tolerance harus 0 sampai 100%."
    If settings.FillRed < 0 Or settings.FillRed > 255 Or _
       settings.FillGreen < 0 Or settings.FillGreen > 255 Or _
       settings.FillBlue < 0 Or settings.FillBlue > 255 Then
        Err.Raise 5, "DXFSettingsStore", "Komponen warna DXF harus 0 sampai 255."
    End If
End Sub

Public Function ParseXML(ByVal savedXML As String) As DXFExportSettings
    Dim xml As Object
    Dim root As Object
    Dim settings As DXFExportSettings
    Set xml = ParseExportPresetXML(savedXML, "DXFSettingsStore")
    Set root = ExportPresetRoot(xml, "DXFSettingsStore")
    If ExportPresetText(root, "format", "") <> "dxf" Then Err.Raise 5, "DXFSettingsStore", "Snapshot bukan pengaturan DXF."
    Set settings = New DXFExportSettings
    settings.DisplayName = ExportPresetText(root, "name", settings.DisplayName)
    settings.AutoCADVersion = ExportPresetInteger(root, "version", settings.AutoCADVersion, 0, 28, "DXFSettingsStore")
    settings.AutoCADUnits = ExportPresetInteger(root, "units", settings.AutoCADUnits, 0, 18, "DXFSettingsStore")
    settings.BitmapType = ExportPresetInteger(root, "bitmap", settings.BitmapType, 0, 3, "DXFSettingsStore")
    settings.CurvesAs = ExportPresetText(root, "curves", settings.CurvesAs)
    settings.Tolerance = ExportPresetInteger(root, "tolerance", settings.Tolerance, 0, 100, "DXFSettingsStore")
    settings.TextAsCurves = ExportPresetBoolean(root, "textAsCurves", settings.TextAsCurves, "DXFSettingsStore")
    settings.FillUnmapped = ExportPresetBoolean(root, "fillUnmapped", settings.FillUnmapped, "DXFSettingsStore")
    settings.FillRed = ExportPresetInteger(root, "red", settings.FillRed, 0, 255, "DXFSettingsStore")
    settings.FillGreen = ExportPresetInteger(root, "green", settings.FillGreen, 0, 255, "DXFSettingsStore")
    settings.FillBlue = ExportPresetInteger(root, "blue", settings.FillBlue, 0, 255, "DXFSettingsStore")
    ValidateSettings settings
    Set ParseXML = settings
End Function

Public Function SerializeSettings(ByVal settings As DXFExportSettings) As String
    Dim xml As Object
    Dim root As Object
    ValidateSettings settings
    Set xml = ParseExportPresetXML("<preset format=""dxf""/>", "DXFSettingsStore")
    Set root = xml.documentElement
    root.setAttribute "name", settings.DisplayName
    root.setAttribute "version", CStr(settings.AutoCADVersion)
    root.setAttribute "units", CStr(settings.AutoCADUnits)
    root.setAttribute "bitmap", CStr(settings.BitmapType)
    root.setAttribute "curves", settings.CurvesAs
    root.setAttribute "tolerance", CStr(settings.Tolerance)
    root.setAttribute "textAsCurves", LCase$(CStr(settings.TextAsCurves))
    root.setAttribute "fillUnmapped", LCase$(CStr(settings.FillUnmapped))
    root.setAttribute "red", CStr(settings.FillRed)
    root.setAttribute "green", CStr(settings.FillGreen)
    root.setAttribute "blue", CStr(settings.FillBlue)
    SerializeSettings = root.xml
End Function
