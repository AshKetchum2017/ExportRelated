Option Explicit

' Class Module (Name): PNGPresetReader.
' Parsing field PNG tetap di sini; discovery dan XML primitives dipakai bersama.
Private pWarnings As Collection

Private Sub Class_Initialize()
    Set pWarnings = New Collection
End Sub

Public Property Get Warnings() As Collection
    Set Warnings = pWarnings
End Property

' Entry point lama dipertahankan untuk caller yang sudah ada.
Public Function DiscoverDirectories() As Collection
    Set DiscoverDirectories = DiscoverExportPresetDirectories()
End Function

Public Function LoadPresets() As Collection
    Dim presets As Collection
    Dim presetFiles As Collection
    Dim preset As PNGPreset
    Dim fileIndex As Long
    Dim filePath As String
    Dim errorNumber As Long
    Dim errorDescription As String

    Set pWarnings = New Collection
    Set presets = New Collection
    Set presetFiles = ListExportPresetXMLFiles(pWarnings)
    For fileIndex = 1 To presetFiles.Count
        filePath = CStr(presetFiles.Item(fileIndex))
        Set preset = Nothing
        On Error Resume Next
        Set preset = LoadFile(filePath)
        errorNumber = Err.Number
        errorDescription = Err.Description
        On Error GoTo 0
        If errorNumber <> 0 Then
            pWarnings.Add filePath & " (" & CStr(errorNumber) & "): " & errorDescription
        ElseIf Not preset Is Nothing Then
            presets.Add preset
        End If
    Next fileIndex
    Set LoadPresets = presets
End Function

' Mengembalikan Nothing untuk preset format lain; XML/nilai invalid memicu Err.
Public Function LoadFile(ByVal filePath As String) As PNGPreset
    Dim xml As Object
    Dim fullPath As String

    fullPath = ExportPresetAbsolutePath(filePath)
    Set xml = LoadExportPresetXMLFile(fullPath, "PNGPresetReader")
    Set LoadFile = ReadDocument(xml, fullPath)
End Function

' Titik masuk untuk validasi parser tanpa membuat file preset di folder Corel.
Public Function ParseXML(ByVal xmlText As String) As PNGPreset
    Dim xml As Object

    Set xml = ParseExportPresetXML(xmlText, "PNGPresetReader")
    Set ParseXML = ReadDocument(xml, vbNullString)
End Function

Private Function ReadDocument(ByVal xml As Object, ByVal filePath As String) As PNGPreset
    Dim root As Object
    Dim preset As PNGPreset
    Dim formatName As String

    Set root = ExportPresetRoot(xml, "PNGPresetReader")
    formatName = LCase$(ExportPresetText(root, "format", vbNullString))
    If Len(formatName) = 0 Then Err.Raise 5, "PNGPresetReader", "Atribut format preset kosong."
    If formatName <> "png" Then Exit Function

    Set preset = New PNGPreset
    preset.SourcePath = filePath
    ' Simpan atribut di luar kontrol UI (mis. palette/dithering) untuk tahap lanjut.
    preset.SourceXML = root.xml
    preset.DisplayName = ExportPresetText(root, "name", vbNullString)
    If Len(preset.DisplayName) = 0 And Len(filePath) > 0 Then preset.DisplayName = ExportPresetFileBaseName(filePath)
    If Len(preset.DisplayName) = 0 Then preset.DisplayName = "PNG"

    preset.ColorMode = NormalizeColorMode(ExportPresetText(root, "colormode", preset.ColorMode))
    preset.Transparency = ExportPresetBoolean(root, "transparency", preset.Transparency, "PNGPresetReader")
    preset.Antialiased = ExportPresetBoolean(root, "antialias", preset.Antialiased, "PNGPresetReader")
    preset.Interlaced = ExportPresetBoolean(root, "interlaced", preset.Interlaced, "PNGPresetReader")
    ReadCommonBitmapPresetNodes root, preset, "PNGPresetReader"
    Set ReadDocument = preset
End Function

Private Function NormalizeColorMode(ByVal mode As String) As String
    Select Case LCase$(Trim$(mode))
        Case "bw", "blackwhite", "blackandwhite", "black-and-white": NormalizeColorMode = "bw"
        Case "gray", "grey", "grayscale", "greyscale": NormalizeColorMode = "grayscale"
        Case "paletted": NormalizeColorMode = "paletted"
        Case "rgb": NormalizeColorMode = "rgb"
        Case Else: Err.Raise 5, "PNGPresetReader", "Color mode PNG tidak dikenal: " & mode
    End Select
End Function
