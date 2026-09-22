Option Explicit

' Class Module (Name): JPEGPresetReader.
' Parsing field JPEG tetap di sini; discovery dan XML primitives dipakai bersama.
Private pWarnings As Collection

Private Sub Class_Initialize()
    Set pWarnings = New Collection
End Sub

Public Property Get Warnings() As Collection
    Set Warnings = pWarnings
End Property

Public Function LoadPresets() As Collection
    Dim presets As Collection
    Dim presetFiles As Collection
    Dim preset As JPEGPreset
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
Public Function LoadFile(ByVal filePath As String) As JPEGPreset
    Dim xml As Object
    Dim fullPath As String

    fullPath = ExportPresetAbsolutePath(filePath)
    Set xml = LoadExportPresetXMLFile(fullPath, "JPEGPresetReader")
    Set LoadFile = ReadDocument(xml, fullPath)
End Function

' Titik masuk untuk snapshot registry tanpa membuat file preset baru.
Public Function ParseXML(ByVal xmlText As String) As JPEGPreset
    Dim xml As Object

    Set xml = ParseExportPresetXML(xmlText, "JPEGPresetReader")
    Set ParseXML = ReadDocument(xml, vbNullString)
End Function

Private Function ReadDocument(ByVal xml As Object, ByVal filePath As String) As JPEGPreset
    Dim root As Object
    Dim node As Object
    Dim preset As JPEGPreset
    Dim formatName As String
    Dim colorSettings As String

    Set root = ExportPresetRoot(xml, "JPEGPresetReader")
    formatName = LCase$(ExportPresetText(root, "format", vbNullString))
    If Len(formatName) = 0 Then Err.Raise 5, "JPEGPresetReader", "Atribut format preset kosong."
    If formatName <> "jpg" Then Exit Function

    Set preset = New JPEGPreset
    preset.SourcePath = filePath
    preset.SourceXML = root.xml
    preset.DisplayName = ExportPresetText(root, "name", vbNullString)
    If Len(preset.DisplayName) = 0 And Len(filePath) > 0 Then preset.DisplayName = ExportPresetFileBaseName(filePath)
    If Len(preset.DisplayName) = 0 Then preset.DisplayName = "JPEG"

    preset.ColorMode = NormalizeColorMode(ExportPresetText(root, "colormode", preset.ColorMode))
    preset.Antialiased = ExportPresetBoolean(root, "antialias", preset.Antialiased, "JPEGPresetReader")
    preset.SubFormat = NormalizeSubFormat(ExportPresetText(root, "subformat", preset.SubFormat))
    preset.Quality = ExportPresetInteger(root, "qualityamount", preset.Quality, 0, 100, "JPEGPresetReader")
    preset.Blur = ExportPresetInteger(root, "bluramount", preset.Blur, 0, 100, "JPEGPresetReader")
    preset.Optimize = ExportPresetBoolean(root, "optimize", preset.Optimize, "JPEGPresetReader")
    preset.Progressive = ExportPresetBoolean(root, "progressive", preset.Progressive, "JPEGPresetReader")

    ReadCommonBitmapPresetNodes root, preset, "JPEGPresetReader"
    Set node = root.selectSingleNode("colorWorkflow")
    If Not node Is Nothing Then
        colorSettings = LCase$(ExportPresetText(node, "colorSettings", "document"))
        Select Case colorSettings
            Case "document": preset.UseColorProof = False
            Case "proof": preset.UseColorProof = True
            Case Else: Err.Raise 5, "JPEGPresetReader", "Color settings JPEG tidak dikenal: " & colorSettings
        End Select
    End If
    Set ReadDocument = preset
End Function

Private Function NormalizeColorMode(ByVal mode As String) As String
    Select Case LCase$(Trim$(mode))
        Case "gray", "grey", "grayscale", "greyscale": NormalizeColorMode = "grayscale"
        Case "rgb": NormalizeColorMode = "rgb"
        Case "cmyk": NormalizeColorMode = "cmyk"
        Case Else: Err.Raise 5, "JPEGPresetReader", "Color mode JPEG tidak dikenal: " & mode
    End Select
End Function

Private Function NormalizeSubFormat(ByVal subFormat As String) As String
    Select Case LCase$(Trim$(subFormat))
        Case "standard": NormalizeSubFormat = "standard"
        Case "optional": NormalizeSubFormat = "optional"
        Case Else: Err.Raise 5, "JPEGPresetReader", "Subformat JPEG tidak dikenal: " & subFormat
    End Select
End Function
