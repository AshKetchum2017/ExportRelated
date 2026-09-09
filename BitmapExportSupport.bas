Option Explicit

' Standard Module (Name): BitmapExportSupport.
' Helper bersama untuk settings/preset bitmap tanpa mengetahui format PNG/JPEG.
Public Const BITMAP_DEFAULT_COLOR_MODE As String = "rgb"
Public Const BITMAP_DEFAULT_RESOLUTION As Long = 300
Public Const BITMAP_MIN_RESOLUTION As Long = 1
Public Const BITMAP_MAX_RESOLUTION As Long = 1200
Public Const BITMAP_DEFAULT_ANTIALIASED As Boolean = True
Public Const BITMAP_DEFAULT_EMBED_PROFILE As Boolean = True
Public Const BITMAP_DEFAULT_CROP_TO_PAGE As Boolean = False
Public Const BITMAP_DEFAULT_MATTE_COMPONENT As Long = 255

Private Const EXPORT_REG_APP_NAME As String = "RinCorelMacros"
Private Const EXPORT_REG_SECTION_NAME As String = "ExportRelatedMacro"

Public Function DiscoverExportPresetDirectories() As Collection
    Dim folders As Collection
    Dim fso As Object
    Dim shell As Object
    Dim documentsPath As String

    Set folders = New Collection
    Set fso = CreateObject("Scripting.FileSystemObject")
    AddDocumentsPresetFolder folders, fso, Environ$("OneDrive")
    AddDocumentsPresetFolder folders, fso, Environ$("OneDriveConsumer")
    AddDocumentsPresetFolder folders, fso, Environ$("OneDriveCommercial")

    On Error Resume Next
    Set shell = CreateObject("WScript.Shell")
    If Not shell Is Nothing Then documentsPath = shell.SpecialFolders("MyDocuments")
    On Error GoTo 0
    AddExportPresetFolder folders, fso, documentsPath
    AddDocumentsPresetFolder folders, fso, Environ$("USERPROFILE")
    Set DiscoverExportPresetDirectories = folders
End Function

Private Sub AddDocumentsPresetFolder(ByVal folders As Collection, ByVal fso As Object, ByVal rootPath As String)
    If Len(Trim$(rootPath)) = 0 Then Exit Sub
    AddExportPresetFolder folders, fso, fso.BuildPath(rootPath, "Documents")
End Sub

Private Sub AddExportPresetFolder(ByVal folders As Collection, ByVal fso As Object, ByVal documentsPath As String)
    Dim folderPath As String
    Dim folderIndex As Long

    If Len(Trim$(documentsPath)) = 0 Then Exit Sub
    folderPath = fso.GetAbsolutePathName(fso.BuildPath(documentsPath, "Corel\Corel Content\Export Presets"))
    If Not fso.FolderExists(folderPath) Then Exit Sub
    For folderIndex = 1 To folders.Count
        If StrComp(folders.Item(folderIndex), folderPath, vbTextCompare) = 0 Then Exit Sub
    Next folderIndex
    folders.Add folderPath
End Sub

Public Function ListExportPresetXMLFiles(ByVal warnings As Collection) As Collection
    Dim filesFound As Collection
    Dim folders As Collection
    Dim folderIndex As Long

    Set filesFound = New Collection
    Set folders = DiscoverExportPresetDirectories()
    If folders.Count = 0 Then
        warnings.Add "Folder Corel\Corel Content\Export Presets tidak ditemukan di Documents atau OneDrive pengguna."
    End If
    For folderIndex = 1 To folders.Count
        AppendPresetXMLFiles CStr(folders.Item(folderIndex)), filesFound, warnings
    Next folderIndex
    Set ListExportPresetXMLFiles = filesFound
End Function

Private Sub AppendPresetXMLFiles(ByVal folderPath As String, ByVal filesFound As Collection, ByVal warnings As Collection)
    Dim fso As Object
    Dim folder As Object
    Dim files As Object
    Dim file As Object
    Dim errorNumber As Long
    Dim errorDescription As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    On Error Resume Next
    Set folder = fso.GetFolder(folderPath)
    If Not folder Is Nothing Then Set files = folder.Files
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error GoTo 0
    If errorNumber <> 0 Or files Is Nothing Then
        warnings.Add folderPath & ": gagal membaca folder (" & CStr(errorNumber) & "). " & errorDescription
        Exit Sub
    End If
    For Each file In files
        If LCase$(fso.GetExtensionName(file.Name)) = "xml" Then filesFound.Add CStr(file.Path)
    Next file
End Sub

Public Function LoadExportPresetXMLFile(ByVal filePath As String, ByVal sourceName As String) As Object
    Dim fso As Object
    Dim xml As Object
    Dim fullPath As String

    Set fso = CreateObject("Scripting.FileSystemObject")
    fullPath = fso.GetAbsolutePathName(filePath)
    If Not fso.FileExists(fullPath) Then Err.Raise 53, sourceName, "File preset tidak ditemukan: " & fullPath
    Set xml = CreateSafeExportPresetXML()
    If Not xml.Load(fullPath) Then RaiseExportPresetXMLParseError xml, sourceName
    Set LoadExportPresetXMLFile = xml
End Function

Public Function ParseExportPresetXML(ByVal xmlText As String, ByVal sourceName As String) As Object
    Dim xml As Object

    Set xml = CreateSafeExportPresetXML()
    If Not xml.LoadXML(xmlText) Then RaiseExportPresetXMLParseError xml, sourceName
    Set ParseExportPresetXML = xml
End Function

Private Function CreateSafeExportPresetXML() As Object
    Dim xml As Object

    Set xml = CreateObject("MSXML2.DOMDocument.6.0")
    xml.async = False
    xml.validateOnParse = False
    xml.resolveExternals = False
    xml.setProperty "ProhibitDTD", True
    Set CreateSafeExportPresetXML = xml
End Function

Private Sub RaiseExportPresetXMLParseError(ByVal xml As Object, ByVal sourceName As String)
    Err.Raise 5, sourceName, "XML preset tidak valid pada baris " & _
        CStr(xml.parseError.Line) & ": " & Trim$(xml.parseError.reason)
End Sub

Public Function ExportPresetRoot(ByVal xml As Object, ByVal sourceName As String) As Object
    Dim root As Object

    Set root = xml.documentElement
    If root Is Nothing Then Err.Raise 5, sourceName, "XML preset kosong."
    If root.nodeName <> "preset" Then Err.Raise 5, sourceName, "Root XML harus berupa elemen preset."
    Set ExportPresetRoot = root
End Function

Public Function ExportPresetText(ByVal node As Object, ByVal attributeName As String, ByVal fallback As String) As String
    Dim attributeNode As Object

    Set attributeNode = node.Attributes.getNamedItem(attributeName)
    If attributeNode Is Nothing Then
        ExportPresetText = fallback
    Else
        ExportPresetText = Trim$(CStr(attributeNode.Text))
    End If
End Function

Public Function ExportPresetBoolean(ByVal node As Object, ByVal attributeName As String, _
                                    ByVal fallback As Boolean, ByVal sourceName As String) As Boolean
    Dim valueText As String

    valueText = LCase$(ExportPresetText(node, attributeName, CStr(fallback)))
    Select Case valueText
        Case "true", "1": ExportPresetBoolean = True
        Case "false", "0": ExportPresetBoolean = False
        Case Else: Err.Raise 5, sourceName, "Nilai " & attributeName & " harus true/false atau 1/0: " & valueText
    End Select
End Function

Public Function ExportPresetNumber(ByVal node As Object, ByVal attributeName As String, _
                                   ByVal fallback As Double, ByVal minimum As Double, ByVal maximum As Double, _
                                   ByVal sourceName As String) As Double
    Dim attributeNode As Object
    Dim valueText As String
    Dim numericValue As Double
    Dim pattern As Object

    Set attributeNode = node.Attributes.getNamedItem(attributeName)
    If attributeNode Is Nothing Then
        ExportPresetNumber = fallback
        Exit Function
    End If
    valueText = Trim$(CStr(attributeNode.Text))
    Set pattern = CreateObject("VBScript.RegExp")
    pattern.Pattern = "^[0-9]+(\.[0-9]+)?$"
    If Not pattern.Test(valueText) Then Err.Raise 5, sourceName, "Angka XML tidak valid untuk " & attributeName & ": " & valueText
    numericValue = Val(valueText)
    If numericValue < minimum Or numericValue > maximum Then
        Err.Raise 5, sourceName, "Nilai " & attributeName & " di luar batas " & CStr(minimum) & " sampai " & CStr(maximum) & ": " & valueText
    End If
    ExportPresetNumber = numericValue
End Function

Public Function ExportPresetInteger(ByVal node As Object, ByVal attributeName As String, _
                                    ByVal fallback As Long, ByVal minimum As Long, ByVal maximum As Long, _
                                    ByVal sourceName As String) As Long
    Dim numericValue As Double

    numericValue = ExportPresetNumber(node, attributeName, fallback, minimum, maximum, sourceName)
    If numericValue <> Fix(numericValue) Then Err.Raise 5, sourceName, "Nilai " & attributeName & " harus berupa bilangan bulat."
    ExportPresetInteger = CLng(numericValue)
End Function

Public Function ExportPresetFileBaseName(ByVal filePath As String) As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    ExportPresetFileBaseName = fso.GetBaseName(filePath)
End Function

Public Function ExportPresetAbsolutePath(ByVal filePath As String) As String
    Dim fso As Object

    Set fso = CreateObject("Scripting.FileSystemObject")
    ExportPresetAbsolutePath = fso.GetAbsolutePathName(filePath)
End Function

Public Sub PopulateBitmapColorModes(ByVal combo As Object, ByVal includeBlackWhite As Boolean, _
                                    ByVal includePaletted As Boolean, ByVal includeCMYK As Boolean, _
                                    ByVal defaultModeKey As String)
    combo.Clear
    combo.Style = fmStyleDropDownList
    combo.ColumnCount = 2
    combo.ColumnWidths = "120 pt;0 pt"
    combo.BoundColumn = 1
    If includeBlackWhite Then AddKeyedComboItem combo, "Black and White", "bw"
    AddKeyedComboItem combo, "Grayscale (8-bit)", "grayscale"
    If includePaletted Then AddKeyedComboItem combo, "Palleted (8-bit)", "paletted"
    AddKeyedComboItem combo, "RGB Color (24-bit)", "rgb"
    If includeCMYK Then AddKeyedComboItem combo, "CMYK Color (32-bit)", "cmyk"
    If Not SelectKeyedComboItem(combo, defaultModeKey) Then Err.Raise 5, "BitmapExportSupport", "Default color mode tidak tersedia."
End Sub

Public Sub AddKeyedComboItem(ByVal combo As Object, ByVal displayName As String, ByVal itemKey As String)
    combo.AddItem displayName
    combo.List(combo.ListCount - 1, 1) = itemKey
End Sub

Public Function SelectKeyedComboItem(ByVal combo As Object, ByVal itemKey As String) As Boolean
    Dim itemIndex As Long

    For itemIndex = 0 To combo.ListCount - 1
        If StrComp(CStr(combo.List(itemIndex, 1)), itemKey, vbTextCompare) = 0 Then
            combo.ListIndex = itemIndex
            SelectKeyedComboItem = True
            Exit Function
        End If
    Next itemIndex
End Function

Public Sub PopulateBitmapResolutions(ByVal combo As Object, ByVal defaultResolution As Long)
    Dim dpi As Variant

    combo.Clear
    combo.Style = fmStyleDropDownCombo
    combo.MatchRequired = False
    For Each dpi In Array(72, 96, 100, 150, 200, 300, 600)
        combo.AddItem CStr(dpi)
    Next dpi
    combo.Value = CStr(defaultResolution)
End Sub

Public Function TryBitmapResolution(ByVal inputText As String, ByRef dpi As Double) As Boolean
    On Error GoTo InvalidResolution
    inputText = Trim$(inputText)
    If Len(inputText) = 0 Then Exit Function
    If Not IsNumeric(inputText) Then Exit Function
    dpi = CDbl(inputText)
    TryBitmapResolution = (dpi >= BITMAP_MIN_RESOLUTION And dpi <= BITMAP_MAX_RESOLUTION And dpi = Fix(dpi))
    Exit Function
InvalidResolution:
    TryBitmapResolution = False
End Function

Public Sub ApplyCommonBitmapPreset(ByVal settingsForm As Object, ByVal preset As Object, ByRef lastValidResolution As Double)
    If Not SelectKeyedComboItem(settingsForm.cmbColorMode, CStr(preset.ColorMode)) Then
        Err.Raise 5, "BitmapExportSupport", "Color mode preset tidak tersedia."
    End If
    settingsForm.chkAntialiased.Value = CBool(preset.Antialiased)
    settingsForm.chkEmbedColorProfile.Value = CBool(preset.EmbedColorProfile)
    settingsForm.chkCropToPageOnExport.Value = CBool(preset.CropToPage)
    settingsForm.cmbResolution.Value = CStr(preset.Resolution)
    lastValidResolution = CDbl(preset.Resolution)
End Sub

Public Sub ConfigureBitmapMatteControl(ByVal combo As Object)
    combo.Clear
    combo.Style = fmStyleDropDownCombo
    combo.Locked = True
    combo.ShowDropButtonWhen = fmShowDropButtonWhenNever
    combo.Value = vbNullString
End Sub

Public Sub UpdateBitmapMatteControl(ByVal combo As Object, ByVal controlEnabled As Boolean, _
                                    ByVal matteRed As Long, ByVal matteGreen As Long, ByVal matteBlue As Long)
    combo.BackColor = RGB(matteRed, matteGreen, matteBlue)
    combo.Enabled = controlEnabled
    combo.ControlTipText = "Matte RGB (" & CStr(matteRed) & ", " & CStr(matteGreen) & ", " & _
        CStr(matteBlue) & "). Klik untuk memilih warna."
End Sub

Public Function ChooseBitmapMatteColor(ByRef matteRed As Long, ByRef matteGreen As Long, _
                                       ByRef matteBlue As Long, ByRef operation As String) As Boolean
    Dim pickedColor As Color

    operation = "CreateRGBColor"
    Set pickedColor = Application.CreateRGBColor(matteRed, matteGreen, matteBlue)
    operation = "Color.UserAssignEx"
    If Not pickedColor.UserAssignEx() Then Exit Function
    operation = "Color.ConvertToRGB"
    pickedColor.ConvertToRGB
    operation = "Membaca komponen RGB"
    matteRed = pickedColor.RGBRed
    matteGreen = pickedColor.RGBGreen
    matteBlue = pickedColor.RGBBlue
    ChooseBitmapMatteColor = True
End Function

Public Function ExportPresetDisplayName(ByVal presets As Collection, ByVal presetIndex As Long) As String
    Dim preset As Object
    Dim otherPreset As Object
    Dim otherIndex As Long

    Set preset = presets.Item(presetIndex)
    ExportPresetDisplayName = CStr(preset.DisplayName)
    For otherIndex = 1 To presets.Count
        If otherIndex <> presetIndex Then
            Set otherPreset = presets.Item(otherIndex)
            If StrComp(CStr(preset.DisplayName), CStr(otherPreset.DisplayName), vbTextCompare) = 0 Then
                ExportPresetDisplayName = CStr(preset.DisplayName) & " - " & CStr(preset.SourcePath)
                Exit Function
            End If
        End If
    Next otherIndex
End Function

Public Function LoadBitmapSettingsXML(ByVal settingsKey As String) As String
    LoadBitmapSettingsXML = GetSetting(EXPORT_REG_APP_NAME, EXPORT_REG_SECTION_NAME, settingsKey, vbNullString)
End Function

Public Sub SaveBitmapSettingsXML(ByVal settingsKey As String, ByVal settingsXML As String)
    SaveSetting EXPORT_REG_APP_NAME, EXPORT_REG_SECTION_NAME, settingsKey, settingsXML
End Sub

Public Sub ReadCommonBitmapPresetNodes(ByVal root As Object, ByVal preset As Object, ByVal sourceName As String)
    Dim node As Object

    Set node = root.selectSingleNode("colorWorkflow")
    If Not node Is Nothing Then
        preset.EmbedColorProfile = ExportPresetBoolean(node, "embedColorProfile", CBool(preset.EmbedColorProfile), sourceName)
    End If
    Set node = root.selectSingleNode("transformation")
    If Not node Is Nothing Then
        preset.Resolution = ExportPresetNumber(node, "dpi", CDbl(preset.Resolution), 0, BITMAP_MAX_RESOLUTION, sourceName)
        If preset.Resolution = 0 Then Err.Raise 5, sourceName, "DPI harus lebih dari 0 dan maksimal 1200."
        preset.CropToPage = ExportPresetBoolean(node, "cropToPage", CBool(preset.CropToPage), sourceName)
    End If
    Set node = root.selectSingleNode("matte")
    If Not node Is Nothing Then
        If LCase$(ExportPresetText(node, "type", vbNullString)) <> "rgb" Then
            Err.Raise 5, sourceName, "Tipe warna matte belum didukung; parser saat ini menerima RGB."
        End If
        preset.MatteRed = ReadBitmapRGBComponent(node, "tint1", sourceName)
        preset.MatteGreen = ReadBitmapRGBComponent(node, "tint2", sourceName)
        preset.MatteBlue = ReadBitmapRGBComponent(node, "tint3", sourceName)
    End If
End Sub

Private Function ReadBitmapRGBComponent(ByVal node As Object, ByVal attributeName As String, ByVal sourceName As String) As Long
    Dim numericValue As Double

    If node.Attributes.getNamedItem(attributeName) Is Nothing Then
        Err.Raise 5, sourceName, "Komponen matte tidak lengkap: " & attributeName
    End If
    numericValue = ExportPresetNumber(node, attributeName, BITMAP_DEFAULT_MATTE_COMPONENT, 0, 255, sourceName)
    If numericValue <> Fix(numericValue) Then Err.Raise 5, sourceName, "Komponen RGB harus berupa bilangan bulat."
    ReadBitmapRGBComponent = CLng(numericValue)
End Function

Public Sub ValidateCommonBitmapSettings(ByVal settings As Object, ByVal sourceName As String)
    If settings Is Nothing Then Err.Raise 5, sourceName, "Pengaturan bitmap belum tersedia."
    If settings.Resolution < BITMAP_MIN_RESOLUTION Or settings.Resolution > BITMAP_MAX_RESOLUTION Or _
       settings.Resolution <> Fix(settings.Resolution) Then
        Err.Raise 5, sourceName, "Resolusi export harus bilangan bulat 1 sampai 1200 dpi."
    End If
    If settings.MatteRed < 0 Or settings.MatteRed > 255 Or _
       settings.MatteGreen < 0 Or settings.MatteGreen > 255 Or _
       settings.MatteBlue < 0 Or settings.MatteBlue > 255 Then
        Err.Raise 5, sourceName, "Komponen warna matte harus 0 sampai 255."
    End If
End Sub

Public Function CreateBitmapSettingsSnapshot(ByVal settings As Object, ByVal formatName As String, _
                                              ByVal sourceName As String) As Object
    Dim xml As Object
    Dim root As Object
    Dim child As Object
    Dim fso As Object

    If Len(settings.SourceXML) > 0 Then
        Set xml = ParseExportPresetXML(CStr(settings.SourceXML), sourceName)
    Else
        Set xml = ParseExportPresetXML("<preset format=""" & formatName & """/>", sourceName)
    End If
    Set root = ExportPresetRoot(xml, sourceName)
    root.setAttribute "format", formatName
    root.setAttribute "name", CStr(settings.DisplayName)
    root.setAttribute "colormode", CStr(settings.ColorMode)
    root.setAttribute "antialias", LCase$(CStr(CBool(settings.Antialiased)))
    root.removeAttribute "macroSourcePath"
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Len(settings.SourcePath) > 0 Then
        If fso.FileExists(CStr(settings.SourcePath)) Then root.setAttribute "macroSourcePath", CStr(settings.SourcePath)
    End If

    Set child = EnsureExportPresetChild(xml, root, "colorWorkflow")
    child.setAttribute "embedColorProfile", LCase$(CStr(CBool(settings.EmbedColorProfile)))
    Set child = EnsureExportPresetChild(xml, root, "transformation")
    child.setAttribute "dpi", CStr(CLng(settings.Resolution))
    child.setAttribute "cropToPage", LCase$(CStr(CBool(settings.CropToPage)))
    Set child = EnsureExportPresetChild(xml, root, "matte")
    child.setAttribute "type", "rgb"
    child.setAttribute "tint1", CStr(CLng(settings.MatteRed))
    child.setAttribute "tint2", CStr(CLng(settings.MatteGreen))
    child.setAttribute "tint3", CStr(CLng(settings.MatteBlue))
    Set CreateBitmapSettingsSnapshot = xml
End Function

Public Function EnsureExportPresetChild(ByVal xml As Object, ByVal root As Object, ByVal childName As String) As Object
    Dim child As Object

    Set child = root.selectSingleNode(childName)
    If child Is Nothing Then
        Set child = xml.createElement(childName)
        root.appendChild child
    End If
    Set EnsureExportPresetChild = child
End Function

Public Sub RestoreBitmapSettingsSourcePath(ByVal settings As Object, ByVal savedXML As String, ByVal sourceName As String)
    Dim xml As Object
    Dim fso As Object
    Dim sourcePath As String

    Set xml = ParseExportPresetXML(savedXML, sourceName)
    sourcePath = CStr(xml.documentElement.getAttribute("macroSourcePath") & vbNullString)
    If Len(sourcePath) = 0 Then Exit Sub
    Set fso = CreateObject("Scripting.FileSystemObject")
    If fso.FileExists(sourcePath) Then settings.SourcePath = sourcePath
End Sub

Public Function BuildCommonBitmapExportOptions(ByVal doc As Document, ByVal settings As Object, _
                                                ByVal transparent As Boolean, ByVal useMatte As Boolean, _
                                                ByRef operation As String) As StructExportOptions
    Dim options As StructExportOptions
    Dim matteColor As Color

    operation = "Membuat StructExportOptions bitmap"
    Set options = New StructExportOptions
    Select Case CStr(settings.ColorMode)
        Case "bw": options.ImageType = cdrBlackAndWhiteImage
        Case "grayscale": options.ImageType = cdrGrayscaleImage
        Case "paletted": options.ImageType = cdrPalettedImage
        Case "rgb": options.ImageType = cdrRGBColorImage
        Case "cmyk": options.ImageType = cdrCMYKColorImage
        Case Else: Err.Raise 5, "BitmapExportSupport", "Color mode bitmap tidak valid."
    End Select
    options.ResolutionX = CLng(settings.Resolution)
    options.ResolutionY = CLng(settings.Resolution)
    If CBool(settings.Antialiased) Then
        options.AntiAliasingType = cdrNormalAntiAliasing
    Else
        options.AntiAliasingType = cdrNoAntiAliasing
    End If
    options.Transparent = transparent
    options.UseColorProfile = CBool(settings.EmbedColorProfile)
    options.Matted = useMatte
    options.MatteMaskedOnly = False
    options.Overwrite = True
    If useMatte Then
        operation = "Mengatur MatteColor bitmap"
        Set matteColor = Application.CreateRGBColor(CLng(settings.MatteRed), CLng(settings.MatteGreen), CLng(settings.MatteBlue))
        ' Typelib Corel memakai PROPERTYPUT, bukan PROPERTYPUTREF, untuk MatteColor.
        CallByName options, "MatteColor", VbLet, matteColor
    End If
    If CBool(settings.CropToPage) Then
        operation = "Mengatur ExportArea ke batas page"
        Set options.ExportArea = doc.ActivePage.BoundingBox
    End If
    Set BuildCommonBitmapExportOptions = options
End Function
