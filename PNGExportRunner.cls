Option Explicit

' Nama berbeda dari class PNGExport bawaan CorelDRAW.
Private pLastErrorNumber As Long
Private pLastErrorDescription As String

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = pLastErrorNumber
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = pLastErrorDescription
End Property

Public Function ExportDocument(ByVal doc As Document, ByVal outputPath As String, _
                               Optional ByVal snapshot As PNGPreset = Nothing) As Boolean
    Dim store As PNGSettingsStore
    Dim settings As PNGPreset
    Dim options As StructExportOptions
    Dim paletteOptions As StructPaletteOptions
    Dim pngFilter As Object
    Dim fso As Object
    Dim pendingPath As String
    Dim operation As String

    On Error GoTo ExportFailed
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    operation = "Membaca pengaturan PNG"
    Set store = New PNGSettingsStore
    If snapshot Is Nothing Then
        Set settings = store.LoadSettings()
        If settings Is Nothing Then Set settings = New PNGPreset
    Else
        Set settings = snapshot
    End If
    store.ValidateSettings settings

    Set options = BuildCommonBitmapExportOptions(doc, settings, settings.Transparency, Not settings.Transparency, operation)
    If settings.ColorMode = "paletted" Then
        operation = "Membaca opsi palette PNG"
        Set paletteOptions = BuildPaletteOptions(settings.SourceXML)
        options.Dithered = (paletteOptions.DitherType <> cdrDitherNone)
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    pendingPath = fso.BuildPath(fso.GetParentFolderName(outputPath), fso.GetTempName & ".png")
    operation = "Document.ExportEx PNG"
    Set pngFilter = doc.ExportEx(pendingPath, cdrPNG, cdrCurrentPage, options, paletteOptions)
    If pngFilter Is Nothing Then Err.Raise 91, "PNGExportRunner", "ExportEx tidak mengembalikan filter PNG."
    operation = "Mengatur interlace dan transparency filter PNG"
    pngFilter.Interlaced = settings.Interlaced
    If settings.Transparency Then
        pngFilter.Transparency = 2 ' FilterPNG.pngMaskedArea, tanpa reference tambahan.
    Else
        pngFilter.Transparency = 0 ' FilterPNG.pngNone.
    End If
    pngFilter.InvertMask = False
    operation = "PNG filter.Finish"
    pngFilter.Finish
    Set pngFilter = Nothing
    operation = "Memeriksa hasil PNG"
    If Not IsCompletePNG(pendingPath) Then Err.Raise 5, "PNGExportRunner", "File PNG belum terbentuk lengkap."
    operation = "Memindahkan hasil ke file tujuan"
    fso.CopyFile pendingPath, outputPath, True
    fso.DeleteFile pendingPath
    ExportDocument = True
    Exit Function

ExportFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = operation & ": " & Err.Description
    On Error Resume Next
    Set pngFilter = Nothing
    If Len(pendingPath) > 0 Then
        If Not fso Is Nothing Then
            If fso.FileExists(pendingPath) Then fso.DeleteFile pendingPath
        End If
    End If
    On Error GoTo 0
End Function

Private Function BuildPaletteOptions(ByVal sourceXML As String) As StructPaletteOptions
    Dim paletteOptions As StructPaletteOptions
    Dim xml As Object
    Dim root As Object
    Dim paletteName As String
    Dim ditherName As String

    Set paletteOptions = New StructPaletteOptions
    paletteOptions.PaletteType = cdrPaletteOptimized
    paletteOptions.NumColors = 256
    paletteOptions.DitherType = cdrDitherNone
    paletteOptions.DitherIntensity = 100
    If Len(sourceXML) > 0 Then
        Set xml = ParseExportPresetXML(sourceXML, "PNGExportRunner")
        Set root = ExportPresetRoot(xml, "PNGExportRunner")
        paletteName = LCase$(ExportPresetText(root, "palette", vbNullString))
        If Len(paletteName) > 0 And paletteName <> "optimized" Then
            Err.Raise 5, "PNGExportRunner", "Palette PNG belum didukung: " & paletteName
        End If
        paletteOptions.NumColors = ExportPresetInteger(root, "numcolors", 256, 2, 256, "PNGExportRunner")
        ditherName = LCase$(ExportPresetText(root, "dithering", vbNullString))
        Select Case ditherName
            Case vbNullString, "none": paletteOptions.DitherType = cdrDitherNone
            Case "floyd-steinberg": paletteOptions.DitherType = cdrDitherFloyd
            Case Else: Err.Raise 5, "PNGExportRunner", "Dithering PNG belum didukung: " & ditherName
        End Select
        paletteOptions.DitherIntensity = ExportPresetInteger(root, "ditheramount", 100, 0, 100, "PNGExportRunner")
    End If
    Set BuildPaletteOptions = paletteOptions
End Function

Private Function IsCompletePNG(ByVal filePath As String) As Boolean
    Dim fileNumber As Integer
    Dim signature(0 To 7) As Byte
    Dim trailer(0 To 11) As Byte
    Dim expectedSignature As Variant
    Dim expectedTrailer As Variant
    Dim byteIndex As Long
    On Error GoTo ReadFailed
    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    If LOF(fileNumber) >= 45 Then
        Get #fileNumber, 1, signature
        Get #fileNumber, LOF(fileNumber) - 11, trailer
        expectedSignature = Array(137, 80, 78, 71, 13, 10, 26, 10)
        expectedTrailer = Array(0, 0, 0, 0, 73, 69, 78, 68, 174, 66, 96, 130)
        For byteIndex = 0 To 7
            If signature(byteIndex) <> expectedSignature(byteIndex) Then GoTo ReadFailed
        Next byteIndex
        For byteIndex = 0 To 11
            If trailer(byteIndex) <> expectedTrailer(byteIndex) Then GoTo ReadFailed
        Next byteIndex
        IsCompletePNG = True
    End If
    Close #fileNumber
    Exit Function
ReadFailed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
End Function
