Option Explicit

' Satu item queue. Summary bukan sumber data; opsi format berupa snapshot per item.
Private pSettings As ExportSettings
Private pSourceDocument As Document
Private pSourceDocumentPath As String
Public PNGSettingsXML As String
Public JPEGSettingsXML As String
Public DXFSettingsXML As String
Public PDFPreset As String
Public PDFPresetDisplay As String
Public PDFCompatibility As String
Public Layer1 As Boolean
Public Layer2 As Boolean
Public Layer3 As Boolean
Private pOutputCount As Long
Private pPageDisplay As String

Private Sub Class_Initialize()
    Set pSettings = New ExportSettings
    Layer1 = True
End Sub

Public Property Set SourceDocument(ByVal value As Document)
    Set pSourceDocument = value

    If value Is Nothing Then
        pSourceDocumentPath = vbNullString
    Else
        pSourceDocumentPath = GetDocumentIdentityPath(value)
    End If
End Property

Public Property Get SourceDocument() As Document
    Set SourceDocument = ResolveSourceDocument()
End Property

Public Property Get SourceDocumentPath() As String
    SourceDocumentPath = pSourceDocumentPath
End Property

Public Function ResolveSourceDocument() As Document
    Dim candidate As Document
    Dim candidatePath As String

    ' Prioritas pertama:
    ' gunakan reference asli selama COM object masih hidup.
    If IsDocumentAlive(pSourceDocument) Then

        candidatePath = GetDocumentIdentityPath(pSourceDocument)

        If Len(candidatePath) > 0 Then
            pSourceDocumentPath = candidatePath
        End If

        Set ResolveSourceDocument = pSourceDocument
        Exit Function

    End If

    ' Reference benar-benar sudah tidak valid.
    Set pSourceDocument = Nothing

    ' Jika dokumen pernah disimpan, coba rebind berdasarkan path.
    If Len(pSourceDocumentPath) = 0 Then Exit Function

    For Each candidate In Application.Documents

        candidatePath = GetDocumentIdentityPath(candidate)

        If Len(candidatePath) > 0 Then

            If StrComp(candidatePath, _
                       pSourceDocumentPath, _
                       vbTextCompare) = 0 Then

                Set pSourceDocument = candidate
                pSourceDocumentPath = candidatePath

                Set ResolveSourceDocument = candidate
                Exit Function

            End If

        End If

    Next candidate
End Function

Private Function IsDocumentAlive(ByVal doc As Document) As Boolean
    Dim pageCount As Long

    On Error GoTo DocumentDead

    If doc Is Nothing Then Exit Function

    ' Paksa COM call ke CorelDRAW.
    ' Jika Document sudah ditutup / invalid,
    ' akses ini seharusnya melempar error.
    pageCount = doc.Pages.Count

    IsDocumentAlive = True
    Exit Function

DocumentDead:
    IsDocumentAlive = False
End Function

Private Function GetDocumentIdentityPath(ByVal doc As Document) As String
    Dim folderPath As String
    Dim fileName As String

    On Error GoTo NoIdentity

    If doc Is Nothing Then Exit Function

    fileName = doc.FileName

    ' Unsaved document.
    If Len(fileName) = 0 Then Exit Function

    folderPath = doc.FilePath
    If Len(folderPath) = 0 Then Exit Function

    If Right$(folderPath, 1) = "\" Or Right$(folderPath, 1) = "/" Then
        GetDocumentIdentityPath = folderPath & fileName
    Else
        GetDocumentIdentityPath = folderPath & "\" & fileName
    End If

    Exit Function

NoIdentity:
    GetDocumentIdentityPath = vbNullString
End Function

Public Property Get Settings() As ExportSettings
    Set Settings = pSettings
End Property

Public Property Get OutputCount() As Long
    OutputCount = pOutputCount
End Property

Public Property Get PageDisplay() As String
    PageDisplay = pPageDisplay
End Property

Public Function Clone() As ExportSettingItem
    Dim result As ExportSettingItem
    Set result = New ExportSettingItem
    result.CopyFrom Me
    Set Clone = result
End Function

Public Sub CopyFrom(ByVal source As ExportSettingItem)
    Dim resolvedDocument As Document

    If source Is Nothing Then
        Err.Raise 91, "ExportSettingItem.CopyFrom", _
            "Source ExportSettingItem tidak tersedia."
    End If

    pSettings.Directory = source.Settings.Directory
    pSettings.UseParentDirectory = source.Settings.UseParentDirectory
    pSettings.FormatText = source.Settings.FormatText
    pSettings.PageText = source.Settings.PageText
    pSettings.TemplateText = source.Settings.TemplateText

    ' Pertahankan identity path meskipun Document sedang tidak tersedia.
    pSourceDocumentPath = source.SourceDocumentPath

    ' Coba ambil/rebind Document sumber.
    Set resolvedDocument = source.ResolveSourceDocument()
    Set pSourceDocument = resolvedDocument

    ' Jika Document valid, refresh identity path.
    If Not resolvedDocument Is Nothing Then
        pSourceDocumentPath = GetDocumentIdentityPath(resolvedDocument)
    End If

    PNGSettingsXML = source.PNGSettingsXML
    JPEGSettingsXML = source.JPEGSettingsXML
    DXFSettingsXML = source.DXFSettingsXML
    PDFPreset = source.PDFPreset
    PDFPresetDisplay = source.PDFPresetDisplay
    PDFCompatibility = source.PDFCompatibility

    Layer1 = source.Layer1
    Layer2 = source.Layer2
    Layer3 = source.Layer3

    pOutputCount = source.OutputCount
    pPageDisplay = source.PageDisplay
End Sub

' Registry hanya dibaca sebagai default ketika format belum memiliki snapshot.
Public Sub EnsureFormatSettings()
    Dim pngStore As PNGSettingsStore
    Dim png As PNGPreset
    Dim jpegStore As JPEGSettingsStore
    Dim jpeg As JPEGPreset
    Dim dxfStore As DXFSettingsStore
    Dim dxf As DXFExportSettings

    Select Case pSettings.FormatText
        Case ".png"
            If Len(PNGSettingsXML) > 0 Then Exit Sub
            Set pngStore = New PNGSettingsStore
            Set png = pngStore.LoadSettings()
            If png Is Nothing Then Set png = New PNGPreset
            PNGSettingsXML = pngStore.SerializeSettings(png)
        Case ".jpg"
            If Len(JPEGSettingsXML) > 0 Then Exit Sub
            Set jpegStore = New JPEGSettingsStore
            Set jpeg = jpegStore.LoadSettings()
            If jpeg Is Nothing Then Set jpeg = New JPEGPreset
            JPEGSettingsXML = jpegStore.SerializeSettings(jpeg)
        Case ".pdf"
            If Len(PDFPreset) = 0 Then
                PDFPreset = GetSetting("RinCorelMacros", "ExportRelatedMacro", "PDFPreset", "Used by Rin's Macro")
                If Len(Trim$(PDFPreset)) = 0 Then PDFPreset = "Used by Rin's Macro"
                PDFPresetDisplay = GetSetting("RinCorelMacros", "ExportRelatedMacro", "PDFPresetDisplay", PDFPreset)
                PDFCompatibility = GetSetting("RinCorelMacros", "ExportRelatedMacro", "PDFCompatibility", "Acrobat 9.0")
            End If
        Case ".dxf"
            If Len(DXFSettingsXML) > 0 Then Exit Sub
            Set dxfStore = New DXFSettingsStore
            Set dxf = dxfStore.LoadSettings()
            If dxf Is Nothing Then Set dxf = New DXFExportSettings
            DXFSettingsXML = dxfStore.SerializeSettings(dxf)
        Case Else
            Err.Raise 5, "ExportSettingItem", "Pilih format PDF, DXF, PNG, atau JPG."
    End Select
End Sub

Public Property Get PresetDisplay() As String
    Dim xml As Object
    Select Case pSettings.FormatText
        Case ".pdf"
            PresetDisplay = PDFPresetDisplay
            If Len(PresetDisplay) = 0 Then PresetDisplay = PDFPreset
        Case ".png", ".jpg"
            If pSettings.FormatText = ".png" Then
                Set xml = ParseExportPresetXML(PNGSettingsXML, "ExportSettingItem")
            Else
                Set xml = ParseExportPresetXML(JPEGSettingsXML, "ExportSettingItem")
            End If
            PresetDisplay = ExportPresetText(xml.documentElement, "name", "Default")
        Case ".dxf"
            Set xml = ParseExportPresetXML(DXFSettingsXML, "ExportSettingItem")
            PresetDisplay = ExportPresetText(xml.documentElement, "name", "Default")
    End Select
    If Len(PresetDisplay) = 0 Then PresetDisplay = "Default"
End Property

Public Property Get Summary() As String
    Dim layerText As String

    If Layer1 Then layerText = "1"
    If Layer2 Then
        If Len(layerText) > 0 Then layerText = layerText & ","
        layerText = layerText & "2"
    End If
    If Layer3 Then
        If Len(layerText) > 0 Then layerText = layerText & ","
        layerText = layerText & "3"
    End If
    If Layer1 And Layer2 And Layer3 Then layerText = "1-3"
    If Len(layerText) = 0 Then layerText = "-"

    Summary = UCase$(Mid$(pSettings.FormatText, 2)) & " | " & PresetDisplay & _
        " | File Qty: " & CStr(pOutputCount) & _
        " | Page: " & pPageDisplay & _
        " | Layer: " & layerText
End Property

' Export queue wajib memakai snapshot; jangan diam-diam mengambil default registry.
Public Function ReadPNGSnapshot() As PNGPreset
    Dim reader As PNGPresetReader
    Dim store As PNGSettingsStore
    Dim snapshot As PNGPreset
    If Len(PNGSettingsXML) = 0 Then Err.Raise 5, "ExportSettingItem", "Snapshot PNG kosong. Buka Modify lalu Save."
    Set reader = New PNGPresetReader
    Set store = New PNGSettingsStore
    Set snapshot = reader.ParseXML(PNGSettingsXML)
    If snapshot Is Nothing Then Err.Raise 5, "ExportSettingItem", "Snapshot item bukan PNG."
    store.ValidateSettings snapshot
    Set ReadPNGSnapshot = snapshot
End Function

Public Function ReadJPEGSnapshot() As JPEGPreset
    Dim reader As JPEGPresetReader
    Dim store As JPEGSettingsStore
    Dim snapshot As JPEGPreset
    If Len(JPEGSettingsXML) = 0 Then Err.Raise 5, "ExportSettingItem", "Snapshot JPEG kosong. Buka Modify lalu Save."
    Set reader = New JPEGPresetReader
    Set store = New JPEGSettingsStore
    Set snapshot = reader.ParseXML(JPEGSettingsXML)
    If snapshot Is Nothing Then Err.Raise 5, "ExportSettingItem", "Snapshot item bukan JPEG."
    store.ValidateSettings snapshot
    Set ReadJPEGSnapshot = snapshot
End Function

' Snapshot DXF wajib tersedia dan valid sebelum export queue.
Public Function ReadDXFSnapshot() As DXFExportSettings
    Dim store As DXFSettingsStore
    If Len(DXFSettingsXML) = 0 Then Err.Raise 5, "ExportSettingItem", "Snapshot DXF kosong. Buka Modify lalu Save."
    Set store = New DXFSettingsStore
    Set ReadDXFSnapshot = store.ParseXML(DXFSettingsXML)
End Function

' Hitung file memakai parser existing, bukan menghitung page atau memecah summary.
Public Sub ValidateForQueue()
    Dim pageParser As ExportPageParser
    Dim nameParser As ExportTemplateParser
    Dim groups As Collection
    Dim group As Variant
    Dim groupIndex As Long
    Dim pageIndex As Long
    Dim fso As Object
    Dim pageDisplayText As String

    If SourceDocument Is Nothing Then Err.Raise 5, "ExportSettingItem", "Tidak ada dokumen sumber."
    pSettings.Directory = pSettings.ResolveDirectory(SourceDocument)
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(pSettings.Directory) Then Err.Raise 5, "ExportSettingItem", "Pilih folder export yang valid."
    EnsureFormatSettings
    ' Input kosong menunjuk page saat item disimpan, bukan page aktif saat export nanti.
    If Len(pSettings.PageText) = 0 Then pSettings.PageText = CStr(SourceDocument.ActivePage.Index)
    Set pageParser = New ExportPageParser
    If Not pageParser.ResolvePageGroups(SourceDocument, pSettings.PageText, groups) Then
        Err.Raise 5, "ExportSettingItem", "Format page tidak valid atau kosong."
    End If
    For groupIndex = 1 To groups.Count
        group = groups.Item(groupIndex)
        If groupIndex > 1 Then pageDisplayText = pageDisplayText & "; "
        pageDisplayText = pageDisplayText & pageParser.DisplayNumberList(group)
        If pSettings.FormatText <> ".pdf" And pageParser.PageGroupCount(group) > 1 Then
            Err.Raise 5, "ExportSettingItem", "Grup multi-page hanya didukung untuk PDF."
        End If
        For pageIndex = LBound(group) To UBound(group)
            If group(pageIndex) < 1 Or group(pageIndex) > SourceDocument.Pages.Count Then
                Err.Raise 5, "ExportSettingItem", "Page " & CStr(group(pageIndex)) & " tidak tersedia pada dokumen sumber."
            End If
        Next pageIndex
    Next groupIndex
    Set nameParser = New ExportTemplateParser
    If Len(pSettings.TemplateText) > 0 Then
        If Not nameParser.ValidateNameTemplate(pSettings.TemplateText, groups.Count, pageParser.SyncGroupCount) Then
            Err.Raise 5, "ExportSettingItem", "Jumlah value nama tidak sesuai dengan grup page."
        End If
    End If
    pOutputCount = groups.Count
    pPageDisplay = pageDisplayText
End Sub
