Option Explicit

Private pLastErrorNumber As Long
Private pLastErrorDescription As String
Private pFailedItemIndex As Long
Private pCompletedFiles As Long

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = pLastErrorNumber
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = pLastErrorDescription
End Property

Public Property Get FailedItemIndex() As Long
    FailedItemIndex = pFailedItemIndex
End Property

Public Property Get CompletedFiles() As Long
    CompletedFiles = pCompletedFiles
End Property

Public Function Run(ByVal items As Collection) As Boolean
    Dim prepared As Collection
    Dim allGroups As Collection
    Dim allPaths As Collection
    Dim groups As Collection
    Dim paths As Variant
    Dim usedPaths As Object
    Dim item As ExportSettingItem
    Dim layers As ExportLayerState
    Dim originalDocument As Document
    Dim itemIndex As Long
    Dim pathIndex As Long
    Dim operation As String
    Dim restoreError As String

    On Error GoTo RunFailed
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    pFailedItemIndex = 0
    pCompletedFiles = 0
    Set originalDocument = ActiveDocument
    Set prepared = New Collection
    Set allGroups = New Collection
    Set allPaths = New Collection
    Set usedPaths = CreateObject("Scripting.Dictionary")
    usedPaths.CompareMode = vbTextCompare
    If items.Count = 0 Then Err.Raise 5, "ExportQueueRunner", "Queue kosong."

    ' Validasi seluruh antrean sebelum layer diubah atau file dibuat.
    For itemIndex = 1 To items.Count
        pFailedItemIndex = itemIndex
        operation = "Validasi item " & CStr(itemIndex)
        Set item = items.Item(itemIndex).Clone()
        RequireOpenDocument item.SourceDocument
        ValidateSnapshot item
        item.ValidateForQueue
        paths = BuildPaths(item, groups)
        Set layers = New ExportLayerState
        layers.ValidateSelection item, groups
        For pathIndex = LBound(paths) To UBound(paths)
            If usedPaths.Exists(CStr(paths(pathIndex))) Then
                Err.Raise 5, "ExportQueueRunner", "File tujuan sama pada item " & _
                    CStr(usedPaths.Item(CStr(paths(pathIndex)))) & " dan " & CStr(itemIndex) & ": " & CStr(paths(pathIndex))
            End If
            usedPaths.Add CStr(paths(pathIndex)), itemIndex
        Next pathIndex
        prepared.Add item
        allGroups.Add groups
        allPaths.Add paths
    Next itemIndex

    For itemIndex = 1 To prepared.Count
        pFailedItemIndex = itemIndex
        Set item = prepared.Item(itemIndex)
        Set groups = allGroups.Item(itemIndex)
        paths = allPaths.Item(itemIndex)
        If Not ExecuteItem(item, groups, paths) Then GoTo RestoreDocument
    Next itemIndex
    Run = True
    pFailedItemIndex = 0

RestoreDocument:
    restoreError = RestoreActiveDocument(originalDocument)
    If Len(restoreError) > 0 Then
        Run = False
        If pLastErrorNumber = 0 Then pLastErrorNumber = 5
        pLastErrorDescription = pLastErrorDescription & vbCrLf & restoreError
    End If
    Exit Function

RunFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = operation & ": " & Err.Description
    Resume RestoreDocument
End Function

Private Sub RequireOpenDocument(ByVal doc As Document)
    Dim candidate As Document
    If doc Is Nothing Then Err.Raise 5, "ExportQueueRunner", "Dokumen sumber kosong."
    For Each candidate In Application.Documents
        If candidate Is doc Then Exit Sub
    Next candidate
    Err.Raise 5, "ExportQueueRunner", "Dokumen sumber item sudah ditutup. Buat item ulang dari dokumen yang terbuka."
End Sub

Private Sub ValidateSnapshot(ByVal item As ExportSettingItem)
    Dim png As PNGPreset
    Dim jpeg As JPEGPreset
    Dim pdf As PDFExport
    Select Case item.Settings.FormatText
        Case ".png": Set png = item.ReadPNGSnapshot()
        Case ".jpg": Set jpeg = item.ReadJPEGSnapshot()
        Case ".pdf"
            Set pdf = New PDFExport
            If Not pdf.ValidateQueueSettings(item.PDFPreset, item.PDFCompatibility) Then
                Err.Raise pdf.LastErrorNumber, "ExportQueueRunner", pdf.LastErrorDescription
            End If
        Case ".dxf"
        Case Else: Err.Raise 5, "ExportQueueRunner", "Format item tidak didukung."
    End Select
End Sub

Private Function BuildPaths(ByVal item As ExportSettingItem, ByRef groups As Collection) As Variant
    Dim pageParser As ExportPageParser
    Dim nameParser As ExportTemplateParser
    Dim fso As Object
    Dim names() As String
    Dim paths() As String
    Dim baseName As String
    Dim outputName As String
    Dim group As Variant
    Dim groupIndex As Long
    Dim firstPage As Long

    Set pageParser = New ExportPageParser
    Set nameParser = New ExportTemplateParser
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not pageParser.ResolvePageGroups(item.SourceDocument, item.Settings.PageText, groups) Then
        Err.Raise 5, "ExportQueueRunner", "Grup page tidak valid."
    End If
    baseName = nameParser.RemoveDocumentExtension(item.SourceDocument.Name)
    If Len(baseName) = 0 Then baseName = "Export"
    ReDim names(0 To groups.Count - 1)
    ReDim paths(0 To groups.Count - 1)
    For groupIndex = 1 To groups.Count
        group = groups.Item(groupIndex)
        firstPage = pageParser.FirstPageInGroup(group)
        names(groupIndex - 1) = nameParser.BuildExportFileName(baseName, item.Settings.TemplateText, _
            groupIndex - 1, firstPage, pageParser.SyncGroupIndex(groupIndex) - 1)
    Next groupIndex
    For groupIndex = 1 To groups.Count
        outputName = nameParser.BuildUniqueExportName(names, groups.Count, groupIndex)
        ValidateFileName outputName
        paths(groupIndex - 1) = fso.GetAbsolutePathName(fso.BuildPath(item.Settings.Directory, outputName & item.Settings.FormatText))
    Next groupIndex
    BuildPaths = paths
End Function

Private Sub ValidateFileName(ByVal value As String)
    Dim character As Variant
    Dim position As Long
    If Len(value) = 0 Then Err.Raise 5, "ExportQueueRunner", "Nama file kosong."
    For Each character In Array("\", "/", ":", "*", "?", Chr$(34), "<", ">", "|")
        If InStr(1, value, CStr(character), vbBinaryCompare) > 0 Then
            Err.Raise 5, "ExportQueueRunner", "Nama file mengandung karakter tidak valid: " & value
        End If
    Next character
    For position = 1 To Len(value)
        If AscW(Mid$(value, position, 1)) >= 0 And AscW(Mid$(value, position, 1)) < 32 Then
            Err.Raise 5, "ExportQueueRunner", "Nama file mengandung karakter kontrol."
        End If
    Next position
End Sub

Private Function ResolveFormat(ByVal formatText As String) As cdrFilter
    Select Case formatText
        Case ".pdf": ResolveFormat = cdrPDF
        Case ".dxf": ResolveFormat = cdrDXF
        Case ".png": ResolveFormat = cdrPNG
        Case ".jpg": ResolveFormat = cdrJPEG
        Case Else: Err.Raise 5, "ExportQueueRunner", "Format item tidak didukung."
    End Select
End Function

Private Function ExecuteItem(ByVal item As ExportSettingItem, ByVal groups As Collection, ByVal paths As Variant) As Boolean
    Dim doc As Document
    Dim originalPage As Page
    Dim originalLayer As Layer
    Dim originalSelection As ShapeRange
    Dim layers As ExportLayerState
    Dim runner As ExportRunner
    Dim parser As ExportPageParser
    Dim fso As Object
    Dim group As Variant
    Dim groupIndex As Long
    Dim pendingPath As String
    Dim outputPath As String
    Dim operation As String
    Dim cleanupError As String
    Dim exportError As Long

    On Error GoTo ExportFailed
    Set doc = item.SourceDocument
    operation = "Memeriksa dokumen sumber"
    RequireOpenDocument doc
    operation = "Menyimpan page, layer aktif, dan selection dokumen sumber"
    Set originalPage = doc.ActivePage
    Set originalLayer = doc.ActiveLayer
    Set originalSelection = doc.SelectionRange
    doc.Activate
    Set layers = New ExportLayerState
    operation = "Menerapkan print/export layer item"
    layers.Apply item, groups
    Set runner = New ExportRunner
    Set parser = New ExportPageParser
    Set fso = CreateObject("Scripting.FileSystemObject")
    For groupIndex = 1 To groups.Count
        group = groups.Item(groupIndex)
        outputPath = CStr(paths(groupIndex - 1))
        operation = "Mengaktifkan page " & CStr(parser.FirstPageInGroup(group))
        doc.Pages(parser.FirstPageInGroup(group)).Activate
        pendingPath = fso.BuildPath(fso.GetParentFolderName(outputPath), fso.GetTempName & item.Settings.FormatText)
        operation = "Export " & outputPath
        If Not runner.ExportDocument(doc, pendingPath, ResolveFormat(item.Settings.FormatText), parser.PageGroupToRangeText(group), item) Then
            exportError = runner.LastErrorNumber
            If exportError = 0 Then exportError = 5
            Err.Raise exportError, "ExportQueueRunner", runner.LastErrorDescription
        End If
        If Not fso.FileExists(pendingPath) Then Err.Raise 53, "ExportQueueRunner", "File hasil belum terbentuk."
        If fso.GetFile(pendingPath).Size = 0 Then Err.Raise 5, "ExportQueueRunner", "File hasil kosong."
        operation = "Menyalin hasil ke " & outputPath
        fso.CopyFile pendingPath, outputPath, True
        pCompletedFiles = pCompletedFiles + 1
        operation = "Membersihkan file sementara " & pendingPath
        fso.DeleteFile pendingPath
        pendingPath = vbNullString
    Next groupIndex
    ExecuteItem = True

RestoreState:
    cleanupError = RestoreItemState(doc, originalPage, originalLayer, originalSelection, layers)
    If Len(pendingPath) > 0 Then
        On Error Resume Next
        Err.Clear
        If Not fso Is Nothing Then
            If fso.FileExists(pendingPath) Then fso.DeleteFile pendingPath
        End If
        If Err.Number <> 0 Then cleanupError = cleanupError & vbCrLf & "Hapus file sementara: " & Err.Description
        On Error GoTo 0
    End If
    If Len(cleanupError) > 0 Then
        ExecuteItem = False
        If pLastErrorNumber = 0 Then pLastErrorNumber = 5
        pLastErrorDescription = pLastErrorDescription & vbCrLf & "Pemulihan state gagal:" & cleanupError
    End If
    Exit Function

ExportFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = operation & ": " & Err.Description
    Resume RestoreState
End Function

Private Function RestoreItemState(ByVal doc As Document, ByVal originalPage As Page, ByVal originalLayer As Layer, _
                                  ByVal originalSelection As ShapeRange, ByVal layers As ExportLayerState) As String
    Dim details As String
    On Error Resume Next
    If Not layers Is Nothing Then details = layers.Restore()
    If Err.Number <> 0 Then details = details & vbCrLf & "Restore print/export layer: " & Err.Description
    Err.Clear
    If Not originalPage Is Nothing Then originalPage.Activate
    If Err.Number <> 0 Then details = details & vbCrLf & "Restore page: " & Err.Description
    Err.Clear
    If Not originalLayer Is Nothing Then originalLayer.Activate
    If Err.Number <> 0 Then details = details & vbCrLf & "Restore layer aktif: " & Err.Description
    Err.Clear
    If Not originalSelection Is Nothing Then
        doc.ClearSelection
        originalSelection.CreateSelection
    End If
    If Err.Number <> 0 Then details = details & vbCrLf & "Restore selection: " & Err.Description
    On Error GoTo 0
    RestoreItemState = details
End Function

Private Function RestoreActiveDocument(ByVal originalDocument As Document) As String
    On Error GoTo RestoreFailed
    If Not originalDocument Is Nothing Then originalDocument.Activate
    Exit Function
RestoreFailed:
    RestoreActiveDocument = "Restore dokumen aktif (" & CStr(Err.Number) & "): " & Err.Description
End Function
