Option Explicit

Private pLastErrorNumber As Long
Private pLastErrorDescription As String

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = pLastErrorNumber
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = pLastErrorDescription
End Property

Public Function ExportDocument(ByVal doc As Document, ByVal outputPath As String, ByVal exportFormat As cdrFilter, _
                               Optional ByVal pageRangeText As String = vbNullString, _
                               Optional ByVal queueItem As ExportSettingItem = Nothing) As Boolean
    Dim pdfExport As PDFExport
    Dim expFilter As ExportFilter
    Dim errorNumber As Long
    Dim errorDescription As String

    If Not queueItem Is Nothing Then
        ExportDocument = ExportQueueDocument(doc, outputPath, exportFormat, pageRangeText, queueItem)
        Exit Function
    End If

    If exportFormat = cdrPNG Then
        ExportDocument = ExportPNGDocument(doc, outputPath)
        Exit Function
    End If
    If exportFormat = cdrJPEG Then
        ExportDocument = ExportJPEGDocument(doc, outputPath)
        Exit Function
    End If

    On Error Resume Next
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString

    If exportFormat = cdrPDF Then
        Set pdfExport = New PDFExport
        ExportDocument = pdfExport.PublishDocument(doc, outputPath, pageRangeText)
        pLastErrorNumber = pdfExport.LastErrorNumber
        pLastErrorDescription = pdfExport.LastErrorDescription
        Exit Function
    End If

    Set expFilter = doc.ExportEx(outputPath, exportFormat, cdrCurrentPage)
    errorNumber = Err.Number
    errorDescription = Err.Description

    If errorNumber = 0 And expFilter Is Nothing Then
        errorNumber = 91
        errorDescription = "ExportEx tidak mengembalikan ExportFilter."
    End If

    If errorNumber = 0 Then
        expFilter.Finish
        errorNumber = Err.Number
        errorDescription = Err.Description
    End If

    Set expFilter = Nothing
    Err.Clear
    pLastErrorNumber = errorNumber
    pLastErrorDescription = errorDescription
    ExportDocument = (errorNumber = 0)
End Function

Private Function ExportQueueDocument(ByVal doc As Document, ByVal outputPath As String, _
                                     ByVal exportFormat As cdrFilter, ByVal pageRangeText As String, _
                                     ByVal item As ExportSettingItem) As Boolean
    Dim pngExporter As PNGExportRunner
    Dim jpegExporter As JPEGExportRunner
    Dim pdfExporter As PDFExport
    Dim png As PNGPreset
    Dim jpeg As JPEGPreset
    Dim filter As ExportFilter

    On Error GoTo ExportFailed
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    Select Case exportFormat
        Case cdrPNG
            Set png = item.ReadPNGSnapshot()
            Set pngExporter = New PNGExportRunner
            ExportQueueDocument = pngExporter.ExportDocument(doc, outputPath, png)
            pLastErrorNumber = pngExporter.LastErrorNumber
            pLastErrorDescription = pngExporter.LastErrorDescription
        Case cdrJPEG
            Set jpeg = item.ReadJPEGSnapshot()
            Set jpegExporter = New JPEGExportRunner
            ExportQueueDocument = jpegExporter.ExportDocument(doc, outputPath, jpeg)
            pLastErrorNumber = jpegExporter.LastErrorNumber
            pLastErrorDescription = jpegExporter.LastErrorDescription
        Case cdrPDF
            Set pdfExporter = New PDFExport
            ExportQueueDocument = pdfExporter.PublishDocument(doc, outputPath, pageRangeText, item)
            pLastErrorNumber = pdfExporter.LastErrorNumber
            pLastErrorDescription = pdfExporter.LastErrorDescription
        Case cdrDXF
            Set filter = doc.ExportEx(outputPath, cdrDXF, cdrCurrentPage)
            If filter Is Nothing Then Err.Raise 91, "ExportRunner", "ExportEx DXF tidak mengembalikan filter."
            filter.Finish
            ExportQueueDocument = True
        Case Else
            Err.Raise 5, "ExportRunner", "Format queue tidak didukung."
    End Select
    Exit Function
ExportFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = "Export item queue: " & Err.Description
End Function

Private Function ExportJPEGDocument(ByVal doc As Document, ByVal outputPath As String) As Boolean
    Dim jpegExporter As JPEGExportRunner

    On Error GoTo ExportFailed
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    Set jpegExporter = New JPEGExportRunner
    ExportJPEGDocument = jpegExporter.ExportDocument(doc, outputPath)
    pLastErrorNumber = jpegExporter.LastErrorNumber
    pLastErrorDescription = jpegExporter.LastErrorDescription
    Exit Function

ExportFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = "Menjalankan JPEGExportRunner: " & Err.Description
End Function

Private Function ExportPNGDocument(ByVal doc As Document, ByVal outputPath As String) As Boolean
    Dim pngExporter As PNGExportRunner

    On Error GoTo ExportFailed
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    Set pngExporter = New PNGExportRunner
    ExportPNGDocument = pngExporter.ExportDocument(doc, outputPath)
    pLastErrorNumber = pngExporter.LastErrorNumber
    pLastErrorDescription = pngExporter.LastErrorDescription
    Exit Function

ExportFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = "Menjalankan PNGExportRunner: " & Err.Description
End Function
