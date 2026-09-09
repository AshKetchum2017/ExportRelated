Option Explicit

' Nama berbeda dari class JPGExport bawaan CorelDRAW.
Private Const JPG_SUBFORMAT_STANDARD As Long = 0
Private Const JPG_SUBFORMAT_OPTIONAL As Long = 1

Private pLastErrorNumber As Long
Private pLastErrorDescription As String

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = pLastErrorNumber
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = pLastErrorDescription
End Property

Public Function ExportDocument(ByVal doc As Document, ByVal outputPath As String, _
                               Optional ByVal snapshot As JPEGPreset = Nothing) As Boolean
    Dim store As JPEGSettingsStore
    Dim settings As JPEGPreset
    Dim options As StructExportOptions
    Dim proofSettings As Object
    Dim jpegFilter As Object
    Dim fso As Object
    Dim pendingPath As String
    Dim operation As String

    On Error GoTo ExportFailed
    pLastErrorNumber = 0
    pLastErrorDescription = vbNullString
    operation = "Membaca pengaturan JPEG"
    Set store = New JPEGSettingsStore
    If snapshot Is Nothing Then
        Set settings = store.LoadSettings()
        If settings Is Nothing Then Set settings = New JPEGPreset
    Else
        Set settings = snapshot
    End If
    store.ValidateSettings settings

    Set options = BuildCommonBitmapExportOptions(doc, settings, False, True, operation)
    If settings.UseColorProof Then
        operation = "Membaca color proof settings dari active view"
        Set proofSettings = Application.ActiveWindow.ActiveView.ProofColorSettings
        If proofSettings Is Nothing Then Err.Raise 91, "JPEGExportRunner", "Color proof settings tidak tersedia pada active view."
        operation = "Mengatur ProofColorSettings JPEG"
        ' Seperti MatteColor, typelib mengekspos object ini sebagai PROPERTYPUT.
        CallByName options, "ProofColorSettings", VbLet, proofSettings
    End If

    Set fso = CreateObject("Scripting.FileSystemObject")
    pendingPath = fso.BuildPath(fso.GetParentFolderName(outputPath), fso.GetTempName & ".jpg")
    operation = "Document.ExportEx JPEG"
    Set jpegFilter = doc.ExportEx(pendingPath, cdrJPEG, cdrCurrentPage, options)
    If jpegFilter Is Nothing Then Err.Raise 91, "JPEGExportRunner", "ExportEx tidak mengembalikan filter JPEG."

    operation = "Mengatur quality dan blur filter JPEG"
    jpegFilter.Compression = settings.Quality
    jpegFilter.Smoothing = settings.Blur
    operation = "Mengatur optimize dan progressive filter JPEG"
    jpegFilter.Optimized = settings.Optimize
    jpegFilter.Progressive = settings.Progressive
    operation = "Mengatur subformat filter JPEG"
    Select Case settings.SubFormat
        Case "standard": jpegFilter.SubFormat = JPG_SUBFORMAT_STANDARD
        Case "optional": jpegFilter.SubFormat = JPG_SUBFORMAT_OPTIONAL
    End Select
    operation = "JPEG filter.Finish"
    jpegFilter.Finish
    Set jpegFilter = Nothing

    operation = "Memeriksa hasil JPEG"
    If Not IsCompleteJPEG(pendingPath) Then Err.Raise 5, "JPEGExportRunner", "File JPEG belum terbentuk lengkap."
    operation = "Memindahkan hasil ke file tujuan"
    fso.CopyFile pendingPath, outputPath, True
    fso.DeleteFile pendingPath
    ExportDocument = True
    Exit Function

ExportFailed:
    pLastErrorNumber = Err.Number
    pLastErrorDescription = operation & ": " & Err.Description
    On Error Resume Next
    Set jpegFilter = Nothing
    If Len(pendingPath) > 0 Then
        If Not fso Is Nothing Then
            If fso.FileExists(pendingPath) Then fso.DeleteFile pendingPath
        End If
    End If
    On Error GoTo 0
End Function

Private Function IsCompleteJPEG(ByVal filePath As String) As Boolean
    Dim fileNumber As Integer
    Dim signature(0 To 1) As Byte
    Dim trailer(0 To 1) As Byte

    On Error GoTo ReadFailed
    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    If LOF(fileNumber) >= 4 Then
        Get #fileNumber, 1, signature
        Get #fileNumber, LOF(fileNumber) - 1, trailer
        IsCompleteJPEG = (signature(0) = 255 And signature(1) = 216 And _
                          trailer(0) = 255 And trailer(1) = 217)
    End If
    Close #fileNumber
    Exit Function

ReadFailed:
    On Error Resume Next
    If fileNumber > 0 Then Close #fileNumber
    On Error GoTo 0
End Function
