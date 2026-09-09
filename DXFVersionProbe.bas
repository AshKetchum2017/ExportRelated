Option Explicit

' Diagnostic only. Tidak dipanggil oleh Export Related dan tidak mengubah mapping.
Public Sub ProbeDXFVersions()
    Dim originalDocument As Document
    Dim testDocument As Document
    Dim fso As Object
    Dim writer As Object
    Dim folder As String
    Dim reportPath As String
    Dim report As String
    Dim candidate As Long
    Dim operation As String
    Dim errorText As String
    Dim restoreText As String

    On Error GoTo ProbeFailed
    If Application.Documents.Count > 0 Then Set originalDocument = ActiveDocument
    Set fso = CreateObject("Scripting.FileSystemObject")
    folder = fso.BuildPath(fso.GetSpecialFolder(2).Path, "DXFVersionProbe-" & fso.GetTempName)
    fso.CreateFolder folder
    reportPath = fso.BuildPath(folder, "DXFVersionReport.txt")
    operation = "Membuat dokumen uji"
    Set testDocument = Application.CreateDocument
    testDocument.ActiveLayer.CreateRectangle2 0, 0, 10, 10
    report = "Nilai setter Version -> header $ACADVER dari file hasil" & vbCrLf
    report = report & "Folder: " & folder & vbCrLf & vbCrLf
    For candidate = 0 To 13
        operation = "Menguji Version = " & CStr(candidate)
        report = report & CStr(candidate) & " -> " & ProbeOneDXFVersion(testDocument, _
            fso.BuildPath(folder, "version-" & CStr(candidate) & ".dxf"), candidate, fso) & vbCrLf
    Next candidate
    GoTo RestoreDocument

ProbeFailed:
    errorText = operation & " (" & CStr(Err.Number) & "): " & Err.Description
    report = report & vbCrLf & errorText
    Resume RestoreDocument

RestoreDocument:
    ' Hanya dokumen baru milik probe yang dibuang; dokumen user tidak disimpan/ditutup.
    On Error Resume Next
    If Not testDocument Is Nothing Then
        Err.Clear
        testDocument.Dirty = False
        If Err.Number = 0 Then
            testDocument.Close
        End If
        If Err.Number <> 0 Then restoreText = "Dokumen uji belum tertutup: " & Err.Description
    End If
    Err.Clear
    If Not originalDocument Is Nothing Then originalDocument.Activate
    If Err.Number <> 0 Then restoreText = restoreText & vbCrLf & "Gagal mengaktifkan dokumen semula: " & Err.Description
    report = report & vbCrLf & restoreText
    Debug.Print report
    Err.Clear
    If Len(reportPath) > 0 Then
        Set writer = fso.CreateTextFile(reportPath, False, True)
        writer.Write report
        writer.Close
    End If
    If Err.Number <> 0 Then
        MsgBox "Report belum tersimpan. Salin hasil dari Immediate Window (Ctrl+G)." & vbCrLf & _
            Err.Description, vbExclamation, "DXF Version Probe"
    Else
        MsgBox "Probe selesai. Report: " & reportPath & vbCrLf & _
            "Hasil juga tersedia di Immediate Window (Ctrl+G)." & vbCrLf & errorText & restoreText, _
            vbInformation, "DXF Version Probe"
    End If
    On Error GoTo 0
End Sub

Private Function ProbeOneDXFVersion(ByVal doc As Document, ByVal outputPath As String, _
                                    ByVal candidate As Long, ByVal fso As Object) As String
    Dim options As StructExportOptions
    Dim filter As ExportFilter
    Dim operation As String

    On Error GoTo ExportFailed
    operation = "CreateStructExportOptions"
    Set options = CreateStructExportOptions
    options.UseColorProfile = False
    operation = "ExportEx"
    Set filter = doc.ExportEx(outputPath, cdrDXF, cdrCurrentPage, options)
    If filter Is Nothing Then Err.Raise 91, "DXFVersionProbe", "Filter kosong."
    operation = "BitmapType/TextAsCurves"
    filter.BitmapType = 0
    filter.TextAsCurves = True
    operation = "Version"
    filter.Version = candidate
    operation = "Units/FillUnmapped"
    filter.Units = 4
    filter.FillUnmapped = False
    operation = "Finish"
    filter.Finish
    operation = "Membaca header file"
    ProbeOneDXFVersion = ProbeDXFHeader(outputPath, fso)
    Exit Function
ExportFailed:
    ProbeOneDXFVersion = "ERROR " & CStr(Err.Number) & " pada " & operation & ": " & Err.Description
End Function

Private Function ProbeDXFHeader(ByVal filePath As String, ByVal fso As Object) As String
    Dim reader As Object
    Dim code As String
    Dim value As String
    Dim pairIndex As Long
    Dim versionNext As Boolean
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ReadFailed
    Set reader = fso.OpenTextFile(filePath, 1, False)
    ProbeDXFHeader = "$ACADVER tidak ditemukan pada 500 pasangan awal; periksa file."
    For pairIndex = 1 To 500
        If reader.AtEndOfStream Then Exit For
        code = Trim$(reader.ReadLine)
        If reader.AtEndOfStream Then Exit For
        value = Trim$(reader.ReadLine)
        If versionNext Then
            If code = "1" Then ProbeDXFHeader = value
            Exit For
        End If
        versionNext = (code = "9" And value = "$ACADVER")
    Next pairIndex
    reader.Close
    Exit Function
ReadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    If Not reader Is Nothing Then reader.Close
    On Error GoTo 0
    Err.Raise errorNumber, "DXFVersionProbe", errorDescription
End Function
