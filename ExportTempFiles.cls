Option Explicit

' Satu instance per Run queue; hanya menerima path sementara buatan runner.
#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal milliseconds As Long)
#End If

Private pFileSystem As Object
Private pPending As Object

Private Sub Class_Initialize()
    Set pFileSystem = CreateObject("Scripting.FileSystemObject")
    Set pPending = CreateObject("Scripting.Dictionary")
    pPending.CompareMode = vbTextCompare
End Sub

Public Sub DeleteOrDefer(ByVal filePath As String)
    Dim details As String
    If Len(filePath) = 0 Then Exit Sub
    details = TryDelete(filePath)
    If Len(details) = 0 Then
        If pPending.Exists(filePath) Then pPending.Remove filePath
    Else
        pPending.Item(filePath) = details
    End If
End Sub

Public Sub RetryPending()
    Dim attempt As Long
    Dim filePath As Variant

    ' Dua retry setelah dokumen dipulihkan. Jeda total maksimal 750 ms,
    ' di luar waktu operasi filesystem; tidak memompa event UI saat export.
    For attempt = 1 To 2
        If pPending.Count = 0 Then Exit Sub
        Sleep 250 * attempt
        For Each filePath In pPending.Keys
            DeleteOrDefer CStr(filePath)
        Next filePath
    Next attempt
End Sub

Public Property Get WarningDescription() As String
    Dim filePath As Variant
    Dim details As String
    If pPending.Count = 0 Then Exit Property
    details = "File sementara belum dapat dihapus (" & CStr(pPending.Count) & "):"
    For Each filePath In pPending.Keys
        details = details & vbCrLf & CStr(filePath) & " - " & CStr(pPending.Item(filePath))
    Next filePath
    WarningDescription = details
End Property

Private Function TryDelete(ByVal filePath As String) As String
    Dim errorNumber As Long
    Dim errorDescription As String
    On Error GoTo DeleteFailed
    ' Jangan paksa atribut read-only atau menganggap FileExists=False sebagai
    ' bukti sukses: share yang tidak dapat diakses juga bisa terlihat tidak ada.
    pFileSystem.DeleteFile filePath
    Exit Function
DeleteFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    Err.Clear
    ' Sudah dihapus proses lain, tidak perlu warning atau percobaan tambahan.
    If errorNumber <> 53 Then
        TryDelete = "Error " & CStr(errorNumber) & ": " & errorDescription
    End If
End Function
