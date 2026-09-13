Option Explicit

Sub ExportRelatedWizard()
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo StartupFailed
    operation = "Load ExportRelatedMenu / UserForm_Initialize"
    Load ExportRelatedMenu
    operation = "ExportRelatedMenu.Show vbModeless"
    ExportRelatedMenu.Show vbModeless
    Exit Sub

StartupFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    MsgBox "Operasi: " & operation & vbCrLf & _
        "Error: " & CStr(errorNumber) & vbCrLf & _
        "Source: " & errorSource & vbCrLf & _
        "Description: " & errorDescription, vbCritical, "Export Queue - Startup"
End Sub
