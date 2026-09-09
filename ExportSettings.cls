Option Explicit

Private pDirectory As String
Private pFormat As String
Private pTemplateText As String
Private pPageText As String

Public Property Get Directory() As String
    Directory = pDirectory
End Property

Public Property Let Directory(ByVal value As String)
    pDirectory = Trim$(value)
End Property

Public Property Get FormatText() As String
    FormatText = pFormat
End Property

Public Property Let FormatText(ByVal value As String)
    pFormat = Trim$(value)
End Property

Public Property Get TemplateText() As String
    TemplateText = pTemplateText
End Property

Public Property Let TemplateText(ByVal value As String)
    pTemplateText = Trim$(value)
End Property

Public Property Get PageText() As String
    PageText = pPageText
End Property

Public Property Let PageText(ByVal value As String)
    pPageText = Trim$(value)
End Property

Public Sub LoadFromForm(ByVal frm As Object)
    Dim operation As String
    Dim errorNumber As Long
    Dim errorSource As String
    Dim errorDescription As String

    On Error GoTo ReadFailed
    operation = "frm.txbDirectory.Text"
    Me.Directory = frm.txbDirectory.Text
    operation = "frm.cmbExFormat.Value"
    Me.FormatText = frm.cmbExFormat.Value
    operation = "frm.txbName.Text"
    Me.TemplateText = frm.txbName.Text
    operation = "frm.txbPage.Text"
    Me.PageText = frm.txbPage.Text
    Exit Sub

ReadFailed:
    errorNumber = Err.Number
    errorSource = Err.Source
    errorDescription = Err.Description
    Err.Raise errorNumber, "ExportSettings.LoadFromForm", _
        operation & " [" & TypeName(frm) & "]" & vbCrLf & _
        "Source asli: " & errorSource & vbCrLf & errorDescription
End Sub

Public Function IsValidDirectory() As Boolean
    IsValidDirectory = (Len(Me.Directory) > 0) And (Len(Dir$(Me.Directory, vbDirectory)) > 0)
End Function
