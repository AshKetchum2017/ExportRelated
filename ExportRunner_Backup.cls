VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExportRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private pLastErrorNumber As Long
Private pLastErrorDescription As String

Public Property Get LastErrorNumber() As Long
    LastErrorNumber = pLastErrorNumber
End Property

Public Property Get LastErrorDescription() As String
    LastErrorDescription = pLastErrorDescription
End Property

Public Function ResolvePageValues(ByVal doc As Document, ByVal pageText As String, ByRef pageValues() As Long, ByRef pageCount As Long) As Boolean
    Dim pageValueText As String

    pageValueText = Trim$(pageText)
    If Len(pageValueText) = 0 Then
        pageValueText = CStr(doc.ActivePage.Index)
        pageCount = 1
        ReDim pageValues(0 To 0)
        pageValues(0) = doc.ActivePage.Index
        ResolvePageValues = True
        Exit Function
    End If

    pageValues = ParseLongList(pageValueText, pageCount)
    If pageCount = 0 Then
        ResolvePageValues = False
        Exit Function
    End If

    ResolvePageValues = True
End Function

Private Function ParseLongList(ByVal inputText As String, ByRef itemCount As Long) As Variant
    Dim rawValues() As String
    Dim values() As Long
    Dim valueIndex As Long
    Dim parseText As String
    Dim itemValue As String

    parseText = Trim$(inputText)
    If Len(parseText) = 0 Then
        itemCount = 0
        ReDim values(0)
        ParseLongList = values
        Exit Function
    End If

    If Left$(parseText, 1) = "(" And Right$(parseText, 1) = ")" Then
        parseText = Trim$(Mid$(parseText, 2, Len(parseText) - 2))
    End If

    rawValues = Split(parseText, ",")
    ReDim values(0 To UBound(rawValues))

    For valueIndex = 0 To UBound(rawValues)
        itemValue = Trim$(rawValues(valueIndex))
        If Len(itemValue) = 0 Then
            Err.Raise 5, , "Format page tidak valid: " & inputText
        End If

        If Not IsNumeric(itemValue) Then
            Err.Raise 5, , "Format page harus berupa angka: " & inputText
        End If

        values(valueIndex) = CLng(itemValue)
    Next valueIndex

    itemCount = UBound(values) + 1
    ParseLongList = values
End Function

Public Function ExportDocument(ByVal doc As Document, ByVal outputPath As String, ByVal exportFormat As cdrFilter) As Boolean
    On Error Resume Next
    doc.ExportEx outputPath, exportFormat, cdrCurrentPage
    pLastErrorNumber = Err.Number
    pLastErrorDescription = Err.Description
    Err.Clear

    ExportDocument = (pLastErrorNumber = 0)
End Function
