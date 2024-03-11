Attribute VB_Name = "Module1"
'
Function GetAttrValue(CellRef As String, AttrName As String) As String
' Parse field such as "DIELECTRIC":"X7R";"DNP":"F";"PARTNO":"1501";;"SOLDERORDER":"18";;"TOLERANCE":"10%";"VALUE":"1uF";;"VOLTAGE":"25V";

Dim AttrNameValue As String
Dim attributes() As String
Dim countArray As Integer

'split cell on separator
attributes = Split(CellRef, ";")
'number of field in attributes array
countArray = ArrayLen(attributes)

For i = 0 To (countArray - 1)
    If InStr(1, attributes(i), AttrName, vbTextCompare) Then
        AttrNameValue = attributes(i)
    End If
Next i
' remove "SOLDERORDER":"0"
Dim lgthChar As Integer

lgthChar = Len(AttrName)
AttrNameValue = Replace(AttrNameValue, AttrName, "")
AttrNameValue = Replace(AttrNameValue, ":", "")
AttrNameValue = Replace(AttrNameValue, """", "")
GetAttrValue = AttrNameValue
End Function


Function SolderOrder(CellRef As String) As Integer
    TxtValue = GetAttrValue(CellRef, "SOLDERORDER")
    If StrComp(TxtValue, "", vbTextCompare) Then
        SolderOrder = CInt(TxtValue)
        Else
        SolderOrder = CInt("-2")
    End If
End Function

Function ColIndex(searchColName As String) As Integer
    'ColIndex = WorksheetFunction.Match(vLookUp, ActiveWorkbook.Sheets("Orders").Range("3:3"), 0)
    Dim Result As Integer
    
'If Not VBA.IsError(Application.Match(vLookUp, ActiveWorkbook.Sheets("Orders").Range("A3:ZZ3"), 0)) Then
    Result = Application.Match(vLookUp, ActiveWorkbook.Sheets("Orders").Range("A3:ZZ3"), 0)
End If

End Function

Private Function ArrayLen(arr As Variant) As Integer
    ArrayLen = UBound(arr) - LBound(arr) + 1
End Function

