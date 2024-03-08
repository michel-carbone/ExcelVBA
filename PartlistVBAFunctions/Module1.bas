Attribute VB_Name = "Module1"
'SolderOrder(CellRef As String)
Function SolderOrder(CellRef As String) As Integer

Dim textOutput As String
Dim attributes() As String
Dim countArray As Integer

'split cell on separator
attributes = Split(CellRef, ";")
'number of field in attributes array
countArray = ArrayLen(attributes)

For i = 0 To (countArray - 1)
    If InStr(1, attributes(i), "SOLDERORDER", vbTextCompare) Then
        textOutput = attributes(i)
    End If
Next i
' remove "SOLDERORDER":"0"
Dim lgthChar As Integer

lgthChar = 15 'Len("SOLDERORDER")
textOutput = Replace(textOutput, "SOLDERORDER", "")
textOutput = Replace(textOutput, ":", "")
textOutput = Replace(textOutput, """", "")
If StrComp(textOutput, "", vbTextCompare) Then
    SolderOrder = CInt(textOutput)
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

