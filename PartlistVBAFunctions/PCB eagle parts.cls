VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Dim sFullName As String
Dim sFileName As String
Dim dateUpdate As Date

Sub Update_eagle_parts_Click()
    sFullName = Application.GetOpenFilename("*_ListByParts.txt,*.txt")
    sFileName = Dir(sFullName)
    Debug.Print sFullName, sFileName, dateUpdate
    ImportFile (sFullName)
    ThisWorkbook.ActiveSheet.Range("F1").Formula = sFullName
    ThisWorkbook.ActiveSheet.Range("F2").Formula = dateUpdate
    ActiveSheet.Calculate
End Sub


Sub ImportFile(fName)
    'https://www.experts-exchange.com/questions/29083977/Excel-VBA-Import-Data-From-text-csv.html
    dateUpdate = Date
    Dim wsData As Worksheet
    Dim FileContents As String, iFF As Long
    Dim fRecords() As String, arrTemp() As Variant, arrFileRecords() As Variant
    Dim RecordCount As Long, FieldCount As Long, i As Long, j As Long
    Dim charSep As String
    
    'Define file
    'fName = "D:\HeatMasterLOG-20190415-1756.txt"
    charSep = vbTab
    'Define sheet for data
    Set wsData = ThisWorkbook.ActiveSheet
    
    'Remove old data on sheet
    wsData.Activate
    wsData.Range("$D4:$L1000").ClearContents
    
    'Pull entire text file into FileContents variable
    iFF = FreeFile
    Open fName For Binary Access Read As #iFF
    FileContents = Space$(LOF(iFF))
    Get #iFF, , FileContents
    Close #iFF
    
    'split by record-delimiter
    fRecords = Split(FileContents, vbCrLf)
    
    'Create array to store entire text file's fields and records
    RecordCount = UBound(fRecords)
    ReDim arrTemp(RecordCount)
    
    'Split each record by field delimiter
    For i = 0 To RecordCount
        arrTemp(i) = Split(fRecords(i), charSep)
        'arrTemp(i) = fRecords(i).Split(vbTab)
    Next i
    
    'Determine max number of fields
    FieldCount = 0
    For i = 0 To RecordCount
        If UBound(arrTemp(i)) > FieldCount Then
            FieldCount = UBound(arrTemp(i))
        End If
    Next i
    
    'Create 2 dimensional array to store all data
    ReDim arrFileRecords(RecordCount, FieldCount)
    For i = 0 To RecordCount
        For j = 0 To UBound(arrTemp(i))
            arrFileRecords(i, j) = arrTemp(i)(j)
        Next j
    Next i
    
    'Transfer array data to sheet
    wsData.Range("D3").Resize(RecordCount + 1, FieldCount + 1).Value = arrFileRecords
End Sub


