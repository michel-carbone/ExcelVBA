Attribute VB_Name = "Module1"
' Function to add description in formula explorer of Excel. This Function must be executed manually as macro 
' this macro function does not add help while entering formula in formula bar
'
' function written for GetAttrValue function

Sub FunctionDescription()

   'Declaring the necessary variables
    Dim FuncName As String
    Dim FuncDesc As String
    Dim FuncCat As Variant
    
    'Depending on the function arguments define the necessary variables on the array.
    Dim ArgDesc(1 To 2) As String

    FuncName = "GetAttrValue"

    FuncDesc = "Parse field such as """"DIELECTRIC"""":""""X7R"""";""""DNP"""":""""F"""";""""PARTNO"""":""""1501"""";;""""SOLDERORDER"""":""""18"""";;""""TOLERANCE"""":""""10%"""";""""VALUE"""":""""1uF"""";;""""VOLTAGE"""":""""25V"""""
    
    'Choose the built-in function category.
    'For example, 14 is the UDF category
    FuncCat = 14
    
    'You can also use instead of numbers the full category name, for example:
    'FuncCat = "Engineering"
    'Or you can define your own custom category:
    'FuncCat = "My VBA Functions"
    
    'Here we add the description for the function's arguments.
    ArgDesc(1) = "Cell of attributes"
    ArgDesc(2) = "Attribute name"

    'Using the MacroOptions method add the function description (and its arguments).
    Application.MacroOptions _
        Macro:=FuncName, _
        Description:=FuncDesc, _
        Category:=FuncCat, _
        ArgumentDescriptions:=ArgDesc

    'Inform the user about the process.
    MsgBox FuncName & " was successfully added to the " & FuncCat & " category!", vbInformation, "Done"
    
End Sub
