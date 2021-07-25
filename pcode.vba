Function PCODE(Code As String) As Variant()
    Dim ThisWorkbook As Workbook
    Dim AllTypeSheet As Worksheet
    Dim AllTypeArr(1 To 111, 2 To 3)
    Dim r As Byte, c As Byte
    
    Dim ProductInfo(0 To 3, 1 To 1)
    Dim counter As Integer
    counter = 0
    
    Set ThisWorkbook = ActiveWorkbook
    Set AllTypeSheet = ThisWorkbook.Worksheets("ALL TYPE")
    
    For r = 1 To 111
        For c = 2 To 3
            If c = 2 And AllTypeSheet.Cells(r, 2).Value = Code Then
                ProductInfo(counter, 1) = AllTypeSheet.Cells(r, 3).Value
                counter = counter + 1
            End If
        Next c
    Next r
    PCODE = ProductInfo
End Function