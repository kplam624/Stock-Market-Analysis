Attribute VB_Name = "Module1"
Sub test()
Dim row_count As Integer
Dim j As Integer
 
row_count = 2

For j = 1 To 10
    If row_count = 11 Then
        j = 10
    ElseIf IsEmpty(Cells(j + 1, 4)) = True Then
        Cells(j + 1, 4).Value = Cells(row_count, 1).Value
        row_count = row_count + 1
        j = 0
    ElseIf IsEmpty(Cells(j + 1, 4)) = False And Cells(j + 1, 4).Value = Cells(row_count, 1).Value Then
        row_count = row_count + 1
        j = 0
    'ElseIf Cells(j, 4).Value = Cells(row_count, 1).Value Then
        'row_count = row_count + 1
    ElseIf Cells(j + 1, 4).Value = Cells(row_count, 1).Value Then
        row_count = row_count + 1
        j = 0
    End If
Next j

End Sub




