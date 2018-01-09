Attribute VB_Name = "Module4"
'This helps protect the worksheets

Sub Thaw()
ActiveSheet.Unprotect "123"
End Sub

Sub Freeze()
If Worksheets(1).Cells(8, 2).value = "Y" Then
ActiveSheet.Protect "123", AllowSorting:=True, AllowFormattingCells:=True

End If
End Sub

Sub Ulock()
Range(Cells(2, 2), Cells(999, 999)).Locked = False
End Sub


































Function BinarySearch(first As Range, last As Range, Target As Range) As Range 'This performs a binary search for a selected column of cells sorted in ascending order
' THIS CODE IS NOT FINISHED
'-----------initialization ----------------
last = last.Offset(1, 0) 'move the last cell down by one (becasue the program will be infinite loop if target = last cell)
Dim fval As String 'value of first cell
Dim lval As String  'Value of last cell
    fval = Asc(first.value)
    lval = Asc(last.value)
Dim l As Integer 'Length of the list
    l = Abs(last.row - (first.row - 1))
Dim c As Integer 'The colum number
Dim piv As Range 'The pivot cell
    Set piv = first.Offset(Round(l / 2 - 1, 0), 0)
    c = first.column
    MsgBox (val(piv.value) & " " & fval & " " & lval)
'-----------comparison--------------------
Do While Asc(piv.value) > fval Or Asc(piv.value) < lval
   
    If val(Target.value) <= piv Then
        Set last = piv.Offset(-1, 0)
        lval = Asc(last.value)
    ElseIf val(targe.value) >= piv Then
        Set first = piv.Offset(1, 0)
        fval = Asc(first.value)
    End If
    l = Abs(last.row - (first.row - 1))
    piv = first.Offset(Round(l / 2 - 1, 0), 0)
    MsgBox (piv)
Loop
Set BinarySearch = first
End Function

Sub bsearch()
'FOR TESTING PURPOSES ONLY
Dim area As Range
Dim i As Integer
    i = 0
Set area = Selection
Dim f As Range
Dim l As Range
Dim Target As Range

    
Set f = Range(Cells(1, 2), Cells(1, 2))
Set l = Range(Cells(50, 2), Cells(50, 2))
Set Target = Selection
MsgBox (BinarySearch(f, l, Target).value)
End Sub
