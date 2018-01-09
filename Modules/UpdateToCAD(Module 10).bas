Attribute VB_Name = "Module10"

Sub MatchData()
    Call OpenFile("Arterial Counting")
End Sub
Sub OpenFile(Destination As String) 'This reads the data exported from AutoCAD. It is probably useless to you...
Dim filePath As String
    filePath = "H:\AutoLisp\CADexport.csv"
    
Open filePath For Input As #1
On Error GoTo 1
Count = 0
Do Until EOF(1)
Count = Count + 1
    Line Input #1, linefromfile
    If Count < 2 Then GoTo 1
    lineitems = Split(linefromfile, " ,")
    With Worksheets(Destination)
    i = 2
        Do While Not IsEmpty(.Cells(i, 2))
         a = 0
         abc = .Cells(i, 5).value 'Watcher...
         def = val(Mid(lineitems(0), 2)) 'watcher....
         j = 8
2:       If abc = def And IsEmpty(Cells(i, j)) Then
                  .Cells(i, j) = Left(lineitems(1), Len(lineitems(1)))
         ElseIf abc = def Then
                  j = j + 1
                  GoTo 2
         End If
        i = i + 1
        Loop
1:    End With
Loop
Close 1

    
End Sub
Sub WriteToFile(Source) 'This exports the traffic counts data to a file readable by autocad
Dim file As String
file = "H:\AutoLisp\output.csv"
Dim content As String
i = 2
Open file For Output As #2
    With Worksheets(Source)
        Do While Not IsEmpty(.Cells(i, 5))
            content = .Cells(i, 5).value & "," & .Cells(i, 7).value
            Print #2, content
            If Not IsEmpty(Cells(i, 8)) Then
                content = .Cells(i, 5).value & "," & .Cells(i, 8).value
                Print #2, content
            End If
        i = i + 1
        Loop
    End With
Close #2
        
        

End Sub
