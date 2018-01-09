Attribute VB_Name = "Module7"
'This module contains the getData sub
Sub getData(sourcesheet As String, Optional r As Integer = 7, Optional c As Integer = 5, Optional again As Boolean = False, Optional r2 As Integer = 28, Optional averagecol As Integer = 4)
'[sourcesheet] enter the name of the source data sheet in the master file
'[r] the first row containing the counts
'[c] the first column containing the counts
'[again] set to true if the source sheet is separated to two segments (e.g. the Arterial is separated to boundary and internal)
'[r2] the first row containing the counts for the second segment
'[averagecol] the column number of the traffic average
If Worksheets("Temp Settings").Cells(3, 3).value = "Y" Then
    MsgBox ("This button will pair up your list with the lastest count data as well as the time when it was done.")
    Exit Sub
End If
Call Thaw
' ----------initialization ------------------
Dim filePath As String
    Call CheckFile
    filePath = Worksheets(1).Cells(5, 2).value & "\" & Worksheets(1).Cells(6, 2).value
Dim targetsheet As Object
    Set targetsheet = ActiveWorkbook.ActiveSheet
Dim masterfile As Workbook
    Set masterfile = Workbooks.Open(filePath)

Dim j As Integer 'counter for looping through columns
Dim ii As Integer 'counter for inner loop
    ii = 1
' ---------Extracting Latest data
Dim val As Double 'traffic data value
Dim t As String   'time retrieved
Dim size As Integer
    size = 2

With masterfile.Worksheets(sourcesheet)
'determine the size of the list
Do While Not IsEmpty(targetsheet.Cells(size, 2))
    size = size + 1
Loop
size = size - 1

'from master to student tool
again: i = r
Do While Not IsEmpty(.Cells(i, 3))
'----looping through columns right -> left, stop and store the first data it hits
    j = c
    Do While Not IsEmpty(.Cells(4, j))
        If Not IsEmpty(.Cells(i, j)) Then
          val = Round(.Cells(i, averagecol).value, 0) 'store the average in val
          t = .Cells(i, j + 1).value & ", " & .Cells(4, j).value 'store time in the form <Month, Year>
          Exit Do
        End If
    j = j + 2
    Loop
'----looping through the list: if match found, then write down val and t
    ii = 1
    Do While ii <= size
        If targetsheet.Cells(ii, 2).value = .Cells(i, 3).value Then
            targetsheet.Cells(ii, 5).value = val
            targetsheet.Cells(ii, 6).value = t
            Exit Do
        End If
    ii = ii + 1
    Loop
i = i + 1
Loop
'---- In the case when there are two segments of data (e.g. Arterial), repeat the process for the second segment
    If again Then
        again = False
        r = r2
        GoTo again
    End If
End With

masterfile.Close (False)
Call Ulock
Call Freeze
End Sub
