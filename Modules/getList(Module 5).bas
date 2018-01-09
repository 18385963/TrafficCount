Attribute VB_Name = "Module5"
'This module contains the getList sub
Sub getList(sourcesheet As String, Optional r As Integer = 7, Optional c As Integer = 5, Optional again As Boolean = False, Optional r2 As Integer = 28) 'i is the first row where data begins in the master file.
'[sourcesheet] enter the name of the source data sheet in the master file
'[r] the first row containing the counts
'[c] the first column containing the counts
'[again] set to true if the source sheet is separated to two segments (e.g. the Arterial is separated to boundary and internal)
'[r2] the first row containing the counts for the second segment

If Worksheets("Temp Settings").Cells(3, 3).value = "Y" Then
    MsgBox ("This button will get you a list of coordinates retrieved from the traffic counting summary file. All the spots will be colored according to when they have been last done. Red means it has not been done in 3 years, yellow 2, and green has been done last year. Spots that has been done this year are not colored")
    Exit Sub
End If
Call Thaw
' ------------Clear Formats ------------------
Dim area As Range
 Set area = Range("A2:Z400")
 area.ClearContents
 area.ClearFormats
' -----------Initializing -------------------
Dim filePath As String
    Call CheckFile
    filePath = Worksheets(1).Cells(5, 2).value & "\" & Worksheets(1).Cells(6, 2).value
Dim Currentwb As Object
    Set targetsheet = ActiveWorkbook.ActiveSheet
Dim masterfile As Workbook
    Set masterfile = Workbooks.Open(filePath)
Dim i As Integer 'Counter for looping through rows in masterfile
Dim j As Integer 'Counter for looping through rows in Student tools file
       j = 2
again: i = r
' -----------Anyalysis ----------------------
With masterfile.Worksheets(sourcesheet)
'----------- Ranking data based on columns----
' The rank is temporarily recorded as cell interior color in the first column(To avoid potential data conflict).
    Do While True
     If IsEmpty(.Cells(i, 2)) Then
        Exit Do
     End If
     .Cells(i, 1).Interior.Color = RGB(0, 0, 0)
     If Not IsEmpty(.Cells(i, c + 4).value) Then
        .Cells(i, 1).Interior.Color = RGB(255, 0, 0)
     End If
     If Not IsEmpty(.Cells(i, c + 2).value) Then
        .Cells(i, 1).Interior.Color = RGB(255, 255, 0)
     End If
     If Not IsEmpty(.Cells(i, c).value) Then
        .Cells(i, 1).Interior.Color = RGB(255, 255, 255)
     End If
     i = i + 1
    Loop
'----------- Transfer data to the target sheet (StudentTools)
    i = r
    Do While True
     If IsEmpty(.Cells(i, 2)) Then
        Exit Do
     End If
     If .Cells(i, 1).Interior.Color = RGB(0, 0, 0) Then
        targetsheet.Cells(j, 2).value = .Cells(i, 3).value
        targetsheet.Cells(j, 2).Interior.Color = RGB(255, 0, 0)
        j = j + 1
     End If
    i = i + 1
    Loop
    
    i = r
    Do While True
     If IsEmpty(.Cells(i, 2)) Then
        Exit Do
     End If
     If .Cells(i, 1).Interior.Color = RGB(255, 0, 0) Then
        targetsheet.Cells(j, 2).value = .Cells(i, 3).value
        targetsheet.Cells(j, 2).Interior.Color = RGB(255, 255, 0)
        j = j + 1
     End If
    i = i + 1
    Loop
    
    i = r
    Do While True
     If IsEmpty(.Cells(i, 2)) Then
        Exit Do
     End If
     If .Cells(i, 1).Interior.Color = RGB(255, 255, 0) Then
        targetsheet.Cells(j, 2).value = .Cells(i, 3).value
        targetsheet.Cells(j, 2).Interior.Color = RGB(0, 255, 0)
        j = j + 1
     End If
    i = i + 1
    Loop
    
    i = r
    Do While True
     If IsEmpty(.Cells(i, 2)) Then
        Exit Do
     End If
     If .Cells(i, 1).Interior.Color = RGB(255, 255, 255) Then
        targetsheet.Cells(j, 2).value = .Cells(i, 3).value
        j = j + 1
     End If
    i = i + 1
    Loop
    
    End With
        
    If again Then
        again = False
        r = r2
        GoTo again
    End If

masterfile.Close (False)

Call Ulock
Call Freeze
End Sub
