Attribute VB_Name = "Module9"
'This module contains the getPic sub
Sub getPic(folder As String, _
                    Optional row As Integer = 12, _
                    Optional maxr As Integer = 500, _
                    Optional col As Integer = 2, _
                    Optional imgcol As Integer = 4, _
                    Optional rheight = 40)

'[folder] the folder containing the pictures
'[row] where first row containing the code start
'[maxr] maximum length of the list
'[col] column containing the OTM code
'[imgcol] column where the image is going to be inserted. *Notice that the image column takes two columns' space
'[rheight] the heigth of each row in the list
Dim pic As String
Dim newpic As String
Dim i As Integer
Dim Exist As String
Dim AR As Double
Dim ARCell As Double

Application.ScreenUpdating = False

i = row
  
Do While True
'---Terminate Conditions ----
 If i > maxr Then
    Exit Do
 End If
 
 If IsEmpty(Cells(i, col)) Then
    GoTo Skip
 End If
'---Formatting & Retrieving file name
 Rows(i).RowHeight = rheight
 pic = Cells(i, col).value
 picpath = folder & pic & ".bmp"
 Exist = Dir(picpath)
 
'---Search for matching files ---
 Dim position As String 'convert signs codes of different sizes
 If Exist = "" Then
    newpic = Left(pic, 3) & "10" & Mid(pic, 4)
    picpath = folder & newpic & ".bmp"
    Exist = Dir(picpath)
    If Exist = "" Then
        newpic = Left(pic, 3) & "1" & Mid(pic, 4)
        picpath = folder & newpic & ".bmp"
        Exist = Dir(picpath)
        If Exist = "" Then
            newpic = Left(pic, 3) & Mid(pic, 5)
            picpath = folder & newpic & ".bmp"
            Exist = Dir(picpath)
            If Exist = "" Then
                newpic = Left(pic, 3) & Mid(pic, 7)
                picpath = folder & newpic & ".bmp"
                Exist = Dir(picpath)
                If Exist = "" Then
                    newpic = Left(pic, 3) & Mid(pic, 6)
                    picpath = folder & newpic & ".bmp"
                    Exist = Dir(picpath)
                End If
            End If
        End If
    End If
 End If
 
'---Insert file ------------
 If Not Exist = "" Then
    With ActiveSheet.Pictures.Insert(picpath)
         With .ShapeRange
             .LockAspectRatio = msoTrue
             AR = .Width / .Height
             ARCell = (Cells(i, imgcol).Width + Cells(i, imgcol + 1).Width) / Cells(i, imgcol).Height
             If AR > ARCell Then
                .Width = Cells(i, imgcol).Width + Cells(i, imgcol + 1).Width - 10
             Else
                .Height = Cells(i, imgcol).Height - 6
             End If
         End With
         .Left = ActiveSheet.Cells(i, imgcol).Left + (Cells(i, imgcol + 1).Width + Cells(i, imgcol).Width - .Width) / 2
         .Top = ActiveSheet.Cells(i, imgcol).Top + (Cells(i, imgcol).Height - .Height) / 2
         .Placement = 1
         .PrintObject = True
     End With
 End If

Skip:  i = i + 1
 Loop
Application.ScreenUpdating = True
End Sub

