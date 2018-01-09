Attribute VB_Name = "Module3"
'This modules checks the existence of the masterfile as well as
'modifies the exisiting filename with wildcard characters to
'accomodate year changes in the file name

Sub CheckFile()
Worksheets(1).Unprotect "123"
Dim filePath As String
Dim ind As Integer
Dim year As Integer
Dim Exist As String
Dim folder As String
Dim file As String
    year = 2017
    folder = Worksheets(1).Cells(5, 2).value
    file = Worksheets(1).Cells(6, 2).value
    filePath = folder & "\" & DigitsAmbi(file) 'modifies the filename with wild card characters
    file = Dir(filePath) 'get the actual file name
    If file = "" Then 'in the case when the file is not found...
        MsgBox ("The masterfile is not found. Check the spelling of the path (Folder & File Name) on the main page.")
        End
    End If
    filePath = folder & "\" & file 'new filepath (not in used at the moment. However, can be useful to have)
    Worksheets(1).Cells(9, 2).value = Date 'writes down the date updated on the main page
    Worksheets(1).Cells(6, 2).value = file 'updates the new file name on the main page
Worksheets(1).Protect "123"
End Sub

Function DigitsAmbi(s As String) As String
    ' Variables needed (remember to use "option explicit").   '
    Dim retval As String    ' This is the return string.      '
    Dim i As Integer        ' Counter for character position. '

    ' Initialise return string to empty                       '
    retval = ""

    ' For every character in input string, copy digits to     '
    '   return string.                                        '
    For i = 1 To Len(s)
        If Mid(s, i, 1) >= "0" And Mid(s, i, 1) <= "9" Then
            s = Left(s, i - 1) + "?" + Mid(s, i + 1)
        End If
    Next
    DigitsAmbi = s
End Function
