Attribute VB_Name = "Module2"
'This Module is used for geocoding the coodinate banks.

Dim gCount As Date
Dim i As Integer
Sub Geocode()
    i = 1
    Call Action
End Sub
'Updateby20140925
Sub Timer()
    MsgBox ("Resting...")
    gCount = Now + TimeValue("00:00:01")
    Application.OnTime gCount, "ResetTime"
End Sub
Sub ResetTime()
Dim xRng As Range
Set xRng = Application.ActiveSheet.Range("A1")
Dim initial As Date
initial = TimeValue("00:00:10")
xRng.value = xRng.value - TimeSerial(0, 0, 1)
If xRng.value <= 0 Then
    xRng.value = initial
    Call Action
    Exit Sub
End If
Call Timer
End Sub
Sub Action()
MsgBox ("In action")
Dim ind As Integer
Dim alter As String
Dim brk1 As Integer
Dim brk2 As Integer
Dim at As Integer
Dim ii As Integer

Do While True
    If IsEmpty(Cells(i, 2).value) Then
        i = 1
        Exit Do
    End If
    If i = ii + 10 Then
        Call Timer
        ii = i
        Exit Sub
    End If
    'If Left(Cells(i, 3).Value, 4) = "!Not" Then
            alter = Cells(i, 2).value
            brk1 = InStr(alter, "(")
            brk2 = InStr(alter, ")")
         If brk1 <> 0 Then
            alter = Left(alter, brk1 - 1) & Mid(alter, brk2 + 1)
         End If
            ind = InStr(alter, "Perth Road")
         If ind <> 0 Then
            alter = Left(Cells(i, 2), ind - 1) & "Hwy 10" & Mid(Cells(i, 2), ind + Len("Perth Road") + 2)
         End If
            'at = InStr(alter, "@")
         'If at <> 0 Then
            'alter = Left(alter, at - 1) & ",South Frontenac, ON, CA"
        ' End If
        Cells(i, 3).value = MyGeocode(alter)
        'Cells(i, 1).Value = "!"
    'End If
    i = i + 1
Loop
End Sub



