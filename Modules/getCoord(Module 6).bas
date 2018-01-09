Attribute VB_Name = "Module6"
Sub getCoord(banksheet As String, Optional comments As Boolean = False)

If Worksheets("Temp Settings").Cells(3, 3).value = "Y" Then
    MsgBox ("This button will pair up the list with coordinates. Most of the coordinates are from the local coordinate bank. For the ones that can't be found in the coordinate bank, online Geocoding via Google Maps API will be used instead.")
    Exit Sub
End If

Call Thaw
Dim targetsheet As Worksheet
    Set targetsheet = ActiveSheet
Dim i As Integer 'Counter for looping through target sheet
Dim j As Integer 'Counter for looping through coordinate bank
' ---- For online geocode
Dim ind As Integer 'index of the road name of interest
Dim found As Boolean
Dim alter As String 'variable to hold the altered road name
Dim brk1 As Integer 'index of left bracket
Dim brk2 As Integer 'index of right bracket
Dim at As Integer 'index of "@"

' ---- Matching coordinates with the sheet
i = 2

Do While True
    If IsEmpty(Cells(i, 2).value) Then
        Exit Do
    End If
    j = 1
    found = False
    Do While True
        If IsEmpty(Worksheets(banksheet).Cells(j, 2).value) Then
            Exit Do
        End If
        If targetsheet.Cells(i, 2).value = Worksheets(banksheet).Cells(j, 2).value Then
            If Worksheets(banksheet).Cells(j, 1).value = "!" Then
                targetsheet.Cells(i, 7).value = "Exact location of this counter spot cannot be found. The coodinates are just for the road"
            End If
            ind = InStr(Worksheets(banksheet).Cells(j, 3).value, ",")
            targetsheet.Cells(i, 3).value = Left(Worksheets(banksheet).Cells(j, 3).value, ind - 1)
            targetsheet.Cells(i, 4).value = Mid(Worksheets(banksheet).Cells(j, 3).value, ind + 1)
            found = True
        End If
    j = j + 1
    Loop
    
' ---- Online geocode
    If Not found Then
        If MsgBox("One entry is not found in the local coordinate bank. The application is about to use the online geocoding system. This can take a few seconds. Please do not click anything within 30s once the process begins (even if excel freezes). Proceed?", vbYesNo, "confirmation") = vbYes Then
            alter = Cells(i, 2).value & ", South Frontenac, ON, CA"
            brk1 = InStr(alter, "(")
            brk2 = InStr(alter, ")")
             If brk1 <> 0 Then
                alter = Left(alter, brk1 - 1) & Mid(alter, brk2 + 1)
             End If
                ind = InStr(alter, "Perth Road")
             If ind <> 0 Then
                alter = Left(Cells(i, 2), ind - 1) & "Hwy 10" & Mid(Cells(i, 2), ind + Len("Perth Road") + 2)
            End If
                at = InStr(alter, "@")
            If at <> 0 Then
               alter = Left(alter, at - 1) & ",South Frontenac, ON, CA"
               Cells(i, 7).value = "Exact location of this counter spot cannot be found. The coodinates are just for the road"
            End If
            alter = MyGeocode(alter)
            ind = InStr(alter, ",")
            Cells(i, 3).value = Left(alter, ind - 1)
            Cells(i, 4).value = Mid(alter, ind + 1)
         End If
    End If
i = i + 1
Loop

Call Ulock
Call Freeze
End Sub
