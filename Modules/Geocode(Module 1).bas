Attribute VB_Name = "Module1"
'This module contains the online Geocoding code. I am not the author of the code. The
'original code can be found on http://grindgis.com/software/microsoft-excel/geocoding-excel-and-google

Function MyGeocode(address As String) As String
  Dim strAddress As String
  Dim strQuery As String
  Dim strLatitude As String
  Dim strLongitude As String
  strAddress = URLEncode(address)
  'Assemble the query string
  strQuery = "http://maps.googleapis.com/maps/api/geocode/xml?"
  strQuery = strQuery & "address=" & strAddress
  strQuery = strQuery & "&sensor=false"
  'define XML and HTTP components
  Dim googleResult As New MSXML2.DOMDocument
  Dim googleService As New MSXML2.XMLHTTP
  Dim oNodes As MSXML2.IXMLDOMNodeList
  Dim oNode As MSXML2.IXMLDOMNode
  'create HTTP request to query URL - make sure to have
  'that last "False" there for synchronous operation
  googleService.Open "GET", strQuery, False
  googleService.send
  googleResult.LoadXML (googleService.responseText)
  Set oNodes = googleResult.getElementsByTagName("geometry")
  If oNodes.length = 1 Then
    For Each oNode In oNodes
      strLatitude = oNode.ChildNodes(0).ChildNodes(0).text
      strLongitude = oNode.ChildNodes(0).ChildNodes(1).text
      MyGeocode = strLatitude & "," & strLongitude
    Next oNode
  Else
    MyGeocode = "Not Found (Slightly alter the address, then try again)"
  End If
End Function
Public Function URLEncode(StringVal As String, Optional SpaceAsPlus As Boolean = False) As String
  Dim StringLen As Long: StringLen = Len(StringVal)
  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String
    If SpaceAsPlus Then Space = "+" Else Space = "%20"
    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
      Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
        result(i) = Char
      Case 32
        result(i) = Space
      Case 0 To 15
        result(i) = "%0" & Hex(CharCode)
      Case Else
        result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

