Attribute VB_Name = "Module8"
'This module contains the ghettomap sub
Sub ghettomap()
'---show help message
    If Worksheets("Temp Settings").Cells(3, 3).value = "Y" Then
        MsgBox ("*Note Ghetto Map has limited accuracy. It is meant to provide a general overveiw for the spots and assistance in updating the Road Classification and Traffic Counting map. For a fully functional map, vist batchgeo.com (the link can be found on the top left corner of the sheet)" & Chr(10) & Chr(10) & "To use the map:" & Chr(10) & "1. Make sure you have a list of roads, coordinates and latest count." & Chr(10) & "2.Select the ROAD NAMES of the spots you want to be displayed on the Map." & Chr(10) & "3.Hit 'Ghetto Map!' button")
        Exit Sub
    End If
'---delete existing series
    For Each s In Charts("Map").SeriesCollection
    s.Delete
    Next s
'---Set x y range. (The numbers are arbitrarily determined according to the map picture)
    Charts("Map").Axes(xlCategory).MinimumScale = -76.9
    Charts("Map").Axes(xlCategory).MaximumScale = -76.2
    Charts("Map").Axes(xlValue).MinimumScale = 44.3
    Charts("Map").Axes(xlValue).MaximumScale = 44.8
'---Creating a series for each data point (because series can have names displayed conviniently
    Dim n As Series
    For Each cell In Selection
        Set n = Charts("Map").SeriesCollection.NewSeries
        n.Values = cell.Offset(0, 1)
        n.XValues = cell.Offset(0, 2)
        n.name = cell.value & " [" & cell.Offset(0, 3).value & "] " & cell.Offset(0, 4).value
        n.MarkerStyle = xlMarkerStyleDiamond
        n.MarkerForegroundColor = RGB(0, 0, 0)
        n.MarkerBackgroundColor = cell.Interior.Color
        n.MarkerSize = 3.5
    Next cell
    Charts("Map").Activate
End Sub
