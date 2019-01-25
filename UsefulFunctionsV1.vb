
Module UsefulFunctionsV1


    Sub RemoveByValue(Of TKey, TValue)(ByVal dictionary As Dictionary(Of TKey, TValue), ByVal someValue As TValue)
        '// How to remove items from a dictionary
        Dim itemsToRemove = (From pair In dictionary
                             Where pair.Value.Equals(someValue)
                             Select pair.Key).ToArray()
        '//
        For Each item As TKey In itemsToRemove
            dictionary.Remove(item)
        Next
    End Sub

    ''Usage:
    ''    Dim dictionary As New Dictionary(Of Int32, String)
    ''dictionary.Add(1, "foo")
    ''dictionary.Add(2, "foo")
    ''dictionary.Add(3, "bar")

    ''RemoveByValue(dictionary, "foo")

    '//========================================================

    ' Public Function IsPointInPolygon(Vector2 point, Vector2[] polygon) As Boolean
    '     {
    'int polygonLength = polygon.Length, i = 0;
    'bool inside = False;
    '// x, y for tested point.
    'float pointX = Point.X, pointY = Point.Y;
    '// start / end point for the current polygon segment.
    'float startX, startY, endX, endY;
    'Vector2 endPoint = polygon[polygonLength-1];           
    'endX = endPoint.x; 
    'endY = endPoint.y;
    'While (i < polygonLength) {
    '   startX = endX;           startY = endY;
    '   endPoint = polygon[i++];
    '   endX = endPoint.x;       endY = endPoint.y;
    '   //
    '   inside ^= ( endY > pointY ^ startY > pointY ) /* ? pointY inside [startY;endY] segment ? */
    '             && /* if so, test if it Is under the segment */
    '             ( (pointX - endX) < (pointY - endY) * (startX - endX) / (startY - endY) ) ;
    '}
    'Return inside;




End Module

Public Class T_InsidePolyGon
    Public constant As New List(Of Double)
    Public multiple As New List(Of Double)
    '//
    '// https://stackoverflow.com/questions/29283871/using-a-raycasting-algorithm-for-point-in-polygon-test-with-latitude-longitude-c
    '// It is running through the full list of 1.8 Million points in under 20 seconds 
    '// (compared to 1 hour And 30 minutes using the DotSpatial.Contains function).
    '// This has been done with ref types on the heap (lists), rather than var tpes on the stack (arrays), so could be slower.
    Private Sub precalc_values(polygon As List(Of INode))
        '// do this once for a given polygon
        Dim i As Integer
        Dim j As Integer = polygon.Count - 1
        Dim V_i_X As Double
        Dim V_i_Y As Double
        Dim V_j_X As Double
        Dim V_j_Y As Double
        '//
        For i = 0 To polygon.Count - 1
            constant.Add(0)
            multiple.Add(0)
        Next
        '//
        For i = 0 To polygon.Count - 1
            V_i_X = polygon.Item(i).X
            V_i_Y = polygon.Item(i).Y
            V_j_X = polygon.Item(j).X
            V_j_Y = polygon.Item(j).Y
            '//
            If V_j_Y = V_i_Y Then
                constant(i) = V_i_X
                multiple(i) = 0
            Else
                constant(i) = V_i_X - (V_i_Y * V_j_X) / (V_j_Y - V_i_Y) + (V_i_Y * V_i_X) / (V_j_Y - V_i_Y)
                multiple(i) = (V_j_X - V_i_X) / (V_j_Y - V_i_Y)
            End If
            j = i
        Next
    End Sub

    Private Function IsPointInPolygon(point As INode, polygon As List(Of INode)) As Boolean
        Dim i As Integer
        Dim j As Integer = polygon.Count - 1
        Dim oddNodes As Boolean = False
        Dim V_i As INode
        Dim V_j As INode
        '//
        For i = 0 To polygon.Count - 1
            V_i = polygon.Item(i)
            V_j = polygon.Item(j)
            If (V_i.Y < point.Y AndAlso V_j.Y >= point.Y OrElse V_j.Y < point.Y AndAlso V_i.Y >= point.Y) Then
                oddNodes = oddNodes Xor (point.Y * multiple(i) + constant(i) < point.X)
            End If
            j = i
        Next
        '//
        Return oddNodes
    End Function

    Public Function PointsInPolygonList(pointList As List(Of INode), polygon As List(Of INode)) As List(Of INode)
        Dim i As Integer
        Dim aPointToTest As INode
        Dim PointsInside As New List(Of INode)
        '//
        precalc_values(polygon)
        '//
        For i = 0 To pointList.Count - 1
            aPointToTest = pointList.Item(i)
            '//
            If IsPointInPolygon(aPointToTest, polygon) Then
                PointsInside.Add(aPointToTest)
            End If
        Next
        Return PointsInside
    End Function
End Class

Public Class TWritePloyGon
    Public Sub WritePolygonToCSV(polygon As List(Of INode), FilePath As String)
        Dim aStringlist As New TStringList
        Dim i As Integer
        Dim aPointToWritr As INode
        Dim line As String
        Dim ErrorStr As String = ""
        '//
        For i = 0 To polygon.Count - 1
            aPointToWritr = polygon.Item(i)
            '//
            line = aPointToWritr.Name + "," + aPointToWritr.X.ToString + "," + aPointToWritr.Y.ToString
            aStringlist.Add(line)
        Next
        If Not aStringlist.SaveToFile(FilePath, ErrorStr) Then
            MsgBox("Could not write edge list " + ErrorStr)
        End If
    End Sub
End Class
