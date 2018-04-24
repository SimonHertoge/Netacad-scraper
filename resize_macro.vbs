Sub resize()
Dim i As Long
With ActiveDocument
    For i = 1 To .InlineShapes.Count
        With .InlineShapes(i)
            .ScaleHeight = 100
            .ScaleWidth = 100
        End With
    Next i
End With
End Sub 
