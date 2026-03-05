Attribute VB_Name = "modText"
Public Function PutText(ByVal txt As String, ByRef P As Point2D, Optional ByVal IsCenter As Boolean = False)

    If Trim$(txt) = "" Then
     Set PutText = Nothing
     Exit Function
    End If
    
    Dim pt(0 To 2) As Double
    pt(0) = P.x
    pt(1) = P.y
    pt(2) = 0

    Dim oText As Object
    Set oText = modelSpace.AddText(txt, pt, 125)

    ' ïùåWêî
     acadDoc.ActiveTextStyle.Width = 0.8
    oText.StyleName = acadDoc.ActiveTextStyle.Name

    ' îzíu
    If IsCenter Then
        oText.Alignment = acAlignmentMiddleCenter
        oText.TextAlignmentPoint = pt
    Else
        oText.Alignment = acAlignmentMiddleLeft
        oText.TextAlignmentPoint = pt
    End If

    oText.Update
    
    '
    Set PutText = oText
    
End Function

