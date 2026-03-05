Attribute VB_Name = "modTextStyle"
Public Sub EnsureTextStyle()

    Dim ts As Object
    On Error Resume Next
    Set ts = acadDoc.TextStyles.Item("ROMANS")
    On Error GoTo 0

    If ts Is Nothing Then
        Set ts = acadDoc.TextStyles.Add("ROMANS")
    End If

    ts.fontFile = "romans.shx"
    ts.BigFontFile = "extfont2.shx"

    acadDoc.ActiveTextStyle = ts
End Sub

