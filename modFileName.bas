Attribute VB_Name = "modFileName"
Public Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "_")
    Next i

    SanitizeFileName = Trim$(s)
End Function


Public Function SanitizeText(ByVal s As String) As String

    If Len(s) = 0 Then
        SanitizeText = ""
        Exit Function
    End If

    ' 怪しい改行を滅ぼす、それらは悪だ
    s = Replace(s, vbCrLf, " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    ' 全角スペース → 半角
    s = Replace(s, "　", " ")

    ' 全角 → 半角（英数字・記号）
    s = StrConv(s, vbNarrow)

    ' 特殊文字除去 バグるっつうの
    Dim badChars As Variant
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")

    Dim i As Long
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, badChars(i), "")
    Next i

    ' 連続スペースを1個に
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    ' 前後トリム
    s = Trim$(s)

    SanitizeText = s
End Function
