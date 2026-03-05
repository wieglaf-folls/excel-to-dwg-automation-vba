Attribute VB_Name = "modPoint"
Public RN As Point2D '部屋名
Public ConLV As Point2D '躯体レベル
Public FinLV As Point2D '仕上レベル
Public FB As Point2D '床下地
Public FF As Point2D '床仕上
Public RH As Point2D '天井高
Public SK As Point2D '幅木
Public WBa As Point2D '壁下地
Public WF1 As Point2D '壁仕上1
Public WF2 As Point2D '壁仕上2
Public MK1 As Point2D '備考1
Public MK2 As Point2D '備考2
Public MK3 As Point2D '備考3
Public Mo As Point2D '廻縁
Public RF As Point2D '天井仕上
Public RB As Point2D '天井下地
Public SB As Point2D    ' ← 追加
Public SBH As Point2D   ' ← 追加

Public Sub InitPoints()
    RN = MakePoint2D(2000, 2137.5)
    ConLV = MakePoint2D(800, 1762.5)
    FinLV = MakePoint2D(1400, 1762.5)
    FB = MakePoint2D(1750, 1937.5)
    FF = MakePoint2D(1750, 1762.5)
    RH = MakePoint2D(800, 537.5)
    SK = MakePoint2D(1150, 1412.5)
    WBa = MakePoint2D(1150, 1414)
    WF1 = MakePoint2D(1150, 1237.5)
    WF2 = MakePoint2D(1150, 1062.5)
    SB = MakePoint2D(550, 1587.5)
    SBH = MakePoint2D(3640, 1587.5)
    Mo = MakePoint2D(550, 885)
    RF = MakePoint2D(1740, 700)
    RB = MakePoint2D(1740, 550)
    MK1 = MakePoint2D(550, 370)
    MK2 = MakePoint2D(550, 225)
    MK3 = MakePoint2D(550, 80)
End Sub
