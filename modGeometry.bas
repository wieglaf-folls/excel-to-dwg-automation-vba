Attribute VB_Name = "modGeometry"
Option Explicit

'========================
' 型定義（Type型が先頭、それ以降にFunction）
'========================
Public Type Point2D
    x As Double
    y As Double
End Type

Public Type Extents2D
    MinX As Double
    MinY As Double
    MaxX As Double
    MaxY As Double
    IsInitialized As Boolean
End Type

'========================
' 関数・処理
'========================
Public Function MakePoint2D(x As Double, y As Double) As Point2D
    MakePoint2D.x = x
    MakePoint2D.y = y
End Function

Public Sub InitExtents(ByRef E As Extents2D)
    E.IsInitialized = False
End Sub

Public Sub UpdateExtents(ByRef E As Extents2D, ByVal obj As Object)

    If obj Is Nothing Then Exit Sub

    Dim minPt As Variant
    Dim maxPt As Variant

    obj.GetBoundingBox minPt, maxPt

    If Not E.IsInitialized Then
        E.MinX = minPt(0)
        E.MinY = minPt(1)
        E.MaxX = maxPt(0)
        E.MaxY = maxPt(1)
        E.IsInitialized = True
    Else
        If minPt(0) < E.MinX Then E.MinX = minPt(0)
        If minPt(1) < E.MinY Then E.MinY = minPt(1)
        If maxPt(0) > E.MaxX Then E.MaxX = maxPt(0)
        If maxPt(1) > E.MaxY Then E.MaxY = maxPt(1)
    End If

End Sub

