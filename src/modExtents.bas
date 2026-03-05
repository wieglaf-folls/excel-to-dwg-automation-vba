Attribute VB_Name = "modExtents"
Option Explicit

Public Sub InitExtents(ByRef E As Extents2D)
    E.IsInitialized = False
End Sub

Public Sub UpdateExtents(ByRef E As Extents2D, ByVal obj As Object)
    Dim minPt As Variant, maxPt As Variant
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

