Attribute VB_Name = "modAddText"
Public Sub AutocadAddText(ByRef D As AutoCADTextData)
    'テキストを選んで消去
    DeleteAllTextOnly
    Dim E As Extents2D
    modGeometry.InitExtents E

    Dim o As Object

    Set o = PutText(D.roomName, RN, True): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Con_Level, ConLV, True): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Finish_Level, FinLV, True): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Roof_H_01, RH, True): modGeometry.UpdateExtents E, o

    Set o = PutText(D.Floor_Base, FB): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Floor_Finish, FF): modGeometry.UpdateExtents E, o
    Set o = PutText(D.S_bord, SB): modGeometry.UpdateExtents E, o
    Set o = PutText(D.S_bord_H, SBH): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Wall_Base, WBa): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Wall_Finish_01, WF1): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Wall_Finish_02, WF2): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Molding, Mo): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Roof_Base, RB): modGeometry.UpdateExtents E, o
    Set o = PutText(D.Roof_Finish, RF): modGeometry.UpdateExtents E, o

    ' 備考 値が入ってなかった場合スキップ
    Set o = PutText(D.Remarks_01, MK1)
    If Not o Is Nothing Then modGeometry.UpdateExtents E, o
    Set o = PutText(D.Remarks_02, MK2)
    If Not o Is Nothing Then modGeometry.UpdateExtents E, o
    Set o = PutText(D.Remarks_03, MK3)
    If Not o Is Nothing Then modGeometry.UpdateExtents E, o
    
    
    '記述したdwgを保存 外部参照ブロックとして呼び出してくれい
    acadDoc.Regen acAllViewports
    SaveRoomDWG D.roomName

'    ' 表示範囲ごとブロック化
'    CreateRoomBlockByExtents D.roomName, E
'
'    ' ブロック出力
'    ExportBlockToDWG D.roomName, "C:\Temp\RoomBlocks\" & D.roomName & ".dwg"

End Sub

Public Sub DeleteAllTextOnly()

    Dim ss As Object

    ' 既存 SelectionSet 削除
    On Error Resume Next
    acadDoc.SelectionSets.Item("SS_DEL_TEXT").Delete
    On Error GoTo 0

    Set ss = acadDoc.SelectionSets.Add("SS_DEL_TEXT")

    ' フィルタ：TEXT のみ
    Dim fType(0) As Integer
    Dim fData(0) As Variant

    fType(0) = 0
    fData(0) = "TEXT"

    ' 全体から TEXT のみ選択
    ss.Select acSelectionSetAll, , , fType, fData

    ' 削除
    Dim i As Long
    For i = ss.Count - 1 To 0 Step -1
        ss.Item(i).Delete
    Next i

    ss.Delete

End Sub
