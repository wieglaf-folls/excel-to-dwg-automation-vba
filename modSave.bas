Attribute VB_Name = "modSave"
Option Explicit

Public Sub SaveRoomDWG(ByVal roomName As String)

    Dim basePath As String
    basePath = ThisWorkbook.Path & "\RoomDWG\"

    ' フォルダ無ければ作成
    If Dir(basePath, vbDirectory) = "" Then
        MkDir basePath
    End If

    ' ファイル名（使用不可文字対策）
    Dim safeName As String
    safeName = SanitizeFileName(roomName)

    Dim fullPath As String
    fullPath = basePath & safeName & ".dwg"

    ' 既存あれば上書き
    acadDoc.SaveAs fullPath

End Sub
