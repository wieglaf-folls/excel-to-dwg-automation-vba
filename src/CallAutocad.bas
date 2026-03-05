Attribute VB_Name = "CallAutocad"
Public Sub Autocad_From_VBA()
    ' AutoCAD オブジェクトを取得
    On Error Resume Next
    Set acadApp = GetObject(, "AutoCAD.Application")
    If acadApp Is Nothing Then Set acadApp = CreateObject("AutoCAD.Application")
    On Error GoTo 0

    acadApp.Visible = True
    Dim dwgPath As String
    'テンプレートファイル読み込み
     dwgPath = ThisWorkbook.Path & "\Block\block_template.dwg"
    On Error Resume Next
    Set acadDoc = acadApp.Documents.Open(dwgPath)
     If Err.Number <> 0 Then
        MsgBox "ファイルを開けませんでした: " & dwgPath & vbCrLf & "エラー: " & Err.Description
        Exit Sub
     End If
     On Error GoTo 0
    
    
'
'    ' Document を取得
'    If acadApp.Documents.Count = 0 Then
'        Set acadDoc = acadApp.Documents.Add
'    Else
'        Set acadDoc = acadApp.ActiveDocument
'    End If

    'モデルスペースの取得
    Set modelSpace = acadDoc.modelSpace
End Sub
