Attribute VB_Name = "Main"
Option Explicit
 Public acadApp As Object
 Public acadDoc As Object
 Public modelSpace As Object
Sub Main()
 Call Autocad_From_VBA
 Call EnsureTextStyle
  DoEvents
  acadDoc.Regen acAllViewports
 Call InitPoints
  Dim D As AutoCADTextData
  Dim i As Long, lastRow As Long
  Dim Split_Array() As String
  Dim Wb As Workbook, Ws As Worksheet
   Set Wb = ThisWorkbook
   Set Ws = Wb.ActiveSheet
  'Excel最終行取得
  lastRow = Ws.Cells(Ws.Rows.Count, 6).End(xlUp).row + 1
   'For i = 3 To 5
   For i = 3 To lastRow Step 2
        
        D.roomName = SanitizeText(Ws.Cells(i, 6).Value)
        If D.roomName = "" Then GoTo NextLoop
        
        D.Con_Level = SanitizeText(Ws.Cells(i + 1, 11).Value)
        D.Finish_Level = SanitizeText(Ws.Cells(i, 11).Value)
        D.Floor_Base = SanitizeText(Ws.Cells(i + 1, 8).Value)
        D.Floor_Finish = SanitizeText(Ws.Cells(i, 8).Value)
        D.S_bord = SanitizeText(Ws.Cells(i, 12).Value)
        D.S_bord_H = SanitizeText(Ws.Cells(i, 14).Value)
        If D.S_bord_H = "" Then D.S_bord_H = "-"

        D.Wall_Base = SanitizeText(Ws.Cells(i + 1, 16).Value)

        ' 壁仕上げ処理
        If Ws.Cells(i, 16).Value <> "" Then
         Split_Array = Split(Ws.Cells(i, 16).Value, vbLf)
         If UBound(Split_Array) = 0 Then
             D.Wall_Finish_01 = Split_Array(0)
             D.Wall_Finish_02 = ""
          Else
             D.Wall_Finish_01 = Split_Array(0)
             D.Wall_Finish_02 = Split_Array(1)
         End If
         Else
          D.Wall_Finish_01 = ""
          D.Wall_Finish_02 = ""
        End If
        D.Molding = SanitizeText(Ws.Cells(i, 21).Value)
        D.Roof_H_01 = SanitizeText(Ws.Cells(i, 22).Value)
        D.Roof_Base = SanitizeText(Ws.Cells(i + 1, 18).Value)
        D.Roof_Finish = SanitizeText(Ws.Cells(i, 18).Value)

        Split_Array = Split(Ws.Cells(i, 23).Value, vbLf)
        D.Remarks_01 = "": D.Remarks_02 = "": D.Remarks_03 = ""
        Select Case UBound(Split_Array)
            Case 0
                D.Remarks_01 = Split_Array(0)
            Case 1
                D.Remarks_01 = Split_Array(0)
                D.Remarks_02 = Split_Array(1)
            Case Is >= 2
                D.Remarks_01 = Split_Array(0)
                D.Remarks_02 = Split_Array(1)
                D.Remarks_03 = Split_Array(2)
        End Select
         'モジュール名と処理名は別にすること
         AutocadAddText D

NextLoop:
   Next i
End Sub
