Attribute VB_Name = "modSelection"
Public Function CreateRoomSelectionSet(ByVal roomName As String) As Object

    Dim ss As Object
    Dim ssName As String
    ssName = "SS_" & roomName

    ' ä˘ë∂Ç™Ç†ÇÍÇŒçÌèú
    On Error Resume Next
    acadDoc.SelectionSets.Item(ssName).Delete
    On Error GoTo 0

    Set ss = acadDoc.SelectionSets.Add(ssName)
    Set CreateRoomSelectionSet = ss

End Function
