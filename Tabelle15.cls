VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tabelle15"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Worksheet_Change(ByVal Target As Range)

    If Intersect(Target, ThisWorkbook.Names("Erhebungsbogen_typ").RefersToRange) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    Update_Uebersicht_Jahre
    Application.EnableEvents = True

End Sub

