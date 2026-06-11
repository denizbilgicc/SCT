Attribute VB_Name = "EHB_Typ"
Option Explicit

Public Sub Update_Uebersicht_Jahre()

    Const WS_OVERVIEW As String = "Übersicht"
    Const WS_TABLE As String = "Übersicht_Tabelle"
    Const NM_EHB_TYP As String = "Erhebungsbogen_typ"
    Const LO_EHB_JAHREN As String = "EHB_Jahren"
    Const START_YEAR As Long = 2025
    Const MAX_YEARS As Long = 5

    Dim wsU As Worksheet
    Dim wsT As Worksheet
    Dim ehbTyp As String
    Dim targetSheets As Variant
    Dim colTab As Long, colKey As Long, colDesc As Long, colYear As Long
    Dim lastRow As Long
    Dim r As Long
    Dim sheetName As String
    Dim yearCount As Long

    On Error GoTo CleanExit

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsU = ThisWorkbook.Worksheets(WS_OVERVIEW)
    Set wsT = ThisWorkbook.Worksheets(WS_TABLE)

    ehbTyp = Trim(CStr(ThisWorkbook.Names(NM_EHB_TYP).RefersToRange.Value))

    If ehbTyp = "" Then
        MsgBox "Bitte zuerst einen Erhebungsbogen auswählen.", vbExclamation
        GoTo CleanExit
    End If

    targetSheets = Array("A2_SaLi", "B_Bilanz", "C_GuV")

    colTab = FindHeaderColumn(wsU, "Tabellenblatt")
    colKey = FindHeaderColumn(wsU, "Kürzel")
    colDesc = FindHeaderColumn(wsU, "Beschreibung")
    colYear = FindHeaderColumn(wsU, "Jahr")

    If colTab = 0 Or colKey = 0 Or colDesc = 0 Or colYear = 0 Then
        MsgBox "Die Spalten Tabellenblatt, Kürzel, Beschreibung oder Jahr wurden nicht gefunden.", vbCritical
        GoTo CleanExit
    End If

    lastRow = wsU.Cells(wsU.Rows.Count, colKey).End(xlUp).Row

    For r = lastRow To 1 Step -1

        sheetName = GetCurrentSheetName(wsU, r, colTab)

        If IsInArray(sheetName, targetSheets) Then

            If Trim(CStr(wsU.Cells(r, colKey).Value)) <> "" _
               And Trim(CStr(wsU.Cells(r, colDesc).Value)) <> "" Then

                yearCount = GetYearCount(wsT, LO_EHB_JAHREN, ehbTyp, sheetName)

                If yearCount > 0 Then
                    ApplyYearVisibility wsU, r, colYear, yearCount, START_YEAR, MAX_YEARS
                End If

            End If

        End If

    Next r

CleanExit:
    Application.EnableEvents = True
    Application.ScreenUpdating = True

End Sub


Private Sub ApplyYearVisibility( _
    ByVal ws As Worksheet, _
    ByVal questionRow As Long, _
    ByVal colYear As Long, _
    ByVal yearCount As Long, _
    ByVal startYear As Long, _
    ByVal maxYears As Long)

    Dim i As Long
    Dim currentRow As Long

    If yearCount > maxYears Then yearCount = maxYears

    For i = 1 To maxYears

        currentRow = questionRow + i - 1

        ws.Cells(currentRow, colYear).Value = startYear - i + 1

        If i <= yearCount Then
            ws.Rows(currentRow).Hidden = False
        Else
            ws.Rows(currentRow).Hidden = True
        End If

    Next i

End Sub


Private Function GetYearCount( _
    ByVal ws As Worksheet, _
    ByVal tableName As String, _
    ByVal ehbTyp As String, _
    ByVal sheetName As String) As Long

    Dim lo As ListObject
    Dim lr As ListRow
    Dim colEHB As Long
    Dim colSeiten As Long
    Dim colJahren As Long
    Dim valueJahren As Variant

    Set lo = ws.ListObjects(tableName)

    colEHB = lo.ListColumns("Erhebungsbogen").Index
    colSeiten = lo.ListColumns("Seiten").Index
    colJahren = lo.ListColumns("Jahren").Index

    For Each lr In lo.ListRows

        If Trim(CStr(lr.Range.Cells(1, colEHB).Value)) = ehbTyp _
           And Trim(CStr(lr.Range.Cells(1, colSeiten).Value)) = sheetName Then

            valueJahren = lr.Range.Cells(1, colJahren).Value

            If IsNumeric(valueJahren) Then
                GetYearCount = CLng(valueJahren)
            Else
                GetYearCount = 0
            End If

            Exit Function

        End If

    Next lr

    GetYearCount = 0

End Function


Private Function FindHeaderColumn(ByVal ws As Worksheet, ByVal headerText As String) As Long

    Dim c As Range

    Set c = ws.Cells.Find( _
        What:=headerText, _
        LookIn:=xlValues, _
        LookAt:=xlWhole, _
        MatchCase:=False)

    If Not c Is Nothing Then
        FindHeaderColumn = c.Column
    Else
        FindHeaderColumn = 0
    End If

End Function


Private Function GetCurrentSheetName(ByVal ws As Worksheet, ByVal rowNumber As Long, ByVal colTab As Long) As String

    Dim r As Long

    For r = rowNumber To 1 Step -1
        If Trim(CStr(ws.Cells(r, colTab).Value)) <> "" Then
            GetCurrentSheetName = Trim(CStr(ws.Cells(r, colTab).Value))
            Exit Function
        End If
    Next r

    GetCurrentSheetName = ""

End Function


Private Function IsInArray(ByVal valueToFind As String, ByVal arr As Variant) As Boolean

    Dim i As Long

    For i = LBound(arr) To UBound(arr)
        If StrComp(valueToFind, CStr(arr(i)), vbTextCompare) = 0 Then
            IsInArray = True
            Exit Function
        End If
    Next i

    IsInArray = False

End Function



