Attribute VB_Name = "Versand"
Option Explicit

Public boInitAutoSv As Boolean

Private Function XlsxPath(ByVal folderPath As String, ByVal fileName As String) As String
    If LCase(Right(fileName, 5)) = ".xlsx" Then
        XlsxPath = folderPath & "\" & fileName
    Else
        XlsxPath = folderPath & "\" & fileName & ".xlsx"
    End If
End Function

Sub createfile()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Dim wbMakro As Workbook
    Set wbMakro = ActiveWorkbook

    Dim nombre_arc As String
    nombre_arc = wbMakro.Worksheets("Bearbeiter").Range("name_datei").Value

    Dim path_data As String
    path_data = XlsxPath(wbMakro.Path, nombre_arc)

    Dim wb As Workbook
    Set wb = Workbooks.Add()

    Dim arrSheets As Variant
    arrSheets = Array( _
        "Übersicht", _
        "Eingabe_Bilanz", _
        "Eingabe_GuV", _
        "Monitor_Bilanz", _
        "Monitor_GuV", _
        "A_Stammdaten", _
        "A1_Fragen", _
        "A2_SaLi", _
        "A3_Hinzu_Kürz", _
        "B_Bilanz", _
        "B1_RSt_Spiegel", _
        "C_GuV", _
        "D1_SAV", _
        "D2_Zuschüsse", _
        "D3_WAV", _
        "E_VL", _
        "Changelog", _
        "Listen" _
    )

    Dim i As Long
    Dim sheetName As String
    Dim wsNew As Worksheet

    For i = LBound(arrSheets) To UBound(arrSheets)

        sheetName = arrSheets(i)

        Set wsNew = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        wsNew.Name = sheetName

       wbMakro.Worksheets(sheetName).Cells.Copy
wsNew.Range("A1").PasteSpecial Paste:=xlPasteFormulas

wbMakro.Worksheets(sheetName).Cells.Copy
wsNew.Range("A1").PasteSpecial Paste:=xlPasteFormats

wbMakro.Worksheets(sheetName).Cells.Copy
wsNew.Range("A1").PasteSpecial Paste:=xlPasteColumnWidths

        wsNew.Activate
        ActiveWindow.DisplayGridlines = False

    Next i

    Application.CutCopyMode = False

    'Standardmäßig von Excel angelegte leere Blätter löschen
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If IsError(Application.Match(ws.Name, arrSheets, 0)) Then
            ws.Delete
        End If
    Next ws

    wb.Worksheets(arrSheets(0)).Activate

    wb.SaveAs fileName:=path_data, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    wb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

Sub sendEmail()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    Call createfile

    Dim ol As Object
    Dim olm As Object

    Dim wd As Object
    Dim doc As Object

    Dim wbMakro As Workbook
    Set wbMakro = ActiveWorkbook

    Set ol = CreateObject("Outlook.Application")

    Call AutoSvOff(wbMakro, boInitAutoSv)

    Dim nombre_arc As String
    nombre_arc = wbMakro.Worksheets("Bearbeiter").Range("name_datei").Value

    Dim Path_word As String
    Path_word = wbMakro.Path & "\20260507_E-Mail-Vorlage für Versand.docx"

    Dim Path_attachment As String
    Dim Path_attachment_2 As String

    Path_attachment = XlsxPath(wbMakro.Path, nombre_arc)
    Path_attachment_2 = Replace(Path_attachment, "%20", " ")

    Set olm = ol.CreateItem(0)

    Set wd = CreateObject("Word.Application")
    wd.Visible = True

    Set doc = wd.Documents.Open(Path_word)

    With wd.Selection.Find
        .Text = "<<Anrede>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_anrede").Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<Name>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_name").Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<arb1>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_arb1").Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<arb1_tel>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_arb1").Offset(1, 0).Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<arb1_mail>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_arb1").Offset(2, 0).Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<arb2>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_arb2").Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<arb2_tel>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_arb2").Offset(1, 0).Value
        .Execute Replace:=2
    End With

    With wd.Selection.Find
        .Text = "<<arb2_mail>>"
        .Replacement.Text = wbMakro.Worksheets("Bearbeiter").Range("e_arb2").Offset(2, 0).Value
        .Execute Replace:=2
    End With

    doc.Content.Copy

    With olm
        .Display
        .SentOnBehalfOfName = "rm@thuega.de"
        .CC = wbMakro.Worksheets("Bearbeiter").Range("e_arb1").Offset(1, 0).Value & ";" & _
              wbMakro.Worksheets("Bearbeiter").Range("e_arb2").Offset(1, 0).Value
        .To = wbMakro.Worksheets("Bearbeiter").Range("e_mail").Value
        .Subject = "Ihr Ergebnis zum Thüga Quick Check zur Kostenmeldung Gas 2025"
        .Attachments.Add Path_attachment_2

        Dim Editor As Object
        Set Editor = .GetInspector.WordEditor
        Editor.Content.Paste
    End With

    Application.CutCopyMode = False

    doc.Application.DisplayAlerts = False
    doc.Close SaveChanges:=False
    Set doc = Nothing

    wd.Quit False
    Set wd = Nothing

    Set olm = Nothing
    Set ol = Nothing

    Call AutoSvOn(wbMakro, boInitAutoSv)

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub

