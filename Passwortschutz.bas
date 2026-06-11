Attribute VB_Name = "Passwortschutz"
Option Explicit

Public Const Pass As String = "ThügaRP5"
Public Const LoeschBlatt As String = "Bearbeiter"
Public Const NameZusatz As String = "_Versand"

Sub PasswortSchutz()

    Dim wb As Workbook
    Dim sheet As Worksheet
    Dim path_data As String
    Dim baseName As String

    Application.DisplayAlerts = False
    Application.ScreenUpdating = False

    Set wb = ActiveWorkbook

    Call AutoSvOff(wb, boInitAutoSv)

    For Each sheet In wb.Worksheets
        sheet.Protect Password:=Pass
    Next sheet

    baseName = Left(wb.Name, InStrRev(wb.Name, ".") - 1)
    path_data = wb.Path & "\" & baseName & NameZusatz & ".xlsx"

    wb.Worksheets(LoeschBlatt).Delete

    wb.SaveAs fileName:=path_data, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False

    Call AutoSvOn(wb, boInitAutoSv)

    MsgBox "Fertig"

    wb.Close SaveChanges:=False

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

End Sub
