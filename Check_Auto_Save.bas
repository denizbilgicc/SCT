Attribute VB_Name = "Check_Auto_Save"
Option Explicit

Public boInitAutoSv As Boolean
Public AutoSv As Boolean

Public Sub AutoSvOff(wb As Workbook, ByRef initAutoSv As Boolean)

    If Val(Application.Version) > 15 Then
        On Error Resume Next
        
        AutoSv = wb.AutoSaveOn
        
        If AutoSv Then
            initAutoSv = True
            wb.AutoSaveOn = False
        Else
            initAutoSv = False
        End If
        
        AutoSv = wb.AutoSaveOn
        
        On Error GoTo 0
    End If

End Sub

Public Sub AutoSvOn(wb As Workbook, ByRef initAutoSv As Boolean)

    If Val(Application.Version) > 15 Then
        On Error Resume Next
        
        If initAutoSv Then
            wb.AutoSaveOn = True
        End If
        
        AutoSv = wb.AutoSaveOn
        
        On Error GoTo 0
    End If

End Sub
