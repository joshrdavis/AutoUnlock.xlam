Option Explicit

Public WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    On Error GoTo 0
    Call UnlockVBAProjects
End Sub
