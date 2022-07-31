Option Explicit

Public WithEvents App As Application

Private Sub App_WorkbookOpen(ByVal Wb As Workbook)
    Call UnlockVBAProjects
End Sub
