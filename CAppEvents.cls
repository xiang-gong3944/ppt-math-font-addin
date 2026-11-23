VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CAppEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public WithEvents App As PowerPoint.Application
Attribute App.VB_VarHelpID = -1

Private Sub App_WindowSelectionChange(ByVal Sel As PowerPoint.Selection)
    On Error Resume Next
    MathFont_HandleSelectionChange Sel
End Sub

Private Sub App_WindowDeactivate(ByVal Pres As PowerPoint.Presentation, ByVal Wn As PowerPoint.DocumentWindow)
    MathFont_FlushAutoFormat Pres, False
End Sub

Private Sub App_PresentationBeforeSave(ByVal Pres As PowerPoint.Presentation, Cancel As Boolean)
    If Not Cancel Then MathFont_FlushAutoFormat Pres, True
End Sub
