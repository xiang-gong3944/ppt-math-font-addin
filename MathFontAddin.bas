Attribute VB_Name = "MathFontAddin"
Option Explicit

Private Const APP_NAME As String = "PptMathFontAddin"
Private Const SETTINGS_SECTION As String = "Settings"
Private Const DEFAULT_MATH_FONT As String = "Latin Modern Math"

Private gAppEvents As CAppEvents
Private gRibbon As Office.IRibbonUI
Private gAutoEnabled As Boolean
Private gBusy As Boolean
Private gTargetFont As String

Public Sub Auto_Open()
    On Error Resume Next
    gTargetFont = GetSetting(APP_NAME, SETTINGS_SECTION, "TargetFont", DEFAULT_MATH_FONT)
    gAutoEnabled = (GetSetting(APP_NAME, SETTINGS_SECTION, "AutoEnabled", "0") = "1")
    Set gAppEvents = New CAppEvents
    Set gAppEvents.App = Application
End Sub

Public Sub Auto_Close()
    On Error Resume Next
    Set gAppEvents.App = Nothing
    Set gAppEvents = Nothing
    Set gRibbon = Nothing
End Sub

Public Sub MathFont_RibbonOnLoad(ByVal ribbon As Office.IRibbonUI)
    Set gRibbon = ribbon
End Sub

Public Sub MathFont_RibbonSelection(ByVal control As Office.IRibbonControl)
    MathFont_ApplySelection
End Sub

Public Sub MathFont_RibbonSlide(ByVal control As Office.IRibbonControl)
    MathFont_ApplySlide
End Sub

Public Sub MathFont_RibbonPresentation(ByVal control As Office.IRibbonControl)
    MathFont_ApplyPresentation
End Sub

Public Sub MathFont_RibbonRestoreCambria(ByVal control As Office.IRibbonControl)
    MathFont_RestoreCambria
End Sub

Public Sub MathFont_RibbonGetFont(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    EnsureSettingsLoaded
    returnedVal = gTargetFont
End Sub

Public Sub MathFont_RibbonFontChanged(ByVal control As Office.IRibbonControl, ByVal text As String)
    On Error GoTo Failed
    MathFont_SetTargetFont text
    Exit Sub
Failed:
    MsgBox "Could not save the target font." & vbCrLf & "PowerPoint returned: " & Err.Description, vbExclamation, "Math Font"
End Sub

Public Sub MathFont_RibbonToggleAuto(ByVal control As Office.IRibbonControl, ByVal pressed As Boolean)
    SetAutoEnabled pressed
End Sub

Public Sub MathFont_RibbonGetAuto(ByVal control As Office.IRibbonControl, ByRef returnedVal)
    returnedVal = gAutoEnabled
End Sub

Public Sub MathFont_ApplySelection()
    EnsureSettingsLoaded
    On Error GoTo Failed
    If ActiveWindow Is Nothing Then Exit Sub

    Select Case ActiveWindow.Selection.Type
        Case ppSelectionText
            Call FormatTextRangeMathZones(ActiveWindow.Selection.TextRange2, gTargetFont)
        Case ppSelectionShapes
            Call ApplyFontToShapeRange(ActiveWindow.Selection.ShapeRange, gTargetFont)
        Case ppSelectionSlides
            Call ApplyFontToSlideRange(ActiveWindow.Selection.SlideRange, gTargetFont)
        Case Else
            Exit Sub
    End Select

    Exit Sub
Failed:
    MsgBox "Could not format the current selection." & vbCrLf & "PowerPoint returned: " & Err.Description, vbExclamation, "Math Font"
End Sub

Public Sub MathFont_ApplySlide()
    Dim sld As PowerPoint.Slide
    EnsureSettingsLoaded
    Set sld = CurrentSlide()
    If sld Is Nothing Then Exit Sub
    Call ApplyFontToSlide(sld, gTargetFont)
End Sub

Public Sub MathFont_ApplyPresentation()
    EnsureSettingsLoaded
    MathFont_ApplyPresentationWithFont gTargetFont, True
End Sub

Public Sub MathFont_RestoreCambria()
    If Application.Presentations.Count = 0 Then Exit Sub
    RestoreCambriaInPresentation ActivePresentation, True
End Sub

Private Sub RestoreCambriaInPresentation(ByVal pres As PowerPoint.Presentation, ByVal showErrors As Boolean)
    Dim sld As PowerPoint.Slide
    On Error GoTo Failed
    gBusy = True
    SetAutoEnabled False
    For Each sld In pres.Slides
        Call ApplyFontToSlide(sld, "Cambria Math")
    Next sld
    gBusy = False
    Exit Sub
Failed:
    gBusy = False
    If showErrors Then MsgBox "Could not restore Cambria Math." & vbCrLf & "PowerPoint returned: " & Err.Description, vbExclamation, "Math Font"
End Sub

Public Sub MathFont_SetTargetFont(ByVal fontName As String)
    fontName = Trim$(fontName)
    If Len(fontName) > 0 Then
        SaveSetting APP_NAME, SETTINGS_SECTION, "TargetFont", fontName
        gTargetFont = fontName
    End If
    ' Blank input restores the displayed value without clearing the setting.
    If Not gRibbon Is Nothing Then gRibbon.InvalidateControl "cmbMathFont"
End Sub

Public Sub MathFont_HandleSelectionChange(ByVal Sel As PowerPoint.Selection)
    Dim sld As PowerPoint.Slide
    If Not gAutoEnabled Then Exit Sub
    If gBusy Then Exit Sub
    On Error GoTo SafeExit
    gBusy = True
    EnsureSettingsLoaded
    Set sld = CurrentSlide()
    If Not sld Is Nothing Then Call ApplyFontToSlide(sld, gTargetFont)
SafeExit:
    gBusy = False
End Sub

Private Sub EnsureSettingsLoaded()
    If Len(gTargetFont) = 0 Then
        gTargetFont = GetSetting(APP_NAME, SETTINGS_SECTION, "TargetFont", DEFAULT_MATH_FONT)
    End If
End Sub

Private Sub SetAutoEnabled(ByVal enabled As Boolean)
    gAutoEnabled = enabled
    SaveSetting APP_NAME, SETTINGS_SECTION, "AutoEnabled", IIf(gAutoEnabled, "1", "0")
    If Not gRibbon Is Nothing Then gRibbon.InvalidateControl "tglMathFontAuto"
End Sub

Private Function CurrentSlide() As PowerPoint.Slide
    On Error Resume Next
    Set CurrentSlide = ActiveWindow.View.Slide
    On Error GoTo 0
End Function

Private Sub MathFont_ApplyPresentationWithFont(ByVal fontName As String, ByVal showErrors As Boolean)
    Dim pres As PowerPoint.Presentation
    Dim sld As PowerPoint.Slide
    On Error GoTo Failed
    If Application.Presentations.Count = 0 Then Exit Sub
    Set pres = ActivePresentation
    gBusy = True
    For Each sld In pres.Slides
        Call ApplyFontToSlide(sld, fontName)
    Next sld
    gBusy = False
    Exit Sub
Failed:
    gBusy = False
    If showErrors Then MsgBox "Could not format the presentation." & vbCrLf & "PowerPoint returned: " & Err.Description, vbExclamation, "Math Font"
End Sub

Private Function ApplyFontToSlideRange(ByVal sr As PowerPoint.SlideRange, ByVal fontName As String) As Long
    Dim i As Long
    For i = 1 To sr.Count
        ApplyFontToSlideRange = ApplyFontToSlideRange + ApplyFontToSlide(sr(i), fontName)
    Next i
End Function

Private Function ApplyFontToShapeRange(ByVal sr As PowerPoint.ShapeRange, ByVal fontName As String) As Long
    Dim i As Long
    For i = 1 To sr.Count
        ApplyFontToShapeRange = ApplyFontToShapeRange + ApplyFontToShape(sr(i), fontName)
    Next i
End Function

Private Function ApplyFontToSlide(ByVal sld As PowerPoint.Slide, ByVal fontName As String) As Long
    Dim shp As PowerPoint.Shape
    For Each shp In sld.Shapes
        ApplyFontToSlide = ApplyFontToSlide + ApplyFontToShape(shp, fontName)
    Next shp
End Function

Private Function ApplyFontToShape(ByVal shp As PowerPoint.Shape, ByVal fontName As String) As Long
    Dim i As Long, r As Long, c As Long
    On Error GoTo Done
    If shp.Type = msoGroup Then
        For i = 1 To shp.GroupItems.Count
            ApplyFontToShape = ApplyFontToShape + ApplyFontToShape(shp.GroupItems(i), fontName)
        Next i
        GoTo Done
    End If
    If shp.HasTable = msoTrue Then
        For r = 1 To shp.Table.Rows.Count
            For c = 1 To shp.Table.Columns.Count
                ApplyFontToShape = ApplyFontToShape + ApplyFontToShape(shp.Table.Cell(r, c).Shape, fontName)
            Next c
        Next r
        GoTo Done
    End If
    If shp.HasTextFrame = msoTrue Then
        If shp.TextFrame2.HasText = msoTrue Then
            ApplyFontToShape = ApplyFontToShape + FormatTextRangeMathZones(shp.TextFrame2.TextRange, fontName)
        End If
    End If
Done:
End Function

Private Function FormatTextRangeMathZones(ByVal tr As Office.TextRange2, ByVal fontName As String) As Long
    Dim zoneIndex As Long
    Dim zone As Office.TextRange2
    Dim prevStart As Long, prevLength As Long
    Dim havePrev As Boolean
    Dim guard As Long
    Dim hadError As Boolean

    If tr Is Nothing Then Exit Function
    If tr.Length <= 0 Then Exit Function

    zoneIndex = 1
    guard = 0
    Do
        guard = guard + 1
        If guard > tr.Length + 2 Then Exit Do
        Set zone = Nothing
        hadError = False

        On Error Resume Next
        Err.Clear
        Set zone = tr.MathZones(zoneIndex, 1)
        If Err.Number <> 0 Then hadError = True
        Err.Clear
        On Error GoTo 0

        If hadError Then Exit Do
        If zone Is Nothing Then Exit Do
        If zone.Length <= 0 Then Exit Do
        If havePrev Then
            If zone.Start = prevStart And zone.Length = prevLength Then Exit Do
        End If

        prevStart = zone.Start
        prevLength = zone.Length
        havePrev = True

        On Error Resume Next
        ' Keep the upright/italic style set by PowerPoint's normal font controls.
        zone.Font.Name = fontName
        If Err.Number = 0 Then FormatTextRangeMathZones = FormatTextRangeMathZones + 1
        Err.Clear
        On Error GoTo 0

        zoneIndex = zoneIndex + 1
    Loop
End Function
