Option Explicit

' Module: FERSDArticleFormat
' Purpose: Apply explicit page/font requirements supplied for the FERSD article guide.
' Dependency: SmartHeadingDetector.bas
' Test status: Static checks only. User validation in Microsoft Word is still required.

Private Const RVA_FERSD_BODY_FONT As String = "B Lotus"
Private Const RVA_FERSD_HEADING_FONT As String = "B Titr"
Private Const RVA_FERSD_LATIN_FONT As String = "Times New Roman"

Public Sub RVA_FERSD_FormatArticle()
    Dim headingCount As Long

    If Application.Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "FERSD Formatter"
        Exit Sub
    End If

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.StatusBar = "Applying FERSD format..."

    RVA_FERSD_ApplyPageSetup ActiveDocument
    RVA_FERSD_ApplyBuiltInStyles ActiveDocument
    headingCount = RVA_ApplySmartArticleHeadingsCore(ActiveDocument, False)

    Application.StatusBar = False
    Application.ScreenUpdating = True
    RVA_FERSD_Audit headingCount
    Exit Sub

FatalError:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Formatting stopped." & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, _
           vbCritical, "FERSD Formatter"
End Sub

Public Sub RVA_FERSD_ApplyPageSetup(ByVal doc As Document)
    Dim sec As Section
    For Each sec In doc.Sections
        With sec.PageSetup
            .Orientation = wdOrientPortrait
            .PageWidth = CentimetersToPoints(17.5)
            .PageHeight = CentimetersToPoints(24.5)
            .TopMargin = CentimetersToPoints(3)
            .BottomMargin = CentimetersToPoints(2.5)
            .LeftMargin = CentimetersToPoints(2.5)
            .RightMargin = CentimetersToPoints(2.5)
        End With
    Next sec
End Sub

Public Sub RVA_FERSD_ApplyBuiltInStyles(ByVal doc As Document)
    RVA_FERSD_SetStyle doc.Styles(wdStyleNormal), RVA_FERSD_BODY_FONT, 12, RVA_FERSD_LATIN_FONT, 11, False, wdAlignParagraphJustify
    RVA_FERSD_SetStyle doc.Styles(wdStyleHeading1), RVA_FERSD_HEADING_FONT, 12, RVA_FERSD_LATIN_FONT, 11, True, wdAlignParagraphRight
    RVA_FERSD_SetStyle doc.Styles(wdStyleHeading2), RVA_FERSD_HEADING_FONT, 10, RVA_FERSD_LATIN_FONT, 10, True, wdAlignParagraphRight
    RVA_FERSD_SetStyle doc.Styles(wdStyleHeading3), RVA_FERSD_HEADING_FONT, 10, RVA_FERSD_LATIN_FONT, 10, True, wdAlignParagraphRight

    doc.Styles(wdStyleHeading1).ParagraphFormat.OutlineLevel = wdOutlineLevel1
    doc.Styles(wdStyleHeading2).ParagraphFormat.OutlineLevel = wdOutlineLevel2
    doc.Styles(wdStyleHeading3).ParagraphFormat.OutlineLevel = wdOutlineLevel3
End Sub

Private Sub RVA_FERSD_SetStyle(ByVal sty As Style, ByVal rtlFont As String, ByVal rtlSize As Single, _
                               ByVal latinFont As String, ByVal latinSize As Single, _
                               ByVal makeBold As Boolean, ByVal align As WdParagraphAlignment)
    On Error Resume Next
    With sty.Font
        .NameBi = rtlFont
        .SizeBi = rtlSize
        .NameAscii = latinFont
        .NameOther = latinFont
        .Size = latinSize
        .Bold = makeBold
        .Italic = False
    End With
    With sty.ParagraphFormat
        .Alignment = align
        .ReadingOrder = wdReadingOrderRtl
        .FirstLineIndent = 0
        .KeepWithNext = (sty.Type = wdStyleTypeParagraph And makeBold)
    End With
    On Error GoTo 0
End Sub

Public Sub RVA_FERSD_Audit(Optional ByVal headingCount As Long = -1)
    Dim doc As Document
    Dim pages As Long
    Dim report As String
    Dim okCount As Long
    Dim warnCount As Long

    If Application.Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    pages = doc.ComputeStatistics(wdStatisticPages)

    report = "FERSD compliance audit" & vbCrLf & String$(28, "-") & vbCrLf

    RVA_FERSD_AddCheck report, okCount, warnCount, (pages >= 16 And pages <= 22), _
        "Page count is within 16-22", "Page count is " & CStr(pages) & "; expected 16-22"

    RVA_FERSD_AddCheck report, okCount, warnCount, RVA_FERSD_PageSetupMatches(doc), _
        "Page size and margins match the supplied guide", "Page size or margins do not match the supplied guide"

    RVA_FERSD_AddCheck report, okCount, warnCount, RVA_FERSD_StyleFontMatches(doc.Styles(wdStyleNormal), RVA_FERSD_BODY_FONT, 12), _
        "Normal style uses B Lotus 12", "Normal style does not use B Lotus 12"

    RVA_FERSD_AddCheck report, okCount, warnCount, RVA_FERSD_StyleFontMatches(doc.Styles(wdStyleHeading1), RVA_FERSD_HEADING_FONT, 12), _
        "Heading 1 uses B Titr 12", "Heading 1 does not use B Titr 12"

    RVA_FERSD_AddCheck report, okCount, warnCount, RVA_FERSD_StyleFontMatches(doc.Styles(wdStyleHeading2), RVA_FERSD_HEADING_FONT, 10), _
        "Heading 2 uses B Titr 10", "Heading 2 does not use B Titr 10"

    If headingCount >= 0 Then
        report = report & "[INFO] Heading styles applied/detected: " & CStr(headingCount) & vbCrLf
    End If

    report = report & vbCrLf & "Passed: " & CStr(okCount) & vbCrLf & "Warnings: " & CStr(warnCount)
    MsgBox report, IIf(warnCount = 0, vbInformation, vbExclamation), "FERSD Audit"
End Sub

Private Function RVA_FERSD_PageSetupMatches(ByVal doc As Document) As Boolean
    Dim sec As Section
    For Each sec In doc.Sections
        With sec.PageSetup
            If Abs(PointsToCentimeters(.PageWidth) - 17.5) > 0.05 Then Exit Function
            If Abs(PointsToCentimeters(.PageHeight) - 24.5) > 0.05 Then Exit Function
            If Abs(PointsToCentimeters(.TopMargin) - 3) > 0.05 Then Exit Function
            If Abs(PointsToCentimeters(.BottomMargin) - 2.5) > 0.05 Then Exit Function
            If Abs(PointsToCentimeters(.LeftMargin) - 2.5) > 0.05 Then Exit Function
            If Abs(PointsToCentimeters(.RightMargin) - 2.5) > 0.05 Then Exit Function
        End With
    Next sec
    RVA_FERSD_PageSetupMatches = True
End Function

Private Function RVA_FERSD_StyleFontMatches(ByVal sty As Style, ByVal fontName As String, ByVal fontSize As Single) As Boolean
    On Error Resume Next
    RVA_FERSD_StyleFontMatches = (LCase$(sty.Font.NameBi) = LCase$(fontName) And Abs(sty.Font.SizeBi - fontSize) < 0.1)
    On Error GoTo 0
End Function

Private Sub RVA_FERSD_AddCheck(ByRef report As String, ByRef okCount As Long, ByRef warnCount As Long, _
                               ByVal condition As Boolean, ByVal okText As String, ByVal warnText As String)
    If condition Then
        report = report & "[OK] " & okText & vbCrLf
        okCount = okCount + 1
    Else
        report = report & "[WARN] " & warnText & vbCrLf
        warnCount = warnCount + 1
    End If
End Sub
