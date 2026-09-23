Option Explicit

' Module: ThesisFontPreset
' Purpose: Apply font sizes visible in the supplied thesis typography reference image.
' Scope: Built-in Normal, Title, Heading 1-3, TOC 1-3, Caption, and Footnote Text styles.
' Test status: Static checks only.

Public Sub RVA_ApplyThesisFontPreset()
    Dim doc As Document
    If Application.Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument

    RVA_ThesisSetStyle doc.Styles(wdStyleNormal), "B Lotus", 14, "Times New Roman", 14, False
    RVA_ThesisSetStyle doc.Styles(wdStyleTitle), "B Titr", 60, "Times New Roman", 14, True
    RVA_ThesisSetStyle doc.Styles(wdStyleHeading1), "B Titr", 14, "Times New Roman", 14, True
    RVA_ThesisSetStyle doc.Styles(wdStyleHeading2), "B Titr", 13, "Times New Roman", 14, True
    RVA_ThesisSetStyle doc.Styles(wdStyleHeading3), "B Titr", 13, "Times New Roman", 14, True
    RVA_ThesisSetStyle doc.Styles(wdStyleTOC1), "B Lotus", 12, "Times New Roman", 12, False
    RVA_ThesisSetStyle doc.Styles(wdStyleTOC2), "B Lotus", 12, "Times New Roman", 12, False
    RVA_ThesisSetStyle doc.Styles(wdStyleTOC3), "B Lotus", 12, "Times New Roman", 12, False
    RVA_ThesisSetStyle doc.Styles(wdStyleCaption), "B Titr", 10, "Times New Roman", 10, False
    RVA_ThesisSetStyle doc.Styles(wdStyleFootnoteText), "B Lotus", 12, "Times New Roman", 12, False

    MsgBox "Thesis font preset applied to built-in Word styles.", vbInformation, "Research VBA"
End Sub

Private Sub RVA_ThesisSetStyle(ByVal sty As Style, ByVal rtlFont As String, ByVal rtlSize As Single, _
                               ByVal latinFont As String, ByVal latinSize As Single, ByVal makeBold As Boolean)
    On Error Resume Next
    With sty.Font
        .NameBi = rtlFont
        .SizeBi = rtlSize
        .NameAscii = latinFont
        .NameOther = latinFont
        .Size = latinSize
        .Bold = makeBold
    End With
    On Error GoTo 0
End Sub
