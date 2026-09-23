Option Explicit

' Module: LatinParenthesesToFootnotes
' Purpose: Move Latin technical terms in parentheses to Word footnotes.
' Legacy status: Production rewrite based on a user-tested legacy macro; this rewrite needs Word validation.

Public Sub RVA_MoveLatinParenthesesToFootnotes()
    Dim doc As Document
    Dim searchRange As Range
    Dim openRange As Range
    Dim closeRange As Range
    Dim fullRange As Range
    Dim anchor As Range
    Dim content As String
    Dim startPos As Long
    Dim endPos As Long
    Dim processed As Long
    Dim skipped As Long
    Dim fn As Footnote

    If Application.Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Research VBA"
        Exit Sub
    End If

    Set doc = ActiveDocument
    Set searchRange = doc.Content.Duplicate
    searchRange.Collapse wdCollapseStart

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.StatusBar = "Moving Latin terms to footnotes..."

    Do While searchRange.Start < doc.Content.End
        Set openRange = doc.Range(searchRange.Start, doc.Content.End)
        With openRange.Find
            .ClearFormatting
            .Text = "("
            .Forward = True
            .Wrap = wdFindStop
        End With
        If Not openRange.Find.Execute Then Exit Do

        startPos = openRange.Start
        Set closeRange = doc.Range(openRange.End, doc.Content.End)
        With closeRange.Find
            .ClearFormatting
            .Text = ")"
            .Forward = True
            .Wrap = wdFindStop
        End With
        If Not closeRange.Find.Execute Then Exit Do

        endPos = closeRange.End
        Set fullRange = doc.Range(startPos, endPos)
        content = Mid$(fullRange.Text, 2, Len(fullRange.Text) - 2)

        If RVA_IsLatinTechnicalTerm(content) Then
            fullRange.Delete
            Set anchor = doc.Range(startPos, startPos)
            Set fn = doc.Footnotes.Add(Range:=anchor, Text:=Trim$(content))
            RVA_FormatLatinFootnote fn
            processed = processed + 1
            searchRange.SetRange Start:=anchor.End, End:=anchor.End
        Else
            skipped = skipped + 1
            searchRange.SetRange Start:=endPos, End:=endPos
        End If
    Loop

    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Footnote conversion completed." & vbCrLf & _
           "Converted: " & CStr(processed) & vbCrLf & _
           "Skipped: " & CStr(skipped), vbInformation, "Research VBA"
    Exit Sub

FatalError:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    MsgBox "Footnote conversion stopped." & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "Research VBA"
End Sub

Private Function RVA_IsLatinTechnicalTerm(ByVal value As String) As Boolean
    Dim i As Long
    Dim code As Long
    Dim ch As String
    Dim hasLetter As Boolean
    Dim s As String

    s = Trim$(value)
    If Len(s) = 0 Or Len(s) > 180 Then Exit Function
    If RVA_LooksLikeLatinCitation(s) Then Exit Function

    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        code = AscW(ch)
        If code < 0 Then code = code + 65536

        If (code >= 65 And code <= 90) Or (code >= 97 And code <= 122) Then
            hasLetter = True
        ElseIf code >= 48 And code <= 57 Then
        ElseIf InStr(1, " -/&.,:+_'", ch, vbBinaryCompare) > 0 Then
        Else
            Exit Function
        End If
    Next i

    RVA_IsLatinTechnicalTerm = hasLetter
End Function

Private Function RVA_LooksLikeLatinCitation(ByVal s As String) As Boolean
    Dim t As String
    t = Trim$(s)
    If Len(t) >= 6 Then
        If Right$(t, 6) Like ", ####" Then
            RVA_LooksLikeLatinCitation = True
            Exit Function
        End If
    End If
    If InStr(1, LCase$(t), "doi", vbTextCompare) > 0 Then RVA_LooksLikeLatinCitation = True
End Function

Private Sub RVA_FormatLatinFootnote(ByVal fn As Footnote)
    On Error Resume Next
    With fn.Range
        .Font.NameAscii = "Times New Roman"
        .Font.NameOther = "Times New Roman"
        .ParagraphFormat.ReadingOrder = wdReadingOrderLtr
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With
    On Error GoTo 0
End Sub
