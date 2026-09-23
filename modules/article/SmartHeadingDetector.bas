Option Explicit

' Module: SmartHeadingDetector
' Purpose: Detect likely article headings and apply Word built-in Heading 1-3 styles.
' Safety: Conservative heuristics are preferred over aggressive guessing.
' Test status: Static checks only. User validation in Microsoft Word is still required.

Private Const RVA_MAX_HEADING_CHARS As Long = 120

Public Sub RVA_ApplySmartArticleHeadings()
    If Application.Documents.Count = 0 Then
        MsgBox "No document is open.", vbExclamation, "Research VBA"
        Exit Sub
    End If
    Call RVA_ApplySmartArticleHeadingsCore(ActiveDocument, True)
End Sub

Public Function RVA_ApplySmartArticleHeadingsCore(ByVal doc As Document, ByVal showReport As Boolean) As Long
    Dim p As Paragraph
    Dim level As Long
    Dim c1 As Long, c2 As Long, c3 As Long
    Dim total As Long

    On Error GoTo FatalError
    Application.ScreenUpdating = False
    Application.StatusBar = "Detecting article headings..."

    For Each p In doc.Paragraphs
        level = RVA_HeadingLevelForParagraph(doc, p)
        If level >= 1 And level <= 3 Then
            RVA_ApplyBuiltInHeading doc, p, level
            total = total + 1
            Select Case level
                Case 1: c1 = c1 + 1
                Case 2: c2 = c2 + 1
                Case 3: c3 = c3 + 1
            End Select
        End If
    Next p

    RVA_ApplySmartArticleHeadingsCore = total
    Application.StatusBar = False
    Application.ScreenUpdating = True

    If showReport Then
        MsgBox "Heading detection completed." & vbCrLf & vbCrLf & _
               "Heading 1: " & CStr(c1) & vbCrLf & _
               "Heading 2: " & CStr(c2) & vbCrLf & _
               "Heading 3: " & CStr(c3), _
               vbInformation, "Research VBA"
    End If
    Exit Function

FatalError:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    If showReport Then
        MsgBox "Heading detection stopped." & vbCrLf & _
               "Error " & CStr(Err.Number) & ": " & Err.Description, _
               vbCritical, "Research VBA"
    End If
End Function

Private Function RVA_HeadingLevelForParagraph(ByVal doc As Document, ByVal p As Paragraph) As Long
    Dim s As String
    Dim level As Long

    If p Is Nothing Then Exit Function
    If p.Range.StoryType <> wdMainTextStory Then Exit Function
    If p.Range.Information(wdWithInTable) Then Exit Function

    s = RVA_NormalizeText(p.Range.Text)
    If Len(s) = 0 Or Len(s) > RVA_MAX_HEADING_CHARS Then Exit Function

    level = RVA_ExistingHeadingLevel(doc, p)
    If level > 0 Then
        RVA_HeadingLevelForParagraph = level
        Exit Function
    End If

    If Not RVA_LooksLikeSentence(s) Then
        level = RVA_NumberedHeadingLevel(s)
        If level > 0 Then
            RVA_HeadingLevelForParagraph = level
            Exit Function
        End If
    End If

    If RVA_IsKnownLevel1(s) Then
        RVA_HeadingLevelForParagraph = 1
        Exit Function
    End If

    If RVA_IsKnownLevel2(s) Then
        RVA_HeadingLevelForParagraph = 2
        Exit Function
    End If

    level = RVA_HeadingLevelFromFormatting(p, s)
    If level > 0 Then RVA_HeadingLevelForParagraph = level
End Function

Private Function RVA_ExistingHeadingLevel(ByVal doc As Document, ByVal p As Paragraph) As Long
    On Error Resume Next
    If p.Range.Style = doc.Styles(wdStyleHeading1) Then
        RVA_ExistingHeadingLevel = 1
    ElseIf p.Range.Style = doc.Styles(wdStyleHeading2) Then
        RVA_ExistingHeadingLevel = 2
    ElseIf p.Range.Style = doc.Styles(wdStyleHeading3) Then
        RVA_ExistingHeadingLevel = 3
    End If
    On Error GoTo 0
End Function

Private Function RVA_NumberedHeadingLevel(ByVal value As String) As Long
    Dim s As String
    Dim i As Long
    Dim groups As Long
    Dim nextPos As Long
    Dim ch As String
    Dim sawHierarchy As Boolean
    Dim sawSpaces As Boolean
    Dim boundaryFound As Boolean

    s = RVA_ToAsciiDigits(Trim$(value))
    If Len(s) < 3 Then Exit Function

    i = 1
    Do While i <= Len(s) And (Mid$(s, i, 1) = "(" Or Mid$(s, i, 1) = "[")
        i = i + 1
    Loop

    Do
        If i > Len(s) Or Not RVA_IsAsciiDigit(Mid$(s, i, 1)) Then Exit Function

        Do While i <= Len(s) And RVA_IsAsciiDigit(Mid$(s, i, 1))
            i = i + 1
        Loop
        groups = groups + 1

        sawSpaces = False
        Do While i <= Len(s) And Mid$(s, i, 1) = " "
            sawSpaces = True
            i = i + 1
        Loop
        If i > Len(s) Then Exit Function

        ch = Mid$(s, i, 1)

        If ch = "-" Or ch = "." Or ch = "/" Then
            nextPos = i + 1
            Do While nextPos <= Len(s) And Mid$(s, nextPos, 1) = " "
                nextPos = nextPos + 1
            Loop

            If nextPos <= Len(s) And RVA_IsAsciiDigit(Mid$(s, nextPos, 1)) Then
                sawHierarchy = True
                i = nextPos
            Else
                boundaryFound = True
                Exit Do
            End If
        ElseIf ch = ")" Or ch = "]" Or ch = ":" Then
            boundaryFound = True
            Exit Do
        ElseIf sawSpaces And sawHierarchy Then
            boundaryFound = True
            Exit Do
        Else
            Exit Function
        End If

        If groups >= 6 Then Exit Function
    Loop

    If Not boundaryFound Then Exit Function
    If groups < 1 Then Exit Function
    If Not RVA_HasHeadingTextAfterPrefix(s, i) Then Exit Function

    If groups > 3 Then groups = 3
    RVA_NumberedHeadingLevel = groups
End Function

Private Function RVA_HasHeadingTextAfterPrefix(ByVal s As String, ByVal pos As Long) As Boolean
    Dim i As Long
    Dim ch As String
    i = pos
    Do While i <= Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Or ch = "-" Or ch = "." Or ch = "/" Or ch = ")" Or ch = "]" Or ch = ":" Then
            i = i + 1
        Else
            Exit Do
        End If
    Loop
    RVA_HasHeadingTextAfterPrefix = (i <= Len(s))
End Function

Private Function RVA_HeadingLevelFromFormatting(ByVal p As Paragraph, ByVal s As String) As Long
    Dim fontName As String
    Dim sizeBi As Single
    Dim isBold As Boolean

    If Len(s) > 90 Then Exit Function
    If RVA_LooksLikeSentence(s) Then Exit Function

    On Error Resume Next
    fontName = LCase$(CStr(p.Range.Font.NameBi))
    If Len(fontName) = 0 Then fontName = LCase$(CStr(p.Range.Font.Name))
    sizeBi = p.Range.Font.SizeBi
    If sizeBi <= 0 Or sizeBi > 100 Then sizeBi = p.Range.Font.Size
    isBold = (p.Range.Font.Bold <> 0)
    On Error GoTo 0

    If Not isBold Then Exit Function
    If InStr(1, fontName, "titr", vbTextCompare) = 0 Then Exit Function

    If sizeBi >= 11.5 Then
        RVA_HeadingLevelFromFormatting = 1
    ElseIf sizeBi >= 9 Then
        RVA_HeadingLevelFromFormatting = 2
    End If
End Function

Private Function RVA_LooksLikeSentence(ByVal s As String) As Boolean
    Dim lastChar As String
    If Len(s) = 0 Then Exit Function
    lastChar = Right$(s, 1)
    RVA_LooksLikeSentence = (lastChar = "." Or lastChar = "!" Or lastChar = "?" Or lastChar = ChrW$(&H61F))
End Function

Private Function RVA_IsKnownLevel1(ByVal s As String) As Boolean
    Dim terms As Variant
    terms = Array( _
        RVA_U("0645,0642,062F,0645,0647"), _
        RVA_U("0645,0628,0627,0646,06CC,0020,0646,0638,0631,06CC"), _
        RVA_U("0631,0648,0634,0020,062A,062D,0642,06CC,0642"), _
        RVA_U("0631,0648,0634,200C,0634,0646,0627,0633,06CC"), _
        RVA_U("0631,0648,0634,0020,0634,0646,0627,0633,06CC"), _
        RVA_U("062A,062C,0632,06CC,0647,0020,0648,0020,062A,062D,0644,06CC,0644,0020,062F,0627,062F,0647,200C,0647,0627,0020,0648,0020,06CC,0627,0641,062A,0647,200C,0647,0627,06CC,0020,062A,062D,0642,06CC,0642"), _
        RVA_U("06CC,0627,0641,062A,0647,200C,0647,0627,06CC,0020,062A,062D,0642,06CC,0642"), _
        RVA_U("0646,062A,06CC,062C,0647,200C,06AF,06CC,0631,06CC,0020,0648,0020,067E,06CC,0634,0646,0647,0627,062F"), _
        RVA_U("0646,062A,06CC,062C,0647,0020,06AF,06CC,0631,06CC,0020,0648,0020,067E,06CC,0634,0646,0647,0627,062F"), _
        RVA_U("0641,0647,0631,0633,062A,0020,0645,0646,0627,0628,0639"), _
        RVA_U("0645,0646,0627,0628,0639") _
    )
    RVA_IsKnownLevel1 = RVA_MatchesAnyTerm(s, terms)
End Function

Private Function RVA_IsKnownLevel2(ByVal s As String) As Boolean
    Dim terms As Variant
    terms = Array( _
        RVA_U("067E,06CC,0634,06CC,0646,0647,0020,067E,0698,0648,0647,0634"), _
        RVA_U("0633,0627,0628,0642,0647,0020,067E,0698,0648,0647,0634"), _
        RVA_U("0645,0641,0647,0648,0645,200C,0634,0646,0627,0633,06CC"), _
        RVA_U("0686,0627,0631,0686,0648,0628,0020,0646,0638,0631,06CC,0020,067E,0698,0648,0647,0634"), _
        RVA_U("0645,062F,0644,0020,0645,0641,0647,0648,0645,06CC"), _
        RVA_U("062C,0627,0645,0639,0647,0020,0622,0645,0627,0631,06CC"), _
        RVA_U("062D,062C,0645,0020,0646,0645,0648,0646,0647"), _
        RVA_U("0631,0648,0634,0020,0646,0645,0648,0646,0647,200C,06AF,06CC,0631,06CC"), _
        RVA_U("0631,0648,0634,0020,06AF,0631,062F,0622,0648,0631,06CC,0020,062F,0627,062F,0647,200C,0647,0627"), _
        RVA_U("0627,0628,0632,0627,0631,0020,06AF,0631,062F,0622,0648,0631,06CC,0020,062F,0627,062F,0647,200C,0647,0627"), _
        RVA_U("0631,0648,0627,06CC,06CC"), _
        RVA_U("067E,0627,06CC,0627,06CC,06CC"), _
        RVA_U("0646,062A,06CC,062C,0647,200C,06AF,06CC,0631,06CC"), _
        RVA_U("067E,06CC,0634,0646,0647,0627,062F"), _
        RVA_U("067E,06CC,0634,0646,0647,0627,062F,0647,0627"), _
        RVA_U("0645,0646,0627,0628,0639,0020,0641,0627,0631,0633,06CC"), _
        RVA_U("0645,0646,0627,0628,0639,0020,0627,0646,06AF,0644,06CC,0633,06CC"), _
        RVA_U("0633,0627,06CC,062A,200C,0647,0627") _
    )
    RVA_IsKnownLevel2 = RVA_MatchesAnyTerm(RVA_RemoveAlphaSubsectionPrefix(s), terms)
End Function

Private Function RVA_RemoveAlphaSubsectionPrefix(ByVal s As String) As String
    Dim t As String
    Dim closePos As Long
    Dim sepPos As Long

    t = Trim$(s)
    If Len(t) < 3 Then
        RVA_RemoveAlphaSubsectionPrefix = t
        Exit Function
    End If

    If Left$(t, 1) = "(" Then
        closePos = InStr(2, t, ")", vbBinaryCompare)
        If closePos >= 3 And closePos <= 8 Then t = Trim$(Mid$(t, closePos + 1))
    Else
        sepPos = InStr(1, Left$(t, 8), ")", vbBinaryCompare)
        If sepPos = 0 Then sepPos = InStr(1, Left$(t, 8), "-", vbBinaryCompare)
        If sepPos >= 2 Then t = Trim$(Mid$(t, sepPos + 1))
    End If

    RVA_RemoveAlphaSubsectionPrefix = t
End Function

Private Function RVA_MatchesAnyTerm(ByVal s As String, ByVal terms As Variant) As Boolean
    Dim item As Variant
    Dim a As String
    Dim b As String

    a = RVA_NormalizeForComparison(s)
    For Each item In terms
        b = RVA_NormalizeForComparison(CStr(item))
        If a = b Then
            RVA_MatchesAnyTerm = True
            Exit Function
        End If
    Next item
End Function

Private Function RVA_NormalizeForComparison(ByVal value As String) As String
    Dim s As String
    s = RVA_NormalizeText(value)
    s = Replace(s, ChrW$(&H200C), " ")
    s = Replace(s, ChrW$(&H64A), ChrW$(&H6CC))
    s = Replace(s, ChrW$(&H643), ChrW$(&H6A9))
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    Do While Len(s) > 0 And InStr(1, ":;.-", Right$(s, 1), vbBinaryCompare) > 0
        s = Trim$(Left$(s, Len(s) - 1))
    Loop
    RVA_NormalizeForComparison = LCase$(Trim$(s))
End Function

Private Function RVA_NormalizeText(ByVal value As String) As String
    Dim s As String
    s = value
    s = Replace(s, Chr$(13), "")
    s = Replace(s, Chr$(7), "")
    s = Replace(s, ChrW$(&HA0), " ")
    s = Replace(s, vbTab, " ")
    s = Replace(s, ChrW$(8211), "-")
    s = Replace(s, ChrW$(8212), "-")
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    RVA_NormalizeText = s
End Function

Private Function RVA_ToAsciiDigits(ByVal value As String) As String
    Dim s As String
    Dim i As Long
    s = value
    For i = 0 To 9
        s = Replace(s, ChrW$(&H6F0 + i), Chr$(48 + i))
        s = Replace(s, ChrW$(&H660 + i), Chr$(48 + i))
    Next i
    RVA_ToAsciiDigits = s
End Function

Private Function RVA_IsAsciiDigit(ByVal ch As String) As Boolean
    RVA_IsAsciiDigit = (Len(ch) = 1 And ch >= "0" And ch <= "9")
End Function

Private Sub RVA_ApplyBuiltInHeading(ByVal doc As Document, ByVal p As Paragraph, ByVal level As Long)
    Select Case level
        Case 1: p.Range.Style = doc.Styles(wdStyleHeading1)
        Case 2: p.Range.Style = doc.Styles(wdStyleHeading2)
        Case 3: p.Range.Style = doc.Styles(wdStyleHeading3)
    End Select
End Sub

Private Function RVA_U(ByVal hexList As String) As String
    Dim parts As Variant
    Dim i As Long
    Dim result As String
    parts = Split(hexList, ",")
    For i = LBound(parts) To UBound(parts)
        result = result & ChrW$(CLng("&H" & parts(i)))
    Next i
    RVA_U = result
End Function
