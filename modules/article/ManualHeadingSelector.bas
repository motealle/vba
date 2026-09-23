Option Explicit

' Module: ManualHeadingSelector
' Purpose: Safe fallback for ambiguous short bold headings.
' Legacy status: Based on a user-tested legacy macro; test date is unknown.

Private Const RVA_MANUAL_MAX_CHARS As Long = 120
Private Const RVA_MANUAL_MAX_LIST As Long = 40

Public Sub RVA_ManuallyClassifyHeadingCandidates()
    Dim doc As Document
    Dim p As Paragraph
    Dim candidates As Collection
    Dim msg As String
    Dim i As Long
    Dim h1 As String, h2 As String, h3 As String
    Dim s As String

    If Application.Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set candidates = New Collection

    For Each p In doc.Paragraphs
        s = Trim$(Replace(Replace(p.Range.Text, Chr$(13), ""), Chr$(7), ""))
        If Len(s) > 0 And Len(s) <= RVA_MANUAL_MAX_CHARS Then
            If Not p.Range.Information(wdWithInTable) Then
                If p.Range.Font.Bold <> 0 Then candidates.Add p
            End If
        End If
    Next p

    If candidates.Count = 0 Then
        MsgBox "No short bold heading candidates found.", vbInformation, "Research VBA"
        Exit Sub
    End If

    msg = "Candidates:" & vbCrLf & vbCrLf
    For i = 1 To candidates.Count
        If i > RVA_MANUAL_MAX_LIST Then
            msg = msg & vbCrLf & "...more candidates exist. Use the Navigation Pane for a second pass."
            Exit For
        End If
        msg = msg & CStr(i) & ". " & Left$(Trim$(Replace(candidates(i).Range.Text, Chr$(13), "")), 90) & vbCrLf
    Next i
    MsgBox msg, vbInformation, "Research VBA"

    h1 = InputBox("Heading 1 candidate numbers, comma-separated:", "Research VBA")
    h2 = InputBox("Heading 2 candidate numbers, comma-separated:", "Research VBA")
    h3 = InputBox("Heading 3 candidate numbers, comma-separated:", "Research VBA")

    RVA_ApplyCandidateList doc, candidates, h1, 1
    RVA_ApplyCandidateList doc, candidates, h2, 2
    RVA_ApplyCandidateList doc, candidates, h3, 3
End Sub

Private Sub RVA_ApplyCandidateList(ByVal doc As Document, ByVal candidates As Collection, _
                                   ByVal numberList As String, ByVal level As Long)
    Dim parts As Variant
    Dim item As Variant
    Dim idx As Long

    If Len(Trim$(numberList)) = 0 Then Exit Sub
    parts = Split(numberList, ",")

    For Each item In parts
        idx = CLng(Val(Trim$(CStr(item))))
        If idx >= 1 And idx <= candidates.Count Then
            Select Case level
                Case 1: candidates(idx).Range.Style = doc.Styles(wdStyleHeading1)
                Case 2: candidates(idx).Range.Style = doc.Styles(wdStyleHeading2)
                Case 3: candidates(idx).Range.Style = doc.Styles(wdStyleHeading3)
            End Select
        End If
    Next item
End Sub
