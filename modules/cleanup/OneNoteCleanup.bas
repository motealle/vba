Option Explicit

' Module: OneNoteCleanup
' Purpose: Normalize a Word document exported from OneNote without changing page setup.
' Safety: Manual page breaks are preserved by default.
' Test status: Static checks only.

Public Sub RVA_CleanOneNoteExport()
    Const FONT_NAME_RTL As String = "Sahel"
    Const FONT_NAME_LATIN As String = "Times New Roman"
    Const FONT_SIZE As Single = 10.5
    Const MAX_EMPTY_PARAGRAPHS As Long = 1

    Dim doc As Document
    Dim p As Paragraph
    Dim i As Long
    Dim emptyRun As Long
    Dim deleteIndexes As Collection

    If Application.Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument
    Set deleteIndexes = New Collection

    On Error GoTo FatalError
    Application.ScreenUpdating = False

    On Error Resume Next
    With doc.Content.Font
        .NameBi = FONT_NAME_RTL
        .SizeBi = FONT_SIZE
        .NameAscii = FONT_NAME_LATIN
        .NameOther = FONT_NAME_LATIN
        .Size = FONT_SIZE
    End With
    On Error GoTo FatalError

    For Each p In doc.Paragraphs
        With p.Format
            .SpaceBefore = 0
            .SpaceAfter = 0
            .LeftIndent = 0
            .RightIndent = 0
            .FirstLineIndent = 0
        End With
    Next p

    For i = 1 To doc.Paragraphs.Count
        If Len(Trim$(Replace(doc.Paragraphs(i).Range.Text, Chr$(13), ""))) = 0 Then
            emptyRun = emptyRun + 1
            If emptyRun > MAX_EMPTY_PARAGRAPHS Then deleteIndexes.Add i
        Else
            emptyRun = 0
        End If
    Next i

    For i = deleteIndexes.Count To 1 Step -1
        doc.Paragraphs(CLng(deleteIndexes(i))).Range.Delete
    Next i

    Application.ScreenUpdating = True
    MsgBox "OneNote cleanup completed.", vbInformation, "Research VBA"
    Exit Sub

FatalError:
    Application.ScreenUpdating = True
    MsgBox "OneNote cleanup stopped." & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "Research VBA"
End Sub
