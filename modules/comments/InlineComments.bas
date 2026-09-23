Option Explicit

' Module: InlineComments
' Purpose: Convert Word comments into visible inline markers before export or AI review.
' Warning: This is destructive because comments are deleted after conversion.
' Test status: Static checks only.

Public Sub RVA_ConvertCommentsToInlineMarkers()
    Dim doc As Document
    Dim selectedText As String
    Dim commentText As String
    Dim markedText As String
    Dim counter As Long

    If Application.Documents.Count = 0 Then Exit Sub
    Set doc = ActiveDocument

    If doc.Comments.Count = 0 Then
        MsgBox "No comments found.", vbInformation, "Research VBA"
        Exit Sub
    End If

    If MsgBox("This will replace commented text with inline markers and delete comments. Continue?", _
              vbYesNo + vbExclamation, "Research VBA") <> vbYes Then Exit Sub

    On Error GoTo FatalError
    Application.ScreenUpdating = False

    Do While doc.Comments.Count > 0
        selectedText = doc.Comments(1).Scope.Text
        commentText = doc.Comments(1).Range.Text
        markedText = "{{" & selectedText & "}} {{{" & commentText & "}}}"
        doc.Comments(1).Scope.Text = markedText
        doc.Comments(1).Delete
        counter = counter + 1
    Loop

    Application.ScreenUpdating = True
    MsgBox "Converted comments: " & CStr(counter), vbInformation, "Research VBA"
    Exit Sub

FatalError:
    Application.ScreenUpdating = True
    MsgBox "Comment conversion stopped." & vbCrLf & _
           "Error " & CStr(Err.Number) & ": " & Err.Description, vbCritical, "Research VBA"
End Sub
