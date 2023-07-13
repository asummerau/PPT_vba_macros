' This script removes all comments from the notes section

Sub Zap()
    Dim osld As Slide
    For Each osld In ActivePresentation.Slides
        With osld.NotesPage.Shapes(2)
            If .HasTextFrame Then
                .TextFrame.DeleteText
            End If
        End With
    Next osld
    MsgBox "Notes cleaup completed!"
End Sub