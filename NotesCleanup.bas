' This script removes all comments from the notes section

Sub NotesCleanup()
    Dim osld As Slide
    For Each osld In ActivePresentation.Slides
        Debug.Print "Slide: " & osld.SlideNumber
        With osld.NotesPage.Shapes(2)
            If .HasTextFrame Then
                Debug.Print "--> Delete Notes"
                .TextFrame.DeleteText
            End If
        End With
    Next osld
    MsgBox "Notes cleaup completed!"
End Sub