' This script deletes all hidden slides
Sub HiddenSlidesCleanup()
    Dim osld As Slide
    Dim i As Integer

    Debug.Print "-----START-----"

    For i = ActivePresentation.Slides.Count To 1 Step -1
        Set osld = ActivePresentation.Slides(i)
        Debug.Print "Slide: " & osld.SlideNumber

        If osld.SlideShowTransition.Hidden = msoTrue Then
            ' If the slide is hidden, delete it
            Debug.Print "Hidden Slide found --> Delete Slide " & osld.SlideNumber
            osld.Delete
        End If
    Next i

    Debug.Print "-----END-------"

    MsgBox "Slide cleanup completed!"
End Sub
