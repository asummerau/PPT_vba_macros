' This script removes all comments from the notes section

Sub NotesCleanup()
    Dim osld As Slide
    Dim oShp As Shape
    Dim shapeCount As Integer
    Dim shapeIndex As Integer
    Dim textContent As String
    Dim firstTenChars As String

    ' Iterate through each slide in the active presentation
    For Each osld In ActivePresentation.Slides
        Debug.Print "--- Slide: " & osld.SlideNumber & "---"
        shapeCount = osld.NotesPage.Shapes.Count
        Debug.Print "Slide " & osld.SlideNumber & " has " & shapeCount & " shapes in the notes page"
        shapeIndex = 0

        For Each oShp In osld.NotesPage.Shapes
            shapeIndex = shapeIndex + 1

            If oShp.HasTextFrame Then
                
                ' Check if there is a TextFrame in the Shape
                If oShp.TextFrame.HasText Then
    
                        textContent = oShp.TextFrame.TextRange.Text
                        If Len(textContent) >= 10 Then
                            firstTenChars = Left(textContent, 10) & "..."
                        Else
                            firstTenChars = textContent & "..."
                        End If

                        Debug.Print "--> **Delete Note """ & firstTenChars & """ in shape number " & shapeIndex & " **"
                        oShp.TextFrame.TextRange.Text = ""
                End If
            End If
        Next oShp
    Next osld
    
    MsgBox "Notes cleanup completed!"
End Sub

