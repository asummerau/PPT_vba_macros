' Slide Analysis: List all slides that were used in the Slide Masters

Sub SlideMasterCleanup()
    Dim oPres As Presentation
    Dim oSlide As Slide
    Dim i As Integer
    
    Set oPres = ActivePresentation
    
    On Error Resume Next
    With oPres
        For i = .Designs.Count To 1 Step -1
            Debug.Print "List of all slides using: " & .Designs(i).slideMaster.Design.Name

            For Each oSlide In .Slides
                If oSlide.Design.Name = .Designs(i).slideMaster.Design.Name Then
                    Debug.Print "PPT Slide #: " & oSlide.SlideNumber
                End If
            Next oSlide
            
        Debug.Print "---"
        Next i
        
    End With
    MsgBox "Finished!"
    
End Sub

