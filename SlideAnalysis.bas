' Slide Analysis: List all slides that were used in the Slide Masters

Sub SlideMasterAnalysis()
    Dim oPres As Presentation
    Dim oSlide As Slide
    Dim i As Integer
    
    Set oPres = ActivePresentation
    
    Debug.Print
    Debug.Print "-----START-----"
    ' iterates through all designs used, then checks for each slide if it uses this design 
    With oPres
        For i = .Designs.Count To 1 Step -1
            Debug.Print "List of all slides using: " & .Designs(i).SlideMaster.Design.Name

            For Each oSlide In .Slides
                If oSlide.design.Name = .Designs(i).SlideMaster.Design.Name Then
                    Debug.Print "PPT Slide #: " & oSlide.SlideNumber
                End If
            Next oSlide
            
        Next i
        
    End With
    Debug.Print "-----END-------"
    Debug.Print

    MsgBox "Finished!"
    
End Sub

