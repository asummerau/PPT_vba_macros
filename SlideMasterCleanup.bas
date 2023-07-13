' This script deletes all unused Masters with all its layout slides

Sub SlideMasterCleanup()
    Dim oPres As Presentation
    Dim oLayout As CustomLayout
    Dim oSlide As Slide
    Dim i As Integer
    Dim j As Integer
    Dim isUsed As Boolean
    
    Set oPres = ActivePresentation
    
    On Error Resume Next
    With oPres
        For i = .Designs.Count To 1 Step -1
            Debug.Print "Check if PPT contains: " & .Designs(i).SlideMaster.Design.Name
            ' Check if any slide is using this slide master
            isUsed = False
            For Each oSlide In .Slides
                If oSlide.Design.Name = .Designs(i).SlideMaster.Design.Name Then
                    Debug.Print "PPT Slide is using this Master: " & oSlide.Design.Name
                    isUsed = True
                    Exit For
                End If
            Next oSlide
            
            ' If no slide is using this slide master, remove its layout slides
            If Not isUsed Then
                Debug.Print "PPT Slide is not using this Master, it will be deleted."
                For j = .Designs(i).SlideMaster.CustomLayouts.Count To 1 Step -1
                    .Designs(i).SlideMaster.CustomLayouts(j).Delete
                Next j
                
                ' Finally delete the slide master itself as well
                .Designs(i).SlideMaster.Delete

            End If
        Debug.Print "---"
        Next i
        
    End With
    Debug.Print "Finished"
    MsgBox "Cleaup completed!"
End Sub
