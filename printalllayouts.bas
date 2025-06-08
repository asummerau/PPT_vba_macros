
Sub printalllayouts()
    Dim oPres As Presentation
    Dim Design As Design
    Dim i As Integer
    Dim newDesignName As String
    Dim newDesign As Design

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"

    With oPres
        ' TODO: Add here the name of the new design
        newDesignName = "DESIGN NAME"

        ' STEP 1: Find the new design in the presentation
        Set newDesign = Nothing
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            If design.Name = newDesignName Then
                Debug.Print "Found new design: " & design.Name
                Set newDesign = design
                Exit For
            End If
        Next i
        
        If newDesign Is Nothing Then
            MsgBox "New design '" & newDesignName & "' not found in the presentation.", vbExclamation
            Exit Sub
        End If

        For Each layout In newDesign.SlideMaster.CustomLayouts
            Debug.Print layout.Name
        Next layout
    End With
    Debug.Print "-----END-----"

End Sub
