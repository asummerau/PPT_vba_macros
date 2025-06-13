
' Printalllayouts: this simple macro prints all custom layouts of the specified master design.
Sub Printalllayouts()
    Dim oPres As Presentation
    Dim Design As Design
    Dim i As Integer
    Dim myDesignName As String
    Dim myDesign As Design

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"

    With oPres
        ' TODO: Add here the name of the new design
        myDesignName = "DESIGN NAME"

        ' STEP 1: Find the new design in the presentation
        Set myDesign = Nothing
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            If design.Name = myDesignName Then
                Debug.Print "Found design: " & design.Name
                Set myDesign = design
                Exit For
            End If
        Next i
        
        If myDesign Is Nothing Then
            MsgBox "New design '" & myDesignName & "' not found in the presentation.", vbExclamation
            Exit Sub
        End If

        For Each layout In myDesign.SlideMaster.CustomLayouts
            Debug.Print layout.Name
        Next layout
    End With
    Debug.Print "-----END-----"

End Sub
