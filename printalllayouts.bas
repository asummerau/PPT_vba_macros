
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
        
        ' === STEP 1: Set your desired master name ===
        myDesignName = "Cisco Light 05-12-2025"
        ' === STEP 2: Try to find that design ===
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
        MsgBox "Master design '" & myDesignName & "' not found.", vbExclamation
            Exit Sub
        End If

        ' === STEP 3: Collect layout names ===
        For Each layout In myDesign.SlideMaster.CustomLayouts
            Debug.Print layout.Name
        Next layout
    End With
    Debug.Print "-----END-----"

End Sub
