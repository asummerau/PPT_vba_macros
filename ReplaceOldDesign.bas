
' ReplaceOldDesign: Replaces the layout of a slide with the layout of the new Slide Master which has to be specified in the code. 
' Iterates through all slides in the powerpoint, but a slide will only be changed if the current layout name matches any layout name in the new Master. If a matching layout is not found, it skips that slide and logs a warning.
' NOTE: this code may mess up your presentation. Only apply on a copy of your presentation!
Sub ReplaceOldDesign()
    Dim oPres As Presentation
    Dim sld As Slide
    Dim Design As Design
    Dim layoutName As String
    Dim i As Integer
    Dim designName As String
    Dim newDesign As Design
    Dim newLayout As CustomLayout
    Dim foundLayout As Boolean
    'Dim currentLayouts As CustomLayouts

    'TODO: Add here the name of the new design
    Const TARGET_MASTER_NAME As String =  "ADD YOUR NEW DESIGN NAME HERE"    

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"
    
    With oPres

        ' STEP 1: Find the new design in the presentation
        Set newDesign = Nothing
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            If design.Name = TARGET_MASTER_NAME Then
                Debug.Print "Found new design: " & design.Name
                Set newDesign = design
                Exit For
            End If
        Next i
        
        If newDesign Is Nothing Then
            MsgBox "New design '" & TARGET_MASTER_NAME & "' not found in the presentation.", vbExclamation
            Exit Sub
        End If

        'For Each layout In newDesign.SlideMaster.CustomLayouts
        '    Debug.Print layout.Name
        'Next layout

        ' STEP 2: Try to replace old designs with the new design if the layout name matches
        For Each sld In oPres.Slides
            layoutName = sld.CustomLayout.Name
            designName = sld.design.Name

            foundLayout = False
            
            If sld.Design.Name = TARGET_MASTER_NAME Then
                Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Design is already '" & TARGET_MASTER_NAME & "', skipping."

            Else
                'Set currentLayouts = sld.Master.design.SlideMaster.CustomLayouts
                'Debug.Print
                'Debug.Print "------"
                'Debug.Print
                'Debug.Print "Found  design: " & sld.Master.design.Name
                'For Each clayout In currentLayouts
                '    Debug.Print clayout.Name
                'Next clayout

                ' Check if the layout name matches any layout in the new design
                For Each newLayout In newDesign.SlideMaster.CustomLayouts
                    If newLayout.Name = layoutName Then
                        Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Changing design from '" & designName & "' to '" & TARGET_MASTER_NAME & "'"
                        sld.Design = newDesign
                        sld.CustomLayout = newLayout
                        foundLayout = True
                        Exit For
                    End If
                Next newLayout

                If Not foundLayout Then
                    Debug.Print "WARNING: Slide " & sld.SlideIndex & " - Could not find matching layout '" & layoutName & "' in new master. Skipping."
                End If

            End If
        Next sld
    End With
    Debug.Print "-----END-----"

    MsgBox "Design replacement complete!", vbInformation
End Sub
