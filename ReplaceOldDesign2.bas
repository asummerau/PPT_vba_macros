
' ReplaceOldDesign2: Replaces the layout of a ppt slide to a new layout from the new Master Design.
' based on a predefined set of mappings.
' NOTE: this code may mess up your presentation.
Sub ReplaceOldDesign2()
    Dim oPres As Presentation
    Dim sld As Slide
    Dim layoutName As String
    Dim i As Integer, j As Integer
    Dim designName As String
    Dim newDesignName As String
    Dim oldDesignName As String
    Dim newDesign As Design
    Dim newLayouts As CustomLayouts
    Dim newLayout As CustomLayout
    Dim foundLayout As Boolean
    Dim currentLayouts As CustomLayouts

    Dim layoutMapping(0 To 5, 0 To 1) As String

    ' TODO: fill in the Slide Master Names
    layoutMapping(0, 0) = "OLD MASTER NAME"
    layoutMapping(0, 1) = "NEW MASTER NAME"

    ' TODO: fill in the Layout Names you want to replace 
    ' (i, 0) are layouts from the old Master, (i, 1) form the new
    layoutMapping(1, 0) = ""
    layoutMapping(1, 1) = ""

    layoutMapping(2, 0) = ""
    layoutMapping(2, 1) = ""

    layoutMapping(3, 0) = ""
    layoutMapping(3, 1) = ""

    layoutMapping(4, 0) = ""
    layoutMapping(4, 1) = ""

    layoutMapping(5, 0) = ""
    layoutMapping(5, 1) = ""

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"

    With oPres
        Debug.Print "New Design: "; layoutMapping(0, 1)
        Debug.Print "Old Design: "; layoutMapping(0, 0)
        Debug.Print
        newDesignName = layoutMapping(0, 1)
        oldDesignName = layoutMapping(0, 0)

        ' STEP 1: Find the new design in the presentation
        Set newDesign = Nothing
        Set newLayouts = Nothing
        For i = .Designs.Count To 1 Step -1
            If .Designs(i).Name = newDesignName Then
                ' Debug.Print "Found new design: " & .Designs(i).Name
                Set newDesign = .Designs(i)
                Set newLayouts = .Designs(i).SlideMaster.CustomLayouts
                Exit For
            End If
        Next i
        
        If newDesign Is Nothing Then
            MsgBox "New design '" & newDesignName & "' not found in the presentation.", vbExclamation
            Exit Sub
        End If

        ' STEP 2: Try to replace old designs with the new design based on the predefined mapping array
        For Each sld In oPres.Slides
            layoutName = sld.CustomLayout.Name
            designName = sld.Design.Name

            foundLayout = False
            
            If designName = oldDesignName Then
                Debug.Print "Find repalcement for: " & layoutName
                ' Check if a mapping exists in the predefined array
                For j = 1 To 5
                    If foundLayout Then Exit For
                    
                     If Trim(layoutName) = Trim(layoutMapping(j, 0)) Then
                    ' if a mapping was found, find right layout from the new design
                        For Each newLayout In newLayouts 
                            If Trim(newLayout.Name) = Trim(layoutMapping(j, 1)) Then
                                
                                Debug.Print "Slide " & sld.SlideIndex & ": Layout '" & layoutName & "' replaced with '" & newLayout.Name & "'"
                                sld.Design = newLayout.Design
                                sld.CustomLayout = newLayout
                                foundLayout = True
                                Exit For
                            End If
                        Next newLayout

                    End If
                Next j
                
                If Not foundLayout Then
                    Debug.Print "WARNING: Slide " & sld.SlideIndex & ": No matching layout found for '" & layoutName & "'. Skipped."
                End If
            
            ElseIf designName = newDesignName Then
                Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Design is already '" & newDesignName & "', skipping."

            Else
                Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Another design was found ('" & designName & "'). Skipping."
            End If

        Next sld
    End With
    Debug.Print "-----END-----"

    MsgBox "Design replacement complete!", vbInformation
End Sub

