
' ReplaceOldDesign2: Replaces the layout of a slide (from a specified old Slide Master) with the layout of the new Slide Master which has to be specified in the code. 
' This macro can be used if the layout names don't match. To make a mapping, a manual mapping has to be done and specified in the code.
' NOTE: this code may mess up your presentation. Only apply on a copy of your presentation!
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
    Dim foundNewLayout As Boolean
    Dim foundOldLayout As Boolean
    Dim currentLayouts As CustomLayouts
    Dim nItems As Integer
    Dim layoutMapping() As String
    Dim underscorePos As Integer

    ' TODO: Specify number of mappings you want to use
    nItems = 87
    ReDim layoutMapping(0 To nItems, 0 To 1)

    ' TODO: fill in the Slide Master Names and update the number of mappings
    layoutMapping(0, 0) = "OLD MASTER NAME"
    layoutMapping(0, 1) = "NEW MASTER NAME"

    ' TODO: fill in the Layout Names you want to replace. E.g.:
    layoutMapping(1, 0) = "Title Slide" 'layout name from old master
    layoutMapping(1, 1) = "New Title Slide" 'layout name from new master
	
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
            layoutname = Trim(.Designs(i).Name)
            underscorePos = InStr(layoutName, "_")
            If underscorePos > 1 And IsNumeric(Left(i, underscorePos - 1)) Then
                layoutName = Mid(layoutName, underscorePos + 1)
            End If

            If layoutname = newDesignName Then
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
            layoutName = Trim(sld.CustomLayout.Name)
            designName = Trim(sld.Design.Name)

            foundNewLayout = False
            foundOldLayout = False

            underscorePos = InStr(designName, "_")
            If underscorePos > 1 And IsNumeric(Left(i, underscorePos - 1)) Then
                designName = Mid(designName, underscorePos + 1)
            End If

            If designName = oldDesignName Then
                Debug.Print "--> Find repalcement for: " & layoutName
                ' Check if a mapping exists in the predefined array

                ' there are tons of duplicate layouts that start with a prefix (e.g. 1_title is the same as title)
                underscorePos = InStr(layoutName, "_")
                If underscorePos > 1 And IsNumeric(Left(layoutName, underscorePos - 1)) Then
                    layoutName = Mid(layoutName, underscorePos + 1)
                End If

                For j = 1 To nItems
                    If foundNewLayout Then Exit For

                     If layoutName = Trim(layoutMapping(j, 0)) Then
                        foundOldLayout = True
                        ' if a mapping was found, find right layout from the new design
                        For Each newLayout In newLayouts
                            If Trim(newLayout.Name) = Trim(layoutMapping(j, 1)) Then
                                
                                'Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Layout '" & layoutName & "' replaced with '" & newLayout.Name & "'"
                                sld.Design = newLayout.Design
                                sld.CustomLayout = newLayout
                                foundNewLayout = True
                                Exit For
                            End If
                        Next newLayout

                    End If
                Next j
                
                If foundOldLayout And Not foundNewLayout Then
                    Debug.Print "WARNING2: Slide " & sld.SlideIndex & ": No matching layout found for '" & layoutName & "' in new design. Skipped."
                ElseIf Not foundOldLayout Then
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

