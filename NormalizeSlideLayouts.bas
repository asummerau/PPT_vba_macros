'NormalizeSlideLayouts: This macro removes slide layouts that have been added over time. 
' Any slide that is using a non-official layout will be updated to use the official layout based on the mapping provided.
'
' Strategy:
' 1. Find all non-official layouts (layouts that come after the "lastLayoutName" in the target design)
' 2. For each non-official layout:
'    - Scenario 2.1: If it's a duplicate with prefix (e.g., "1_Title Slide"), try to use canonical layout ("Title Slide")
'    - Scenario 2.2: If no canonical layout exists, use mapping table to find replacement layout
' 3. Delete the non-official layout after moving all slides to official layouts
'
' TODO: Verify scenario 2.2 logic for imported layouts from other presentations
Sub NormalizeSlideLayouts()
    Dim oPres As Presentation
    Dim targetDesign As Design
    Dim layout As CustomLayout
    Dim layoutName As String
    Dim lastLayoutName As String
    Dim lastLayoutIndex As Integer
    Dim j As Integer
    Dim numLayouts As Integer

    ' Modify this to your actual master name
    Const TARGET_MASTER_NAME As String =  "ADD YOUR NEW DESIGN NAME HERE"    
    lastLayoutName = "Closing 1"

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"
    
    With oPres
        ' STEP 1: Find the target slide master in the presentation
        Set targetDesign = Nothing
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            If design.Name = TARGET_MASTER_NAME Then
                Debug.Print "Found target design: " & design.Name
                Set targetDesign = design
                Exit For
            End If
        Next i

        If targetDesign Is Nothing Then
            MsgBox "Target design '" & TARGET_MASTER_NAME & "' not found in the presentation.", vbExclamation
            Exit Sub
        End If

        lastLayoutIndex = 0
        numLayouts = targetDesign.SlideMaster.CustomLayouts.Count

        ' STEP 2: Process each layout in the target design
        ' Strategy: Find the lastLayoutIndex first, then process all layouts that come after it
        For j = 0 To numLayouts - 1
            Set layout = targetDesign.SlideMaster.CustomLayouts(j)
            layoutName = layout.Name

            If layoutName = lastLayoutName Then
                lastLayoutIndex = j
            End If

            ' Process non-official layouts (those that come after the last official layout)
            If lastLayoutIndex > 0 And layoutName <> lastLayoutName Then
                Debug.Print "Non-official layout found: " & layoutName

                Dim sld As Slide
                Dim layoutWasUsed As Boolean
                Dim foundMappingLayout As Boolean
                Dim foundCanonicalLayout As Boolean
                layoutWasUsed = False

                ' STEP 2a: Check all slides to see if they use this non-official layout
                For Each sld In oPres.Slides                
                    If sld.CustomLayout.Name = layoutName Then
                        Debug.Print "-Non-official layout '" & layoutName & "' is currently being used by slide " & sld.SlideIndex & "."
                        layoutWasUsed = True
                        
                        ' SCENARIO 2.1: -------- Try to find canonical layout (remove numeric prefix) -------
                        ' Example: "1_Title Slide" becomes "Title Slide"
                        Dim canonicalLayoutName As String
                        Dim targetDesignLayout As CustomLayout
                        canonicalLayoutName = GetCanonicalName(layoutName)
                        foundCanonicalLayout = False

                        If canonicalLayoutName = layoutName Then
                            Debug.Print "-" & layoutName & "' is already in canonical form."
                        Else
                            ' Search for canonical layout in target design (e.)
                            For Each targetDesignLayout In targetDesign.SlideMaster.CustomLayouts
                                If targetDesignLayout.Name = canonicalLayoutName Then
                                    sld.CustomLayout = targetDesignLayout
                                    foundCanonicalLayout = True
                                    Debug.Print "-Moved slide """ & sld.SlideIndex & """ to canonical layout: " & canonicalLayoutName
                                    Exit For
                                End If

                                If foundCanonicalLayout Then
                                    GoTo NextSlide
                                End If

                            Next targetDesignLayout
  
                        End If
               
                        ' SCENARIO 2.2:  -------- Use mapping table to find replacement layout -------
                        ' This handles layouts imported from other presentations
                        Dim newLayoutName As String
                        newLayoutName = FindMapping(canonicalLayoutName)
                        If newLayoutName = "" Then
                            MsgBox "Non-official layout '" & layoutName & "' not found in Mapping.", vbExclamation
                            Exit Sub
                        End If

                        ' Apply the mapped layout to the slide
                        foundMappingLayout = False
                        For Each targetDesignLayout In targetDesign.SlideMaster.CustomLayouts
                            If targetDesignLayout.Name = newLayoutName Then
                                sld.CustomLayout = targetDesignLayout
                                foundMappingLayout = True
                                Debug.Print "-Moved slide """ & sld.SlideIndex & """ to new layout: " & newLayoutName
                                Exit For
                            End If
                        Next targetDesignLayout

                        If Not foundMappingLayout Then
                            MsgBox "Layout '" & newLayoutName & "' not found in TargetDesign.", vbExclamation
                            Exit Sub
                        End If

NextSlide:
                    End If
                Next sld

                ' STEP 3: Clean up - delete the non-official layout after processing all slides
                If Not layoutWasUsed Then
                    Debug.Print "**Deleting unused non-official layout: " & layoutName
                    layout.Delete
                    j = j - 1 ' Adjust index since we just deleted a layout
                    numLayouts = targetDesign.SlideMaster.CustomLayouts.Count
                Else
                    Debug.Print "**Deleting non-official layout """ & layoutName & """ after moving all slides: "
                    layout.Delete
                    j = j - 1 ' Adjust index since we just deleted a layout
                    numLayouts = targetDesign.SlideMaster.CustomLayouts.Count
                End If
                
                ' Safety check: exit loop if we've processed all layouts
                If j > numLayouts - 1 Then
                    Debug.Print "Exiting loop since we deleted the last layout."
                    Exit For
                End If
                Debug.Print ""
            End If
        Next j

        If lastLayoutIndex = 0 Then
            Debug.Print "Layout '" & lastLayoutName & "' not found in the target design. Last layout name was: " & layoutName
            MsgBox "Layout '" & lastLayoutName & "' not found in the target design.", vbExclamation
            Exit Sub
        End If
        
        Debug.Print "-----END-----"

        MsgBox "Cleanup complete."
    End With


End Sub


Function loadMapping() As Variant
    Dim nItems As Integer
    Dim layoutMapping() As String
    ' TODO: Specify number of mappings you want to use
    nItems = 96
    ReDim layoutMapping(0 To nItems, 0 To 1)

    ' TODO: fill in the Slide Master Names and update the number of mappings
    layoutMapping(0, 0) = "OLD MASTER NAME"
    layoutMapping(0, 1) = "NEW MASTER NAME"

    ' TODO: fill in the Layout Names you want to replace. E.g.:
    layoutMapping(1, 0) = "Title Slide" 'layout name from old master
    layoutMapping(1, 1) = "New Title Slide" 'layout name from new master

    ' Return the populated array
    loadMapping = layoutMapping
End Function

Function FindMapping(layoutName As String) As String
    Dim layoutMapping() As String
    Dim i As Integer
    Dim nItems As Integer

    layoutMapping = loadMapping()
    nItems = UBound(layoutMapping, 1)

    For i = 1 To nItems
        If layoutMapping(i, 0) = layoutName Then
            FindMapping = layoutMapping(i, 1)
            Exit Function
        End If
    Next i
    ' If no mapping found, return an empty string
    FindMapping = ""
End Function

Function GetCanonicalName(name As String) As String
    ' If name has a prefix, remove it (e.g. "23_name" -> "name")
    Dim underscorePos As Integer
    underscorePos = InStr(name, "_")
    If underscorePos > 1 Then 
        If IsNumeric(Left(name, underscorePos - 1)) Then
            GetCanonicalName = Mid(name, underscorePos + 1)
        Else
            GetCanonicalName = name
        End If
    Else
        GetCanonicalName = name
    End If
End Function