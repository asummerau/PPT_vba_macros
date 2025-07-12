'NormalizeSlideLayouts: This macro removes slide layouts that have been added over time. 
' Any slide that is using a non-official layout will be updated to use the official layout based on the mapping provided.
Sub NormalizeSlideLayouts()
    Dim oPres As Presentation
    Dim targetDesign As Design
    Dim layout As CustomLayout
    Dim layoutDict As Object
    Dim layoutName As String
    Dim normLayoutName As String
    Dim layoutCanonical As CustomLayout
    Dim layoutCandidate As CustomLayout
    Dim slide As slide
    Dim sldLayout As CustomLayout
    Dim k As Variant
    Dim lastLayoutName As String
    Dim lastLayoutIndex As Integer

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
        ' find the index of the layout called "Closing Midnight"
        Dim j As Integer
        Dim numLayouts As Integer
        numLayouts = targetDesign.SlideMaster.CustomLayouts.Count

        For j = 1 To numLayouts
            Set layout = targetDesign.SlideMaster.CustomLayouts(j)
            If layout.Name = lastLayoutName Then
                lastLayoutIndex = j
            End If

            If lastLayoutIndex > 0 And layout.Name <> lastLayoutName Then
                Debug.Print "Non-official layout found: " & layout.Name

                ' 1. check if this slide was used in the presentation, if not delete it
                ' 2. if it was used, move the slides to the official layout and delete the layout
                ' 2.1. layout is imported from another presentation, check for the most similar layout
                ' 2.2. layout is a duplicate, e.g., has a prefix, move the slides to the official layout and delete the layout
            
                Dim sld As Slide
                Dim layoutWasUsed As Boolean
                Dim foundLayout As Boolean

                layoutWasUsed = False
                ' replace all slides that are using this layout  
                For Each sld In oPres.Slides                
                    If sld.CustomLayout.Name = layout.Name Then
                        Debug.Print "Non-official layout '" & layout.Name & "' is currently being used."
                        layoutWasUsed = True

                        ' find the new layout name based on the mapping
                        Dim newLayoutName As String
                        newLayoutName = FindMapping(GetCanonicalName(layout.Name))
                        If newLayoutName = "" Then
                            MsgBox "Non-official layout '" & layout.Name & "' not found in Mapping.", vbExclamation
                            Exit Sub
                        End If

                        ' replace the slide layout with the new layout from the target design
                        Dim targetDesignLayout As CustomLayout
                        foundLayout = False
                        For Each targetDesignLayout In targetDesign.SlideMaster.CustomLayouts
                            If targetDesignLayout.Name = newLayoutName Then
                                sld.CustomLayout = targetDesignLayout
                                foundLayout = True
                                Debug.Print "Moved slide to new layout: " & newLayoutName
                                Exit For
                            End If
                        Next targetDesignLayout

                        If Not foundLayout Then
                            MsgBox "Layout '" & newLayoutName & "' not found in TargetDesign.", vbExclamation
                            Exit Sub
                        End If
                        ' TODO: Parse the BAS file content to extract objects/mappings
                        ' Example: look for specific patterns or execute the code if it contains functions
                        
                    End If
                Next sld

                ' after having gone through all slides, delete the layout
                If Not layoutWasUsed Then
                    Debug.Print "Deleting unused non-official layout: " & layout.Name
                    layout.Delete
                Else
                    Debug.Print "Deleting non-official layout after moving slides: " & layout.Name
                    layout.Delete
                End If
            End If
        Next j

        If lastLayoutIndex = 0 Then
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
    Dim newLayoutName As String
    Dim i As Integer
    Dim nItems As Integer

    layoutMapping = loadMapping()
    nItems = UBound(layoutMapping, 1)

    For i = 1 To nItems
        If layoutMapping(i, 0) = layoutName Then
            newLayoutName = layoutMapping(i, 1)
            FindMapping = newLayoutName
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