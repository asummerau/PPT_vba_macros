'NomralizedSlideDesigns: Thi macro moves slides to a canonical master design if there are multiple designs with similar names in the presentation.
' If the same Master Design is imported multiple times (e.g. `"23_Blue_theme"`, `"22_Blue_theme"`, `"Blue_theme"`), it ensures that all slides are moved to the canonical design (e.g. `"Blue_theme"`), while preserving the layout used on each slide.
Sub NormalizeSlideDesigns()
    Dim oPres As Presentation
    Dim sld As Slide
    Dim design As Design
    Dim layoutName As String
    Dim foundNormalized As Boolean
    Dim i As Integer, j As Integer
    Dim designName As String
    Dim normalizedName As String
    Dim underscorePos As Integer
    Dim foundIndex As Long
    Dim normalizedNameArray() As String
    Dim nameCountArray() As Long
    Dim origNameArray() As String
    Dim designRefArray() As Design

    ReDim normalizedNameArray(1 To 1)
    ReDim origNameArray(1 To 1)
    ReDim nameCountArray(1 To 1)
    ReDim designRefArray(1 To 1)

    Set oPres = ActivePresentation

    ' STEP 1: Build list of normalized design names and count usage
    On Error Resume Next
    With oPres
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            designName = .Designs(i).SlideMaster.Design.Name
            underscorePos = InStr(designName, "_")

            ' If design has a prefix, remove it (e.g. "23_Blue_theme" -> "Blue_theme")
            If underscorePos > 1 And IsNumeric(Left(designName, underscorePos - 1)) Then
                normalizedName = Mid(designName, underscorePos + 1)
            Else
                normalizedName = designName
            End If

            ' Check if the normalized name has been added before
            foundIndex = 0
            For j = LBound(normalizedNameArray) To UBound(normalizedNameArray)
                If normalizedNameArray(j) = normalizedName Then
                    foundIndex = j
                    Exit For
                End If
            Next j

            If foundIndex > 0 Then
                nameCountArray(foundIndex) = nameCountArray(foundIndex) + 1
                ' If original name is non-normalized, and current designName is normalized, update the original name
                If designName = normalizedName And origNameArray(foundIndex) <> designName Then
                    origNameArray(foundIndex) = designName
                    Set designRefArray(foundIndex) = design
                End If
            Else
                ' Add new entry
                If normalizedNameArray(1) = "" Then
                    normalizedNameArray(1) = normalizedName
                    nameCountArray(1) = 1
                    origNameArray(1) = designName
                    Set designRefArray(1) = design
                Else
                    ReDim Preserve normalizedNameArray(1 To UBound(normalizedNameArray) + 1)
                    ReDim Preserve nameCountArray(1 To UBound(nameCountArray) + 1)
                    ReDim Preserve origNameArray(1 To UBound(origNameArray) + 1)
                    ReDim Preserve designRefArray(1 To UBound(designRefArray) + 1)

                    normalizedNameArray(UBound(normalizedNameArray)) = normalizedName
                    nameCountArray(UBound(nameCountArray)) = 1
                    origNameArray(UBound(origNameArray)) = designName
                    Set designRefArray(UBound(designRefArray)) = design
                End If
            End If

        Next i

        ' STEP 2: Go through slides and update layout if better matching design exists
        For Each sld In oPres.Slides
            layoutName = Trim(sld.CustomLayout.Name)
            designName = Trim(sld.Design.Name)
            underscorePos = InStr(designName, "_")

            ' Remove number prefix and underscore if present (e.g. "23_Blue_theme" -> "Blue_theme")
            If underscorePos > 1 And IsNumeric(Left(designName, underscorePos - 1)) Then
                normalizedName = Mid(designName, underscorePos + 1)
            Else
                normalizedName = designName
            End If

            ' if the current design is already normalized, skip it
            If normalizedName = designName Then
                Debug.Print "Slide " & sld.SlideIndex & " already uses normalized design: " & designName
                GoTo NextSlide
            Else
            ' else find canonical design with the same normalized name
                For j = LBound(normalizedNameArray) To UBound(normalizedNameArray)
                    If normalizedNameArray(j) = normalizedName Then
                        ' verify that the design of the slide does not match the canonical design (should already be a given)
                        If Not sld.design.Name = designRefArray(j).Name Then
                            Debug.Print "+++Updating Slide " & sld.SlideIndex & ": from '" & sld.design.Name & "' to '" & designRefArray(j).Name & "'"
                            sld.design = designRefArray(j)

                            ' Try to find matching layout by name
                            Dim newLayout As CustomLayout
                            Dim foundLayout As Boolean
                            foundLayout = False

                            ' Check if the layout name matches any layout in the new design
                            For Each newLayout In designRefArray(j).SlideMaster.CustomLayouts
                                If newLayout.Name = layoutName Then
                                    sld.CustomLayout = newLayout
                                    foundLayout = True
                                    Exit For
                                End If
                            Next newLayout

                            If Not foundLayout Then
                                Debug.Print "WARNING: Could not find matching layout '" & layoutName & "' in new master. Using default layout."
                            End If
                            
                        Else
                            Debug.Print "No need to update Slide " & sld.SlideIndex & ": from '" & sld.design.Name & "' to '" & designRefArray(j).Name & "'"
                        
                        End If
                        Exit For
                    End If
                Next j
            End If
        NextSlide:
        Next sld
    End With

    MsgBox "Design normalization complete!", vbInformation

End Sub
