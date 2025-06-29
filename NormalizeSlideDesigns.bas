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
    Dim normDesignName As String
    Dim normLayoutName As String
    Dim foundIndex As Long
    Dim normNameArray() As String
    Dim nameCountArray() As Long
    Dim nameArray() As String
    Dim designRefArray() As Design

    ReDim normNameArray(1 To 1)
    ReDim nameArray(1 To 1)
    ReDim nameCountArray(1 To 1)
    ReDim designRefArray(1 To 1)

    Set oPres = ActivePresentation

    ' STEP 1: Build list of normalized design names and count usage
    On Error Resume Next
    With oPres
        ' (e.g. "23_Blue_theme", "Blue_theme", "1_Green_theme", "2_Green_theme", "2_Red_theme", "Black_theme")
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            designName = .Designs(i).Name
            normDesignName = GetCanonicalName(designName)

            ' Check if the normalized name has been added before
            foundIndex = 0
            For j = LBound(normNameArray) To UBound(normNameArray)
                If normNameArray(j) = normDesignName Then
                    foundIndex = j
                    Exit For
                End If
            Next j

            If foundIndex > 0 Then
                nameCountArray(foundIndex) = nameCountArray(foundIndex) + 1
                ' If original name is non-normalized, and current designName is normalized, update the original name
                If designName = normDesignName And nameArray(foundIndex) <> normDesignName Then
                    nameArray(foundIndex) = designName
                    Set designRefArray(foundIndex) = design
                End If
            Else
                ' Add new entry
                If normNameArray(1) = "" Then
                    normNameArray(1) = normDesignName
                    nameCountArray(1) = 1
                    nameArray(1) = designName
                    Set designRefArray(1) = design
                Else
                    ReDim Preserve normNameArray(1 To UBound(normNameArray) + 1)
                    ReDim Preserve nameCountArray(1 To UBound(nameCountArray) + 1)
                    ReDim Preserve nameArray(1 To UBound(nameArray) + 1)
                    ReDim Preserve designRefArray(1 To UBound(designRefArray) + 1)

                    normNameArray(UBound(normNameArray)) = normDesignName
                    nameCountArray(UBound(nameCountArray)) = 1
                    nameArray(UBound(nameArray)) = designName
                    Set designRefArray(UBound(designRefArray)) = design
                End If
            End If

        Next i
        ' Output will be 
        ' normNameArray=("Blue_theme", "Green_theme", "Red_theme", "Black_theme")
        ' nameArray=("Blue_theme", "1_Green_theme", "2_Red_theme", "Black_theme")
        ' nameCountArray=(2, 2, 1, 1)
        ' designRefArray=(Design("Blue_theme"), Design("1_Green_theme"), Design("2_Red_theme"), Design("Black_theme"))
        ' --> designRefArray(j).Name == nameArray(j), for all j

        ' STEP 2: Go through slides and update layout if better matching design exists
        For Each sld In oPres.Slides
            layoutName = Trim(sld.CustomLayout.Name)
            designName = Trim(sld.Design.Name)
            normDesignName = GetCanonicalName(designName)
            normLayoutName = GetCanonicalName(layoutName)

            ' if the current design is already using the normalized design, skip it
            If normDesignName = designName Then
                Debug.Print "Slide " & sld.SlideIndex & " already uses normalized design: " & designName
                GoTo NextSlide
            Else
            ' else find canonical design with the same normalized name (bc designName has a prefix like "23_", "22_", etc.)
                For j = LBound(normNameArray) To UBound(normNameArray)
                    If normNameArray(j) = normDesignName Then                         
                        ' no need to replace design because even though it has a prefix, there was not normalized design in the presentation
                        ' (e.g. keep "1_Green_theme" -> "1_Green_theme")
                        If designName = nameArray(j) Then 
                            Debug.Print "No need to update Slide " & sld.SlideIndex & ": from '" & designName & "' to '" & designRefArray(j).Name & "'"
                        
                        'if designName <> nameArray(j) --> replace with nameArray(j)'s design (e.g replace "23_Blue_theme" with "Blue_theme")
                        Else 
                            ' Try to find matching layout by name
                            Dim newLayout As CustomLayout
                            Dim foundLayout As Boolean
                            foundLayout = False

                            ' Check if the layout name matches any layout in the new design
                            For Each newLayout In designRefArray(j).SlideMaster.CustomLayouts
                                Dim normNewLayoutName As String
                                normNewLayoutName = GetCanonicalName(newLayout.Name)
                                If normNewLayoutName = normLayoutName Then
                                    Debug.Print "+++Updating Slide " & sld.SlideIndex & ": from '" & designName & "' to '" & designRefArray(j).Name & "'"
                                    foundLayout = True

                                    sld.CustomLayout = newLayout
                                    sld.design = designRefArray(j) 
                                    Exit For
                                End If
                            Next newLayout

                            If Not foundLayout Then
                                Debug.Print "WARNING: Slide " & sld.SlideIndex & " Could not find layout '" & normLayoutName & "' in " & designRefArray(j).Name & ", copying the layout to it."
                                ' sld.design = designRefArray(j) -> This line not only copies over this layout but also all other layouts from the old design
                            
                                ' Trick: Create a temporary slide using the original layout, then assign normalized master to it
                                Dim tempSlide As Slide
                                Set tempSlide = oPres.Slides.Add(oPres.Slides.Count + 1, ppLayoutBlank)
                                tempSlide.CustomLayout = sld.CustomLayout

                                ' This assigns the new master, which auto-imports the necessary layout
                                tempSlide.Design = designRefArray(j)

                                ' Use the imported layout on the original slide
                                sld.CustomLayout = tempSlide.CustomLayout                                       
                                
                                ' Remove the temp slide after use
                                tempSlide.Delete                            
                                End If
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


Function GetCanonicalName(name As String) As String
    Dim underscorePos As Integer
    Dim prefix As String
    Dim cleanedName As String

    cleanedName = Trim(name)
    underscorePos = InStr(cleanedName, "_")
    
    If underscorePos > 1 Then
        prefix = Left(cleanedName, underscorePos - 1)
        If IsNumeric(prefix) Then
            GetCanonicalName = Mid(cleanedName, underscorePos + 1)
            Exit Function
        End If
    End If
    
    GetCanonicalName = cleanedName
End Function
