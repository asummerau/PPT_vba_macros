' Slide Analysis2: Lists all used (unique) Slide Designs and how often the same Design Layout was imported.
Sub SlideMasterAnalysis2()
    Dim oPres As Presentation
    Dim sld As Slide
    Dim design As Design
    Dim i As Integer, j As Integer
    Dim designName As String
    Dim normalizedName As String
    Dim underscorePos As Integer
    Dim foundIndex As Long
    Dim nameArray() As String
    Dim nameCountArray() As Long
    Dim outputLen As Integer

    ReDim nameArray(1 To 1)
    ReDim nameCountArray(1 To 1)

    Set oPres = ActivePresentation

    On Error Resume Next
    With oPres
        For i = .Designs.Count To 1 Step -1
            designName = .Designs(i).slideMaster.Design.Name
            underscorePos = InStr(designName, "_")

            ' Remove number prefix and underscore if present (e.g. "23_Blue_theme" -> "Blue_theme")
            If underscorePos > 1 And IsNumeric(Left(designName, underscorePos - 1)) Then
                normalizedName = Mid(designName, underscorePos + 1)
            Else
                normalizedName = designName
            End If

            ' Debug.Print "Design Name: " & designName
            ' Debug.Print "Normalized Name: " & normalizedName

            ' Check if name already exists in array
            foundIndex = 0
            For j = LBound(nameArray) To UBound(nameArray)
                If nameArray(j) = normalizedName Then
                    foundIndex = j
                    Exit For
                End If
            Next j

            If foundIndex > 0 Then
                nameCountArray(foundIndex) = nameCountArray(foundIndex) + 1
            ' Add new entry if foundindex = 0 (normalizedName was not found in nameArray)
            Else
                ' if nameArray is completely empty, add it at position 1
                If nameArray(1) = "" Then
                    nameArray(1) = normalizedName
                    nameCountArray(1) = 1
                ' increase size of array by 1 and append new item
                Else
                    ReDim Preserve nameArray(1 To UBound(nameArray) + 1)
                    ReDim Preserve nameCountArray(1 To UBound(nameCountArray) + 1)
                    nameArray(UBound(nameArray)) = normalizedName
                    nameCountArray(UBound(nameCountArray)) = 1
                End If
            End If

        Next i
    End With


    ' max output len = 20
    outputLen = 35
    Debug.Print String(outputLen, "-")
    Debug.Print "Design Name -------- Design Count"
    Debug.Print String(outputLen, "-")
    
    ' max output len = 20
    outputLen = 30
    For i = LBound(nameArray) To UBound(nameArray)
        Debug.Print nameArray(i) & String(outputLen - Len(nameArray(i)), "-") & nameCountArray(i)
    Next i

    MsgBox "Finished!"
    
End Sub
