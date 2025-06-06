' Slide Analysis2: Lists all used (unique) Slide Designs and how often the same Design Layout was imported.
' There are usually many duplicate slide masters that have the same name but with a different number prefix. (e.g. "23_Blue_theme" and "24_Blue_theme"). 
' This script analyes how often the same design was used, regardless of the number prefix.
Sub SlideMasterAnalysis2()
    Dim oPres As Presentation
    Dim sld As Slide
    Dim design As Design
    Dim i As Integer, j As Integer
    Dim designName As String
    Dim normalizedName As String
    Dim underscorePos As Integer
    Dim foundIndex As Long
    Dim normalizedNameArray() As String
    Dim nameCountArray() As Long
    Dim origNameArray() As String
    Dim outputLen As Integer

    ReDim normalizedNameArray(1 To 1)
    ReDim origNameArray(1 To 1)
    ReDim nameCountArray(1 To 1)

    Set oPres = ActivePresentation

    On Error Resume Next
    With oPres
        For i = .Designs.Count To 1 Step -1
            designName = .Designs(i).SlideMaster.Design.Name
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
            For j = LBound(normalizedNameArray) To UBound(normalizedNameArray)
                If normalizedNameArray(j) = normalizedName Then
                    foundIndex = j
                    Exit For
                End If
            Next j

            If foundIndex > 0 Then
                nameCountArray(foundIndex) = nameCountArray(foundIndex) + 1
                ' if orignalName is non-normalized, and the current designName is normalized, update the original name
                If designName = normalizedName And origNameArray(foundIndex) <> designName Then
                    origNameArray(foundIndex) = designName
                End If
            ' Add new entry if foundindex = 0 (normalizedName was not found in normalizedNameArray)
            Else
                ' if normalizedNameArray is completely empty, add it at position 1
                If normalizedNameArray(1) = "" Then
                    normalizedNameArray(1) = normalizedName
                    nameCountArray(1) = 1
                    origNameArray(1) = designName
                ' increase size of array by 1 and append new item
                Else
                    ReDim Preserve normalizedNameArray(1 To UBound(normalizedNameArray) + 1)
                    ReDim Preserve nameCountArray(1 To UBound(nameCountArray) + 1)
                    ReDim Preserve origNameArray(1 To UBound(origNameArray) + 1)
                
                    normalizedNameArray(UBound(normalizedNameArray)) = normalizedName
                    nameCountArray(UBound(nameCountArray)) = 1
                    origNameArray(UBound(origNameArray)) = designName
                End If
            End If
        Next i
    End With


    ' max output len = 20
    outputLen = 35
    Debug.Print String(outputLen, "-")
    Debug.Print "Design Name ---------------- Count"
    Debug.Print String(outputLen, "-")
    
    ' max output len = 20
    outputLen = 30
    For i = LBound(origNameArray) To UBound(origNameArray)
        Debug.Print origNameArray(i) & String(outputLen - Len(origNameArray(i)), "-") & nameCountArray(i)
    Next i

    MsgBox "Finished!"
    
End Sub
