
' ReplaceOldDesign2: Replaces the layout of a slide (from a specified old Slide Master) with the layout of the new Slide Master which has to be specified in the code. 
' This macro can be used if the layout names don't match. To make a mapping, a manual mapping has to be done and specified in the code.
' NOTE: this code may mess up your presentation. Only apply on a copy of your presentation!
Sub ReplaceOldDesign2()
    Dim oPres As Presentation
    Dim sld As Slide
    Dim layoutName As String
    Dim i As Integer, j As Integer
    Dim designName As String
    Dim newDesign As Design
    Dim newLayouts As CustomLayouts
    Dim newLayout As CustomLayout
    Dim foundNewLayout As Boolean
    Dim foundOldLayout As Boolean
    Dim currentLayouts As CustomLayouts
    Dim nItems As Integer
    Dim layoutMapping() As String

    ' Modify this to your actual master name
    Const TARGET_MASTER_NAME As String =  "ADD YOUR NEW DESIGN NAME HERE"    
    Const OLD_MASTER_NAME As String = "ADD YOUR OLD DESIGN NAME HERE"

    Const CSV_FILE_NAME As String = "layoutmapping.csv"
    layoutMapping = loadMapping(CSV_FILE_NAME)
    nItems = UBound(layoutMapping, 1)

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"

    With oPres
        Debug.Print "New Design: "; TARGET_MASTER_NAME
        Debug.Print "Old Design: "; OLD_MASTER_NAME
        Debug.Print

        ' STEP 1: Find the new design in the presentation
        Set newDesign = Nothing
        Set newLayouts = Nothing
        For i = .Designs.Count To 0 Step -1
            layoutName = Trim(.Designs(i).Name)
            layoutName = GetCanonicalName(layoutName)

            If layoutName = TARGET_MASTER_NAME Then
                ' Debug.Print "Found new design: " & .Designs(i).Name
                Set newDesign = .Designs(i)
                Set newLayouts = .Designs(i).SlideMaster.CustomLayouts
                Exit For
            End If
        Next i
        
        If newDesign Is Nothing Then
            MsgBox "New design '" & TARGET_MASTER_NAME & "' not found in the presentation.", vbExclamation
            Exit Sub
        End If

        ' STEP 2: Try to replace old designs with the new design based on the predefined mapping array
        For Each sld In oPres.Slides
            layoutName = Trim(sld.CustomLayout.Name)
            designName = Trim(sld.Design.Name)

            foundNewLayout = False
            foundOldLayout = False

            designName = GetCanonicalName(designName)

            If designName = OLD_MASTER_NAME Then
                Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Find replacement for Layout '" & layoutName & "'"
                ' Check if a mapping exists in the predefined array

                ' there are tons of duplicate layouts that start with a prefix (e.g. 1_title is the same as title)
                layoutName = GetCanonicalName(layoutName)

                For j = 0 To nItems
                    If foundNewLayout Then Exit For

                     If layoutName = Trim(layoutMapping(j, 0)) Then
                        foundOldLayout = True
                        ' if a mapping was found, find right layout from the new design
                        For Each newLayout In newLayouts
                            If Trim(newLayout.Name) = Trim(layoutMapping(j, 1)) Then
                                
                                Debug.Print "--> Layout '" & layoutName & "' replaced with '" & newLayout.name & "'"
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
            
            ElseIf designName = TARGET_MASTER_NAME Then
                Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Design is already '" & TARGET_MASTER_NAME & "', skipping."

            Else
                Debug.Print "PPT Slide #: " & sld.SlideIndex & ": Another design was found ('" & designName & "'). Skipping."
            End If

        Next sld
    End With
    Debug.Print "-----END-----"

    MsgBox "Design replacement complete!", vbInformation
End Sub

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

Function loadMapping(fileName) As Variant
    ' This function reads layout mappings from a CSV file instead of hardcoded values
    ' CSV Format: OldLayoutName,NewLayoutName (with header row)
    ' File should be in the same directory as the presentation
    Dim nItems As Integer
    Dim layoutMapping() As String
    Dim filePath As String
    Dim fileNum As Integer
    Dim lineText As String
    Dim splitData As Variant
    Dim i As Integer
    
    ' Set the path to the CSV file (same directory as presentation)
    filePath = ActivePresentation.Path & "/" & fileName
    fileNum = FreeFile
    
    ' First, count the number of lines to size the array
    On Error GoTo FileError
    Open filePath For Input As #fileNum
    nItems = 0
    While Not EOF(fileNum)
        Line Input #fileNum, lineText
        If Trim(lineText) <> "" Then ' Skip empty lines
            nItems = nItems + 1
        End If
    Wend
    Close #fileNum
    
    ' Subtract 1 for header row
    nItems = nItems - 1
    If nItems < 1 Then
        MsgBox "No data found in CSV file: " & filePath, vbExclamation
        Exit Function
    End If
    
    ' Size the array
    ReDim layoutMapping(0 To nItems, 0 To 1)
    
    ' Read the CSV file and populate the array
    Open filePath For Input As #fileNum
    Line Input #fileNum, lineText ' Skip header row
    
    i = 0
    While Not EOF(fileNum) And i <= nItems
        Line Input #fileNum, lineText
        If Trim(lineText) <> "" Then ' Skip empty lines
            ' Split by comma and clean up quotes
            splitData = Split(lineText, """,""")
            If UBound(splitData) >= 1 Then
                layoutMapping(i, 0) = Trim(Replace(splitData(0), """", ""))
                layoutMapping(i, 1) = Trim(Replace(splitData(1), """", ""))
                i = i + 1
            End If
        End If
    Wend
    Close #fileNum
    
    ' Return the populated array
    loadMapping = layoutMapping
    Exit Function
    
FileError:
    MsgBox "Error reading CSV file: " & filePath & vbCrLf & "Error: " & Err.Description, vbCritical
    If fileNum > 0 Then Close #fileNum
    ' Return empty array on error
    ReDim layoutMapping(0 To 0, 0 To 1)
    loadMapping = layoutMapping
End Function
