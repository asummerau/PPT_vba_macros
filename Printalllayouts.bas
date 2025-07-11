
' Printalllayouts: this simple macro prints all custom layouts of the specified master design.
Sub Printalllayouts()
    Dim oPres As Presentation
    Dim myDesign As Design
    Dim Design As Design
    Dim layout As CustomLayout
    Dim myDesignName As String
    Dim layoutNames As String
    Dim outputFile As String
    Dim fileName As String
    Dim fNum As Integer
    Dim i As Integer
    Dim shouldExportToFile As Boolean
    Dim oldOrNew As String

    Set oPres = ActivePresentation
    layoutNames = ""

    On Error Resume Next
    Debug.Print "-----START-----"

    With oPres
        
        ' === STEP 1: Set your desired master name ===
        myDesignName = "DESIGN NAME" ' <-- Replace this with your actual master name
        shouldExportToFile = False ' Set to True if you want to export to file
        oldOrNew = "new" ' Set to "old" if the specified design is an old design (will be replaced), if its a new design, set to "new"
        fileName = myDesignName & "_PowerPoint_Layouts.txt" 

        ' === STEP 2: Try to find that design ===
        Set myDesign = Nothing
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            If design.Name = myDesignName Then
                Debug.Print "Found design: " & design.Name
                Set myDesign = design
                Exit For
            End If
        Next i
        
        If myDesign Is Nothing Then
        MsgBox "Master design '" & myDesignName & "' not found.", vbExclamation
            Exit Sub
        End If

        ' === STEP 3: Collect layout names ===
        if oldOrNew  = "old" Then
            layoutNames = "layoutMapping(0, 0) = " & myDesign.Name & vbCrLf
        Else
            layoutNames = "layoutMapping(0, 1) = " & myDesign.Name & vbCrLf
        End If
        Dim j As Integer
        Dim numLayouts As Integer
        numLayouts = myDesign.SlideMaster.CustomLayouts.Count

        For j = 1 To numLayouts
            Set layout = myDesign.SlideMaster.CustomLayouts(j)
            Debug.Print layout.Name
            if oldOrNew  = "old" Then
                layoutNames = layoutNames & "layoutMapping(" & j & ", 0) = " & layout.Name & vbCrLf
            Else
                layoutNames = layoutNames & "layoutMapping(" & j & ", 1) = " & layout.Name & vbCrLf
            End If
        Next j

    If shouldExportToFile Then
        ' === STEP 4: Set output file path on Mac Desktop ===
        outputFile = MacScript("return POSIX path of (path to desktop folder)") & fileName

        ' === STEP 5: Write to file ===
        fNum = FreeFile
        Open outputFile For Output As #fNum
        Print #fNum, layoutNames
        Close #fNum
        MsgBox "Exported layout names to: " & outputFile, vbInformation
    End If
    
    End With
    Debug.Print "-----END-----"


End Sub

