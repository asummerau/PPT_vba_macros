
' Printalllayouts: this simple macro prints all custom layouts of the specified master design.
Sub Printalllayouts()
    Dim oPres As Presentation
    Dim myDesign As Design
    Dim Design As Design
    Dim layout As CustomLayout
    Dim outputFile As String
    Dim fileName As String
    Dim fNum As Integer
    Dim i As Integer
    Dim shouldExportToFile As Boolean
    Dim oldOrNew As String

    ' === STEP 1: Set your desired master name ===
    Const MY_DESING_NAME As String = "ADD YOUR NEW DESIGN NAME HERE" ' <-- Replace this with your actual master name
    Const CSV_FILENAME As String = "layoutmapping.csv" 
    shouldExportToFile = False ' Set to True if you want to export to file
    oldOrNew = "new" ' Set to "old" if the specified design is an old design (will be replaced), if its a new design, set to "new"

    Set oPres = ActivePresentation

    On Error Resume Next
    Debug.Print "-----START-----"

    With oPres

        ' === STEP 2: Try to find that design ===
        Set myDesign = Nothing
        For i = .Designs.Count To 1 Step -1
            Set design = .Designs(i)
            If design.Name = MY_DESING_NAME Then
                Debug.Print "Found design: " & design.Name
                Set myDesign = design
                Exit For
            End If
        Next i
        
        If myDesign Is Nothing Then
        MsgBox "Master design '" & MY_DESING_NAME & "' not found.", vbExclamation
            Exit Sub
        End If

        Dim j As Integer
        Dim numLayouts As Integer
        numLayouts = myDesign.SlideMaster.CustomLayouts.Count

        ' === STEP 3: Print layout names ===
        For j = 1 To numLayouts
            Set layout = myDesign.SlideMaster.CustomLayouts(j)
            Debug.Print layout.Name
        Next j

        ' === STEP 4: Collect layout names in CSV format ===
        If shouldExportToFile Then
            outputFile = oPres.Path & "/" & CSV_FILENAME

            ' === STEP 5: Write to CSV file ===
            fNum = FreeFile
            Open outputFile For Output As #fNum
            
            ' Write CSV header
            Print #fNum, "OldLayoutName,NewLayoutName"
            
            ' Write each layout as a separate line
            For j = 1 To numLayouts
                Set layout = myDesign.SlideMaster.CustomLayouts(j)
                
                ' Create CSV row based on oldOrNew setting
                If oldOrNew = "old" Then
                    ' Add layout name to OldLayoutName column, leave NewLayoutName empty
                    Print #fNum, """" & layout.Name & """,""""" 
                Else
                    ' Add layout name to NewLayoutName column, leave OldLayoutName empty  
                    Print #fNum, """""," & """" & layout.Name & """"
                End If
            Next j
            
            Close #fNum
            MsgBox "Exported layout mappings to CSV: " & outputFile, vbInformation
        End If
    
    End With
    Debug.Print "-----END-----"


End Sub

