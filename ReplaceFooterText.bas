' Replace all old years defined in the 'years' array with the current year
Sub ReplaceFooterText()
    Dim oPres As Presentation
    Dim slideMaster As Master
    Dim shape As shape
    Dim shapes As shapes
    Dim text As String
    Dim thisYear As String
    Dim layout As CustomLayout

    Set oPres = ActivePresentation
    
    Dim years
    years = Array("2017", "2018", "2019", "2020", "2021", "2022", "2023", "2024")
    thisYear = "2025"
    Dim y
    
    With oPres
        For i = .Designs.Count To 1 Step -1
            Debug.Print "Current Master: " & .Designs(i).slideMaster.Design.Name
            
            Set slideMaster = .Designs(i).slideMaster
        
            ' Most slide layouts inerit the Footer directly from the Master, thus change it there
            On Error Resume Next
            For Each shape In slideMaster.shapes
                text = shape.TextEffect.text
                
                For y = LBound(years) To UBound(years)
                    If InStr(text, years(y)) > 0 Then
                        Debug.Print "Replace Footer!" & text
                        shape.TextEffect.text = Replace(shape.TextEffect.text, years(y), thisYear)
                        Debug.Print shape.TextEffect.text
                    End If
                Next
            Next shape
            
            ' Some slide layouts do NOT inherit the Footer from the master slide
            ' For those, iterate through all layouts of the current Master and search for the respective field
            For Each layout In .Designs(i).slideMaster.CustomLayouts
                ' In some layouts the shape.TextEffect is not readable thus skip the error
                On Error Resume Next
                For Each shape In layout.shapes
                    If IsObject(shape.TextEffect) Then

                        text = shape.TextEffect.text
                        For y = LBound(years) To UBound(years)
                            If InStr(text, years(y)) > 0 Then
                                Debug.Print "Replace Footer! " & text
                                shape.TextEffect.text = Replace(shape.TextEffect.text, years(y), thisYear)
                                Debug.Print shape.TextEffect.text
                            End If
                        Next
                    End If
                Next shape
            Next layout

            
        Debug.Print "---"
        Next i
    End With
    MsgBox "Replacement completed!"
End Sub

