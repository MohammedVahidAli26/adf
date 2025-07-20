Sub ValidateAllVerdicts()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.count, 1).End(xlUp).row

    ' ------------------ COLUMN INDEX MAPPING ------------------
    Dim colTenantVerdict As Long: colTenantVerdict = GetCol(ws, "Tenant verdict")
    Dim colTenantAttr As Long: colTenantAttr = GetCol(ws, "Attribute", colTenantVerdict)
    Dim colTenantOtherAttr As Long: colTenantOtherAttr = GetCol(ws, "Other attribute", colTenantVerdict)
    Dim colTenantErrType As Long: colTenantErrType = GetCol(ws, "tenant verdict error type")
    Dim colTenantComments As Long: colTenantComments = GetCol(ws, "Conflict comments", colTenantVerdict)

    Dim colSourceVerdict As Long: colSourceVerdict = GetCol(ws, "Source verdict")
    Dim colSourceAttr As Long: colSourceAttr = GetCol(ws, "Attribute", colSourceVerdict)
    Dim colSourceOtherAttr As Long: colSourceOtherAttr = GetCol(ws, "Other attribute", colSourceVerdict)
    Dim colSourceErrType As Long: colSourceErrType = GetCol(ws, "Source verdict error type")
    Dim colSourceComments As Long: colSourceComments = GetCol(ws, "Conflict comments", colSourceVerdict)

    Dim colTVSVerdict As Long: colTVSVerdict = GetCol(ws, "Tenant vs Source Verdict")
    Dim colTVSAttr As Long: colTVSAttr = GetCol(ws, "Attribute", colTVSVerdict)
    Dim colTVSOtherAttr As Long: colTVSOtherAttr = GetCol(ws, "Other Attribute Name")
    Dim colTVSErrType As Long: colTVSErrType = GetCol(ws, "Tenant vs Source verdict error type")
    Dim colTVSComments As Long: colTVSComments = GetCol(ws, "Tenant vs Source verdict Comments")

    Dim colTenantImgValid As Long: colTenantImgValid = GetCol(ws, "Image validation", colTenantVerdict)
    Dim colSourceImgValid As Long: colSourceImgValid = GetCol(ws, "Image validation", colSourceVerdict)
    Dim colTenantImageValid As Long: colTenantImageValid = GetCol(ws, "Image validation", colTenantVerdict)
Dim colSourceImageValid As Long: colSourceImageValid = GetCol(ws, "Image validation", colSourceVerdict)


    ' ------------------ PROCESS ROWS ------------------
    Dim i As Long
    Dim errorCount As Long: errorCount = 0

    For i = 2 To lastRow
        ' Validate Tenant Verdict
        errorCount = errorCount + ValidateSection(ws, i, colTenantVerdict, colTenantAttr, colTenantOtherAttr, colTenantErrType, colTenantComments)

        ' Validate Source Verdict
        errorCount = errorCount + ValidateSection(ws, i, colSourceVerdict, colSourceAttr, colSourceOtherAttr, colSourceErrType, colSourceComments)

        ' Verdict value capture
        Dim tenantVerdict As String: tenantVerdict = Trim(ws.Cells(i, colTenantVerdict).Value)
        Dim sourceVerdict As String: sourceVerdict = Trim(ws.Cells(i, colSourceVerdict).Value)
        Dim tvsVerdict As String: tvsVerdict = Trim(ws.Cells(i, colTVSVerdict).Value)

        ' Rule: If any verdict is "Unable to validate", TVS must be "Unable to validate"
        If tenantVerdict = "Unable to validate" Or sourceVerdict = "Unable to validate" Then
            If tvsVerdict <> "Unable to validate" Then
                Highlight ws, i, colTVSVerdict: errorCount = errorCount + 1
            End If
        End If

        ' Rule: Comments required if verdict is "Unable to validate"
        If tenantVerdict = "Unable to validate" And Trim(ws.Cells(i, colTenantComments).Value) = "" Then
            Highlight ws, i, colTenantComments: errorCount = errorCount + 1
        End If
        If sourceVerdict = "Unable to validate" And Trim(ws.Cells(i, colSourceComments).Value) = "" Then
            Highlight ws, i, colSourceComments: errorCount = errorCount + 1
        End If
        If tvsVerdict = "Unable to validate" And Trim(ws.Cells(i, colTVSComments).Value) = "" Then
            Highlight ws, i, colTVSComments: errorCount = errorCount + 1
        End If

        ' ----------- TVS Verdict Decision Rules -------------
        If LCase(tenantVerdict) = "no conflict" And LCase(sourceVerdict) = "no conflict" Then
            If tvsVerdict = "" Then
                Highlight ws, i, colTVSVerdict: errorCount = errorCount + 1
            Else
                errorCount = errorCount + ValidateTVSSection(ws, i, colTVSVerdict, colTVSAttr, colTVSOtherAttr, colTVSErrType, colTVSComments)
            End If

        ElseIf LCase(tenantVerdict) = "no conflict" And sourceVerdict <> "" And LCase(sourceVerdict) <> "no conflict" Then
            If tvsVerdict <> "" Then Highlight ws, i, colTVSVerdict: errorCount = errorCount + 1

        ElseIf LCase(sourceVerdict) = "no conflict" And tenantVerdict <> "" And LCase(tenantVerdict) <> "no conflict" Then
            If tvsVerdict <> "" Then Highlight ws, i, colTVSVerdict: errorCount = errorCount + 1

        ElseIf tenantVerdict <> "" And sourceVerdict <> "" And _
               LCase(tenantVerdict) <> "no conflict" And LCase(sourceVerdict) <> "no conflict" And _
               tenantVerdict <> "Unable to validate" And sourceVerdict <> "Unable to validate" Then
            If tvsVerdict <> "" Then Highlight ws, i, colTVSVerdict: errorCount = errorCount + 1

        ElseIf tenantVerdict <> "" And sourceVerdict <> "" Then
            errorCount = errorCount + ValidateTVSSection(ws, i, colTVSVerdict, colTVSAttr, colTVSOtherAttr, colTVSErrType, colTVSComments)

        Else
            ClearHighlight ws, i, colTVSAttr
            ClearHighlight ws, i, colTVSOtherAttr
            ClearHighlight ws, i, colTVSErrType
            ClearHighlight ws, i, colTVSComments
        End If

' ------------------ IMAGE VALIDATION (Tenant & Source) ------------------
Dim tenantImgValidRaw As String: tenantImgValidRaw = LCase(Trim(ws.Cells(i, colTenantImageValid).Value))
Dim sourceImgValidRaw As String: sourceImgValidRaw = LCase(Trim(ws.Cells(i, colSourceImageValid).Value))



Dim normTenantImgValid As String: normTenantImgValid = NormalizeImageValidation(tenantImgValidRaw)
Dim normSourceImgValid As String: normSourceImgValid = NormalizeImageValidation(sourceImgValidRaw)

' --- Tenant Image Validation ---
If normTenantImgValid <> "" Then
    ws.Cells(i, colTenantImageValid).Value = normTenantImgValid
    ws.Cells(i, colTenantImageValid).Interior.ColorIndex = xlNone
Else
    Highlight ws, i, colTenantImageValid
    errorCount = errorCount + 1
End If

' --- Source Image Validation ---
If normSourceImgValid <> "" Then
    ws.Cells(i, colSourceImageValid).Value = normSourceImgValid
    ws.Cells(i, colSourceImageValid).Interior.ColorIndex = xlNone
Else
    Highlight ws, i, colSourceImageValid
    errorCount = errorCount + 1
End If



    Next i

  MsgBox "Verdict Validation Complete " & vbCrLf & vbCrLf & _
       "Total Errors Highlighted: " & errorCount & vbCrLf & vbCrLf & _
       "Check the red-highlighted cells for issues.", _
       vbInformation + vbOKOnly, " Validation Summary"


End Sub

' --- Helper Function to Normalize Value ---
Function NormalizeImageValidation(val As String) As String
    Select Case val
        Case "yes"
            NormalizeImageValidation = "Yes"
        Case "unable to validate", "unble to validate", "unable to validte", "unable validate", "unabl"
            NormalizeImageValidation = "Unable to validate"
        Case Else
            NormalizeImageValidation = ""
    End Select
End Function


' ------------------ TENANT/SOURCE ------------------
Function ValidateSection(ws As Worksheet, rowNum As Long, _
    colVerdict As Long, colAttr As Long, colOtherAttr As Long, _
    colErrType As Long, colComment As Long) As Long

    Dim errorCount As Long: errorCount = 0

    Dim verdict As String
    verdict = Trim(ws.Cells(rowNum, colVerdict).Value)

    Dim attrVal As String: attrVal = Trim(ws.Cells(rowNum, colAttr).Value)
    Dim otherAttrVal As String: otherAttrVal = Trim(ws.Cells(rowNum, colOtherAttr).Value)
    Dim errTypeVal As String: errTypeVal = Trim(ws.Cells(rowNum, colErrType).Value)
    Dim commentVal As String: commentVal = Trim(ws.Cells(rowNum, colComment).Value)

    ' Clear previous highlights
    ClearHighlight ws, rowNum, colAttr
    ClearHighlight ws, rowNum, colOtherAttr
    ClearHighlight ws, rowNum, colErrType
    ClearHighlight ws, rowNum, colComment

    If verdict = "" Then
        Highlight ws, rowNum, colVerdict: errorCount = errorCount + 1
        Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        Highlight ws, rowNum, colComment: errorCount = errorCount + 1

       ElseIf verdict = "Unable to validate" Then
        If attrVal <> "" Then Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        If otherAttrVal <> "" Then Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        If errTypeVal <> "" Then Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        If commentVal = "" Then Highlight ws, rowNum, colComment: errorCount = errorCount + 1

    ElseIf verdict = "No conflict" Then
        If attrVal <> "" Then Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        If otherAttrVal <> "" Then Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        If errTypeVal <> "" Then Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        If commentVal <> "" Then Highlight ws, rowNum, colComment: errorCount = errorCount + 1

    Else
        If attrVal = "" Then Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        If errTypeVal = "" Then Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        If commentVal = "" Then Highlight ws, rowNum, colComment: errorCount = errorCount + 1
        If LCase(attrVal) = "other_attribute" And otherAttrVal = "" Then
            Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        End If
    End If

    ValidateSection = errorCount

End Function

' ------------------ TVS CUSTOM LOGIC ------------------
Function ValidateTVSSection(ws As Worksheet, rowNum As Long, _
    colVerdict As Long, colAttr As Long, colOtherAttr As Long, _
    colErrType As Long, colComment As Long) As Long

    Dim errorCount As Long: errorCount = 0

    Dim verdict As String: verdict = Trim(ws.Cells(rowNum, colVerdict).Value)
    Dim attrVal As String: attrVal = Trim(ws.Cells(rowNum, colAttr).Value)
    Dim otherAttrVal As String: otherAttrVal = Trim(ws.Cells(rowNum, colOtherAttr).Value)
    Dim errTypeVal As String: errTypeVal = Trim(ws.Cells(rowNum, colErrType).Value)
    Dim commentVal As String: commentVal = Trim(ws.Cells(rowNum, colComment).Value)

    ' Clear previous highlights
    ClearHighlight ws, rowNum, colAttr
    ClearHighlight ws, rowNum, colOtherAttr
    ClearHighlight ws, rowNum, colErrType
    ClearHighlight ws, rowNum, colComment

    If verdict = "" Then
        Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        Highlight ws, rowNum, colComment: errorCount = errorCount + 1

          ElseIf verdict = "Unable to validate" Then
        If attrVal <> "" Then Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        If otherAttrVal <> "" Then Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        If errTypeVal <> "" Then Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        If commentVal = "" Then Highlight ws, rowNum, colComment: errorCount = errorCount + 1

    ElseIf verdict = "Match" Then
        If attrVal <> "" Then Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        If otherAttrVal <> "" Then Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        If errTypeVal <> "" Then Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        If commentVal <> "" Then Highlight ws, rowNum, colComment: errorCount = errorCount + 1

    Else
        If attrVal = "" Then Highlight ws, rowNum, colAttr: errorCount = errorCount + 1
        If errTypeVal = "" Then Highlight ws, rowNum, colErrType: errorCount = errorCount + 1
        If commentVal = "" Then Highlight ws, rowNum, colComment: errorCount = errorCount + 1
        If LCase(attrVal) = "other_attribute" And otherAttrVal = "" Then
            Highlight ws, rowNum, colOtherAttr: errorCount = errorCount + 1
        End If
    End If

    ValidateTVSSection = errorCount

End Function


' ------------------ HELPERS ------------------
Function GetCol(ws As Worksheet, header As String, Optional afterCol As Long = 0) As Long
    Dim c As Range
    For Each c In ws.Rows(1).Cells
        If c.Column > afterCol And LCase(Trim(c.Value)) = LCase(header) Then
            GetCol = c.Column
            Exit Function
        End If
    Next
    MsgBox "Column not found: " & header, vbCritical
    End
End Function

Sub Highlight(ws As Worksheet, row As Long, col As Long)
    ws.Cells(row, col).Interior.Color = RGB(255, 0, 0)
End Sub

Sub ClearHighlight(ws As Worksheet, row As Long, col As Long)
    ws.Cells(row, col).Interior.ColorIndex = xlNone
End Sub




