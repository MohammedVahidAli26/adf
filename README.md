Sub MismatchModel01()
    Dim ws As Worksheet
    Set ws = ActiveSheet

    Dim lastRow As Long, i As Long
    Dim tenantPNCol As Long, sourcePNCol As Long
    Dim tenantSDCol As Long, sourceSDCol As Long
    Dim tenantLDCol As Long, sourceLDCol As Long
    Dim itemIDCol As Long
    Dim serialNum As Long
    Dim tenantPrimaryCol As Long
    Dim sourcePrimaryCol As Long
    
    ' Find column Headers
    tenantPNCol = ws.Rows(1).Find("tenant_product_name").Column
    sourcePNCol = ws.Rows(1).Find("source_product_name").Column
    tenantSDCol = ws.Rows(1).Find("tenant_short_description").Column
    sourceSDCol = ws.Rows(1).Find("source_short_description").Column
    tenantLDCol = ws.Rows(1).Find("tenant_long_description").Column
    sourceLDCol = ws.Rows(1).Find("source_long_description").Column
    
    ' Item ID section
    
    
    
    itemIDCol = ws.Rows(1).Find("item_id").Column
    serialNum = ws.Rows(1).Find("Sl No.").Column
    
    ' Primary Image Sections
    tenantPrimaryCol = ws.Rows(1).Find("tenant_primary_image").Column
    sourcePrimaryCol = ws.Rows(1).Find("source_primary_image").Column
  

    lastRow = ws.Cells(ws.Rows.Count, itemIDCol).End(xlUp).Row
    
    '!!! Warning !!!
    '!!! Don't change anything !!!
    For i = 2 To lastRow
        Dim tenantPN As String, sourcePN As String
        Dim tenantSD As String, sourceSD As String
        Dim tenantLD As String, sourceLD As String
        Dim tenantPrimary As String, tenantSecondary As String
        Dim sourcePrimary As String, sourceSecondary As String
        Dim itemID As String
        Dim serial As Long
tenantPN = HighlightTerms(ws.Cells(i, tenantPNCol).Value, ws.Cells(i, sourcePNCol).Value)
sourcePN = HighlightTerms(ws.Cells(i, sourcePNCol).Value, ws.Cells(i, tenantPNCol).Value)

tenantSD = HighlightTerms(ws.Cells(i, tenantSDCol).Value, ws.Cells(i, sourceSDCol).Value)
sourceSD = HighlightTerms(ws.Cells(i, sourceSDCol).Value, ws.Cells(i, tenantSDCol).Value)

tenantLD = HighlightTerms(ws.Cells(i, tenantLDCol).Value, ws.Cells(i, sourceLDCol).Value)
sourceLD = HighlightTerms(ws.Cells(i, sourceLDCol).Value, ws.Cells(i, tenantLDCol).Value)

itemID = ws.Cells(i, itemIDCol).Value

                serial = ws.Cells(i, serialNum).Value

        tenantPrimary = Replace(Replace(Replace(ws.Cells(i, tenantPrimaryCol).Value, "[", ""), "]", ""), "'", "")
        sourcePrimary = Replace(Replace(Replace(ws.Cells(i, sourcePrimaryCol).Value, "[", ""), "]", ""), "'", "")
      
        Dim tpImages As String, tsImages As String
        Dim spImages As String, ssImages As String
        Dim urls() As String, j As Long

        ' Tenant Primary Images
        urls = Split(tenantPrimary, ",")
        tpImages = "<div style='display:flex; flex-wrap:wrap;'>"
        For j = LBound(urls) To UBound(urls)
            tpImages = tpImages & "<div style='flex:1 0 21%; margin:5px;'><img src='" & Trim(urls(j)) & "' style='height:300px; width:300px;' alt='Tenant Primary " & (j + 1) & "'></div>"
        Next j
        tpImages = tpImages & "</div>"

        ' Tenant Secondary Images
        urls = Split(tenantSecondary, ",")
        tsImages = "<div style='display:flex; flex-wrap:wrap;'>"
        For j = LBound(urls) To UBound(urls)
            tsImages = tsImages & "<div style='flex:1 0 21%; margin:5px;'><img src='" & Trim(urls(j)) & "' style='height:300px; width:300px;' alt='Tenant Secondary " & (j + 1) & "'></div>"
        Next j
        tsImages = tsImages & "</div>"

        ' Source Primary Images
        urls = Split(sourcePrimary, ",")
        spImages = "<div style='display:flex; flex-wrap:wrap;'>"
        For j = LBound(urls) To UBound(urls)
            spImages = spImages & "<div style='flex:1 0 21%; margin:5px;'><img src='" & Trim(urls(j)) & "' style='height:350px; width:350px;' alt='Source Primary " & (j + 1) & "'></div>"
        Next j
        spImages = spImages & "</div>"

        ' Source Secondary Images
        urls = Split(sourceSecondary, ",")
        ssImages = "<div style='display:flex; flex-wrap:wrap;'>"
        For j = LBound(urls) To UBound(urls)
            ssImages = ssImages & "<div style='flex:1 0 21%; margin:5px;'><img src='" & Trim(urls(j)) & "' style='height:350px; width:350px;' alt='Source Secondary " & (j + 1) & "'></div>"
        Next j
        ssImages = ssImages & "</div>"

        ' Create the HTML table
        Dim result As String
        result = "<table border='1' style='width:100%; border-collapse:collapse; table-layout:fixed; font-family:Poppins, sans-serif; font-weight:500; box-shadow: 0 4px 8px rgba(0, 0, 0, 0.1); border-radius: 8px; overflow: hidden; background-color:#2c2c2c; color:#e0e0e0;margin-bottom:20px;'>" & _
                 "<tr style='background-color:#1a1a1a; color:#ffffff; text-align:center;'>" & _
                 "<th colspan='2' style='padding:10px;'><h3>Row No: " & i & "</h3><h3>Serial No: " & serial & "</h3><h4>Item ID: " & itemID & "</h4></th></tr>" & _
                 "<tr style='background-color:#333333; text-align:center;'>" & _
                 "<th style='padding:5px;'>Tenant PN</th><th style='padding:5px;'>Source PN</th></tr>" & _
                 "<tr><td style='padding:5px;'>" & tenantPN & "</td><td style='padding:5px;'>" & sourcePN & "</td></tr>" & _
                 "<tr style='background-color:#333333; text-align:center;'>" & _
                 "<th style='padding:5px;'>Tenant SD</th><th style='padding:5px;'>Source SD</th></tr>" & _
                 "<tr><td style='padding:5px;'>" & tenantSD & "</td><td style='padding:5px;'>" & sourceSD & "</td></tr>" & _
                 "<tr style='background-color:#333333; text-align:center;'>" & _
                 "<th style='padding:5px;'>Tenant LD</th><th style='padding:5px;'>Source LD</th></tr>" & _
                 "<tr><td style='padding:5px;'>" & tenantLD & "</td><td style='padding:5px;'>" & sourceLD & "</td></tr>" & _
                 "<tr style='background-color:#1a1a1a; color:#ffffff; text-align:center;'>" & _
                 "<th style='padding:5px;'>Tnt Primary Images</th><th style='padding:5px;'>Src Primary Image</th></tr>" & _
                 "<th style='padding:10px;'>" & tpImages & "</th><th style='padding:10px;'>" & spImages & "</th></tr>" & _
                 "<tr style='background-color:#1a1a1a; color:#ffffff; text-align:center;'>" & _
                 "</table>"

        ' Output Column
        ws.Cells(i, "AJ").Value = result
    Next i
End Sub








Function HighlightTerms(tenantText As String, sourceText As String) As String
    Dim words() As Variant
    Dim word As Variant
    Dim highlightedText As String
    highlightedText = tenantText

    ' Split tenant text into words
    words = Split(tenantText, " ")

    ' Loop through each word and check if it exists in source text
    For Each word In words
        If InStr(1, sourceText, word, vbTextCompare) > 0 Then
            highlightedText = Replace(highlightedText, word, "<span style='background-color:yellow;'>" & word & "</span>", , , vbTextCompare)
        End If
    Next word

    HighlightTerms = highlightedText
End Function














