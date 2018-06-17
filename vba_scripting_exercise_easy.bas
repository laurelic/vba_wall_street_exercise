Attribute VB_Name = "Module1"
Sub easy_vba()
'PURPOSE: To summarize the total volume for each stock on the sheet for the given year
'SOURCE: Lauren Creatura

Dim ws As Worksheet
Dim lrow As Long
Dim c As Range
Dim tList As Object

Set tList = CreateObject("Scripting.Dictionary")

'loop through each worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets
    With ws
        'determine the last row of data
        lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'create headings for the data summary
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Total Stock Volume"
        .Range("I1:J1").Font.Underline = True
        
        'loop through the ticker values and add unqiue values to a scripting dictionary
        For Each c In .Range("A2:A" & lrow)
            If Not tList.exists(c.Value) Then tList.Add c.Value, Nothing
        Next
        
        'transpose the keys so that a unique list of keys populates
        .Range("I2").Resize(tList.Count) = Application.Transpose(tList.keys)
        
        'find the last row of data in the "ticker" column
        lrow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        
        'use sumif formulas to populate the yearly volume next to each ticker
        .Range("J2:J" & lrow).FormulaR1C1 = "=SUMIF(C1,RC[-1],C7)"
        
        'clear the scripting dictionary to be used in the next loop
        tList.RemoveAll
    End With
Next
    
End Sub
