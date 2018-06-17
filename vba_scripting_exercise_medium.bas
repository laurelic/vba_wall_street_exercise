Attribute VB_Name = "Module1"
Sub vba_scripting_exercise_medium()
'PURPOSE: To summarize the total volume for each stock based on yearly change, percent change, and volume
    'color code the changes according to positive or negative change.
'SOURCE: Lauren Creatura

Dim ws As Worksheet
Dim lrow As Long
Dim srow As Long
Dim c As Range, yChange As Range
Dim tList As Object
Dim pos As FormatCondition, neg As FormatCondition
Dim oValue As Variant
Dim cValue As Variant

Set tList = CreateObject("Scripting.Dictionary")

'loop through each worksheet in the workbook
For Each ws In ThisWorkbook.Worksheets
    With ws
        'determine the last row of data
        lrow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
        
        'create headings for the data summary
        .Range("I1").Value = "Ticker"
        .Range("J1").Value = "Yearly Change"
        .Range("K1").Value = "Percent Change"
        .Range("L1").Value = "Total Stock Volume"
        .Range("I1:L1").Font.Underline = True
        
        'loop through the ticker values and add unqiue values to a scripting dictionary
        For Each c In .Range("A2:A" & lrow)
            If Not tList.exists(c.Value) Then tList.Add c.Value, Nothing
        Next
        
        'transpose the keys so that a unique list of keys populates
        .Range("I2").Resize(tList.Count) = Application.Transpose(tList.keys)
        
        'find the last row of data in the "ticker" column
        srow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
        For i = 2 To srow
            'determine the opening value of the stock based on the lowest date value in the "date" column
            oValue = .Evaluate("INDEX(C2:C" & lrow & ",MATCH(I" & i & "&MINIFS(B2:B" & lrow & ",A2:A" & lrow & ",I" & i & "),A2:A" & lrow & "&B2:B" & lrow & ",0))")
            
            'determin the closing value of stock based on the maximum date value in the "date" column
            cValue = .Evaluate("INDEX(F2:F" & lrow & ",MATCH(I" & i & "&MAXIFS(B2:B" & lrow & ",A2:A" & lrow & ",I" & i & "),A2:A" & lrow & "&B2:B" & lrow & ",0))")
            
            'store the volume change between the opening and closing stock values in column J
            .Cells(i, 10).Value = cValue - oValue
            
            'store the percent change between opening and closing stock values in column k
            If oValue = 0 Then
                .Cells(i, 11).Value = ""
            Else
                .Cells(i, 11).Value = (cValue - oValue) / oValue
            End If
            
            'use the evaluate function to populate the yearly volume in column L
            .Cells(i, 12).Value = .Evaluate("SUMIF(A2:A" & lrow & ",I" & i & ",G2:G" & lrow & ")")
        Next i
            
        'adjust the number formatting for columns J and K
        .Range("K2:K" & srow).NumberFormat = "0.00%"
        
         'clear any existing conditional formatting
        Set yChange = .Range("J2:J" & srow)
        yChange.FormatConditions.Delete
        
        'define conditional formatting rules: green is positive, red it negative
        Set pos = yChange.FormatConditions.Add(xlCellValue, xlGreater, "0")
        pos.Interior.Color = vbGreen
        Set neg = yChange.FormatConditions.Add(xlCellValue, xlLess, "0")
        neg.Interior.Color = vbRed
        
        'clear the scripting dictionary to be used on the next sheet
        tList.RemoveAll
    End With
Next
    
End Sub
