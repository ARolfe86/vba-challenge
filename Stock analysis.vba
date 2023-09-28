Sub Ticker_Anaylsis()
    Const DATA_COL_TICKER As Integer = 1
    Const DATA_COL_OPEN As Integer = 3
    Const DATA_COL_CLOSE As Integer = 6
    Const DATA_COL_VOLUME As Integer = 7
    
    Const OUTPUT_COL_TICKER As Integer = 9
    Const OUTPUT_COL_YEARLY_CHANGE As Integer = 10
    Const OUTPUT_COL_PERCENT_CHANGE As Integer = 11
    Const OUTPUT_COL_TOTAL_VOLUME As Integer = 12
    
    Const OUTPUT_TABLE_COL_LABELS As Integer = 15
    Const OUTPUT_TABLE_COL_TICKER As Integer = 16
    Const OUTPUT_TABLE_COL_VALUE As Integer = 17
    
    Const OUTPUT_TABLE_ROW_INCREASE As Integer = 2
    Const OUTPUT_TABLE_ROW_DECREASE As Integer = 3
    Const OUTPUT_TABLE_ROW_TOTAL_VOLUME As Integer = 4
    
    Const HEADER_ROW As Integer = 1
    
    
    Dim sheet As Worksheet
    Dim data_row As Double
    Dim output_row As Integer
    Dim total_volume As Double
    Dim yearly_open As Double
    Dim greatest_increase As Double
    Dim greatest_increase_ticker As String
    Dim greatest_decrease As Double
    Dim greatest_decrease_ticker As String
    Dim greatest_total_volume As Double
    Dim greatest_total_volume_ticker As String
    
    Dim format_range As Range
    Dim conditon1 As FormatCondition
    
    
    For Each sheet In ThisWorkbook.Worksheets
    
        sheet.Cells(HEADER_ROW, OUTPUT_COL_TICKER).Value = "Ticker"
        sheet.Cells(HEADER_ROW, OUTPUT_COL_YEARLY_CHANGE).Value = "Yearly Change"
        sheet.Cells(HEADER_ROW, OUTPUT_COL_PERCENT_CHANGE).Value = "Percent Change"
        sheet.Cells(HEADER_ROW, OUTPUT_COL_TOTAL_VOLUME).Value = "Total Stock Volume"
        
        sheet.Cells(HEADER_ROW, OUTPUT_TABLE_COL_TICKER).Value = "Ticker"
        sheet.Cells(HEADER_ROW, OUTPUT_TABLE_COL_VALUE).Value = "Value"
        
        sheet.Cells(OUTPUT_TABLE_ROW_INCREASE, OUTPUT_TABLE_COL_LABELS).Value = "Greatest % Increase"
        sheet.Cells(OUTPUT_TABLE_ROW_DECREASE, OUTPUT_TABLE_COL_LABELS).Value = "Greatest % Decrease"
        sheet.Cells(OUTPUT_TABLE_ROW_TOTAL_VOLUME, OUTPUT_TABLE_COL_LABELS).Value = "Greatest Total Volume"
    
        
        data_row = 2
        output_row = 2
        total_volume = 0
        greatest_increase = 0
        greatest_decrease = 0
        greatest_total_volume = 0
        greatest_increase_ticker = ""
        greatest_decrease_ticker = ""
        greatest_total_volume_ticker = ""
        
        yearly_open = sheet.Cells(data_row, DATA_COL_OPEN).Value
        
        Do Until IsEmpty(sheet.Cells(data_row, DATA_COL_TICKER).Value)
        
            total_volume = total_volume + sheet.Cells(data_row, DATA_COL_VOLUME).Value
            
            If sheet.Cells(data_row, DATA_COL_TICKER).Value <> sheet.Cells(data_row + 1, DATA_COL_TICKER).Value Then
                sheet.Cells(output_row, OUTPUT_COL_TICKER).Value = sheet.Cells(data_row, DATA_COL_TICKER).Value
                
                yearly_change = sheet.Cells(data_row, DATA_COL_CLOSE).Value - yearly_open
                sheet.Cells(output_row, OUTPUT_COL_YEARLY_CHANGE).Value = yearly_change
                
                percent_change = yearly_change / yearly_open
                sheet.Cells(output_row, OUTPUT_COL_PERCENT_CHANGE).Value = percent_change
                
                sheet.Cells(output_row, OUTPUT_COL_TOTAL_VOLUME).Value = total_volume
                
                output_row = output_row + 1
                
                If percent_change > greatest_increase Then
                    greatest_increase = percent_change
                    greatest_increase_ticker = sheet.Cells(data_row, DATA_COL_TICKER).Value
                End If
                
                If percent_change < greatest_decrease Then
                    greatest_decrease = percent_change
                    greatest_decrease_ticker = sheet.Cells(data_row, DATA_COL_TICKER).Value
                End If
                
                If total_volume > greatest_total_volume Then
                    greatest_total_volume = total_volume
                    greatest_total_volume_ticker = sheet.Cells(data_row, DATA_COL_TICKER).Value
                End If
                
                total_volume = 0
                yearly_open = sheet.Cells(data_row + 1, DATA_COL_OPEN).Value
                
            End If
            
            data_row = data_row + 1
            
        Loop
           
        sheet.Cells(OUTPUT_TABLE_ROW_INCREASE, OUTPUT_TABLE_COL_TICKER).Value = greatest_increase_ticker
        sheet.Cells(OUTPUT_TABLE_ROW_INCREASE, OUTPUT_TABLE_COL_VALUE).Value = greatest_increase
        sheet.Cells(OUTPUT_TABLE_ROW_DECREASE, OUTPUT_TABLE_COL_TICKER).Value = greatest_decrease_ticker
        sheet.Cells(OUTPUT_TABLE_ROW_DECREASE, OUTPUT_TABLE_COL_VALUE).Value = greatest_decrease
        sheet.Cells(OUTPUT_TABLE_ROW_TOTAL_VOLUME, OUTPUT_TABLE_COL_TICKER).Value = greatest_total_volume_ticker
        sheet.Cells(OUTPUT_TABLE_ROW_TOTAL_VOLUME, OUTPUT_TABLE_COL_VALUE).Value = greatest_total_volume
               
        
        Set format_range = sheet.Range("J2:K" & output_row)
        format_range.FormatConditions.Delete
        Set condition1 = format_range.FormatConditions.Add(xlCellValue, xlGreater, "=0")
        With condition1
            .Interior.ColorIndex = 4
        End With
            
        Set condition1 = format_range.FormatConditions.Add(xlCellValue, xlLess, "=0")
        With condition1
            .Interior.ColorIndex = 3
        End With
        
        sheet.Columns("K").NumberFormat = "0.00%"
        sheet.Range("Q2:Q3").NumberFormat = "0.00%"
        sheet.Columns("I:Q").AutoFit
    Next sheet

  
    

End Sub