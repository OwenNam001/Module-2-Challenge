Attribute VB_Name = "Module1"
Sub Stock_Analyze()

    Dim ws As Worksheet
    Dim WorksheetName As String
    Dim ticker As String
    Dim ticker_next As String
    Dim ticker_list As String
    Dim ticker_list_row As Double
    Dim lastRow As Double
    Dim stock_volume As Double
    Dim open_date As String
    Dim close_date As String
    Dim open_price As Double
    Dim close_price As Double
    Dim greatPerIncreaseVal As Double
    Dim greatPerDecreaseVal As Double
    Dim greatTotalVolume As Double
    Dim greatPerIncreaseTicker As String
    Dim greatPerDecreaseTicker As String
    Dim greatTotalVolumeTicker As String
        
    
    ' << ======================================= Looping Across Worksheet ================================================>
    For Each ws In ThisWorkbook.Worksheets
    
        'last row number of each sheet
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        'variable initialised in each sheet
        ticker = " "
        ticker_list = " "
        ticket_next = " "
        stock_volume = 0
        open_date = "99991231"
        close_date = "19000101"
        open_price = 0
        close_price = 0
        
        
        ' << ======================================= Retreival of Data ================================================>
        'unique ticker list row number in each sheet
        ticker_list_row = 1
        
        'set summary table header
        ws.Cells(ticker_list_row, "I") = "Ticker"
        ws.Cells(ticker_list_row, "J") = "Yearly Change"
        ws.Cells(ticker_list_row, "K") = "Percent Change"
        ws.Cells(ticker_list_row, "L") = "Total Stock Volume"

        'loops through the data in the sheet from row 2
        For i = 2 To lastRow
        
            ticker = ws.Cells(i, "A")           'current ticker assignment
            ticker_next = ws.Cells(i + 1, "A")  'next row ticker assignment
            
            'min date, open_price on min date
            If ws.Cells(i, "B") < open_date Then
               open_date = ws.Cells(i, "B")
               open_price = ws.Cells(i, "C")
            End If
            
            'max date, close_price on max date
            If ws.Cells(i, "B") > close_date Then
               close_date = ws.Cells(i, "B")
               close_price = ws.Cells(i, "F")
            End If
            
            If ticker = ticker_next Then                       'when ticker symbol = next ticket symbol
               stock_volume = stock_volume + ws.Cells(i, "G")  'add up stock volume
            Else                                               'when ticker symbol <> next ticket symbol
               
               ticker_list_row = ticker_list_row + 1           'summary table row
               
               ticker_list = ticker                            'get ticker symbol
               stock_volume = stock_volume + ws.Cells(i, "G")  'add up stock volume
               
               ws.Cells(ticker_list_row, "I") = ticker_list                             'Ticker symbol
               ws.Cells(ticker_list_row, "J") = close_price - open_price                'Yearly Change
               ws.Cells(ticker_list_row, "K") = (close_price - open_price) / open_price 'Percent Change
               ws.Cells(ticker_list_row, "L") = stock_volume                            'Total Stock Volume
                
               ws.Cells(ticker_list_row, "J").NumberFormat = "0.00" 'Yearly change Column J change to .00 format
               ws.Cells(ticker_list_row, "K").NumberFormat = "0.00%" 'Percent change Column J change to percent format
               
               'Yearly Change >= 0 green, Yearly Change < 0 green red
               If ws.Cells(ticker_list_row, "J") >= 0 Then
                  ws.Cells(ticker_list_row, "J").Interior.ColorIndex = 4
               Else
                  ws.Cells(ticker_list_row, "J").Interior.ColorIndex = 3
               End If
                
               stock_volume = 0 'when next ticker is different, stock_volume initialised
               open_price = 0   'when next ticker is different, stock_volume initialised
               open_date = "99991231"  'when next ticker is different, open_date initialised
               close_date = "19000101" 'when next ticker is different, close_date initialised
            End If
        Next i
        ' << ======================================= Retreival of Data ================================================>
        
        
        ' << ======================================= Column Creation & Conditional Formatting ================================================>
        'set Greatest summar table header
        ws.Cells(1, "P") = "Ticker"
        ws.Cells(1, "Q") = "Value"
        
        'Great % Increase initialised
        greatPerIncreaseVal = -1
        greatPerDecreaseVal = 1
        greatTotalVolume = 0
        
        
        For i = 2 To ticker_list_row
        
            'Get greatest Increase
            If ws.Cells(i, "K").Value >= greatPerIncreaseVal Then
               greatPerIncreaseVal = ws.Cells(i, "K").Value
               greatPerIncreaseTicker = ws.Cells(i, "I").Value
            End If
            
            'Get greatest Decrease
            If ws.Cells(i, "K").Value < greatPerDecreaseVal Then
               greatPerDecreaseVal = ws.Cells(i, "K").Value
               greatPerDecreaseTicker = ws.Cells(i, "I").Value
            End If
            
            'Get greatest Total Volume
            If ws.Cells(i, "L").Value >= greatTotalVolume Then
               greatTotalVolume = ws.Cells(i, "L").Value
               greatTotalVolumeTicker = ws.Cells(i, "I").Value
            End If
            
        Next i
        ' << ======================================= Column Creation & Conditional Formatting ================================================>
        
        
        ' << ======================================= Block Arrow ================================================>
        ' Add a block arrow shape
        Dim arrowShape As Shape
        Dim startCell As Range
        Dim endCell As Range
        Dim arrowWidth As Double
        Dim arrowHeight As Double
        
        Set startCell = ws.Range("M7")
        Set endCell = ws.Range("N6")
        
        ' Calculate width and height for the arrow
        ' Ensure width and height are positive by taking the absolute value
        arrowWidth = Abs((endCell.Left + endCell.Width / 2) - (startCell.Left + startCell.Width / 2))
        arrowHeight = Abs((endCell.Top + endCell.Height / 2) - (startCell.Top + startCell.Height / 2))
        
        ' Add the arrow shape to the worksheet
        Set arrowShape = ws.Shapes.AddShape(msoShapeRightArrow, startCell.Left + startCell.Width / 2, startCell.Top + startCell.Height / 2, arrowWidth, arrowHeight)
                    
        arrowShape.Fill.ForeColor.RGB = RGB(255, 0, 0) 'Red Color
        arrowShape.Fill.Solid
        arrowShape.Line.Visible = msoFalse
        
        arrowShape.Rotation = -45
        ' << ======================================= Block Arrow ================================================>

        
        ' << ======================================= Calculated Values ================================================>
        'Greatest % Increase
        ws.Cells(2, "O") = "Greatest % Increase"
        ws.Cells(2, "P") = greatPerIncreaseTicker
        ws.Cells(2, "Q") = greatPerIncreaseVal
        ws.Cells(2, "Q").NumberFormat = "0.00%" 'Greatest % Increase Column Q change to percent format
        
        'Greatest % Decrease
        ws.Cells(3, "O") = "Greatest % Decrease"
        ws.Cells(3, "P") = greatPerDecreaseTicker
        ws.Cells(3, "Q") = greatPerDecreaseVal
        ws.Cells(3, "Q").NumberFormat = "0.00%" 'Greatest % Decrease Column Q change to percent format
        
        'Greatest Total Volume
        ws.Cells(4, "O") = "Greatest Total Volume"
        ws.Cells(4, "P") = greatTotalVolumeTicker
        ws.Cells(4, "Q") = greatTotalVolume
        ws.Cells(4, "Q").NumberFormat = "0.00E+00" 'Greatest Total Volume Column Q change to Scientific notation
        ' << ======================================= Calculated Values ================================================>
        
        
        ' Autofit to display data
        ws.Columns("A:Q").AutoFit
    
    Next ws
    ' << ======================================= Looping Across Worksheet ================================================>

End Sub

Sub Clear_Data()

For Each ws In ThisWorkbook.Sheets
    Dim lastRow As Double
    
    lastRow = ws.Cells(ws.Rows.Count, "I").End(xlUp).Row
    
    For i = 1 To lastRow
        ws.Cells(i, "I").Clear
        ws.Cells(i, "J").Clear
        ws.Cells(i, "K").Clear
        ws.Cells(i, "L").Clear
        ws.Cells(i, "M").Clear
        ws.Cells(i, "L").Clear
        ws.Cells(i, "O").Clear
        ws.Cells(i, "P").Clear
        ws.Cells(i, "Q").Clear
    Next i
      
    Dim sh As Shape

    For Each sh In ws.Shapes
        If sh.AutoShapeType = msoShapeRightArrow Then
        sh.Delete
        End If
    Next sh

Next ws

End Sub


