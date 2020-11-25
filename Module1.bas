Attribute VB_Name = "Module1"
 Sub vba_stock()

'Attribute VB_Name = "Module5"

'1. Set Variables

'Loop through every worksheet and select the state contents.

'2. find lastrow each sheet
  
'3. print summary table headers
  
'4. set the cell location for the opening price value

    ' Loop through all all rows on the worksheet

'5.  Groups stock symbols by ticker symbol

    ' Set the location for the Stock Ticker value

    ' Set the location for the stock closing price value
   
    ' Calculated the value for the Percent Change
    
    
    ' Print the Stock Ticker in the Summary Table
    
    ' Add one to the summary table row

    ' Reset the Stock Total Volume
    
    ' If the cell immediately following a row is the same Stock.

'6. Identify and set Stock opening value

    '( 1 row aafter the previops symbol closing price (Cell (i+1 , 1))

    'Add to the Stock Total Volume

'7 Entering summary table cell values for

    'ticker symbol
    'yearly change
    'percent change

'8. added Conditional Format to Yearly Change column

    'Add to the Stock Total Volume

'9 --------------------------

    'Auto fit the summery table columns
    
    'format Currency and percentage columns

'1.--------------------------------------
    
            Dim Ws As Worksheet

    For Each Ws In Worksheets
    
            Dim Stock_Ticker As String
            
            Stock_Ticker = " "
            
            Dim Stock_Volume_Total As LongLong
            
            Stock_Volume_Total = 0
            
            Dim SOPrce  As Double 'Stock Opening Price
            
            SOPrce = 0
            
            Dim SCPrce As Double 'Stock Closing Price
            
            SCPrce = 0
            
            Dim Yearly_change As Double
            
            Yearly_change = 0
            
            Dim Percent_change As Double
            
            Pecent_Change = 0
            
            Dim Summary_Table_Row As Long
            
            Dim lastrow As Long
                    
            Dim i As Long
            
            Dim Percent_Change_calc As Double 'Calcuated Percent Change
            
            Percent_Change_calc = 0
        
'2--------------------------------------------------
        
            lastrow = Ws.Cells(Rows.Count, 1).End(xlUp).Row

'3---------------------------------------

            Ws.Cells(1, 10).Value = "Ticker"
        
            Ws.Cells(1, 11).Value = "Yearly Change"
             
            Ws.Cells(1, 12).Value = "Percent Change"
              
            Ws.Cells(1, 13).Value = "Total Stock Volume"
            
            Summary_Table_Row = 2
            
  

'4.---------------------

            SOPrce = Ws.Cells(2, 3).Value
    
    For i = 2 To lastrow
    

'5. -----------
  
   
    If Ws.Cells(i + 1, 1).Value <> Ws.Cells(i, 1).Value Then
  
            Stock_Ticker = Ws.Cells(i, 1).Value
            
            SCPrce = Ws.Cells(i, 6).Value
            
            Yearly_change = (SCPrce - SOPrce)

'6.-------------------

    If SOPrce <> 0 Then
        
            Percent_change = (Yearly_change / SOPrce)

    Else

            MsgBox ("For " & Stock_Ticker & ", Row " & CStr(i) & ": SOPrce =" & SOPrce & ". Fix <open> field manually and save the spreadsheet.")
            
    End If
      
           Stock_Volume_Total = Stock_Volume_Total + Ws.Cells(i, 7).Value
           
'7.-------------------

           Ws.Range("J" & Summary_Table_Row).Value = Stock_Ticker
   
           Ws.Range("K" & Summary_Table_Row).Value = Yearly_change
   
           Ws.Range("L" & Summary_Table_Row).Value = Percent_change

'8.------------------------------------
         
    If Yearly_change > 0 Then
             
            Ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
             
    ElseIf Yearly_change <= 0 Then
             
            Ws.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
        
    End If
        
            Ws.Range("M" & Summary_Table_Row).Value = Stock_Volume_Total
            
           'Ws.Range("N" & Summary_Table_Row).Value = SOPrce
              
           'Ws.Range("O" & Summary_Table_Row).Value = SCPrce
            
'8.------------------------------------
 
             Summary_Table_Row = Summary_Table_Row + 1
         
             Yearly_change = 0
         
             SCPrce = 0
    
             SOPrce = Ws.Cells(i + 1, 3).Value
    
             Percent_change = 0
       
             Stock_Volume_Total = 0

     Else

             Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

    End If

    Next i

'9.------------------------------------
           
           Ws.Columns("K").NumberFormat = "$#,##0.00"                                             'Style = "Currency"
           
           Ws.Columns("L").NumberFormat = "#0.00%"                             'Style = "Percent"
                      
           'Ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_change) & "0.00%")
           
           'Ws.Range("L" & Summary_Table_Row).Value = (CStr(Yearly_change) & "$0.00")
           
           Columns("J:M").AutoFit
     


    Next Ws

End Sub




