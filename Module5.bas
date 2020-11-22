Attribute VB_Name = "Module5"
 Sub vba_stock()


'1. Set Variables

'Loop through every worksheet and select the state contents.

'2. print summary table headers
  
'3. find lastrow each sheet
  
'4.   Set an initial variable for holding the Stock Ticker

    ' Set an initial variable for holding the total per Stock Ticker
    
    ' Keep track of the location for each Stock Ticker in the summary table
     
    ' Loop through all Stock Volume

    ' Check if we are still within the same Stock Ticker, if it is not...

'5. ' Set the Stock Ticker

    ' Add to the Stock Total Volume
    
    ' Print the Stock Ticker in the Summary Table
    
    ' Add one to the summary table row

    ' Reset the Stock Total Volume
    
    ' If the cell immediately following a row is the same Stock.

'6.  Add to the Stock Total Volume

'7  Formatting 'Autofit to display data
'1.--------------------------------------
    
    For Each WS In Worksheets
    
    Dim lastrow As Long
    
    Dim i As Long
    
    Dim WorksheetName As String
 
    Dim Stock_Ticker As String
    
    Dim SOPrce  As Double 'Stock Opening Price
    
    Dim SCPrce As Double 'Stock Closing Price
    
    Dim Yearly_change As Double
    
    Dim Percent_change As Double
        
    Dim Stock_Volume_Total As LongLong
    
    Dim Summary_Table_Row As Integer
       
    
    

    
    WorksheetName = WS.Name
 '2.-----------------

    WS.Cells(1, 10).Value = "Ticker"

    WS.Cells(1, 11).Value = "Yearly Change"
     
    WS.Cells(1, 12).Value = "Percent Change"
      
    WS.Cells(1, 13).Value = "Total Stock Volume"
    

'3.---------------------

    lastrow = WS.Cells(Rows.Count, 1).End(xlUp).Row

'4. -----------
  
  

  
  
  Summary_Table_Row = 2

  SOPrce = WS.Cells(2, 3).Value
  
  For i = 2 To lastrow
        

        
  
  If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
  
  SCPrce = WS.Cells(i, 6).Value
    
  Yearly_change = (SCPrce - SOPrce)
    
  Percent_change = (Yearly_change / SOPrce)

 '5.-------------------
      Stock_Ticker = WS.Cells(i, 1).Value

      Stock_Volume_Total = Stock_Volume_Total + WS.Cells(i, 7).Value

     
     
     WS.Range("J" & Summary_Table_Row).Value = Stock_Ticker

      
     WS.Range("K" & Summary_Table_Row).Value = SOPrce
      
      
      WS.Range("L" & Summary_Table_Row).Value = SCPrce
   
      
      WS.Range("M" & Summary_Table_Row).Value = Stock_Volume_Total
    
       WS.Range("N" & Summary_Table_Row).Value = SOPrce
              
       WS.Range("O" & Summary_Table_Row).Value = SCPrce
      
         
     If Yearly_change > 0 Then
             
        WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
             
     ElseIf Yearly_change <= 0 Then
             
        WS.Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
     
     Summary_Table_Row = Summary_Table_Row + 1
     
     SOPrce = WS.Cells(i + 1, 3).Value
     
     End If
     
     
     Stock_Volume_Total = 0
     
     Yearly_change = 0
    
     Pervcent_change = 0
     
     SCPrce = 0
     
     

'6.------------------------------------
      Stock_Volume_Total = Stock_Volume_Total + Cells(i, 7).Value

    End If
    
        

  Next i
  
  
'7.------------------------------------
        Columns("J:M").AutoFit
 
        WS.Range("L2:L" & lastrow).NumberFormat = "0.00%"
     
        WS.Range("K2:K" & lastrow).NumberFormat = "$0.00"
 
 Next WS
 
 

End Sub


