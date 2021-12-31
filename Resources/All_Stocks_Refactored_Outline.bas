Attribute VB_Name = "Module7"
Sub AllStocksAnalysisRefactored()

    'Initialize variables for starting price and ending price
    
    'Input box

    'Start clock
   
    
    'Format the output sheet on All Stocks Analysis worksheet
 
    
    'Create a header row

    
    'Initialize array of all tickers

    
    'Activate data worksheet
    
    
    'Get number of rows to count
   
    

    '1a) Create a ticker Index and set it equal to zero
  
       
    '1b)Create Output arrays
      
     
    '2a)Initialize tickervolume equal zero
      
       
       'Activate sheet
       
       
    '2b)Loop through all rows to find if ticker matches
     

    
    '3a)Increase ticker volume and add ticker volume to current ticker
               
    
        'End if
        
    
    '3b)If this row is first row with current ticker then
          
        
        'Assign ticker starting price to corresponding variable


    '3c) If current row is last row to include ticker then
          
        
        'Assign ticker ending price to variable
              
        
        'Increase ticker index if next row doesn't match current row
        
        
        'End if
           
       
    
    '4) Loop through your arrays and output the Ticker, Total Daily Volume, and Return.
        

    
    'Format script
       
    
    'Stop clock


End Sub

