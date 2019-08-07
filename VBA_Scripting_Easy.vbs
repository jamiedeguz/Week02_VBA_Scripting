Sub RunOnAllWOrksheets()
    ' Runs the code on all worksheets at once
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call Test
    Next
    Application.ScreenUpdating = True
End Sub

Sub Test()

' Variable to set ticker as string
Dim ticker As String

' Set an initial variable for holding the total volume per ticker
Dim volume As Double
volume = 0

' Keep track of the location for each ticker in the summary table
Dim Summary_table_row As Integer
Summary_table_row = 2

Range("I1").Value = "Ticker"
Range("J1").Value = "Total Stock Volume"

' Find last row in a column
Dim lastRow As Double
lastRow = Cells(Rows.Count, 1).End(xlUp).Row

' Loop through all tickers
For i = 2 To lastRow

' Check if ticker is still within the same type
If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
    
        ' Set the ticker
        ticker = Cells(i, 1).Value
        
        ' Add to the ticker count
        volume = volume + Cells(i, 7).Value
        
        ' print the ticker type in the summary table
        Range("I" & Summary_table_row).Value = ticker
        
        ' Print the volume to the summary table window
        Range("J" & Summary_table_row).Value = volume
        
        'Add on to the summary table row
        Summary_table_row = Summary_table_row + 1
        
        'Reset the ticker type
        volume = 0
        
      ' If the cell immediately following a row is the same ticker type
      Else
        ' Add to the volume
        volume = volume + Cells(i, 7).Value
        
      End If
      
    Next i

End Sub

