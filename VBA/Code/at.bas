Attribute VB_Name = "Module1"
Sub alphasort()
'1: This should go through all the lines in the table, and sum up stock volume
'     per each ticker value. I should also assume that I don't know the final
'     row of the data as well!
'2: When it finds a new ticker, this should add the sum to a row on the sheet,
'     and print the ticker value next to it, then move to the next unique ticker.
'3: Finally, try to add code to make it automatically move to the next sheet in
'     this document. Good practice for the larger code!

' Notes - I'm assuming all the tables have the same number of colmuns, and are sorted.
'    - Code runs well, can't compare to results on the gitlab account since these are 2016.
'    - Make it automatically go through all the sheets, so find a way to access and switch
'         between the sheets, and rerun this working code.


' ===== Initializing Variables =====
  Dim ticks As String ' ticker variable
  Dim vol As Double ' volume storage variable (Col 7/G)
  Dim step As Integer ' row number for totaled ticks/vol
  Dim cl As Integer ' column value for results
  Dim last As Double ' last row
  Dim sheets As Integer ' number of sheets
  
  cl = 9 ' Col 9/I
  step = 2 ' row under header
  sheets = ActiveWorkbook.Worksheets.Count ' number of sheets in doc
  
  For j = 1 To sheets
      step = 2
      vol = 0 ' reset vol every time
      ActiveWorkbook.Worksheets(j).Activate
      last = Cells(Rows.Count, 1).End(xlUp).Row ' final row code
      ' Add headers so users know what's what!
      Cells(1, cl) = "TICKER"
      Cells(1, cl + 1) = "TOTAL STOCK VOLUME"
  
    ' ===== Making it do what I need it to do =====
      For i = 2 To last
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticks = Cells(i, 1).Value
            vol = vol + Cells(i, 7)
            Cells(step, cl) = ticks
            Cells(step, cl + 1) = vol
            step = step + 1
            vol = 0
        Else
            vol = vol + Cells(i, 7)
        End If
      Next i
  Next j
End Sub

