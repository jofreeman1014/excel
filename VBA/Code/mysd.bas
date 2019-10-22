Attribute VB_Name = "Module1"
Sub betasort()
' Modifications:
'  - All the letters are on one sheet, so just keep rolling through it
'  - The sheets are years now instead of letter groups, but this code should work through it regardless

' ===== Initializing Variables =====
  Dim ticks As String ' ticker variable
  Dim vol As Double ' volume storage variable (Col 7/G)
  Dim step As Integer ' row number for totaled ticks/vol
  Dim cl As Integer ' column value for results
  Dim last As Double ' last row
  Dim sheets As Integer ' number of sheets
  Dim beg As Double ' the first vol yearly
  Dim fin As Double ' the last vol yearly
  
  beg = 0
  fin = 0
  cl = 9 ' Col 9/I
  sheets = ActiveWorkbook.Worksheets.Count ' number of sheets in doc
  'j = 2 ' For single-sheet testing
  For j = 1 To sheets
      ActiveWorkbook.Worksheets(j).Activate
      vol = 0 ' Reset vol every time
      step = 2 ' row under header
      last = Cells(Rows.Count, 1).End(xlUp).Row ' Final row code
      beg = Cells(2, 3) ' Every sheet starts at the 2nd row anyway
                  
    
      ' Add headers so users know what's what!
      Cells(1, cl) = "TICKER"
      Cells(1, cl + 1) = "YEARLY CHANGE"
      Cells(1, cl + 2) = "% CHANGE"
      Cells(1, cl + 3) = "TOTAL STOCK VOLUME"
    ' ===== Making it do what I need it to do =====
      For i = 2 To last
        If Cells(i + 1, 1).Value <> Cells(i, 1).Value Then
            ticks = Cells(i, 1).Value
            vol = vol + Cells(i, 7)
            
            fin = Cells(i, 6)
            Cells(step, cl + 1) = fin - beg 'calculating the difference
            If Not (beg = 0) Then
                Cells(step, cl + 2) = (fin - beg) / beg ' % diff
                Cells(step, cl + 2).Style = "Percent"
                If ((fin - beg) <= 0) Then
                    Cells(step, cl + 1).Interior.ColorIndex = 3
                Else
                    Cells(step, cl + 1).Interior.ColorIndex = 4
                End If
            End If
            ' Populate cells and move on
            Cells(step, cl) = ticks
            Cells(step, cl + 3) = vol
            step = step + 1
            vol = 0
            beg = Cells(i + 1, 3)
        Else
            vol = vol + Cells(i, 7)
            
        End If
      Next i
  Next j
End Sub


