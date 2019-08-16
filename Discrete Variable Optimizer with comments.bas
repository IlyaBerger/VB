Attribute VB_Name = "Module2"
Sub Solve()

'initialize the variables and set max to the entire time budget
Utility = 0
Max = Cells(4, 2)

'go through all possibilities for work in period 1
For Work = 0 To Max

Cells(6, 2) = Work

'go through all possibilities for investment
    For Investment = 0 To Max - Work
     
     Cells(6, 3) = Investment
     
'go through all possibilibites for work in period 2
        For Work2 = 0 To Max
           
           Cells(7, 2) = Work2
           
'see if the work2 variable yields higher utility and, if so, update the max values of work, investment and leisure

           If Cells(8, 19) > Utility Then
           Utility = Cells(8, 19)
           Cells(20, 19) = Work
           Cells(20, 20) = Investment
           Cells(20, 21) = Max - Work - Investment
           Cells(20, 22) = Work2
           End If
           
        Next
        
'see if the investment variable yields higher utility and, if so, update the max values of work, investment and leisure
     
     If Cells(8, 19) > Utility Then
     Utility = Cells(8, 19)
     Cells(20, 19) = Work
     Cells(20, 20) = Investment
     Cells(20, 21) = Max - Work - Investment
     End If
     
     Next
     
'see if the work 1 variable yields higher utility and, if so, update the max values of work, investment and leisure

If Cells(8, 19) > Utility Then
Utility = Cells(8, 19)
Cells(20, 19) = Work
Cells(20, 20) = Investment
Cells(20, 21) = Max - Work - Investment
End If

Next

'make a record of the results

Cells(6, 2) = Cells(20, 19)
Cells(6, 3) = Cells(20, 20)
Cells(7, 2) = Cells(20, 22)


End Sub
