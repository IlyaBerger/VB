Attribute VB_Name = "Module2"
Sub Solve()

Utility = 0
Max = Cells(4, 2)

For Work = 0 To Max

Cells(6, 2) = Work

    For Investment = 0 To Max - Work
     
     Cells(6, 3) = Investment
     
        For Work2 = 0 To Max
           
           Cells(7, 2) = Work2
           
            
           If Cells(8, 19) > Utility Then
           Utility = Cells(8, 19)
           Cells(20, 19) = Work
           Cells(20, 20) = Investment
           Cells(20, 21) = Max - Work - Investment
           Cells(20, 22) = Work2
           End If
           
        Next
        

     
     If Cells(8, 19) > Utility Then
     Utility = Cells(8, 19)
     Cells(20, 19) = Work
     Cells(20, 20) = Investment
     Cells(20, 21) = Max - Work - Investment
     End If
     
     Next
     


If Cells(8, 19) > Utility Then
Utility = Cells(8, 19)
Cells(20, 19) = Work
Cells(20, 20) = Investment
Cells(20, 21) = Max - Work - Investment
End If

Next


Cells(6, 2) = Cells(20, 19)
Cells(6, 3) = Cells(20, 20)
Cells(7, 2) = Cells(20, 22)


End Sub
