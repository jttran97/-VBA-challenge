VBA Challenge Module 2 of Jennifer Tran

## Acknowledgments
Special thanks to Reza Abasaltian and Sunil Khambaita for their valuable assistance and guidance. 

' Greatest Value code - Sunil Khambaita 
ws.Range("Q2") = "%" & WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row)) * 100
            ws.Range("Q3") = "%" & WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row)) * 100
            ws.Range("Q4") = WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row))
            
            increase_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("K2:K" & Summary_Table_Row)), ws.Range("K2:K" & Summary_Table_Row), 0)
            decrease_number = WorksheetFunction.Match(WorksheetFunction.Min(ws.Range("K2:K" & Summary_Table_Row)), ws.Range("K2:K" & Summary_Table_Row), 0)
            volume_number = WorksheetFunction.Match(WorksheetFunction.Max(ws.Range("L2:L" & Summary_Table_Row)), ws.Range("L2:L" & Summary_Table_Row), 0)
            
            ws.Range("P2") = ws.Cells(increase_number + 1, 9)
            ws.Range("P3") = ws.Cells(decrease_number + 1, 9)
            ws.Range("P4") = ws.Cells(volume_number + 1, 9)


            
