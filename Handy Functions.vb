Public Function lrowsetter(trg As Worksheet) As Long 'find lowest row in spreadsheet and returns it
 lrowsetter = trg.Cells.Find(What:="*", _
                            After:=Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Row
               
                            
End Function

Public Function lcolsetter(trg As Worksheet) As Long 'finds highest columm and returns it
 lcolsetter = trg.Cells.Find(What:="*", _
                            After:=Range("A1"), _
                            LookAt:=xlPart, _
                            LookIn:=xlFormulas, _
                            SearchOrder:=xlByColumns, _
                            SearchDirection:=xlPrevious, _
                            MatchCase:=False).Column
               
                            
End Function

Public Function lrowbycol(trg As Worksheet, col As Integer) As Long 'finds lowest row in a given column
lrowbycol = trg.Cells(Rows.Count, col).End(xlUp).Row
End Function






