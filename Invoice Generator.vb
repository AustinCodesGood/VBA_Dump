Sub RegionAndCompanyInvoicer() 'YOU CHANGED REGION COUNTER #6 FROM ALLTECK TO POTELCO. YOU CHANGED REGION COUNTER TO ONLY COUNT TO 6
Dim sWkbName As String: sWkbName = "C:\Austin_Kimes\CostingTemplate\TransCosting.xlsx"
Dim sSubBkName As String: sSubBkName = "C:\Austin_Kimes\CostingTemplate\TransCostingSubStations.xlsx"
Dim wkb As Workbook
Dim sht As Worksheet
Dim sRegion As String
Dim d
Dim d2
Dim StartDate As String
Dim EndDate As String
Dim rng As Range
Dim conn As New ADODB.Connection
Dim rstnew As New ADODB.Recordset
Dim ra As Range
Dim LaborCreditCount As Integer
Dim CreditSum As Double
Dim CreditRange As Range
Application.ScreenUpdating = False
Application.EnableEvents = False
Application.DisplayAlerts = False
Dim clearrange As Range
Dim YardCounter As Integer
Dim CrewCreditCost As Double

Set d2 = CreateObject("Scripting.Dictionary")
d2.Add 0, "Circuit Repair & Maintenance"
d2.Add 1, "Microsite Support"
d2.Add 2, "Program Support"
d2.Add 3, "Circuit Inspectors"
d2.Add 4, "Inspectors"
d2.Add 5, "Substation"

conn.Open "Provider=SQLOLEDB; Data Source=WIN-A3SISNE4TPB; Initial Catalog=ExcelDemo; User ID=Austin_Kimes; Password=dupa00"

rstnew.Open "select count (*) as count from transmissionregion where active = 1", conn
YardCounter = rstnew.Fields("count").Value
rstnew.Close



' you can use this to make the daterange input be a clickable calendar http://www.fontstuff.com/vba/vbatut07.htm
StartDater:
StartDate = InputBox("Please Input StartDate for Costing Document.", "Start Date", Format(Now(), "dd/mm/yy"))
If Not IsDate(StartDate) Then
    MsgBox "Wrong Date Format"
    GoTo StartDater:
End If
If Not Weekday(CDate(StartDate), 2) = 1 Then
    MsgBox "StartDate must be a Monday, you entered a " & WeekdayName(Weekday(StartDate, 2), False, 2)
    GoTo StartDater:
End If

EndDater:
EndDate = DateAdd("d", 6, CDate(StartDate))

    
For RegionCounter = 1 To 14
    If RegionCounter = 1 Then
    sRegion = "Central Coast"
        ElseIf RegionCounter = 2 Then sRegion = "Central Valley"
        ElseIf RegionCounter = 3 Then sRegion = "North Valley"
        ElseIf RegionCounter = 4 Then sRegion = "North Coast"
        ElseIf RegionCounter = 5 Then sRegion = "Master"
        ElseIf RegionCounter = 6 Then sRegion = "Allteck"
        ElseIf RegionCounter = 7 Then sRegion = "LongFellow"
        ElseIf RegionCounter = 8 Then sRegion = "JW Didado"
        ElseIf RegionCounter = 9 Then sRegion = "MJ Electric"
        ElseIf RegionCounter = 10 Then sRegion = "Par Electric"
        ElseIf RegionCounter = 11 Then sRegion = "Potelco"
        ElseIf RegionCounter = 12 Then sRegion = "Pro Energy Services Group"
        ElseIf RegionCounter = 13 Then sRegion = "Summit"
        ElseIf RegionCounter = 14 Then sRegion = "SubStations"
    End If

    If RegionCounter = 1 Then 'this is how we name our pages. This creates a dictionary object and then changes the 1st entry to match the region every other iteration.
        Set d = CreateObject("Scripting.Dictionary")
        d.Add 1, sRegion
        d.Add 2, "Summary"
        d.Add 3, "Crew Rates"
        d.Add 4, "Crew Equipment"
        d.Add 5, "Support_Inspection Summary"
        d.Add 6, "Microsite Support"
        d.Add 7, "Program Support"
        d.Add 8, "Circuit Inspections"
        d.Add 9, "Inspectors"
        d.Add 10, "Substations"
        d.Add 11, "R&M"
        d.Add 12, "MS"
        d.Add 13, "PS"
        d.Add 14, "CL"
        d.Add 15, "INS"
        d.Add 16, "Crew Labor Overage Breakdown"
        d.Add 17, "Crew EQ Overage Breakdown"
        d.Add 18, "PM Cost Summary"
    Else: d(1) = sRegion
    End If
    

     
    
If RegionCounter < 14 Then 'starts the loop for main documents. 14 makes it the substation document.
    Set wkb = Workbooks.Add(sWkbName)

'counter = page number.
     For counter = 1 To 17
        Set sht = wkb.Worksheets(counter)
        sht.Name = d(counter)
        If counter = 1 Then
            sht.Range("E1").Value = sRegion
            sht.Range("B5").Value = StartDate
            ElseIf counter = 5 Then sht.Range("A3").Value = sRegion
        Else: sht.Range("A1").Value = sRegion
        End If
        If counter = 1 Then 'this is the cover page block. Usess the covercount to set coverpage values based on yard and region. Uses RegionCounter to make changes based on if Document is Master or Regional.
            For covercount = 0 To YardCounter '<---------------------------------------------------------------------------this value needs to be changed if more yards are added.
                If covercount = 0 And RegionCounter < 5 Then
                    rstnew.Open "exec TransCrewCover @startdate = '" & StartDate & "', @region = '" & sRegion & "', @yard = '%'", conn
                    sht.Range(sht.Cells(8, 2), sht.Cells(8, 2)).CopyFromRecordset rstnew
                    rstnew.Close
                    rstnew.Open "exec TransEquipCover @startdate = '" & StartDate & "', @region = '" & sRegion & "', @yard = '%'", conn
                    sht.Range(sht.Cells(16, 2), sht.Cells(16, 2)).CopyFromRecordset rstnew
                    rstnew.Close
                ElseIf covercount = 0 And RegionCounter = 5 Then
                    rstnew.Open "exec TransCrewCover @startdate = '" & StartDate & "', @region = '%', @yard = '%'", conn
                    sht.Range(sht.Cells(8, 2), sht.Cells(8, 2)).CopyFromRecordset rstnew
                    rstnew.Close
                    rstnew.Open "exec TransEquipCover @startdate = '" & StartDate & "', @region = '%', @yard = '%'", conn
                    sht.Range(sht.Cells(16, 2), sht.Cells(16, 2)).CopyFromRecordset rstnew
                    rstnew.Close
                ElseIf covercount = 0 And RegionCounter > 5 Then
                    rstnew.Open "exec TransCrewCover @startdate = '" & StartDate & "', @region = '%', @yard = '%', @company = '" & sRegion & "'", conn
                    sht.Range(sht.Cells(8, 2), sht.Cells(8, 2)).CopyFromRecordset rstnew
                    rstnew.Close
                    rstnew.Open "exec TransEquipCover @startdate = '" & StartDate & "', @region = '%', @yard = '%', @company = '" & sRegion & "'", conn
                    sht.Range(sht.Cells(16, 2), sht.Cells(16, 2)).CopyFromRecordset rstnew
                    rstnew.Close
                ElseIf covercount > 0 And RegionCounter < 5 Then rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.YardFinder('" & sRegion & "'," & covercount & ", '%'))set @str = 'exec TransCrewCover @startdate = ''" & StartDate & "'', @region = ''" & sRegion & "'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                ElseIf covercount > 0 And RegionCounter = 5 Then rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.YardFinder('%'," & covercount & ", '%'))set @str = 'exec TransCrewCover @startdate = ''" & StartDate & "'', @region = ''%'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                ElseIf covercount > 0 And RegionCounter > 5 Then rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.YardFinder('%'," & covercount & ", '%'))set @str = 'exec TransCrewCover @startdate = ''" & StartDate & "'', @region = ''%'', @yard = ' + '''' + @yarder + '''' + ',@company = ''" & sRegion & "''' exec (@str)", conn
                End If
                If covercount > 0 And rstnew.State = 1 Then
                    sht.Range(sht.Cells(8 + (19 * covercount), 2), sht.Cells(8 + (19 * covercount), 2)).CopyFromRecordset rstnew
                    rstnew.Close
                    If RegionCounter < 5 Then
                        rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.yardfinder('" & sRegion & "'," & covercount & ", '%'))set @str = 'exec TransEquipCover @startdate = ''" & StartDate & "'', @region = ''" & sRegion & "'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                    ElseIf RegionCounter = 5 Then
                        rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.yardfinder('%'," & covercount & ", '%'))set @str = 'exec TransEquipCover @startdate = ''" & StartDate & "'', @region = ''%'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                    ElseIf RegionCounter > 5 Then
                        rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.yardfinder('%'," & covercount & ", '" & sRegion & "'))set @str = 'exec TransEquipCover @startdate = ''" & StartDate & "'', @region = ''%'', @yard = ' + '''' + @yarder + '''' + ',@company = ''" & sRegion & "''' exec (@str)", conn
                    End If
                    If rstnew.State = 1 Then
                    sht.Range(sht.Cells(16 + (19 * covercount), 2), sht.Cells(16 + (19 * covercount), 2)).CopyFromRecordset rstnew
                    rstnew.Close
                    End If
                    If RegionCounter < 5 Then
                        rstnew.Open "select dbo.yardfinder('" & sRegion & "'," & covercount & ", '%')"
                        sht.Range(sht.Cells(25 + (19 * (covercount - 1)), 2), sht.Cells(25 + (19 * (covercount - 1)), 2)).CopyFromRecordset rstnew
                        rstnew.Close
                    ElseIf RegionCounter = 5 Then
                        rstnew.Open "select dbo.yardfinder('%'," & covercount & ", '%')"
                        sht.Range(sht.Cells(25 + (19 * (covercount - 1)), 2), sht.Cells(25 + (19 * (covercount - 1)), 2)).CopyFromRecordset rstnew
                        rstnew.Close
                    ElseIf RegionCounter > 5 Then
                        rstnew.Open "select dbo.yardfinder('%'," & covercount & ", '" & sRegion & "')"
                        If rstnew.State = 1 Then
                        sht.Range(sht.Cells(25 + (19 * (covercount - 1)), 2), sht.Cells(25 + (19 * (covercount - 1)), 2)).CopyFromRecordset rstnew
                        rstnew.Close
                        End If
                    End If
                    If RegionCounter < 5 Then
                        rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.yardfinder('" & sRegion & "'," & covercount & ", '%')) set @str = 'exec CrewCounter @startdate = ''" & StartDate & "'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                    ElseIf RegionCounter = 5 Then
                        rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.yardfinder('%'," & covercount & ", '%')) set @str = 'exec CrewCounter @startdate = ''" & StartDate & "'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                    ElseIf RegionCounter > 5 Then
                        rstnew.Open "SET NOCOUNT ON SET ANSI_WARNINGS OFF declare @str varchar(max) declare @yarder varchar(max) = (select dbo.yardfinder('%'," & covercount & ", '" & sRegion & "')) set @str = 'exec CrewCounter @startdate = ''" & StartDate & "'', @yard = ' + '''' + @yarder + ''''exec (@str)", conn
                    End If
                    If rstnew.State = 1 Then
                        sht.Range("H1").CopyFromRecordset rstnew
                        rstnew.Close
                    End If
                    sht.Cells(25 + (19 * (covercount - 1)), 4) = sht.Range("H1")
                    sht.Cells(25 + (19 * (covercount - 1)), 10) = sht.Range("I1")
                    sht.Cells(25 + (19 * (covercount - 1)), 16) = sht.Range("J1")
                    sht.Cells(25 + (19 * (covercount - 1)), 22) = sht.Range("K1")
                    sht.Cells(25 + (19 * (covercount - 1)), 28) = sht.Range("L1")
                    sht.Cells(25 + (19 * (covercount - 1)), 34) = sht.Range("M1")
                    sht.Cells(25 + (19 * (covercount - 1)), 40) = sht.Range("N1")
                    sht.Range(sht.Cells(1, 8), sht.Cells(1, 14)).Value = ""
                    If sht.Cells(25 + (19 * (covercount - 1)), 2).Value = "" Then
                        Set ra = sht.Range(sht.Cells(23 + (19 * (covercount - 1)), 1), sht.Cells(194, 1))
                        ra.EntireRow.Hidden = True
                        Exit For
                    End If
                End If
            Next covercount
        End If
        
            
        If counter = 2 Then
            Set sht = wkb.Worksheets(counter)
            sht.Range("B1").Value = "Billing Summary - " & StartDate & " to " & EndDate
            sht.Range("B26").Value = "Total Billing Summary - " & StartDate & " to " & EndDate
        ElseIf counter = 3 Then
            Set sht = wkb.Worksheets(counter)
            If RegionCounter < 5 Then rstnew.Open "exec TransCrewCosting '" & StartDate & "','" & sRegion & "'", conn
            If RegionCounter = 5 Then rstnew.Open "exec TransCrewCosting '" & StartDate & "','%'", conn
            If RegionCounter > 5 Then rstnew.Open "exec TransCrewCosting '" & StartDate & "','%','" & sRegion & "'", conn
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
            End If
        'If regioncounter < 5 Then
            'Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(lrowsetter(sht), 1))
            'rng.AutoFilter Field:=1, Criteria1:=sRegion, VisibleDropDown:=False
        'End If
        ElseIf counter = 4 Then
            If RegionCounter = 5 Then rstnew.Open "exec TransEQCredit @startdate = '" & StartDate & "'", conn
            If RegionCounter < 5 Then rstnew.Open "exec TransEQCredit '" & StartDate & "','" & sRegion & "'", conn
            If RegionCounter > 5 Then rstnew.Open "exec TransEQCredit '" & StartDate & "', '%', '" & sRegion & "'", conn
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
                Set clearrange = sht.Range(sht.Cells(lrowsetter(sht) + 1, 9).Address)
            End If
            If RegionCounter < 5 Then rstnew.Open "execute CrewEQCreditCompanyDayBreakDown '" & StartDate & "', '" & sRegion & "'", conn
            If RegionCounter = 5 Then rstnew.Open "execute CrewEQCreditCompanyDayBreakDown '" & StartDate & "'", conn
            If RegionCounter > 5 Then rstnew.Open "execute CrewEQCreditCompanyDayBreakDown '" & StartDate & "', '%','%', '" & sRegion & "'", conn
            If rstnew.State = 1 Then
                sht.Range(sht.Cells(lrowsetter(sht) + 1, 1).Address).CopyFromRecordset rstnew
                rstnew.Close
            End If
           ' For EquipCreditYardCount = 0 To YardCounter - 1 '<-------------------------------------------------This needs to be changed if we add or subtract yards from transyards. It's one less than the cover one (9) because it doesn't have to account for the master row.
            '    If Worksheets(1).Cells(25 + (19 * EquipCreditYardCount), 2).Value = "" Then Exit For
             '   For EquipCreditDayCount = 0 To 6
              '      Set CreditRange = Worksheets(1).Range(Worksheets(1).Cells(35 + (19 * EquipCreditYardCount), 6 + (6 * EquipCreditDayCount)), Worksheets(1).Cells(41 + (19 * EquipCreditYardCount), 6 + (6 * EquipCreditDayCount)))
               '     Set creditinsert = sht.Range(sht.Cells(lrowbycol(sht, 16) + 1, 16).Address) ' need lrow for column 16 here
                '    For creditinsertloop = 1 To sht.Cells(creditinsert.Row, 9)
                 '       sht.Cells(creditinsert.Row - 1 + creditinsertloop, 16).Value = WorksheetFunction.Sum(CreditRange) / sht.Cells(creditinsert.Row, 9).Value * -1
                  '  Next creditinsertloop
                'Next EquipCreditDayCount
            'Next EquipCreditYardCount
            'Set clearrange = sht.Range(clearrange, clearrange.End(xlDown))
            'clearrange.Delete (xlShiftUp)
        'If regioncounter < 5 Then
         '   Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(lrowsetter(sht), 1))
          '  rng.AutoFilter Field:=1, Criteria1:=sRegion, VisibleDropDown:=False
        'End If
       '------------------------------------we pulled this because we don't have a page5 anymore. We delete page 5 below. If we reinstate this code, we need to get rid of the delete code at the end of the counter
       ' ElseIf counter = 5 Then
       ' rstnew.Open "exec TransSummaryPivot @startdate = '" & StartDate & "', @Region = '" & sRegion & "'", conn
       ' If rstnew.State = 1 Then
       '     sht.Range("B3").CopyFromRecordset rstnew
       '     rstnew.Close
       ' End If
        ElseIf counter = 6 Then
            If RegionCounter = 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where ordername = 'Microsite Support' AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter < 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where region = '" & sRegion & "' AND ordername = 'Microsite Support' AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter > 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where ordername = 'Microsite Support' AND c.yard <> 'Substation' AND Company like '" & sRegion & "' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
            End If
            'For LaborCreditYardCount = 0 To YardCounter - 1
             '   If Worksheets(1).Cells(25 + (19 * LaborCreditYardCount), 2).Value = "" Then Exit For
              '  For LaborCreditDayCount = 0 To 6
               '     Set CreditRange = Worksheets(1).Range(Worksheets(1).Cells(27 + (19 * LaborCreditYardCount), 6 + (6 * LaborCreditDayCount)), Worksheets(1).Cells(32 + (19 * LaborCreditYardCount), 6 + (6 * LaborCreditDayCount)))
            Set creditinsert = sht.Range(sht.Cells(lrowsetter(sht) + 1, 1).Address)
            If RegionCounter < 5 Then rstnew.Open "exec LaborCreditCompanyDayBreakDown '" & StartDate & "', '" & sRegion & "'", conn
            If RegionCounter = 5 Then rstnew.Open "exec LaborCreditCompanyDayBreakDown '" & StartDate & "'", conn
            If RegionCounter > 5 Then rstnew.Open "exec LaborCreditCompanyDayBreakDown '" & StartDate & "','%','%','" & sRegion & "'", conn
            If rstnew.State = 1 Then
                creditinsert.CopyFromRecordset rstnew
                rstnew.Close
            End If
                    'creditinsert.Value = WorksheetFunction.Sum(CreditRange) * -1
                    'sht.Range(sht.Cells(lrowsetter(sht), 1).Address).Value = sRegion 'this wont work for master. Might need to use sql to do a join off yard and region
                    'sht.Cells(lrowsetter(sht), 3).Value = Worksheets(1).Cells(25 + (19 * LaborCreditYardCount), 2).Value
                    'sht.Cells(lrowsetter(sht), 11).Value = "Microsite Support"
                    'rstnew.Open "Select YardPM From MicrositeSupportPMs Where Yard = '" & sht.Cells(lrowsetter(sht), 3).Value & "'", conn
                    'If rstnew.State = 1 Then
                    '    sht.Cells(lrowsetter(sht), 12).CopyFromRecordset rstnew
                    '    rstnew.Close
                    'End If
                'Next LaborCreditDayCount
            'Next LaborCreditYardCount
            'If RegionCounter < 5 Then
             '   Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(lrowsetter(sht), 1))
              '  rng.AutoFilter Field:=1, Criteria1:=sRegion, VisibleDropDown:=False
            'End If
        ElseIf counter = 7 Then
            If RegionCounter = 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where ordername = 'Program Support' AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter < 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where Region like '" & sRegion & "' AND  ordername = 'Program Support' AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter > 5 Then
               rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where ordername = 'Program Support' AND c.yard <> 'Substation' AND company = '" & sRegion & "' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
            End If
            'If RegionCounter < 5 Then
             '   Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(lrowsetter(sht), 1))
              '  rng.AutoFilter Field:=1, Criteria1:=sRegion, VisibleDropDown:=False
            'End If
        ElseIf counter = 8 Then
            If RegionCounter = 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where ordername = 'Circuit Inspections' AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter < 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where Region like '" & sRegion & "' AND ordername = 'Circuit Inspections' AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter > 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where ordername = 'Circuit Inspections' AND c.yard <> 'Substation' AND company = '" & sRegion & "' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
            End If
            'If RegionCounter < 5 Then
                'Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(lrowsetter(sht), 1))
                'rng.AutoFilter Field:=1, Criteria1:=sRegion, VisibleDropDown:=False
            'End If
        ElseIf counter = 9 Then
            If RegionCounter = 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where C.PM = 2049045 AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter < 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where Region like '" & sRegion & "' AND C.PM = 2049045 AND c.yard <> 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If RegionCounter > 5 Then
                rstnew.Open "Select c.rowid, tr.region, c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
                & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
                & " Where C.PM = 2049045 AND c.yard <> 'Substation' AND company = '" & sRegion & "' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
            End If
            'If RegionCounter < 5 Then
             '   Set rng = sht.Range(sht.Cells(2, 1), sht.Cells(lrowsetter(sht), 1))
              '  rng.AutoFilter Field:=1, Criteria1:=sRegion, VisibleDropDown:=False
            'End If
        ElseIf counter = 10 Then
            If RegionCounter = 5 Then 'creates substation page for the Master Document
            rstnew.Open "Select c.rowid, 'Substation', c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
           & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
           & " Where c.yard = 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            ElseIf RegionCounter > 5 Then 'creates substation page for the Master Document
            rstnew.Open "Select c.rowid, 'Substation', c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
           & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
           & " Where c.yard = 'Substation' AND company = '" & sRegion & "' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
            End If
            If rstnew.State = 1 Then
                sht.Range("A3").CopyFromRecordset rstnew
                rstnew.Close
            End If
        ElseIf counter = 11 Then 'Moves R&M to front, then the other invoice category sheets behind it.
            sht.Move Before:=Worksheets(1)
        ElseIf counter = 12 Then
            sht.Move After:=Worksheets(1)
        ElseIf counter = 13 Then
            sht.Move After:=Worksheets(2)
        ElseIf counter = 14 Then
            sht.Move After:=Worksheets(3)
        ElseIf counter = 15 Then
            sht.Move After:=Worksheets(4)
        ElseIf counter = 16 Then
            If RegionCounter = 5 Then rstnew.Open "exec crewcostlaboroveragebreakdown '" & StartDate & "', '%'", conn
            If RegionCounter < 5 Then rstnew.Open "exec crewcostlaboroveragebreakdown '" & StartDate & "', '" & sRegion & "'", conn
            If RegionCounter > 5 Then rstnew.Open "exec crewcostlaboroveragebreakdown '" & StartDate & "', '%', '" & sRegion & "'", conn
            If rstnew.State = 1 Then
                sht.Range("A2").CopyFromRecordset rstnew
                rstnew.Close
            End If
            sht.Move After:=Worksheets(8)
        ElseIf counter = 17 Then
            If RegionCounter = 5 Then rstnew.Open "exec crewcostEquipmentoveragebreakdown '" & StartDate & "'", conn
            If RegionCounter < 5 Then rstnew.Open "exec crewcostEquipmentoveragebreakdown '" & StartDate & "', '" & sRegion & "'", conn
            If RegionCounter > 5 Then rstnew.Open "exec crewcostEquipmentoveragebreakdown '" & StartDate & "', '%', '" & sRegion & "'", conn
            If rstnew.State = 1 Then
                sht.Range("A2").CopyFromRecordset rstnew
                rstnew.Close
            End If
            sht.Move After:=Worksheets(9)
       End If
        
    Next counter
    
 End If
    
    If RegionCounter < 5 Then
        Worksheets("Substations").Delete 'deletes the substation page if the worksheet isn't a MasterSheet
        Worksheets("Summary").Rows("22:23").Delete 'deletes the substation rows on the summary page
        Worksheets.Add After:=Worksheets(15)
        Set sht = Worksheets(16)
        sht.Name = d(18)
        sht.Range("A1").Value = "PM"
        sht.Range("B1").Value = "Cost Per PM"
        sht.Range("D1").Value = "LC"
        sht.Range("E1").Value = "Cost Per LC"
        sht.Range("A1:E1").Font.Bold = True
        rstnew.Open "Execute RegionalPMCostSummary '" & StartDate & "','" & sRegion & "'", conn
        If rstnew.State = 1 Then
            sht.Range("A2").CopyFromRecordset rstnew
            rstnew.Close
        End If
        rstnew.Open "Execute RegionalPMLCCostSummary '" & StartDate & "','" & sRegion & "'", conn
        If rstnew.State = 1 Then
            sht.Range("D2").CopyFromRecordset rstnew
            rstnew.Close
        End If
    ElseIf RegionCounter = 5 Then
            'For i = 1 To 7-----------------------------------------------------------------------------this was an old way to get crewcost for the week based off of excel values. Now we get it from a stored procedure in SQL
                'CrewCreditCost = CrewCreditCost + Worksheets("Master").Cells(6, i * 6).Value
            'Next i
        Worksheets.Add After:=Worksheets(15)
        Set sht = Worksheets(16)
        sht.Name = d(18)
        sht.Range("A1").Value = "PM"
        sht.Range("B1").Value = "Cost Per PM"
        sht.Range("D1").Value = "LC"
        sht.Range("E1").Value = "Cost Per LC"
        sht.Range("A1:E1").Font.Bold = True
        rstnew.Open "Execute PMCostSummary '" & StartDate & "'", conn
        If rstnew.State = 1 Then
            sht.Range("A2").CopyFromRecordset rstnew
            rstnew.Close
        End If
        rstnew.Open "Execute PMLCCostSummary '" & StartDate & "'", conn
        If rstnew.State = 1 Then
            sht.Range("D2").CopyFromRecordset rstnew
            rstnew.Close
        End If
        Set sht = Worksheets("Progress Billing")
        rstnew.Open "Execute IntercompanySummaryDebits '" & StartDate & "'", conn
        If rstnew.State = 1 Then
            sht.Range("B90").CopyFromRecordset rstnew
            rstnew.Close
        End If
        rstnew.Open "Execute IntercompanySummaryCredits '" & StartDate & "'", conn
        If rstnew.State = 1 Then
            sht.Range("B101").CopyFromRecordset rstnew
            rstnew.Close
        End If
    ElseIf RegionCounter < 14 Then
    Worksheets("Progress Billing").Delete
    End If
If RegionCounter < 14 Then Worksheets("Support_Inspection Summary").Delete 'this is the delete line that needs to be deleted if we uncomment out the above code for page 5
   
If RegionCounter > 4 And RegionCounter < 14 Then
    Set sht = wkb.Worksheets("Intercompany Transfer")
    rstnew.Open "exec intercompanysummarydebits '" & StartDate & "'", conn
    If rstnew.State = 1 Then
            sht.Range("A3").CopyFromRecordset rstnew
            rstnew.Close
    End If
    rstnew.Open "exec intercompanysummarycredits '" & StartDate & "'", conn
    If rstnew.State = 1 Then
            sht.Range("A14").CopyFromRecordset rstnew
            rstnew.Close
    End If

ElseIf Not RegionCounter > 4 And RegionCounter < 14 Then
    wkb.Worksheets("Intercompany Transfer").Delete
 
    
ElseIf RegionCounter = 14 Then 'this creates the master substation document. It's much smaller than the main documents, so it gets less code.
        Set wkb = Workbooks.Add(sSubBkName)
        Set sht = wkb.Worksheets(1)
        sht.Range("B1").Value = "Billing Summary - " & StartDate & " to " & EndDate
        sht.Range("B7").Value = "Total Billing Summary - " & StartDate & " to " & EndDate
        Set sht = wkb.Worksheets(2)
        rstnew.Open "Select 'SubStation', c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
        & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
        & " Where c.yard = 'Substation' AND workdate between '" & StartDate & "' AND '" & EndDate & "'", conn
'ElseIf RegionCounter < 14 Then 'this creates the company substation documents
 '       Set wkb = Workbooks.Add(sSubBkName)
  '      Set sht = wkb.Worksheets(1)
   '     sht.Range("B1").Value = "Billing Summary - " & startdate & " to " & EndDate
    '    sht.Range("B25").Value = "Total Billing Summary - " & startdate & " to " & EndDate
     '   Set sht = wkb.Worksheets(2)
      '  rstnew.Open "Select 'SubStation', c.Company, c.yard, c.workdate, c.lanid, c.empname, c.classid, c.classdescription, c.crewleadlan, c.crew, c.ordername, c.pm, c.lc, c.sthours, c.dthours, c.meals, c.subsistance, c.isunion, c.nonunion, c.unitid, c.equiphours, eqr.HourlyRate, eqr.CO6RateGroup, eqr.CO6EquipmentDescription, c.equipowner, lr.strate, lr.dtrate, 15*c.Meals as meals , 114.66*c.subsistance as sub, 100*isunion as UnionPerDiem, 200*nonunion as NonUnionPerDiem, strate*sthours as stcost, dtrate*dthours as dtcost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) as LaborCost, isnull(c.equiphours*eqr.HourlyRate,0) as EquipCost, isnull(15*c.meals + 114.66*c.subsistance + 100*isunion + 200*nonunion + strate*sthours + dtrate*dthours,0) + isnull(c.equiphours*eqr.HourlyRate,0) as LineTotalCost " _
       ' & "from CostingHelper c left join laborrates lr on lr.classid = c.classid left join equiprates eqr on eqr.co5rategroup = c.rategroup left join TransmissionRegion tr on c.yard = tr.yard" _
        '& " Where c.yard = 'Substation' AND company = '" & sRegion & "' AND workdate between '" & startdate & "' AND '" & EndDate & "'", conn
    'End If
    If rstnew.State = 1 Then
        sht.Range("A3").CopyFromRecordset rstnew
        rstnew.Close
    End If
End If

Set sht = Worksheets("summary")
If RegionCounter > 5 And RegionCounter < 14 Then
    
    For categorycounter = 0 To 5
        rstnew.Open "exec intercompanytotals '" & StartDate & "', '" & sRegion & "', '" & d2(categorycounter) & "'", conn
        If rstnew.State = 1 Then
            If rstnew.EOF = False Then
                sht.Range(sht.Cells(12 + (categorycounter * 2), 5).Address).Value = rstnew.Fields("Total")
            End If
            rstnew.Close
        End If
    Next categorycounter
    sht.Range("E2").Value = "Intercompany Totals"
    sht.Range("F2").Value = "Grand Total"
    sht.Range("E2:F2").Font.Bold = True
End If

If RegionCounter = 5 Then
    rstnew.Open "exec InvoiceTandE '" & StartDate & "'", conn
    If rstnew.State = 1 Then
        sht.Range("I6").CopyFromRecordset rstnew
        rstnew.Close
    End If
    rstnew.Open "Exec CrewCountsByGuysPerCrew '" & StartDate & "'", conn
    If rstnew.State = 1 Then
        sht.Range("H12").CopyFromRecordset rstnew
        rstnew.Close
    End If
End If

If RegionCounter > 5 And RegionCounter < 14 Then
    rstnew.Open "exec InvoiceTandE '" & StartDate & "', '" & sRegion & "'", conn
    If rstnew.State = 1 Then
        sht.Range("I6").CopyFromRecordset rstnew
        rstnew.Close
    End If
    rstnew.Open "Exec CrewCountsByGuysPerCrew '" & StartDate & "', '" & sRegion & "'", conn
    If rstnew.State = 1 Then
        sht.Range("H12").CopyFromRecordset rstnew
        rstnew.Close
    End If
End If

For Each sht In ThisWorkbook.Worksheets
    sht.Cells.EntireColumn.AutoFit
  Next sht
'a = Worksheets("summary").Cells(26, 4).Value
'b = Worksheets("summary").Cells(24, 4).Value
'If a + b > 0 Or sRegion = "SubStations" Then
    wkb.SaveAs ("C:\Austin_Kimes\CostingTemplate\TransCosting " & sRegion & Month(StartDate) & "." & Day(StartDate) & " to " & Month(EndDate) & "." & Day(EndDate) & ".xlsx")
'End If
wkb.Close



Next RegionCounter

Application.ScreenUpdating = True
Application.EnableEvents = True
Application.DisplayAlerts = True

MsgBox "done!"
End Sub




















