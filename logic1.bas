'create worksheets as global variables
Dim wb As Workbook
Dim D1schedule As Worksheet
Dim D1Default As Worksheet
Dim D2Schedule As Worksheet
Dim D2Default As Worksheet
Dim DBSchedule As Worksheet
Dim PPCanSchedule As Worksheet
Dim PPTippingStation As Worksheet
Dim Silos As Worksheet
Dim D1DefaultOriginal As Worksheet
Dim D2DefaultOriginal As Worksheet

Sub resetAll()
    initializeWorksheets
    
    D1Default.Range("A:N").Value = D1DefaultOriginal.Range("A:N").Value
    D2Default.Range("A:N").Value = D2DefaultOriginal.Range("A:N").Value
    
    PPCanSchedule.Range("A:N").Value = PPCanSchedule.Range("R:AD").Value
    DBSchedule.Range("A:O").Value = DBSchedule.Range("Q:AE").Value
    
    D1schedule.Range("A:N").Value = D1Default.Range("A:N").Value
    D2Schedule.Range("A:N").Value = D2Default.Range("A:N").Value
    
    Application.Calculate
    
End Sub

Sub main()
    'turn off autosave
    Application.AutoRecover.Enabled = False
    
    initializeWorksheets
    
    Dim isLogic1Feasible As Boolean
    isLogic1Feasible = logic1()
    If isLogic1Feasible = False Then
        MsgBox "PP-Can and 100DB Campaigns cannot be inserted even after setting silo constraint to 22(6). Terminating Program."
        End
    End If
     
    
End Sub

Sub initializeWorksheets()
    'note that the worksheets have to be in the same workbook
    'have the PPCan and 100DB schedules in the same workbook
    Set wb = ThisWorkbook
    setWorksheet D1schedule, "D1B1L65T"
    setWorksheet D1Default, "D1Sched"
    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet D2Default, "D2Sched"
    setWorksheet DBSchedule, "DBSCH Reorder Select"
    setWorksheet PPCanSchedule, "PP CAN"
    setWorksheet Silos, "Silos"
    setWorksheet PPTippingStation, "PP"
    setWorksheet D1DefaultOriginal, "D1Sched (2)"
    setWorksheet D2DefaultOriginal, "D2Sched (2)"
    ' update pivot table to correct setting PP sheet
    Dim PT As PivotTable
    For Each PT In PPTippingStation.PivotTables
        On Error Resume Next
        For Each PI In PT.PivotFields("Source (DR, DB, PP)").PivotItems
            Select Case PI.Name
                Case Is = "PP"
                    PI.Visible = True
                Case Else
                    PI.Visible = False
            End Select
        Next PI
    Next PT

    'Include Silo Constraint presense for SG
    Silos.Range("R8:S8").Value = "PE"
    Silos.Range("T8:U8").Value = "SG"
    Silos.Range("T9").Formula = "=MAXIFS(D1B1L65T!AJ:AJ,D1B1L65T!AJ:AJ,""<=""&Silos!$K$2,D1B1L65T!AP:AP,"">=1"")"
    Silos.Range("T10").Formula =  "=MAXIFS(D2B1L3B3B4L45T!AJ:AJ,D2B1L3B3B4L45T!AJ:AJ,""<=""&Silos!$K$2,D2B1L3B3B4L45T!AP:AP,"">=1"")"
    Silos.Range("U9").Formula = "=IF(K2-T9<0.5,""YES"",""NO"")"
    Silos.Range("U10").Formula = "=IF(K2-T10<0.5,""YES"",""NO"")"
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wb.Sheets(worksheetName)
    Exit Sub
Err:
    MsgBox worksheetName & " is not in current workbook"
    End
End Sub

Function logic1()
    Dim mainSilo As Integer
    Dim otherSilo As Integer
    mainSilo = 16
    otherSilo = 6
    
    Dim isFeasible As Boolean
    isFeasible = False
    Do While mainSilo <= 22
        isFeasible = insertPPCan100DBCampaigns(mainSilo, otherSilo)
        If isFeasible = True Then
            Exit Do
        End If
        mainSilo = mainSilo + 1
    Loop
    logic1 = isFeasible
End Function
    
Function insertPPCan100DBCampaigns(mainSilo, otherSilo) As Boolean
    
    ' arrays for determining which can starve to skip
    Dim d1Skip() As Integer
    Dim d2Skip() As Integer
    ReDim d1Skip(1)
    ReDim d2Skip(1)
    d1Skip(0) = 0
    d2Skip(0) = 0
    
    Do While True
        ' get row of campaign to insert
        ' -1 if there is no campaign
        Dim PPCampaignToInsert As Double
        Dim DBCampaignToInsert As Double
        PPCampaignToInsert = findNextCampaignToInsert(PPCanSchedule)
        DBCampaignToInsert = findNextCampaignToInsert(DBSchedule)
        
        ' get row of insertion in schedule
        ' -1 if there is no can starve
        Dim D1FirstCanStarveTime As Double
        Dim D2FirstCanStarveTime As Double
        D1FirstCanStarveTime = findFirstCanStarveTime(D1schedule, d1Skip)
        D2FirstCanStarveTime = findFirstCanStarveTime(D2Schedule, d2Skip)
        
        ' get initial silo constraint violation time
        Dim initialSiloConstraintViolation
        initialSiloConstraintViolation = Silos.Range("K1").Value
        
'        ' if the min of the time of insertion is after the initial silo constraint violation, run dryer blockage
'        If D1FirstCanStarveTime <> -1 And D2FirstCanStarveTime <> -1 Then
'            If D1FirstCanStarveTime < D2FirstCanStarveTime Then
'                If checkSiloConstraint(mainSilo, otherSilo, D1schedule, D1FirstCanStarveTime, initialSiloConstraintViolation) = False Then
'                    Module4.dryerBlockDelayMain D1FirstCanStarveTime
'                    GoTo continueLoop
'                End If
'            Else
'                If checkSiloConstraint(mainSilo, otherSilo, D2Schedule, D2FirstCanStarveTime, initialSiloConstraintViolation) = False Then
'                    Module4.dryerBlockDelayMain D2FirstCanStarveTime
'                    GoTo continueLoop
'                End If
'            End If
'        ElseIf D2FirstCanStarveTime <> -1 Then
'            If checkSiloConstraint(mainSilo, otherSilo, D2Schedule, D2FirstCanStarveTime, initialSiloConstraintViolation) = False Then
'                    Module4.dryerBlockDelayMain D2FirstCanStarveTime
'                    GoTo continueLoop
'            End If
'        ElseIf D1FirstCanStarveTime <> -1 Then
'            If checkSiloConstraint(mainSilo, otherSilo, D1schedule, D1FirstCanStarveTime, initialSiloConstraintViolation) = False Then
'                    Module4.dryerBlockDelayMain D1FirstCanStarveTime
'                    GoTo continueLoop
'            End If
'        End If
        
        ' get which dryer and which campaign to insert
        Dim dryerCampaign As Integer
        dryerCampaign = determineDryerCampaign(D1FirstCanStarveTime, D2FirstCanStarveTime, PPCampaignToInsert, DBCampaignToInsert)
            
        If dryerCampaign = -2 Then 'case: db campaigns but no more d2 slots (infeasible solution)
            MsgBox "DB campaigns remaining but no more can starvation slots in dryer 2. Exiting Program."
            End
        ElseIf dryerCampaign = -1 Then 'case: no more campaigns left
            MsgBox "All campaigns Inserted"
            insertPPCan100DBCampaigns = True
            Exit Function
        ElseIf dryerCampaign = 0 Then 'case: no more dryer slots
            MsgBox "All can starvation slots used. Increasing silo constraint"
            insertPPCan100DBCampaigns = False
            Exit Function
        ElseIf dryerCampaign = 1 Then 'case: d1 pp campaign
            If D1schedule.Range("BI" & D1FirstCanStarveTime - 1).Value > initialSiloConstraintViolation Then
                    Module4.dryerBlockDelayMain D1schedule.Range("BI" & D1FirstCanStarveTime - 1).Value
                    GoTo continueLoop
            End If
            MsgBox "Adding PP campaign to dryer 1"
            d1Skip = addPPCampaign(PPCampaignToInsert, D1schedule, D1Default, D1FirstCanStarveTime, mainSilo, otherSilo, d1Skip, initialSiloConstraintViolation)
        ElseIf dryerCampaign = 2 Then 'case: d2 pp campaign
           If D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value > initialSiloConstraintViolation Then
                    Module4.dryerBlockDelayMain D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value
                    GoTo continueLoop
            End If
            MsgBox "Adding PP campaign to dryer 2"
            d2Skip = addPPCampaign(PPCampaignToInsert, D2Schedule, D2Default, D2FirstCanStarveTime, mainSilo, otherSilo, d2Skip, initialSiloConstraintViolation)
        ElseIf dryerCampaign = 3 Then 'case: d2 db campaign
            If D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value > initialSiloConstraintViolation Then
                    Module4.dryerBlockDelayMain D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value
                    GoTo continueLoop
            End If
            MsgBox "Adding DB campaign to dryer 2"
            d2Skip = addDBCampaign(DBCampaignToInsert, D2Schedule, D2Default, D2FirstCanStarveTime, mainSilo, otherSilo, d2Skip, initialSiloConstraintViolation)
        ElseIf dryerCampaign = 4 Then 'case: skip d1 can starve time
            d1Skip = addItemToArray(D1FirstCanStarveTime, d1Skip)
        ElseIf dryerCampaign = 5 Then 'case: skip d2 can starve time
            d2Skip = addItemToArray(D2FirstCanStarveTime, d2Skip)
        ElseIf dryerCampaign = 6 Then 'case: skip d1 and d2 can starve time
            d1Skip = addItemToArray(D1FirstCanStarveTime, d1Skip)
            d2Skip = addItemToArray(D2FirstCanStarveTime, d2Skip)
        End If
continueLoop:
    Loop
    insertPPCan100DBCampaigns = True
End Function

Function addDBCampaign(DBCampaignToInsert, dryerSchedule, dryerDefaultSchedule, dryerFirstCanStarveTime, mainSilo, otherSilo, dryerSkipArray, initialSiloConstraintViolation) As Integer()
    ' the window to add campaigns from
    Dim dbWindow As Integer
    dbWindow = DBSchedule.Range("O" & DBCampaignToInsert).Value
    
    ' get the last row with same window
    Dim lastRow As Integer
    lastRow = DBCampaignToInsert
    Do While True
        If DBSchedule.Range("O" & lastRow).Value <> dbWindow Then
            lastRow = lastRow - 1
            Exit Do
        Else
            lastRow = lastRow + 1
        End If
    Loop
    
    Dim i As Integer
    For i = lastRow To DBCampaignToInsert Step -1
        ' insert DB campaign
        DBSchedule.Range("A" & DBCampaignToInsert, "M" & i).Copy
        dryerDefaultSchedule.Range("A" & dryerFirstCanStarveTime).Insert xlShiftDown
        dryerSchedule.Range("A:M").Value = dryerDefaultSchedule.Range("A:M").Value
        Application.Calculate
        
        ' check if the added campaign satisfies silo constraint
        canAdd = checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerFirstCanStarveTime, initialSiloConstraintViolation)
        If canAdd = True Then
            DBSchedule.Range("A" & DBCampaignToInsert, "O" & i).Delete xlShiftUp
            Exit For
        End If
        
        dryerDefaultSchedule.Rows(dryerFirstCanStarveTime & ":" & (dryerFirstCanStarveTime + (i - DBCampaignToInsert))).EntireRow.Delete
        ' case nothing can be added
        If i <= DBCampaignToInsert Then
            dryerSkipArray = addItemToArray(dryerFirstCanStarveTime, dryerSkipArray)
            dryerSchedule.Range("A:M").Value = dryerDefaultSchedule.Range("A:M").Value
            Exit For
        End If
    Next
        
    Application.Calculate
    addDBCampaign = dryerSkipArray
End Function

Function addPPCampaign(PPCampaignToInsert, dryerSchedule, dryerDefaultSchedule, dryerFirstCanStarveTime, mainSilo, otherSilo, dryerSkipArray, initialSiloConstraintViolation) As Integer()
    
    ' decrement counter can be modified to determine the "steps" to reduce campaign load when it can't be inserted
    Dim decrementCounter As Double
    decrementCounter = 0.5

    ' boolean flag to determine if silo constraint is being violated
    Dim canAdd As Boolean
    canAdd = False

    Dim i As Double
    For i = 1 To decrementCounter Step -decrementCounter
        ' insert to the row before the can starvation time
        PPCanSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Copy
        dryerDefaultSchedule.Range("A" & dryerFirstCanStarveTime).Insert xlShiftDown
        dryerDefaultSchedule.Range("J" & dryerFirstCanStarveTime).Value = dryerDefaultSchedule.Range("J" & dryerFirstCanStarveTime).Value * i
        dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
        Application.Calculate

        canAdd = checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerFirstCanStarveTime, initialSiloConstraintViolation)
        If canAdd = True Then
            If i = 1 Then
                PPCanSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Delete
            Else
                PPCanSchedule.Range("J" & PPCampaignToInsert).Value = PPCanSchedule.Range("J" & PPCampaignToInsert).Value * (1 - i)
            End If
            Exit For
        End If
        dryerDefaultSchedule.Rows(dryerFirstCanStarveTime).EntireRow.Delete
        If i <= decrementCounter Then
            dryerSkipArray = addItemToArray(dryerFirstCanStarveTime, dryerSkipArray)
            dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
        End If
    Next
    Application.Calculate
    
    ' this is to ensure that the pivot table is updated after adding pp campaigns
    wb.RefreshAll
    
    addPPCampaign = dryerSkipArray
End Function

Function checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerInsertRow, initialSiloConstraintViolation) As Boolean
    Dim siloCheckStartTime As Double
    Dim siloCheckTimeEnd As Double
    siloCheckTimeStart = dryerSchedule.Range("BY" & dryerInsertRow).Value 'silo entry hour
    'siloCheckTimeEnd = dryerSchedule.Range("BB" & dryerInsertRow).Value
    
    ' iterate through silos sheet to find if the silo constraint is being violated by the campaign insertion
    Dim i As Double
    For i = 2 To (2 ^ 15) - 1 Step 1
        If Silos.Range("A" & i).Value >= siloCheckTimeStart And Silos.Range("A" & i).Value < initialSiloConstraintViolation Then
            If Silos.Range("D" & i).Value > mainSilo Or Silos.Range("G" & i).Value > otherSilo Then
                checkSiloConstraint = False
                Exit Function
            End If
        End If
    Next
    
    checkSiloConstraint = True
End Function
      
Function determineDryerCampaign(D1FirstCanStarveTime, D2FirstCanStarveTime, PPCampaignToInsert, DBCampaignToInsert) As Integer
    If PPCampaignToInsert = -1 And DBCampaignToInsert = -1 Then
        determineDryerCampaign = -1
        Exit Function
    End If
    
    If D1FirstCanStarveTime = -1 And D2FirstCanStarveTime = -1 Then
        determineDryerCampaign = 0
        Exit Function
    End If
    
    ' check PP sheet pivot table to determine tipping station availability
    Dim tippingStationAvailableTime As Double
    tippingStationAvailableTime = 0
    tippingStationAvailableTime = getTippingStationAvailableStartTime
    
    Dim D1CanStarveStartTime As Double
    Dim D2CanStarveStartTime As Double
    If D1FirstCanStarveTime <> -1 Then
        D1CanStarveStartTime = D1schedule.Range("BK" & D1FirstCanStarveTime - 1).Value
    End If
    If D2FirstCanStarveTime <> -1 Then
        D2CanStarveStartTime = D2Schedule.Range("BK" & D2FirstCanStarveTime - 1).Value
    End If


    If D1FirstCanStarveTime <> -1 And D2FirstCanStarveTime <> -1 Then 'case d1 and d2 both have slots
        If PPCampaignToInsert <> -1 And DBCampaignToInsert <> -1 Then 'case both pp and db campaigns available
            If D1CanStarveStartTime < D2CanStarveStartTime + 150 Then
                If D1CanStarveStartTime > tippingStationAvailableTime Then
                    determineDryerCampaign = 1 'd1pp
                Else
                    If D2CanStarveStartTime > tippingStationAvailableTime Then
                        determineDryerCampaign = 2 'd2pp
                    Else
                        determineDryerCampaign = 3 'd2db
                    End If
                End If
            Else
                If D2CanStarveStartTime > tippingStationAvailableTime Then
                    determineDryerCampaign = 2 'd2pp
                Else
                    determineDryerCampaign = 3 'd2db
                End If
            End If
        ElseIf PPCampaignToInsert <> -1 And DBCampaignToInsert = -1 Then 'case only pp campaign available
            If D1CanStarveStartTime < D2CanStarveStartTime Then
                If D1CanStarveStartTime > tippingStationAvailableTime Then
                    determineDryerCampaign = 1 'd1pp
                Else
                    If D2CanStarveStartTime > tippingStationAvailableTime Then
                        determineDryerCampaign = 2 'd2pp
                    Else
                        determineDryerCampaign = 6 'can't do pp on d1 and d2, no more db campaign so skip can starve time
                    End If
                End If
            Else
                If D2CanStarveStartTime > tippingStationAvailableTime Then
                    determineDryerCampaign = 2 'd2pp
                Else
                    If D1CanStarveStartTime > tippingStationAvailableTime Then
                        determineDryerCampaign = 1 'd1pp
                    Else
                        determineDryerCampaign = 6 'can't do pp on d1 and d2, no more db campaign so skip can starve time
                    End If
                End If
            End If
        ElseIf PPCampaignToInsert = -1 And DBCampaignToInsert <> -1 Then 'case only db campaign available
            determineDryerCampaign = 3 'd2db
        End If
    ElseIf D1FirstCanStarveTime <> -1 And D2FirstCanStarveTime = -1 Then 'case only d1 has slots
        If PPCampaignToInsert <> -1 And DBCampaignToInsert <> -1 Then 'case both pp and db campaigns available
            If D1CanStarveStartTime > tippingStationAvailableTime Then
                determineDryerCampaign = 1 'd1pp
            Else
                determineDryerCampaign = 4 'can't do pp on d1 and d2 is not available so skip can starve time
            End If
        ElseIf PPCampaignToInsert <> -1 And DBCampaignToInsert = -1 Then 'case only pp campaign available
            If D1CanStarveStartTime > tippingStationAvailableTime Then
                determineDryerCampaign = 1 'd1pp
            Else
                determineDryerCampaign = 4 'can't do pp on d1 and d2 is not available so skip can starve time
            End If
        ElseIf PPCampaignToInsert = -1 And DBCampaignToInsert <> -1 Then 'case only db campaign available
            determineDryerCampaign = -2 'there are no d2 can starve times but db campaigns remaning
        End If
    ElseIf D1FirstCanStarveTime = -1 And D2FirstCanStarveTime <> -1 Then 'case only d2 has slots
        If PPCampaignToInsert <> -1 And DBCampaignToInsert <> -1 Then 'case both pp and db campaigns available
            If D2CanStarveStartTime > tippingStationAvailableTime Then
                determineDryerCampaign = 2 'd2pp
            Else
                determineDryerCampaign = 3 'd2db
            End If
        ElseIf PPCampaignToInsert <> -1 And DBCampaignToInsert = -1 Then 'case only pp campaign available
            If D2CanStarveStartTime > tippingStationAvailableTime Then
                determineDryerCampaign = 2 'd2pp
            Else
                determineDryerCampaign = 5 'can't insert pp can and there are no more db campaigns so skip d2 can starve time
            End If
        ElseIf PPCampaignToInsert = -1 And DBCampaignToInsert <> -1 Then 'case only db campaign available
            determineDryerCampaign = 3 'd2db
        End If
    End If
End Function

Function getTippingStationAvailableStartTime() As Double
    Dim tippingStationAvailableTime As Double
    tippingStationAvailableTime = 0
    Dim PT As PivotTable
    For Each PT In PPTippingStation.PivotTables
        For Each Column In PT.ColumnRange
             If Column.Value = "Sum of Can After CO Hrs" Then
                For Each Row In PT.RowRange
                    If IsNumeric(Row.Value) Then
                        If PPTippingStation.Cells(Row.Row, Column.Column).Value > tippingStationAvailableTime Then
                            tippingStationAvailableTime = PPTippingStation.Cells(Row.Row, Column.Column).Value
                        End If
                    End If
                Next
            End If
        Next
    Next PT
    If tippingStationAvailableTime <> 0 Then
        tippingStationAvailableTime = tippingStationAvailableTime + 40
    End If
    getTippingStationAvailableStartTime = tippingStationAvailableTime
End Function

' checks the FP Loading Per Batch (if 0 then the campaign has been inserted)
Function findNextCampaignToInsert(Worksheet) As Integer
    If Worksheet.Range("A1").End(xlDown).Value = "" Then
        findNextCampaignToInsert = -1
        Exit Function
    End If
    Dim cell As Range
    For Each cell In Worksheet.Range("A2:A" & Worksheet.Range("A" & Rows.Count).End(xlUp).Row)
        If cell.Value <> "" Then
            findNextCampaignToInsert = cell.Row
            Exit Function
        End If
    Next cell
End Function

Function isCanStarveInArray(canStarve, dryerSkipArray) As Boolean
    Dim i As Integer
    For i = LBound(dryerSkipArray) To UBound(dryerSkipArray)
        If dryerSkipArray(i) = canStarve Then
            isCanStarveInArray = True
            Exit Function
        End If
    Next
    isCanStarveInArray = False
End Function

Function addItemToArray(item, dryerSkipArray) As Integer()
    ReDim Preserve dryerSkipArray(LBound(dryerSkipArray) To UBound(dryerSkipArray) + 1)
    dryerSkipArray(UBound(dryerSkipArray)) = item
    addItemToArray = dryerSkipArray
End Function

Function findFirstCanStarveTime(Worksheet, dryerSkipArray) As Double
    'ensure column CI is Can Starve
    If IsNumeric("CI1") Or Worksheet.Range("CI1").Value <> "Can Starve" Then
            MsgBox "Cell CI1 is not set to Can Starve for " & Worksheet.Name
        End
    End If
    
    ' return first can starve time
    Dim cell As Range
    For Each cell In Worksheet.Range("CI:CI")
        If cell.Value > 0 And IsNumeric(cell.Value) And isCanStarveInArray(cell.Row, dryerSkipArray) = False Then
            findFirstCanStarveTime = cell.Row
            Exit Function
        End If
        If cell.Value = "" Then
            Exit For
        End If
    Next cell
    
    'no can starve time found
    findFirstCanStarveTime = -1
End Function
