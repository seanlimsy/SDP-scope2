'create worksheets as global variables
Dim wb As Workbook
Dim D1Schedule As Worksheet
Dim D1Default As Worksheet
Dim D2Schedule As Worksheet
Dim D2Default As Worksheet
Dim DBSchedule As Worksheet
Dim PPCanSchedule As Worksheet
Dim PPTippingStation As Worksheet
Dim Silos As Worksheet
Dim D1DefaultOriginal As Worksheet
Dim D2DefaultOriginal As Worksheet
Dim D1TipStatPivotTable As pivotTable
Dim D2TipStatPivotTable As pivotTable

Sub calculateAll()
    Application.CalculateFull
    If Not Application.CalculationState = xlDone Then 
        DoEvents
    End If
    D1TipStatPivotTable.RefreshTable
    D2TipStatPivotTable.RefreshTable
End Sub

Sub resetAll()
    ' Print #logic1TextFile, " ": Space 0
    ' Print #logic1TextFile, "================ Insert fail due to Silo Constraint ================": Space 0
    ' Print #logic1TextFile, "++++ Resetting Sheets for relaxed Constraint ++++": Space 0
    initializeWorksheets
    
    ' Print #logic1TextFile, "--- Reverting Schedules..."
    D1Default.Range("A:N").Value = D1DefaultOriginal.Range("A:N").Value
    D2Default.Range("A:N").Value = D2DefaultOriginal.Range("A:N").Value
    
    PPCanSchedule.Range("A:N").Value = PPCanSchedule.Range("R:AD").Value
    DBSchedule.Range("A:O").Value = DBSchedule.Range("Q:AE").Value
    
    D1Schedule.Range("A:N").Value = D1Default.Range("A:N").Value
    D2Schedule.Range("A:N").Value = D2Default.Range("A:N").Value
    calculateAll
    ' Print #logic1TextFile, "Done."

    ' Print #logic1TextFile, "--- Reverting CIP & Blockage..."
    ' ===== reset cip and dryer blockage cells =====
    Dim lastRowD1 As Integer
    Dim lastRowD2 As Integer
    lastRowD1 = D1Schedule.Range("AF1").End(xlDown).Row
    lastRowD2 = D2Schedule.Range("AF1").End(xlDown).Row

    D1Schedule.Range("AF2:AF" & lastRowD1).Formula = "=If(ISBLANK(A2),"""",IF(G2=""DR"",IF(SUMIFS(V:V,O:O,"">""&AE2,O:O,""<=""&O2)>='Evap DryCIP'!$T$2,'Evap DryCIP'!$T$3,0),0))"
    D2Schedule.Range("AF2:AF" & lastRowD2).Formula = "=IF(ISBLANK(A2),"""",IF(G2=""DR"",IF(SUMIFS(V:V,O:O,"">""&AE2,O:O,""<=""&O2)>='Evap DryCIP'!$T$5,'Evap DryCIP'!$T$6,0),0))"
    calculateAll

    D1Schedule.Range("AI2:AI" & lastRowD1).Value = 0
    D2Schedule.Range("AI2:AI" & lastRowD2).Value = 0
    calculateAll
    wb.refreshAll
    wb.Save
    ' Print #logic1TextFile, "Done."
    ' Print #logic1TextFile, "++++ Reset Done. Reattempting ++++": Space 0
    ' Print #logic1TextFile, " ": Space 0
End Sub

Sub main()
    ' turn off autosave
    Application.AutoRecover.Enabled = False
    ' Print #logic1TextFile, "======== Initializing ========"
    ' Print #logic1TextFile, "Program Started @ " & Now
    initializeWorksheets
    ' Print #logic1TextFile, "Done."
    
    ' Print #logic1TextFile, "======== Main Logic ========"
    ' Dim isLogic1Feasible As Boolean
    isLogic1Feasible = logic1()
    If isLogic1Feasible = False Then
        resetAll
        ' Print #logic1TextFile, "PP-Can and 100DB Campaigns cannot be inserted even after setting silo constraint to 22(6)."
        reasonForStop = "Max PE Silo Constraint Reached."
        ' Print #logic1Textfile, "Terminating Program.": Space 0
    End If
    
    ' Print #logic1TextFile, "logic1 Ended @ " & Now
    Close #logic1TextFile
End Sub

Sub initializeWorksheets()
    'note that the worksheets have to be in the same workbook
    'have the PPCan and 100DB schedules in the same workbook
    Set wb = ThisWorkbook
    setWorksheet D1Schedule, "D1B1L65T"
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
    
    Set D1TipStatPivotTable = PPTippingStation.PivotTables("PivotTableD1")
    Set D2TipStatPivotTable = PPTippingStation.PivotTables("PivotTableD2")

    'Include Silo Constraint presense for PE and SG
    Silos.Range("R8:S8").Value = "PE"
    Silos.Range("R9").Formula = "=MAXIFS(D1B1L65T!AJ:AJ,D1B1L65T!AJ:AJ,""<=""&Silos!$K$1,D1B1L65T!AP:AP,"">=1"")"
    Silos.Range("R10").Formula = "=MAXIFS(D2B1L3B3B4L45T!AJ:AJ,D2B1L3B3B4L45T!AJ:AJ,""<=""&Silos!$K$1,D2B1L3B3B4L45T!AP:AP,"">=1"")"
    Silos.Range("S9").Formula = "=IF(K1-R9<0.5,""YES"",""NO"")"
    Silos.Range("S10").Formula = "=IF(K1-R10<0.5,""YES"",""NO"")"
    
    Silos.Range("T8:U8").Value = "SG"
    Silos.Range("T9").Formula = "=MAXIFS(D1B1L65T!AJ:AJ,D1B1L65T!AJ:AJ,""<=""&Silos!$K$2,D1B1L65T!AP:AP,"">=1"")"
    Silos.Range("T10").Formula = "=MAXIFS(D2B1L3B3B4L45T!AJ:AJ,D2B1L3B3B4L45T!AJ:AJ,""<=""&Silos!$K$2,D2B1L3B3B4L45T!AP:AP,"">=1"")"
    Silos.Range("U9").Formula = "=IF(K2-T9<0.5,""YES"",""NO"")"
    Silos.Range("U10").Formula = "=IF(K2-T10<0.5,""YES"",""NO"")"
    
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wb.Sheets(worksheetName)
    Exit Sub
Err:
    reasonForStop = worksheetName & " is not in current workbook"
    End
End Sub

Function logic1()
    ' Dim mainSilo As Integer
    ' Dim otherSilo As Integer
    mainSilo = 16
    otherSilo = 6
    
    Dim isFeasible As Boolean
    isFeasible = False

    Dim reportWS As Worksheet
    Dim maxPESilos As Integer
    Set reportWS = wb.Worksheets("Program Report Page")
    maxPESilos = reportWS.range("B10").Value

    Do While mainSilo <= maxPESilos
        ' Print #logic1TextFile, "Current PE Silo Allowance: " & mainSilo: Space 0
        ' Print #logic1TextFile, "Current SG Silo Allowance: " & otherSilo: Space 0
        isFeasible = insertPPCan100DBCampaigns(mainSilo, otherSilo)
        If isFeasible = True Then
            Exit Do
        End If
        resetAll
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
    
    Dim D1PrevInsertTime as Double
    Dim D2PrevInsertTime as Double
    D1PrevInsertTime = -1
    D2PrevInsertTime = -1

    Dim count As Integer
    count = 1

    Do While True
        ' Print #logic1TextFile, "======== Attempt " & count & " ========"
        count = count + 1
        
        ' Print #logic1TextFile, "-- Finding PP / DB Campaign to insert..."
        ' get row of campaign to insert
        ' -1 if there is no campaign
        Dim PPCampaignToInsert As Double
        Dim DBCampaignToInsert As Double
        PPCampaignToInsert = findNextCampaignToInsert(PPCanSchedule)
        DBCampaignToInsert = findNextCampaignToInsert(DBSchedule)
        ' Print #logic1TextFile, "Done."
        ' Print #logic1TextFile, "-------"
        ' Print #logic1TextFile, "PP Campaign to insert: " & PPCampaignToInsert: Space 0
        ' Print #logic1TextFile, "DB Campaign to insert: " & DBCampaignToInsert: Space 0

        ' Print #logic1TextFile, "-- Finding CanStarveTime..."
        ' get row of insertion in schedule
        ' -1 if there is no can starve
        Dim D1FirstCanStarveTime As Double
        Dim D2FirstCanStarveTime As Double
        D1FirstCanStarveTime = findFirstCanStarveTime(D1Schedule, d1Skip)
        D2FirstCanStarveTime = findFirstCanStarveTime(D2Schedule, d2Skip)
        ' Print #logic1TextFile, "Done."
        ' Print #logic1TextFile, "-------"
        ' Print #logic1TextFile, "D1 First Can Starve Time Index: " & D1FirstCanStarveTime: Space 0
        ' Print #logic1TextFile, "D2 First Can Starve Time Index: " & D2FirstCanStarveTime: Space 0

        ' Print #logic1TextFile, "-- Finding initial silo constraint..."
        ' get initial silo constraint violation time
        Dim initialSiloConstraintViolation As Double
        if Silos.Range("K1").Value <> 0 and silos.range("K2").value <> 0 then
            if silos.range("K1").value > silos.range("K2").value then
                initialSiloConstraintViolation = silos.range("K2").value
            Else
                initialSiloConstraintViolation = silos.range("K1").value
            end if
        ElseIf Silos.Range("K1").Value = 0 then
            initialSiloConstraintViolation = Silos.Range("K2").Value
        ElseIf Silos.Range("K2").Value = 0 then
            initialSiloConstraintViolation = Silos.Range("K1").Value
        else
            initialSiloConstraintViolation = 0 
        end if
        ' Print #logic1TextFile, "Done."
        ' Print #logic1TextFile, "-------"
        ' Print #logic1TextFile, "Initial Silo Constraint Violation: " & initialSiloConstraintViolation: Space 0

        ' Print #logic1TextFile, "-- Finding dryer campaign value..."
        ' get which dryer and which campaign to insert
        Dim dryerCampaign As Integer
        dryerCampaign = determineDryerCampaign(D1FirstCanStarveTime, D2FirstCanStarveTime, PPCampaignToInsert, DBCampaignToInsert, D1PrevInsertTime, D2PrevInsertTime)
        ' Print #logic1TextFile, "Done."
        ' Print #logic1TextFile, "-------"
        ' Print #logic1TextFile, "Dryer Campaign Value: " & dryerCampaign
        
        If dryerCampaign = -2 Then 'case: db campaigns but no more d2 slots (infeasible solution)
            ' Print #logic1TextFile, "DB campaigns remaining but no more can starvation slots in dryer 2. Exiting Program.": Space 0
            ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========": Space 0
            ' Print #logic1TextFile, "+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-+-": Space 0
            ' Print #logic1TextFile, "DB Campaigns remaining but there are no more can starvation slots in D2.": Space 0
            ' Print #logic1TextFile, "Resetting Schedules and increasing dryer allowances": Space 0
            End
        ElseIf dryerCampaign = -1 Then 'case: no more campaigns left
            ' Print #logic1TextFile, "All campaigns Inserted. Running dryer blockage on all remaining silo constraint violations. ": Space 0
            ' run dryer blockage on remaining silo constraint violations
            programModule2.dryerBlockDelayMain 9999999
            ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========": Space 0
            insertPPCan100DBCampaigns = True
            Exit Function
        ElseIf dryerCampaign = 0 Then 'case: no more dryer slots
            ' Print #logic1TextFile, "All can starvation slots used. Increasing silo constraint": Space 0
            ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========": Space 0
            insertPPCan100DBCampaigns = False
            Exit Function
        ElseIf dryerCampaign = 1 Then 'case: d1 pp campaign
            If D1Schedule.Range("BI" & D1FirstCanStarveTime - 1).Value > initialSiloConstraintViolation and initialsiloconstraintviolation <> 0 Then
                    ' Print #logic1TextFile, "Effect: Encountered silo constraint violation prior to insertion point. Moving to solve violation first.": Space 0
                    programModule2.dryerBlockDelayMain D1Schedule.Range("BI" & D1FirstCanStarveTime - 1).Value
                    ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========": Space 0
                    ' Print #logic1TextFile, " "
                    GoTo continueLoop
            End If
            ' Print #logic1TextFile, "Adding PP campaign to dryer 1"
            d1Skip = addPPCampaign(PPCampaignToInsert, D1Schedule, D1Default, D1FirstCanStarveTime, mainSilo, otherSilo, d1Skip, initialSiloConstraintViolation)
            D1PrevInsertTime = D1FirstCanStarveTime
            D2PrevInsertTime = -1
        ElseIf dryerCampaign = 2 Then 'case: d2 pp campaign
           If D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value > initialSiloConstraintViolation and initialsiloconstraintviolation <> 0 Then
                    ' Print #logic1TextFile, "Effect: Encountered silo constraint violation prior to insertion point. Moving to solve violation first.": Space 0
                    programModule2.dryerBlockDelayMain D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value
                    ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========": Space 0
                    ' Print #logic1TextFile, " "
                    GoTo continueLoop
            End If
            ' Print #logic1TextFile, "Adding PP campaign to dryer 2"
            d2Skip = addPPCampaign(PPCampaignToInsert, D2Schedule, D2Default, D2FirstCanStarveTime, mainSilo, otherSilo, d2Skip, initialSiloConstraintViolation)
            D2PrevInsertTime = D2FirstCanStarveTime
            D1PrevInsertTime = -1
        ElseIf dryerCampaign = 3 Then 'case: d2 db campaign
            If D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value > initialSiloConstraintViolation and initialsiloconstraintviolation <> 0 Then
                    ' Print #logic1TextFile, "Effect: Encountered silo constraint violation prior to insertion point. Moving to solve violation first.": Space 0
                    programModule2.dryerBlockDelayMain D2Schedule.Range("BI" & D2FirstCanStarveTime - 1).Value
                    ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========": Space 0
                    ' Print #logic1TextFile, " "
                    GoTo continueLoop
            End If
            ' Print #logic1TextFile, "Adding DB campaign to dryer 2"
            d2Skip = addDBCampaign(DBCampaignToInsert, D2Schedule, D2Default, D2FirstCanStarveTime, mainSilo, otherSilo, d2Skip, initialSiloConstraintViolation)
        ElseIf dryerCampaign = 4 Then 'case: skip d1 can starve time
            ' Print #logic1TextFile, "Skipping D1"
            d1Skip = addItemToArray(D1FirstCanStarveTime, d1Skip)
        ElseIf dryerCampaign = 5 Then 'case: skip d2 can starve time
            ' Print #logic1TextFile, "Skipping D2"
            d2Skip = addItemToArray(D2FirstCanStarveTime, d2Skip)
        ElseIf dryerCampaign = 6 Then 'case: skip d1 and d2 can starve time
            ' Print #logic1TextFile, "Skipping D1 or D1 or Both D1 and D2 slots"
            d1Skip = addItemToArray(D1FirstCanStarveTime, d1Skip)
            d2Skip = addItemToArray(D2FirstCanStarveTime, d2Skip)
        End If
        ' Print #logic1TextFile, "======== Attempt " & (count-1) & " Concluded ========"
        ' Print #logic1TextFile, " "
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
    ' Print #logic1TextFile, "++++++++++++++++++++++++"
    For i = lastRow To DBCampaignToInsert Step -1
        ' insert DB campaign
        DBSchedule.Range("A" & DBCampaignToInsert, "M" & i).Copy
        dryerDefaultSchedule.Range("A" & dryerFirstCanStarveTime).Insert xlShiftDown
        dryerSchedule.Range("A:M").Value = dryerDefaultSchedule.Range("A:M").Value
        calculateAll
        
        ' check if the added campaign satisfies silo constraint
        canAdd = checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerFirstCanStarveTime, initialSiloConstraintViolation)
        If canAdd = True Then
            DBSchedule.Range("A" & DBCampaignToInsert, "O" & i).Delete xlShiftUp
            ' Print #logic1TextFile, "-----------"
            ' Print #logic1TextFile, "Inserted @ " & dryerFirstCanStarveTime: Space 0
            ' Print #logic1TextFile, "Inserted " & (i-1) & " campaign(s) from window": Space 0
            ' Print #logic1TextFile, "-----------"
            ' case not 16(6) - run dryer blockage
            If mainSilo <> 16 Then
                ' Print #logic1TextFile, "Silo allowance attained. Inducing dryer blockage/delay.": Space 0
                ' Print #logic1TextFile, "Induced Delay/Block @ " & initialSiloConstraintViolation: Space 0
                If initialSiloConstraintViolation = Silos.Range("K1").Value Or initialSiloConstraintViolation = Silos.Range("K2").Value Then
                    Exit For
                Else
                    If Silos.Range("K1").Value > Silos.Range("K2").Value Then
                        programModule2.dryerBlockDelayMain Silos.Range("K1").Value + 1
                    Else
                        programModule2.dryerBlockDelayMain Silos.Range("K2").Value + 1
                    End If
                End If
            End If
            ' Print #logic1TextFile, "++++++++++++++++++++++++"
            Exit For
        End If
        
        ' Print #logic1TextFile, "Reducing amount to " & (i - 2)
        dryerDefaultSchedule.Rows(dryerFirstCanStarveTime & ":" & (dryerFirstCanStarveTime + (i - DBCampaignToInsert))).EntireRow.Delete xlShiftUp
        ' case nothing can be added
        If i <= DBCampaignToInsert Then
            dryerSkipArray = addItemToArray(dryerFirstCanStarveTime, dryerSkipArray)
            dryerSchedule.Range("A:M").Value = dryerDefaultSchedule.Range("A:M").Value
            ' Print #logic1TextFile, "Cannot be inserted at slot. Skipping."
            ' Print #logic1TextFile, "++++++++++++++++++++++++"
            Exit For
        End If
    Next
        
    calculateAll    
    addDBCampaign = dryerSkipArray
End Function

Function addPPCampaign(PPCampaignToInsert, dryerSchedule, dryerDefaultSchedule, dryerFirstCanStarveTime, mainSilo, otherSilo, dryerSkipArray, initialSiloConstraintViolation) As Integer()
    
    ' decrement counter can be modified to determine the "steps" to reduce campaign load when it can't be inserted
    Dim decrementCounter As Double
    decrementCounter = 0.2

    ' boolean flag to determine if silo constraint is being violated
    Dim canAdd As Boolean
    canAdd = False

    Dim i As Double
    ' Print #logic1TextFile, "++++++++++++++++++++++++"
    For i = 1 To decrementCounter Step -decrementCounter
        ' insert to the row before the can starvation time
        PPCanSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Copy
        dryerDefaultSchedule.Range("A" & dryerFirstCanStarveTime).Insert xlShiftDown
        dryerDefaultSchedule.Range("J" & dryerFirstCanStarveTime).Value = dryerDefaultSchedule.Range("J" & dryerFirstCanStarveTime).Value * i
        dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
        calculateAll

        canAdd = checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerFirstCanStarveTime, initialSiloConstraintViolation)
        If canAdd = True Then
            ' Print #logic1TextFile, "-----------"
            ' Print #logic1TextFile, "Inserted @ " & dryerFirstCanStarveTime: Space 0
            ' Print #logic1TextFile, "Inserted " & i & "th amount of campaign": Space 0
            ' Print #logic1TextFile, "-----------"
            If i = 1 Then
                PPCanSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Delete xlShiftUp
            Else
                PPCanSchedule.Range("J" & PPCampaignToInsert).Value = PPCanSchedule.Range("J" & PPCampaignToInsert).Value * (1 - i)
            End If
            ' case not 16(6) - run dryer blockage
            If mainSilo <> 16 Then
                ' Print #logic1TextFile, "Silo allowance retained. Inducing dryer blockage/delay.": Space 0
                ' Print #logic1TextFile, "Induced Delay/Block @ " & initialSiloConstraintViolation: Space 0
                If initialSiloConstraintViolation = Silos.Range("K1").Value Or initialSiloConstraintViolation = Silos.Range("K2").Value Then
                    Exit For
                Else
                    If Silos.Range("K1").Value > Silos.Range("K2").Value Then
                        programModule2.dryerBlockDelayMain Silos.Range("K1").Value + 1
                    Else
                        programModule2.dryerBlockDelayMain Silos.Range("K2").Value + 1
                    End If
                End If
            End If
            ' Print #logic1TextFile, "++++++++++++++++++++++++"
            Exit For
        End If

        ' Print #logic1TextFile, "Reducing amount to " & (i - decrementCounter)
        dryerDefaultSchedule.Rows(dryerFirstCanStarveTime).EntireRow.Delete xlShiftUp
        If i - decrementCounter < decrementCounter Then
            dryerSkipArray = addItemToArray(dryerFirstCanStarveTime, dryerSkipArray)
            dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
            ' Print #logic1TextFile, "Cannot be inserted at slot. Skipping."
            ' Print #logic1TextFile, "++++++++++++++++++++++++"
        End If
    Next

    calculateAll
    ' this is to ensure that the pivot table is updated after adding pp campaigns
    wb.refreshAll
    
    addPPCampaign = dryerSkipArray
End Function

Function checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerInsertRow, initialSiloConstraintViolation) As Boolean
    If initialSiloConstraintViolation = 0 then
        If Silos.Range("K1").Value <> 0 or Silos.Range("K2").Value <> 0 then
            checkSiloConstraint = False
            ' Print #logic1TextFile, "Effect: Silo Constraint violated by insertion."
            ' Print #logic1TextFile, "PE Silo: " & Silos.Range("J1").Value & "; SG Silo: " & Silos.Range("J2").Value
            Exit Function
        End If
    End If
    Dim siloCheckTimeStart As Double
    Dim siloCheckTimeEnd As Double
    siloCheckTimeStart = dryerSchedule.Range("BY" & dryerInsertRow).Value 'silo entry hour
    'siloCheckTimeEnd = dryerSchedule.Range("BB" & dryerInsertRow).Value
    
    ' iterate through silos sheet to find If the silo constraint is being violated by the campaign insertion
    Dim i As Double
    For i = 2 To (2 ^ 15) - 1 Step 1
        If Silos.Range("A" & i).Value >= siloCheckTimeStart And Silos.Range("A" & i).Value < initialSiloConstraintViolation Then
            If Silos.Range("D" & i).Value > mainSilo Or Silos.Range("G" & i).Value > otherSilo Then
                ' Print #logic1TextFile, "Effect: Silo Constraint violated by insertion."
                ' Print #logic1TextFile, "PE Silo: " & Silos.Range("J1").Value & "; SG Silo: " & Silos.Range("J2").Value
                checkSiloConstraint = False
                Exit Function
            End If
        End If
    Next
    
    checkSiloConstraint = True
End Function
      
Function determineDryerCampaign(D1FirstCanStarveTime, D2FirstCanStarveTime, PPCampaignToInsert, DBCampaignToInsert, D1PrevInsertTime, D2PrevInsertTime) As Integer
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
    tippingStationAvailableTime = getTippingStationAvailableStartTime(D1FirstCanStarveTime, D2FirstCanStarveTime, D1PrevInsertTime, D2PrevInsertTime)
    
    ' Print #logic1TextFile, "Tipping Station Available Time: " & tippingStationAvailableTime: Space 0

    Dim D1CanStarveStartTime As Double
    Dim D2CanStarveStartTime As Double
    If D1FirstCanStarveTime <> -1 Then
        D1CanStarveStartTime = D1Schedule.Range("BK" & D1FirstCanStarveTime - 1).Value
    End If
    If D2FirstCanStarveTime <> -1 Then
        D2CanStarveStartTime = D2Schedule.Range("BK" & D2FirstCanStarveTime - 1).Value
    End If

    ' Print #logic1TextFile, "D1CanStarveStartTime: " & D1CanStarveStartTime: Space 0
    ' Print #logic1TextFile, "D2CanStarveStartTime: " & D2CanStarveStartTime: Space 0

    If D1CanStarveStartTime < tippingStationAvailableTime And D1CanStarveStartTime <> 0 Then
        determineDryerCampaign = 4 'If d1 can starve is before tipping station start then skip d1 time
        Exit Function
    End If

    If D1FirstCanStarveTime <> -1 And D2FirstCanStarveTime <> -1 Then 'case d1 and d2 both have slots
        If PPCampaignToInsert <> -1 And DBCampaignToInsert <> -1 Then 'case both pp and db campaigns available
            If D1CanStarveStartTime < D2CanStarveStartTime + 50 Then
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

Function getTippingStationAvailableStartTime(D1FirstCanStarveTime, D2FirstCanStarveTime, D1PrevInsertTime, D2PrevInsertTime) As Double
    Dim tippingStationAvailableTime As Double
    Dim Column As Range, row As Range

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
        If D1PrevInsertTime <> -1 And D1FirstCanStarveTime = D1PrevInsertTime + 1 Then
            getTippingStationAvailableStartTime = tippingStationAvailableTime
        ElseIf D2PrevInsertTime <> -1 And D2FirstCanStarveTime = D2PrevInsertTime + 1 Then
            getTippingStationAvailableStartTime = tippingStationAvailableTime
        Else
            tippingStationAvailableTime = tippingStationAvailableTime + 40
        End If
    End If
    getTippingStationAvailableStartTime = tippingStationAvailableTime
End Function

' checks the FP Loading Per Batch (If 0 then the campaign has been inserted)
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
            reasonForStop = "Cell CI1 is not set to Can Starve for " & Worksheet.Name
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