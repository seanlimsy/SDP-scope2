Option Explicit
Dim wb As Workbook
Dim D1Schedule As Worksheet
Dim D1Default As Worksheet
Dim D2Schedule As Worksheet
Dim D2Default As Worksheet
Dim PPThreshold As Worksheet
Dim PPTippingStation As Worksheet
Dim PPRateDSSheet As Worksheet
Dim workingDryerSchedule As Worksheet
Dim Silos As Worksheet

Dim logic4File as String
Dim logic4TextFile As Integer

Sub PPCanStretchMain()
    'Debugging:
    logic4File = "/Users/ben/Desktop/logic4.txt"
    logic4TextFile = FreeFile
    Open logic4File For Output As logic4TextFile 

    Application.AutoRecover.Enabled = False
    Print #logic4TextFile, "======== Initializing ========"
    initializeWorksheets
    'runOrDuplicateFile
    initializePPRateDS
    Print #logic4TextFile, "Done."

    Print #logic4TextFile, "======== Main Logic ========"
    ' Dim isLogic4Feasible As Boolean
    isLogic4Feasible = logic4()
    If isLogic4Feasible = True Then
        Print #logic4TextFile, "Worst Case PP Can inserted. Terminating Program"
    End If

    Close #logic4TextFile

End Sub

' ============================================= Setup Logic =============================================
Sub initializeWorksheets()
    Set wb = ThisWorkbook

    setWorksheet D1Schedule, "D1B1L65T"
    setWorksheet D1Default, "D1Sched"
    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet D2Default, "D2Sched"
    setWorksheet PPThreshold, "PP CAN ADDED THRESHOLD"
    setWorksheet PPTippingStation, "PP"
    setWorksheet Silos, "Silos"
    setWorksheet PPRateDSSheet, "PPRateDS"

    ' Update pivot table to correct setting PP sheet
    Dim PT As pivotTable, PI As PivotItem
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
    
    Application.CalculateFull
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wb.Sheets(worksheetName)
    Exit Sub
Err:
    reasonForStop = worksheetName & " is not in current workbook"
    End
End Sub

Sub runOrDuplicateFile()
    If InStr(wb.Name, " - (Original LTP w Additional PPCAN)") Then
        MsgBox "Running PPCan Stretching on this file."
    ElseIf InStr(wb.Name, " - (Original LTP wo Additional PPCAN)") = False Then
        MsgBox "Making a copy of BaseFile and Saving into an alternate file."
        makeCopy wb
        MsgBox "Duplication complete. Running PPCan Stretching on this file."
    End If
End Sub

Sub makeCopy(file)
    Dim p As Long
    ' duplicating base LTP
    With file
        p = InStrRev(.FullName, ".")
        .SaveCopyAs Left(.FullName, p - 1) & " - (Original LTP wo Additional PPCAN)" & Mid(.FullName, p)
    End With

    ' Saving this file for PP Can Stretching
    With file
        p = InStrRev(.FullName, ".")
        .SaveAs Left(.FullName, p - 1) & " - (Original LTP w Additional PPCAN)" & Mid(.FullName, p)
    End With
End Sub

Function addToArray(item, valueArray) As Double()
    ReDim Preserve valueArray(LBound(valueArray) To UBound(valueArray) + 1)
    valueArray(UBound(valueArray)) = item
    addToArray = valueArray
End Function

Sub initializePPRateDS()
    Dim lastRow As Integer, canStretchRow As Integer
    Dim tonPerHrOEEs() As Double, FPLbsPerSilos() As Double
    ReDim tonPerHrOEEs(1)
    ReDim FPLbsPerSilos(1)

    lastRow = PPRateDSSheet.Range("B1").End(xlDown).row
    If PPRateDSSheet.Range("A" & lastRow).Value = "PP - CAN - 5" Then 
        Exit Sub 
    End If

    canStretchRow = lastRow + 1

    Dim cell As Range
    Dim canRow As Integer
    For Each cell In PPRateDSSheet.Range("A2:A" & lastRow)
        If InStr(cell.Value, "CAN") Then
            canRow = cell.row
            tonPerHrOEEs = addToArray(PPRateDSSheet.Range("D" & canRow).Value, tonPerHrOEEs)
            FPLbsPerSilos = addToArray(PPRateDSSheet.Range("E" & canRow).Value, FPLbsPerSilos)
        End If
    Next cell

    Dim worstTonPerHourPOEE As Double, worstSA As Double, worstTonPerHourOEE As Double, worstFPLbsPerSilo As Double
    Dim indexWorstTonePerHourOEE As Integer

    ' worse case Ton per Hour OEE (smallest value)
    worstTonPerHourOEE = findMinNonZero(tonPerHrOEEs)
    indexWorstTonePerHourOEE = Application.Match(worstTonPerHourOEE, PPRateDSSheet.Range("D1:D" & lastRow), 0)
    worstTonPerHourPOEE = PPRateDSSheet.Range("B" & indexWorstTonePerHourOEE).Value
    worstSA = PPRateDSSheet.Range("C" & indexWorstTonePerHourOEE).Value

    ' worse case FP lbs per silo (smallest value)
    worstFPLbsPerSilo = findMinNonZero(FPLbsPerSilos)

    ' build PP - CAN - 5
    PPRateDSSheet.Range("A" & canStretchRow).Value = "PP - CAN - 5"
    PPRateDSSheet.Range("B" & canStretchRow).Value = worstTonPerHourPOEE
    PPRateDSSheet.Range("C" & canStretchRow).Value = FormatPercent(worstSA)
    PPRateDSSheet.Range("D" & canStretchRow).Value = worstTonPerHourOEE
    PPRateDSSheet.Range("E" & canStretchRow).Value = worstFPLbsPerSilo

    If worstTonPerHourPOEE * worstSA = worstTonPerHourOEE Then
        PPRateDSSheet.Range("A2:E2").Copy
        PPRateDSSheet.Range("A" & canStretchRow & ":" & "E" & canStretchRow).PasteSpecial xlFormats
    Else
        DeBug.Print "Error in Determining PPRateDS for PP-Can-5 (Stretching add). Check code-base ""initializePPRateDS"". Ending program."
        End
    End If
End Sub

Function findMinNonZero(arrayValues) As Double
    Dim smallest As Double, item As Variant
    smallest = Application.Max(arrayValues)
    For Each item In arrayValues
        If item <> 0 And smallest > item Then
            smallest = item
        End If
    Next item
    findMinNonZero = smallest
End Function

' ============================================= Main Logic =============================================
Function logic4()
    Dim mainSilo As Integer
    Dim otherSilo As Integer
     
    mainSilo = 16
    otherSilo = 6

    Dim isFeasible As Boolean
    isFeasible = stretchingCampaigns(mainSilo, otherSilo)
    logic4 = isFeasible

End Function

Function stretchingCampaigns(mainSilo, otherSilo)
    Dim PPCanStretching As Range
    Dim D1TipStatPivotTable As pivotTable, D2TipStatPivotTable As pivotTable

    Set D1TipStatPivotTable = pivotFromDryers(1)
    Set D2TipStatPivotTable = pivotFromDryers(2)
    D1TipStatPivotTable.RefreshTable
    D2TipStatPivotTable.RefreshTable

    Application.CalculateFull
    wb.RefreshAll

    ' Dim D1TipStatCanCOMax As Long, D2TipStatCanCOMax As Long
    ' D1TipStatCanCOMax = infoFromDryers(D1TipStatPivotTable)
    ' D2TipStatCanCOMax = infoFromDryers(D2TipStatPivotTable)
    
    ' arrays for determining which can starve to skip
    Dim d1Skip() As Integer
    Dim d2Skip() As Integer
    ReDim d1Skip(1)
    ReDim d2Skip(1)
    d1Skip(0) = 0
    d2Skip(0) = 0

    Dim D1PrevInsertTime As Double
    Dim D2PrevInsertTime As Double
    D1PrevInsertTime = -1
    D2PrevInsertTime = -2

    Dim count As Integer
    count = 1

    Do While True
        Print #logic4TextFile, "======== Attempt " & count & " ========"
        count = count + 1
        ' get row of campaign to insert
        Dim PPCampaignToInsert As Double
        PPCampaignToInsert = 2 ' fixed
        
        Print #logic4TextFile, "-- Finding CanStarveTime..."
        ' get row of insertion in schedule
        ' -1 if there is no can starve
        Dim D1FirstCanStarveTime As Double
        Dim D2FirstCanStarveTime As Double
        D1FirstCanStarveTime = findFirstCanStarveTime(D1Schedule, d1Skip)
        D2FirstCanStarveTime = findFirstCanStarveTime(D2Schedule, d2Skip)
        Print #logic4TextFile, "Done."

        Print #logic4TextFile, "-- Finding initial silo constraint..."
        ' get initial silo constraint violation time
        Dim initialSiloConstraintViolation
        If Silos.Range("K1").Value <> 0 And Silos.Range("K2").Value <> 0 Then
            If Silos.Range("K1").Value > Silos.Range("K2").Value Then
                initialSiloConstraintViolation = Silos.Range("K2").Value
            Else
                initialSiloConstraintViolation = Silos.Range("K1").Value
            End If
        ElseIf Silos.Range("K1").Value = 0 Then
            initialSiloConstraintViolation = Silos.Range("K2").Value
        ElseIf Silos.Range("K2").Value = 0 Then
            initialSiloConstraintViolation = Silos.Range("K1").Value
        Else
            initialSiloConstraintViolation = 0
        End If
        Print #logic4TextFile, "Done."
        Print #logic4TextFile, "-------"
        Print #logic4TextFile, "D1 First Can Starve Time Index: " & D1FirstCanStarveTime: Space 0
        Print #logic4TextFile, "D2 First Can Starve Time Index: " & D2FirstCanStarveTime

        Dim dryerCampaign As Integer
        dryerCampaign = determineDryerCampaignCanStretch(D1FirstCanStarveTime, D2FirstCanStarveTime, D1PrevInsertTime, D2PrevInsertTime)
        Print #logic4TextFile, "Dryer Campaign Value: " & dryerCampaign
        
        If dryerCampaign = 0 Then 'case: no more dryer slots
            Print #logic4TextFile, "All can starve slots used. Terminating Program"
            Print #logic4TextFile, "======== Attempt " & (count-1) & " Concluded ========"
            stretchingCampaigns = True
            Exit Function
        ElseIf dryerCampaign = 1 Then 'case: d1 PP campaign
            Print #logic4TextFile, "Add PPCan to Dryer 1"
            d1Skip = addPPCampaign(PPCampaignToInsert, D1Schedule, D1Default, D1FirstCanStarveTime, mainSilo, otherSilo, d1Skip, initialSiloConstraintViolation)
            D1PrevInsertTime = D1FirstCanStarveTime
            D2PrevInsertTime = -1
        ElseIf dryerCampaign = 2 Then 'case: d2 PP campaign
            Print #logic4TextFile, "Add PPCan to Dryer 2"
            d2Skip = addPPCampaign(PPCampaignToInsert, D2Schedule, D2Default, D2FirstCanStarveTime, mainSilo, otherSilo, d2Skip, initialSiloConstraintViolation)
            D1PrevInsertTime = -1
            D2PrevInsertTime = D2FirstCanStarveTime
        ElseIf dryerCampaign = 4 Then 'case: skip d1 can starve time
            Print #logic4TextFile, "Skipping D1"
            d1Skip = addItemToArray(D1FirstCanStarveTime, d1Skip)
        ElseIf dryerCampaign = 5 Then 'case: skip d2 can starve time
            Print #logic4TextFile, "Skipping D1"
            d2Skip = addItemToArray(D2FirstCanStarveTime, d2Skip)
        ElseIf dryerCampaign = 6 Then 'case: skip d1 and d2 can starve time
            Print #logic4TextFile, "Skipping D1 or D1 or Both D1 and D2 slots"
            d1Skip = addItemToArray(D1FirstCanStarveTime, d1Skip)
            d2Skip = addItemToArray(D2FirstCanStarveTime, d2Skip)
        End If
        Print #logic4TextFile, "======== Attempt " & (count-1) & " Concluded ========"
        Print #logic4TextFile, " "
continueLoop:
    Loop
    stretchingCampaigns = True
End Function

Function pivotFromDryers(identity)
    If identity = 1 Then
        'D1 - Tip Station (40H Gap)
        Set pivotFromDryers = PPTippingStation.PivotTables("PivotTableD1")
    ElseIf identity = 2 Then
        'D2 - Tip Station (40H Gap)
        Set pivotFromDryers = PPTippingStation.PivotTables("PivotTableD2")
    End If
End Function

Function infoFromDryers(pivotTable)
    Dim TipStatCanCORange As Range
    Dim TipStatCanCOMax As Long
    
    Set TipStatCanCORange = pivotTable.PivotFields("Sum of Can After CO Hrs").DataRange
    TipStatCanCOMax = Application.WorksheetFunction.Max(TipStatCanCORange)
    
    infoFromDryers = TipStatCanCOMax
End Function

Function determineDryerCampaignCanStretch(D1FirstCanStarveTime, D2FirstCanStarveTime, D1PrevInsertTime, D2PrevInsertTime) As Integer
    If D1FirstCanStarveTime = -1 And D2FirstCanStarveTime = -1 Then
        determineDryerCampaignCanStretch = 0
        Exit Function
    End If
    
    ' check PP sheet pivot table to determine tipping station availability
    Dim tippingStationAvailableTime As Double
    tippingStationAvailableTime = 0
    tippingStationAvailableTime = getTippingStationAvailableStartTime(D1FirstCanStarveTime, D2FirstCanStarveTime, D1PrevInsertTime, D2PrevInsertTime)
    
    Print #logic4TextFile, "Tipping Station Available Time: " & tippingStationAvailableTime: Space 0

    Dim D1CanStarveStartTime As Double
    Dim D2CanStarveStartTime As Double
    If D1FirstCanStarveTime <> -1 Then
        D1CanStarveStartTime = D1Schedule.Range("BK" & D1FirstCanStarveTime - 1).Value
    End If
    If D2FirstCanStarveTime <> -1 Then
        D2CanStarveStartTime = D2Schedule.Range("BK" & D2FirstCanStarveTime - 1).Value
    End If

    Print #logic4TextFile, "D1CanStarveStartTime: " & D1CanStarveStartTime: Space 0
    Print #logic4TextFile, "D2CanStarveStartTime: " & D2CanStarveStartTime

    ' If D1CanStarveStartTime >= tippingStationAvailableTime OR D2CanStarveStartTime >= tippingStationAvailableTime Then 
    '     End
    ' End If

    ' If D1CanStarveStartTime < tippingStationAvailableTime Then
    '     determineDryerCampaignCanStretch = 4 'if d1 can starve if before tipping station start then skip d1 time
    '     Exit Function
    ' End If

    If D1FirstCanStarveTime <> -1 And D2FirstCanStarveTime <> -1 Then 'case d1 and d2 both have slots
        If D1CanStarveStartTime < D2CanStarveStartTime Then
            If D1CanStarveStartTime > tippingStationAvailableTime Then
                determineDryerCampaignCanStretch = 1 'd1pp
            Else
                If D2CanStarveStartTime > tippingStationAvailableTime Then
                    determineDryerCampaignCanStretch = 2 'd2pp
                Else
                    determineDryerCampaignCanStretch = 6 'can't do pp on d1 and d2, no more db campaign so skip can starve time
                End If
            End If
        Else
            If D2CanStarveStartTime > tippingStationAvailableTime Then
                determineDryerCampaignCanStretch = 2 'd2pp
            Else
                If D1CanStarveStartTime > tippingStationAvailableTime Then
                    determineDryerCampaignCanStretch = 1 'd1pp
                Else
                    determineDryerCampaignCanStretch = 6 'can't do pp on d1 and d2, no more db campaign so skip can starve time
                End If
            End If
        End If
    ElseIf D1FirstCanStarveTime <> -1 And D2FirstCanStarveTime = -1 Then 'case only d1 has slots
        If D1CanStarveStartTime > tippingStationAvailableTime Then
            determineDryerCampaignCanStretch = 1 'd1pp
        Else
            determineDryerCampaignCanStretch = 4 'can't do pp on d1 and d2 is not available so skip can starve time
        End If
    ElseIf D1FirstCanStarveTime = -1 And D2FirstCanStarveTime <> -1 Then 'case only d2 has slots
        If D2CanStarveStartTime > tippingStationAvailableTime Then
            determineDryerCampaignCanStretch = 2 'd2pp
        Else
            determineDryerCampaignCanStretch = 5 'can't insert pp can and there are no more db campaigns so skip d2 can starve time
        End If

    End If
End Function

Function getTippingStationAvailableStartTime(D1FirstCanStarveTime, D2FirstCanStarveTime, D1PrevInsertTime, D2PrevInsertTime) As Double
    Dim tippingStationAvailableTime As Double
    Dim Column As Range, row As Range
    tippingStationAvailableTime = 0
    Dim PT As pivotTable
    For Each PT In PPTippingStation.PivotTables
        For Each Column In PT.ColumnRange
             If Column.Value = "Sum of Can After CO Hrs" Then
                For Each row In PT.RowRange
                    If IsNumeric(row.Value) Then
                        If PPTippingStation.Cells(row.row, Column.Column).Value > tippingStationAvailableTime Then
                            tippingStationAvailableTime = PPTippingStation.Cells(row.row, Column.Column).Value
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

Function addPPCampaign(PPCampaignToInsert, dryerSchedule, dryerDefaultSchedule, dryerFirstCanStarveTime, mainSilo, otherSilo, dryerSkipArray, initialSiloConstraintViolation) As Integer()
    ' PPCampaignToInsert = 2 fixed & no need to delete sample campaign

    ' decrement counter can be modified to determine the "steps" to reduce campaign load when it can't be inserted
    Dim decrementCounter As Double
    decrementCounter = 0.1

    ' boolean flag to determine if silo constraint is being violated
    Dim canAdd As Boolean
    canAdd = False

    Dim i As Double
    Dim FPLoadingWeight As Double
    FPLoadingWeight = PPThreshold.Range("J" & PPCampaignToInsert).Value
    Print #logic4TextFile, "++++++++++++++++++++++++"
    For i = 1 To 0.05 Step -decrementCounter
        ' insert to the row before the can starvation time
        PPThreshold.Range("A" & PPCampaignToInsert, "N" & PPCampaignToInsert).Copy
        dryerDefaultSchedule.Range("A" & dryerFirstCanStarveTime).Insert xlShiftDown
        dryerDefaultSchedule.Range("J" & dryerFirstCanStarveTime).Value = FPLoadingWeight * i
        dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
        Application.CalculateFull

        canAdd = checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerFirstCanStarveTime, initialSiloConstraintViolation)
        If canAdd = True Then
            Print #logic4TextFile, "Inserted @ " & dryerFirstCanStarveTime
            Print #logic4TextFile, "Inserted " & i & "th amount of 1 worst-case PPCan Campaign"
            Print #logic4TextFile, "++++++++++++++++++++++++"
            Exit For
        End If

        Print #logic4TextFile, "Reducing amount to " & (i - decrementCounter)
        dryerDefaultSchedule.Rows(dryerFirstCanStarveTime).EntireRow.Delete
        If i - decrementCounter < decrementCounter Then
            dryerSkipArray = addItemToArray(dryerFirstCanStarveTime, dryerSkipArray)
            dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
            Print #logic4TextFile, "Cannot be inserted at slot. Skipping."
            Print #logic4TextFile, "++++++++++++++++++++++++"
        End If
    Next
    Application.CalculateFull
    
    ' this is to ensure that the pivot table is updated after adding pp campaigns
    wb.RefreshAll
    
    addPPCampaign = dryerSkipArray
End Function

Function checkSiloConstraint(mainSilo, otherSilo, dryerSchedule, dryerInsertRow, initialSiloConstraintViolation) As Boolean
    If Silos.Range("J1").Value > 16 Or Silos.Range("J2").Value > 6 Then
        checkSiloConstraint = False
        Print #logic4TextFile, "Effect: Silo Constraint violated by insertion."
        Print #logic4TestFile, "PE Silo: " & Silos.Range("J1").Value & "; SG Silo: " & Silos.Range("J2").Value
        Exit Function
    End If
    checkSiloConstraint = True
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
        If cell.Value > 0 And IsNumeric(cell.Value) And isCanStarveInArray(cell.row, dryerSkipArray) = False Then
            findFirstCanStarveTime = cell.row
            Exit Function
        End If
        If cell.Value = "" Then
            Exit For
        End If
    Next cell
    
    'no can starve time found
    findFirstCanStarveTime = -1
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

