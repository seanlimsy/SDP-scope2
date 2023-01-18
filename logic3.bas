Option Explicit
Dim wb As Workbook
Dim D2Schedule As Worksheet
Dim PPPouchSchedule As Worksheet
Dim PPTippingStation As Worksheet
Dim PPRateDSSheet As Worksheet
Dim pouchInsertSpace As Worksheet
Dim D2Default As Worksheet
Dim Silos As Worksheet

' Dim logic3File as String
' Dim logic3TextFile As Integer

Sub ppPouchMain()
    'Debugging:
    Open logic3File For Output As logic3TextFile 

    Application.AutoRecover.Enabled = False
    Print #logic3TextFile, "======== Initializing ========"
    initializeWorksheets

    Dim numberPouchCampaigns As Integer
    numberPouchCampaigns = initializePouchInsertion 
    Print #logic3TextFile, "Done."

    ' 'To Remove
    ' PPPouchSchedule.Select
    reportWS.Select

    Print #logic3TextFile, "======== Main Logic ========"
    ' Dim isLogic3Feasible As Boolean
    isLogic3Feasible = logic3(numberPouchCampaigns)
    If isLogic3Feasible = False Then
        Print #logic3TextFile, "PP-Pouch Campaigns cannot be inserted by automated process. Terminating Program."
    ElseIf isLogic3Feasible = True Then 
        Print #logic3TextFile, "All PP-Pouch Campaigns inserted. Ending Stage 3."
    End If

    Close #logic3TextFile

End Sub

' ============================================= Setup Logic =============================================
Function initializePouchInsertion()
    Dim countPouchCampaigns As Long
    
    countPouchCampaigns = initializePouchWorksheets
    getPotentialSlots countPouchCampaigns

    initializePouchInsertion = countPouchCampaigns

End Function

Sub initializeWorksheets()
    'Without Initialising into same workbook
    Set wb = ThisWorkbook

    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet D2Default, "D2Sched"
    setWorksheet Silos, "Silos"
    setWorksheet PPPouchSchedule, "PP PCH"
    setWorksheet PPTippingStation, "PP"
    setWorksheet PPRateDSSheet, "PPRateDS"
    setWorksheet pouchInsertSpace, "PP PCH SPACE" 'Create new sheet for finding idle times based on pivot tables

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

Function initializePouchWorksheets()
    Dim Pouch_OriginalDetails As Range
    Dim lastRow As Long, countPouches As Long
    
    'Copy and Paste original data to the side
    countPouches = PPPouchSchedule.Cells(2, 1).End(xlDown).Row
    lastRow = PPPouchSchedule.Cells(2, 19).End(xlDown).Row
    PPPouchSchedule.Range("S2:AF" & lastRow).ClearContents
    Set Pouch_OriginalDetails = PPPouchSchedule.Range("A2:N" & countPouches)
    Pouch_OriginalDetails.Copy
    PPPouchSchedule.Range("S2:AF" & countPouches).PasteSpecial xlPasteValues
    
    'Calculate Pouch Fill Times
    Dim effective_fp_tonnes_perhr As Double
    Dim Pouch_Rates As Range
    Set Pouch_Rates = PPRateDSSheet.Range(PPRateDSSheet.Range("D2"), PPRateDSSheet.Range("D2").End(xlDown))
    
    effective_fp_tonnes_perhr = Application.WorksheetFunction.Min(Pouch_Rates)
    PPPouchSchedule.Range("Q1").Value = "Effective FP Tonnes per Hour"
    PPPouchSchedule.Range("Q2:Q" & countPouches).Formula = "=J2/2.2/1000/" & effective_fp_tonnes_perhr
    PPPouchSchedule.Range("Q2:Q" & countPouches).Copy
    PPPouchSchedule.Range("Q2:Q" & countPouches).PasteSpecial xlPasteValues
    
    initializePouchWorksheets = countPouches
End Function

Sub getPotentialSlots(countPouches)
    Dim D1TipStatPivotTable As pivotTable
    Dim D2TipStatPivotTable As pivotTable
    Dim D1TipStatStart As Range, D1TipStatEnd As Range, D2TipStatStart As Range, D2TipStatEnd As Range
    
    Set D1TipStatPivotTable = PPTippingStation.PivotTables("PivotTableD1")
    Set D2TipStatPivotTable = PPTippingStation.PivotTables("PivotTableD2")
    
    D1TipStatPivotTable.RefreshTable
    D2TipStatPivotTable.RefreshTable
    
    Set D1TipStatStart = getPivotEntry(D1TipStatPivotTable, 1)
    Set D1TipStatEnd = getPivotEntry(D1TipStatPivotTable, 2)
    Set D2TipStatStart = getPivotEntry(D2TipStatPivotTable, 1)
    Set D2TipStatEnd = getPivotEntry(D2TipStatPivotTable, 2)

    getTipStatIdleTimes countPouches, D1TipStatStart, D1TipStatEnd, D2TipStatStart, D2TipStatEnd
    getPchLineIdleTimes
    findIntersectionsOfIdleTimes countPouches
    
End Sub

Function getPivotEntry(pivotTable, identity)
    If identity = 1 Then
       Set getPivotEntry = pivotTable.PivotFields("Sum of Silo Entry Hr").DataRange
    ElseIf identity = 2 Then
        Set getPivotEntry = pivotTable.PivotFields("Sum of Can After CO Hrs").DataRange
    End If

End Function

Sub getTipStatIdleTimes(countPouches, D1TipStatStart, D1TipStatEnd, D2TipStatStart, D2TipStartEnd)
    Dim startRow As Integer, endRow As Integer
    
    startRow = 5
    endRow = startRow + D1TipStatStart.Count - 1
    D1TipStatStart.Copy
    pouchInsertSpace.Range("AA4:AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D1TipStatEnd.Count
    D1TipStatEnd.Copy
    pouchInsertSpace.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D2TipStatStart.Count
    D2TipStatStart.Copy
    pouchInsertSpace.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D2TipStartEnd.Count
    D2TipStartEnd.Copy
    pouchInsertSpace.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    Dim PPStatInUse As Range
    Set PPStatInUse = pouchInsertSpace.Range("AA3:AA" & endRow)
    PPStatInUse.Sort Key1:=pouchInsertSpace.Range("AA3"), Order1:=xlAscending, Header:=xlYes
    
    pouchInsertSpace.Range("A2").Value = "TipStation Idle"
    pouchInsertSpace.Range("A3").Value = "Start"
    pouchInsertSpace.Range("B3").Value = "End"
 
    Dim i As Integer, j As Integer
    Dim positivetime_start As Double, positivetime_end As Double
    i = 4
    j = 4
    Do Until i >= endRow + 1
        positivetime_start = pouchInsertSpace.Range("AA" & i).Value
        If positivetime_start >= 0 Then
            If i = 4 Then
                positivetime_end = positivetime_start
                positivetime_start = 0
                i = i + 1
            Else
                positivetime_end = pouchInsertSpace.Range("AA" & i + 1).Value
                i = i + 2
            End If
            pouchInsertSpace.Range("A" & j).Value = positivetime_start
            pouchInsertSpace.Range("B" & j).Value = positivetime_end
            j = j + 1
        Else
            i = i + 1
        End If
    Loop
    pouchInsertSpace.Range("AA4:AA" & endRow).Clear
    pouchInsertSpace.Range("B" & j - 1).Value = 5000
    
End Sub

Sub getPchLineIdleTimes()
    Dim PchLine_Starts As Range, PchLine_Ends As Range
    Dim ColumnNumber_PchStart As Long, ColumnNumber_PchEnd As Long
    Dim PchLine_Start_ColLetter As String, PchLine_End_ColLetter As String
    
    'Enter error checker for retrieving value
    ColumnNumber_PchStart = WorksheetFunction.Match("Pch Start", D2Schedule.Range("A1:CI1"), 0)
    PchLine_Start_ColLetter = Split(Cells(1, ColumnNumber_PchStart).Address, "$")(1) & "2"
    ColumnNumber_PchEnd = WorksheetFunction.Match("Pch End", D2Schedule.Range("A1:CI1"), 0)
    PchLine_End_ColLetter = Split(Cells(1, ColumnNumber_PchEnd).Address, "$")(1) & "2"

    Set PchLine_Starts = D2Schedule.Range(D2Schedule.Range(PchLine_Start_ColLetter), D2Schedule.Range(PchLine_Start_ColLetter).End(xlDown))
    Set PchLine_Ends = D2Schedule.Range(D2Schedule.Range(PchLine_End_ColLetter), D2Schedule.Range(PchLine_End_ColLetter).End(xlDown))
    
    PchLine_Starts.Copy
    pouchInsertSpace.Range("AA1").Value = "PouchLineInUse_Start"
    pouchInsertSpace.Range("AA2:AA" & PchLine_Starts.Count + 1).PasteSpecial xlPasteValues
    PchLine_Ends.Copy
    pouchInsertSpace.Range("AB1").Value = "PouchLineInUse_End"
    pouchInsertSpace.Range("AB2:AB" & PchLine_Ends.Count + 1).PasteSpecial xlPasteValues
    
    pouchInsertSpace.Select
    pouchInsertSpace.Range("AA1:AB1").Select
    Selection.AutoFilter Field:=1, Criteria1:="<>#N/A", Criteria2:="<> ", Operator:=xlAnd
    pouchInsertSpace.Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    
    pouchInsertSpace.Range("W8").PasteSpecial xlPasteValues
    pouchInsertSpace.Range("AA1:AB1").Select
    Selection.AutoFilter
    pouchInsertSpace.Range("AA:AB").ClearContents

    Dim PchLineIdle_Start As Range, PchLineIdle_End As Range
    Set PchLineIdle_Start = pouchInsertSpace.Range(pouchInsertSpace.Range("X9"), pouchInsertSpace.Range("X9").End(xlDown))
    Set PchLineIdle_End = pouchInsertSpace.Range(pouchInsertSpace.Range("W9"), pouchInsertSpace.Range("W9").End(xlDown))

    pouchInsertSpace.Range("P8").Value = "Start"
    pouchInsertSpace.Range("Q8").Value = "End"
    pouchInsertSpace.Range("R8").Value = "same"
    pouchInsertSpace.Range("P9").Value = 0

    PchLineIdle_End.Copy
    pouchInsertSpace.Range("Q9:Q" & PchLineIdle_End.Count).PasteSpecial xlPasteValues
    PchLineIdle_Start.Copy
    pouchInsertSpace.Range("P10:P" & PchLineIdle_Start.Count).PasteSpecial xlPasteValues
    pouchInsertSpace.Range("I" & PchLineIdle_Start.Count + 9).Value = wb.Worksheets("Silos").Range("A1").End(xlDown)
    pouchInsertSpace.Range("W:X").ClearContents

    pouchInsertSpace.Range("R9:R" & PchLineIdle_Start.Count + 9).Formula = "=IF(P9=Q9, ""Yes"", ""No"")"
    pouchInsertSpace.Range("P8:R8").Select
    Selection.AutoFilter Field:=3, Criteria1:="No"
    pouchInsertSpace.Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    pouchInsertSpace.Range("D2").Value = "PouchLine Idle"
    pouchInsertSpace.Range("D3").PasteSpecial xlPasteValues
    pouchInsertSpace.Range("P8:R8").Select
    Selection.AutoFilter
    pouchInsertSpace.Range("O:R").ClearContents
    pouchInsertSpace.Range("F:F").ClearContents

End Sub

Sub findIntersectionsOfIdleTimes(countPouches)
    pouchInsertSpace.Range("H1").Value = "Total Pouch Campaigns: " & countPouches
    pouchInsertSpace.Range("H2").Value = "Both Tip Station & Pouchline Idle"
    pouchInsertSpace.Range("H3").Value = "Potential Slot Point i"
    pouchInsertSpace.Range("I3").Value = "Start"
    pouchInsertSpace.Range("J3").Value = "End"
    
    Dim tipIdleStart As Double, tipIdleEnd As Double
    Dim pchIdleStart As Double, pchIdleEnd As Double
    Dim intersectIdleStart As Double, intersectIdleEnd As Double
    Dim tipRow As Integer, pchRow As Integer
    Dim tipLastRow As Integer, pchLastRow As Integer
    Dim lenTipIdle As Integer, lenPchIdle As Integer
    Dim potentialSlotCount As Integer
    
    potentialSlotCount = 4
    tipLastRow = pouchInsertSpace.Range("A4").End(xlDown).Row
    pchLastRow = pouchInsertSpace.Range("D4").End(xlDown).Row
    lenTipIdle = pouchInsertSpace.Range("A4:A" & tipLastRow).Count
    lenPchIdle = pouchInsertSpace.Range("D4:D" & pchLastRow).Count
    tipRow = 4
    pchRow = 4

    Do While tipRow <= lenTipIdle And pchRow <= lenPchIdle
        tipIdleStart = pouchInsertSpace.Range("A" & tipRow)
        pchIdleStart = pouchInsertSpace.Range("D" & pchRow)
        intersectIdleStart = WorksheetFunction.Max(tipIdleStart, pchIdleStart)

        tipIdleEnd = pouchInsertSpace.Range("B" & tipRow)
        pchIdleEnd = pouchInsertSpace.Range("E" & pchRow)
        intersectIdleEnd = WorksheetFunction.Min(tipIdleEnd, pchIdleEnd)

        If intersectIdleStart <= intersectIdleEnd Then
            pouchInsertSpace.Range("H" & potentialSlotCount).Value = potentialSlotCount - 3
            pouchInsertSpace.Range("I" & potentialSlotCount).Value = intersectIdleStart
            pouchInsertSpace.Range("J" & potentialSlotCount).Value = intersectIdleEnd
            
            potentialSlotCount = potentialSlotCount + 1
        End If

        If tipIdleEnd < pchIdleEnd Then
            tipRow = tipRow + 1
        Else
            pchRow = pchRow + 1
        End If
    Loop
End Sub

' ============================================= Main Logic =============================================
Function logic3(countPouchCampaigns)
    Dim mainSilo As Integer
    Dim otherSilo As Integer
    mainSilo = 16
    otherSilo = 6

    Dim isFeasible As Boolean
    isFeasible = insertPPPouchCampaigns(mainSilo, otherSilo)
    logic3 = isFeasible

End Function

Function insertPPPouchCampaigns(mainSilo, otherSilo) As Boolean

    ' arrays for determining which can starve to skip
    Dim d2Skip() As Integer
    ReDim d2Skip(1)
    d2Skip(0) = 0

    Dim count as Integer
    count = 1

    Do While True
        Print #logic3TextFile, "======== Attempt " & count & " ========"
        count = count + 1

        Print #logic3TextFile, "-- Finding PP Pouch Campaign to Insert..."
        ' get row of campaign to insert
        ' -1 if there is no campaign
        Dim PPCampaignToInsert As Double
        PPCampaignToInsert = findNextCampaignToInsert(PPPouchSchedule)
        Print #logic3TextFile, "To insert Campaign: " & PPCampaignToInsert
        Print #logic3TextFile, "Done."

        Print #logic3TextFile, "-- Finding Pouch Line Availability..."
        ' get row of insertion in schedule
        ' -1 if there is no intersection of idle times
        Dim D2FirstPchAvailHrs As Integer
        D2FirstPchAvailHrs = findFirstPchAvailHrs(D2Schedule, d2Skip, PPCampaignToInsert)
        Print #logic3TextFile, "First Pouch Availability: " & D2FirstPchAvailHrs
        Print #logic3TextFile, "Done." 
        Print #logic3TextFile, "-------"

        ' get which index to skip in d2Skip
        Dim dryerCampaign As Integer
        dryerCampaign = determineDryerCampaign(D2FirstPchAvailHrs, PPCampaignToInsert)
        Print #logic3TextFile, "Dryer Campaign Value: " & dryerCampaign

        If dryerCampaign = -2 Then 'Case: pouch campaigns but no more d2 slots (infeasible solution)
            Print #logic3TextFile, "PP Pouch Campaigns remaining but no more insertion points in Dryer 2. Exiting Program."
            insertPPPouchCampaigns = False
            reasonForStop = "PP-Pouch campaigns remaining but no more insertion points in dryer 2."
            Print #logic3TextFile, "======== Attempt " & (count-1) & " Concluded ========"
            Exit Function
        ElseIf dryerCampaign = -1 Then 'Case: no more campaigns left
            Print #logic3TextFile, "All PP Pouch Campaigns inserted."
            Print #logic3TextFile, "======== Attempt " & (count-1) & " Concluded ========"
            insertPPPouchCampaigns = True
            Exit Function
        Else
            Print #logic3TextFile, "Adding PP Pouch Campaign campaign"
            d2Skip = addPouchCampaign(PPCampaignToInsert, D2Schedule, D2Default, D2FirstPchAvailHrs, mainSilo, otherSilo, d2Skip)
        End If
        Print #logic3TextFile, "======== Attempt " & (count-1) & " Concluded ========"
        Print #logic3TextFile, " "
    Loop
End Function

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

Function findFirstPchAvailHrs(Worksheet, dryerSkipArray, PPCampaignToInsert) As Double
    ' ensure column BX is Pch Avail Hrs
    If IsNumeric("BX1") Or Worksheet.Range("BX1").Value <> "Pch Avail Hrs" Then
        reasonForStop = "Cell BX1 is not set to Pch Avail Hrs for " & Worksheet.Name
    End If
    ' ensure column BL is Pch Start
    If IsNumeric("BL1") Or Worksheet.Range("BL1").Value <> "Pch Start" Then
        reasonForStop = "Cell BL1 is not set to Pch Start for " & Worksheet.Name
    End If

    ' Stop Condition
    If PPCampaignToInsert = -1 Then 
        findFirstPchAvailHrs = -2
        Exit Function
    End If 

    'To Remove
    PPPouchSchedule.Range("R9").Value = "Current Pouch Avail Hours Row"

    ' return first pouch available hours
    Dim pchAvailHrsCell As Range
    Dim nextPchStartCell As Range
    For Each pchAvailHrsCell In Worksheet.Range("BX:BX")
        If pchAvailHrsCell.Value > 0 And IsNumeric(pchAvailHrsCell.Value) And isPchAvailInArray(pchAvailHrsCell.Row, dryerSkipArray) = False Then
            Set nextPchStartCell = Worksheet.Range("BL" & pchAvailHrsCell.Row + 1)
            If IsNumeric(nextPchStartCell.Value) Then
                If nextPchStartCell.Value <> pchAvailHrsCell.Value Then
                    If containedInIntersection(pchAvailHrsCell.Value, nextPchStartCell.Value, PPCampaignToInsert) Then
                        findFirstPchAvailHrs = pchAvailHrsCell.Row
                        
                        'To Remove
                        PPPouchSchedule.Range("S9").Value = pchAvailHrsCell.Row

                        Exit Function
                    End If
                End If
            End If
        End If
        If pchAvailHrsCell.Value = "" Then
            Exit For
        End If
    Next pchAvailHrsCell

    'No Can Starve Time Found
    findFirstPchAvailHrs = -1
End Function

Function isPchAvailInArray(pchAvail, dryerSkipArray) As Boolean
    Dim i As Integer
    For i = LBound(dryerSkipArray) To UBound(dryerSkipArray)
        If dryerSkipArray(i) = pchAvail Then
            isPchAvailInArray = True
            Exit Function
        End If
    Next
    isPchAvailInArray = False
End Function

Function containedInIntersection(pchAvailHrs, nextPchStart, PPCampaignToInsert) As Boolean
    Dim idleStartCell As Range, idleEndCell As Range, nextIdleStartCell As Range
    Dim lastRow As Integer
    Dim afterStart As Boolean, beforeEnd As Boolean
    lastRow = pouchInsertSpace.Range("I4").End(xlDown).Row

    For Each idleStartCell In pouchInsertSpace.Range("I4:I" & lastRow)
        Set idleEndCell = pouchInsertSpace.Range("J" & idleStartCell.Row)
        Set nextIdleStartCell = pouchInsertSpace.Range("I" & idleStartCell.Row + 1)
        
        Dim pouchFillTime As Range
        Set pouchFillTime = PPPouchSchedule.Range("Q" & PPCampaignToInsert)
        
        If betweenIntersected(idleStartCell.Value, idleEndCell.Value, pchAvailHrs, pouchFillTime) Then
            containedInIntersection = True
            Exit Function
        End If

        If nextIdleStartCell.Value > nextPchStart Then
            Exit For
        End If

        If idleStartCell.Value = "" Then
            Exit For
        End If
    Next idleStartCell
    
    containedInIntersection = False
End Function

Function betweenIntersected(idleStart, idleEnd, pchAvailHrs, pchFillTime) As Boolean
    Dim pchTimeRequired As Double
    pchTimeRequired = idleStart + pchFillTime

    If pchAvailHrs >= idleStart And pchAvailHrs <= idleEnd And pchTimeRequired <= idleEnd Then
        betweenIntersected = True
    Else
        betweenIntersected = False
    End If
End Function

Function determineDryerCampaign(D2FirstPchAvailHrs, PPCampaignToInsert)
    If PPCampaignToInsert = -1 Then
        determineDryerCampaign = -1
        Exit Function
    ElseIf D2FirstPchAvailHrs = -1 Then
        determineDryerCampaign = -2
        Exit Function
    Else
        determineDryerCampaign = 1
    End If
End Function

Function addPouchCampaign(PPCampaignToInsert, dryerSchedule, dryerDefaultSchedule, D2FirstPchAvailHrs, mainSilo, otherSilo, dryerSkipArray) As Integer()
    Print #logic3TextFile, "++++++++++++++++++++++++"
    PPPouchSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Copy
    dryerDefaultSchedule.Range("A" & D2FirstPchAvailHrs).Insert xlShiftDown
    dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
    Application.CalculateFull

    Dim canAdd As Boolean
    canAdd = checkSiloConstraint(mainSilo, otherSilo)
    If canAdd = True Then
        Print #logic3TextFile, "Inserted @ " & D2FirstPchAvailHrs
        PPPouchSchedule.Range("A" & PPCampaignToInsert, "N" & PPCampaignToInsert).Delete xlShiftUp
        PPPouchSchedule.Range("Q" & PPCampaignToInsert).Delete xlShiftUp
        dryerSkipArray = addItemToArray((D2FirstPchAvailHrs), dryerSkipArray)
        dryerSkipArray = addItemToArray((D2FirstPchAvailHrs + 1), dryerSkipArray)
        Print #logic3TextFile, "++++++++++++++++++++++++"
    Else
        Print #logic3TextFile, "Cannot be inserted at slot. Skipping."
        dryerDefaultSchedule.Rows(D2FirstPchAvailHrs).EntireRow.Delete xlShiftUp
        dryerSkipArray = addItemToArray(D2FirstPchAvailHrs, dryerSkipArray)
        Application.CalculateFull
        Print #logic3TextFile, "++++++++++++++++++++++++"
    End If


    addPouchCampaign = dryerSkipArray
    ' to ensure all pivottables are updated after adding pouch campaigns
    wb.RefreshAll
End Function

Function checkSiloConstraint(mainSilo, otherSilo) As Boolean
    Dim effectOnMainSilo As Double
    Dim effectOnOtherSilo As Double

    effectOnMainSilo = Silos.Range("J1").Value
    effectOnOtherSilo = Silos.Range("J2").Value

    If effectOnMainSilo <= mainSilo And effectOnOtherSilo <= otherSilo Then
        checkSiloConstraint = True
    Else
        checkSiloConstraint = False
        Print #logic3TextFile, "Effect: Silo Constraint violated by insertion"
    End If
End Function

Function addItemToArray(item, dryerSkipArray) As Integer()
    ReDim Preserve dryerSkipArray(LBound(dryerSkipArray) To UBound(dryerSkipArray) + 1)
    dryerSkipArray(UBound(dryerSkipArray)) = item
    addItemToArray = dryerSkipArray
End Function