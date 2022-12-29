Option Explicit
Dim wb As Workbook
Dim D2Schedule As Worksheet 
Dim PPPouchSchedule As Worksheet
Dim PPTippingStation As Worksheet
Dim PPRateDSSheet As Worksheet
Dim pouchInsertSpace as Worksheet
Dim D2Default As Worksheet

Sub ppPouchMain()
    Application.AutoRecover.Enabled = False
    initializeWorksheets

    Dim numberPouchCampaigns As Integer
    numberPouchCampaigns = initializePouchInsertion

    Dim isLogic3Feasible As Boolean
    isLogic3Feasible = logic3(numberPouchCampaigns)
    If isLogic3Feasible = False Then 
        MsgBox "PP-Pouch Campaigns cannot be inserted by automated process. Terminating Program."
        End
    End If

End Sub

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
    setWorksheet PPPouchSchedule, "PP PCH"
    setWorksheet PPTippingStation, "PP"
    setWorksheet PPRateDSSheet, "PPRateDS"
    setWorksheet pouchInsertSpace, "PP PCH SPACE" 'Create new sheet for finding idle times based on pivot tables

    ' Update pivot table to correct setting PP sheet
    Dim PT as PivotTable
    For Each PT in PPTippingStation.PivotTables
        On Error Resume Next
        For Each PI in PT.PivotFields("Source (DR, DB, PP)").PivotItems
            Select Case PI.Name
                Case Is = "PP"
                    PI.Visible = True
                Case Else
                    PI.Visible = False
            End Select
        Next PI
    Next PT
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wb.Sheets(worksheetName)
    Exit Sub
Err:
    MsgBox worksheetName & " is not in current workbook"
    End
End Sub

Function initializePouchWorksheets()
    Dim Pouch_OriginalDetails As Range
    Dim lastrow As Long, countPouches As Long
    
    'Copy and Paste original data to the side
    countPouches = PPPouchSchedule.Cells(2, 1).End(xlDown).Row
    lastrow = PPPouchSchedule.Cells(2, 19).End(xlDown).Row
    PPPouchSchedule.Range("S2:AF" & lastrow).ClearContents
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
    
    initializePouchWorksheets = countPouches
End Function

Sub getPotentialSlots(countPouches)
    Dim D1TipStat_pivotTable As pivotTable, D2TipStat_pivotTable As pivotTable
    Dim D1TipStat_start As Range, D1TipStat_end As Range, D2TipStat_start As Range, D2TipStat_end As Range
    
    Set D1TipStat_pivotTable = PPTippingStation.PivotTables("PivotTable16")
    Set D2TipStat_pivotTable = PPTippingStation.PivotTables("PivotTable15")
    
    D1TipStat_pivotTable.RefreshTable
    D2TipStat_pivotTable.RefreshTable
    
    Set D1TipStat_start = getPivotEntry(D1TipStat_pivotTable, 1)
    Set D1TipStat_end = getPivotEntry(D1TipStat_pivotTable, 2)
    Set D2TipStat_start = getPivotEntry(D2TipStat_pivotTable, 1)
    Set D2TipStat_end = getPivotEntry(D2TipStat_pivotTable, 2)

    getTipStatIdleTimes countPouches, D1TipStat_start, D1TipStat_end, D2TipStat_start, D2TipStat_end
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
    
    pouchInsertSpace.Range("J2").Value = "TipStation Idle"
    pouchInsertSpace.Range("J3").Value = "Start"
    pouchInsertSpace.Range("K3").Value = "End"
 
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
            pouchInsertSpace.Range("J" & j).Value = positivetime_start
            pouchInsertSpace.Range("K" & j).Value = positivetime_end
            j = j + 1
        Else
            i = i + 1
        End If
    Loop
    pouchInsertSpace.Range("AA4:AA" & endRow).Clear
    pouchInsertSpace.Range("K" & j - 1).Value = 5000
    
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
    pouchInsertSpace.Range("Q" & PchLineIdle_Start.Count + 9).Value = wb.Worksheets("Silos").Range("A1").End(xlDown)
    pouchInsertSpace.Range("W:X").ClearContents

    pouchInsertSpace.Range("R9:R" & PchLineIdle_Start.Count + 9).Formula = "=IF(P9=Q9, ""Yes"", ""No"")"
    pouchInsertSpace.Range("P8:R8").Select
    Selection.AutoFilter Field:=3, Criteria1:="No"
    pouchInsertSpace.Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    pouchInsertSpace.Range("M2").Value = "PouchLine Idle"
    pouchInsertSpace.Range("M3").PasteSpecial xlPasteValues
    pouchInsertSpace.Range("P8:R8").Select
    Selection.AutoFilter
    pouchInsertSpace.Range("O:R").ClearContents

End Sub

Sub findIntersectionsOfIdleTimes(countPouches)
    pouchInsertSpace.Range("P1").Value = "Total Pouch Campaigns: " & countPouches
    pouchInsertSpace.Range("P2").Value = "Both Tip Station & Pouchline Idle"
    pouchInsertSpace.Range("P3").Value = "Potential Slot Point i"
    pouchInsertSpace.Range("Q3").Value = "Start"
    pouchInsertSpace.Range("R3").Value = "End"
    
    Dim TipIdleStart As Double, TipIdleEnd As Double, PchIdleStart_next As Double, PchIdleEnd_next As Double
    Dim i As Integer, j As Integer, k As Integer
    Dim PchLineIdle_Start As Range
    Dim PchIdleStart As Double, PchIdleEnd As Double
    
    Set PchLineIdle_Start = pouchInsertSpace.Range(pouchInsertSpace.Range("N4"), pouchInsertSpace.Range("N4").End(xlDown))
    
    i = 1
    Do Until i > pouchInsertSpace.Range(pouchInsertSpace.Range("K4"), pouchInsertSpace.Range("K4").End(xlDown)).Count
        j = i + 3
        pouchInsertSpace.Range("P" & j).Value = i
    
        TipIdleStart = pouchInsertSpace.Range("J" & j)
        TipIdleEnd = pouchInsertSpace.Range("K" & j)
    
        k = 4
        Do Until k > PchLineIdle_Start.Count + 4
            PchIdleStart = pouchInsertSpace.Range("M" & k)
            PchIdleStart_next = pouchInsertSpace.Range("M" & k + 1)
            PchIdleEnd = pouchInsertSpace.Range("N" & k)
            PchIdleEnd_next = pouchInsertSpace.Range("N" & k + 1)
    
            If TipIdleStart >= PchIdleStart And TipIdleStart < PchIdleStart_next Then
                If TipIdleStart > PchIdleEnd Then
                    pouchInsertSpace.Range("Q" & j).Value = PchIdleStart_next
                    pouchInsertSpace.Range("R" & j).Value = WorksheetFunction.Min(PchIdleEnd_next, TipIdleEnd)
                    Exit Do
                ElseIf TipIdleEnd < PchIdleEnd Then
                    pouchInsertSpace.Range("Q" & j).Value = TipIdleStart
                    pouchInsertSpace.Range("R" & j).Value = TipIdleEnd
                    Exit Do
                ElseIf TipIdleEnd > PchIdleEnd Then
                    pouchInsertSpace.Range("Q" & j).Value = TipIdleStart
                    pouchInsertSpace.Range("R" & j).Value = PchIdleEnd
                    Exit Do
                End If
            End If
            k = k + 1
        Loop
        i = i + 1
    Loop
    
    k = k + 1
    j = j + 1
    Dim Count_PchIdleRemaining As Long, PchIdleRemaining As Range
    Set PchIdleRemaining = pouchInsertSpace.Range(pouchInsertSpace.Range("M" & k), pouchInsertSpace.Range("M" & k).End(xlDown))
    Count_PchIdleRemaining = PchIdleRemaining.Count

    Do Until Count_PchIdleRemaining = 0
        pouchInsertSpace.Range("P" & j).Value = i
        pouchInsertSpace.Range("Q" & j).Value = pouchInsertSpace.Range("M" & k)
        pouchInsertSpace.Range("R" & j).Value = pouchInsertSpace.Range("N" & k)

        i = i + 1
        j = j + 1
        k = k + 1
        Count_PchIdleRemaining = Count_PchIdleRemaining - 1
    Loop

End Sub

Function logic3(countPouchCampaigns)
    Dim mainSilo as Integer
    Dim otherSilo as Integer
    mainSilo = 16
    otherSilo = 6

    Dim isFeasible As Boolean
    isFeasible = insertPPPouchCampaigns(mainSilo, otherSilo)
    logic3 = isFeasible

End Function

Function insertPPPouchCampaigns(mainSilo, otherSilo) As Boolean
    Dim d2Skip() As Integer
    ReDim d2Skip(1)
    d2Skip(0) = 0

    Do While True
        ' get row of campaign to insert
        ' -1 if there is no campaign
        Dim PPCampaignToInsert As Double
        PPCampaignToInsert = findNextCampaignToInsert(PPPouchSchedule)

        ' get row of insertion in schedule
        ' -1 if there is no intersection of idle times
        Dim D2FirstPchAvailHrs as Integer 
        D2FirstPchAvailHrs = findFirstPchAvailHrs(D2Schedule, d2Skip, PPCampaignToInsert)
        
        ' get which index to skip in d2Skip
        Dim dryerCampaign as Integer
        dryerCampaign = determineDryerCampaign(D2FirstPchAvailHrs, PPCampaignToInsert)


        If dryerCampaign = -2 Then 'Case: pouch campaigns but no more d2 slots (infeasible solution)
            insertPPPouchCampaigns = False
            MsgBox "PP-Pouch campaigns remaining but no more insertion points in dryer 2. Exiting Program."
            Exit Function
        ElseIf dryerCampaign = -1 Then 'Case: no more campaigns left
            MsgBox "All pouches inserted"
            insertPPPouchCampaigns = True
            Exit Function
        Else
            d2Skip = addPouchCampaign(PPCampaignToInsert, D2Schedule, D2Default, D2FirstPchAvailHrs, mainSilo, otherSilo, d2Skip)
        End If        
    Loop
End Function

Function findNextCampaignToInsert(Worksheet) As Integer
    If Worksheet.Range("A1").End(xlDown).Value = "" Then
        findNextCampaignToInsert = -1
        Exit Function
    End If
    Dim cell As Range
        For Each cell in Worksheet.Range("A2:A" & Worksheete.Range("A" & Rows.Count).End(xlUp).Row)
            If cell.Value <> "" Then 
                findNextCampaignToInsert = cell.Row
                Exit Function
            End If
        Next cell
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

Function addItemToArray(item, dryerSkipArray) As Integer()
    ReDim Preserve dryerSkipArray(LBound(dryerSkipArray) To UBound(dryerSkipArray) + 1)
    dryerSkipArray(UBound(dryerSkipArray)) = item
    addItemToArray = dryerSkipArray
End Function

Function findFirstPchAvailHrs(Worksheet, dryerSkipArray, PPCampaignToInsert) As Double
    ' ensure column BX is Pch Avail Hrs
    If IsNumeric("BX1") Or Worksheet.Range("BX1").Value <> "Pch Avail Hrs" Then 
        MsgBox "Cell BX1 is not set to Pch Avail Hrs for " & Worksheet.Name
    End If
    ' ensure column BL is Pch Start
    If IsNumeric("BL1") Or Worksheet.Range("BL1").Value <> "Pch Start" Then 
        MsgBox "Cell BL1 is not set to Pch Start for " & Worksheet.Name
    End If

    ' return first pouch available hours 
    Dim pchAvailHrsCell as Range
    Dim nextPchStartCell as Range
    For Each pchAvailHrsCell in Worksheet.Range("BX:BX")
        If pchAvailHrsCell.Value > 0 and IsNumeric(pchAvailHrsCell.Value) And isPchAvailInArray(pchAvailHrsCell.Row, dryerSkipArray) = False Then 
            Set nextPchStartCell = Worksheet.Range("BL" & pchAvailHrsCell.Row + 1)
            If nextPchStartCell.Value <> pchAvailHrsCell.Value And IsNumeric(nextPchStartCell.Value) Then
                If containedInIntersection(pchAvailHrsCell.Value, nextPchStartCell.Value) And moreThanPouchFill(pchAvailHrsCell.Value, nextPchStartCell.Value, PPCampaignToInsert) Then
                    findFirstPchAvailHrs = pchAvailHrsCell.Row
                    Exit Function
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

Function containedInIntersection(pchAvailHrs, nextPchStart) As Boolean
    Dim idleStartCell As Range, idleEndCell As Range
    Dim lastRow As Integer
    Dim afterStart as Boolean, beforeEnd as Boolean
    lastRow = pouchInsertSpace.Range("Q4").End(xlDown).Row

    For Each idleStartCell In pouchInsertSpace.Range("Q4:Q" & lastRow)
        Set idleEndCell = pouchInsertSpace.Range("R" & idleStartCell.Row)
        If betweenIntersected(idleStartCell.Value, idleEndCell.Value, pchAvailHrs, nextPchStart) Then
            containedInIntersection = True
            Exit Function
        End If
        
        If idleStartCell.Value = "" Then
            Exit For
        End If
    Next idleStartCell
    containedInIntersection = False
End Function

Function betweenIntersected(idleStart, idleEnd, pchAvailHrs, nextPchStart) As Boolean
    If pchAvailHrs >= idleStart And nextPchStart <= idleEnd Then 
        betweenIntersected = True
    Else
        betweenIntersected = False
    End If
End Function

Function moreThanPouchFill(pchAvailHrs, nextPchStart, PPCampaignToInsert) As Boolean
    Dim pouchFillTime as Range
    Set pouchFillTime = PPPouchSchedule.Range("Q" & PPCampaignToInsert)
    
    If nextPchStart - pchAvailHrs > pouchFillTime.Value Then 
        moreThanPouchFill = True
    Else
        moreThanPouchFill = False
    End If
End Function

Function determineDryerCampaign(D2FirstPchAvailHrs, PPCampaignToInsert)
    If PPCampaignToInsert = -1 Then 
        determineDryerCampaign = -1
    ElseIf D2FirstPchAvailHrs = -1 Then 
        determineDryerCampaign = -2
    Else
        determineDryerCampaign = 1
    End If
End Function

Function addPouchCampaign(PPCampaignToInsert, dryerSchedule, dryerDefaultSchedule, D2FirstPchAvailHrs, mainSilo, otherSilo, dryerSkipArray) as Integer()
    PPPouchSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Copy
    dryerDefaultSchedule.Range("A" & D2FirstPchAvailHrs).Insert xlShiftDown
    dryerSchedule.Range("A:N").Value = dryerDefaultSchedule.Range("A:N").Value
    Application.CalculateFull

    canAdd = checkSiloConstraint(mainSilo, otherSilo)
    If canAdd = True Then 
        PPPouchSchedule.Range("A" & PPCampaignToInsert, "M" & PPCampaignToInsert).Delete
    Else
        dryerDefaultSchedule.Rows(D2FirstPchAvailHrs).EntireRow.Delete
        
        dryerSkipArray = addItemToArray(D2FirstPchAvailHrs, dryerSkipArray)
        addPouchCampaign = dryerSkipArray
        Application.CalculateFull
    End If
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
    End If
End Function

Function addItemToArray(item, dryerSkipArray) As Integer()
    ReDim Preserve dryerSkipArray(LBound(dryerSkipArray) To UBound(dryerSkipArray) + 1)
    dryerSkipArray(UBound(dryerSkipArray)) = item
    addItemToArray = dryerSkipArray
End Function