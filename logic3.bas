Attribute VB_Name = "PPPouchInitialisation"
Option Explicit
Dim wb As Workbook
Dim D2Schedule As Worksheet, PPPouchSchedule As Worksheet, PPTipStatSheet As Worksheet, PPRateDSSheet As Worksheet

Sub initializePouchInsertion()
    Dim countPouchCampaigns As Long
    
    Application.AutoRecover.Enabled = False
    
    initializeWorksheets
    countPouchCampaigns = initializePouchWorksheets
    getPotentialSlots countPouchCampaigns
    
End Sub

Sub initializeWorksheets()
    'Without Initialising into same workbook
    Set wb = ThisWorkbook

    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet PPPouchSchedule, "PP PCH"
    setWorksheet PPTipStatSheet, "PP"
    setWorksheet PPRateDSSheet, "PPRateDS"
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
    
    Set D1TipStat_pivotTable = PPTipStatSheet.PivotTables("PivotTable16")
    Set D2TipStat_pivotTable = PPTipStatSheet.PivotTables("PivotTable15")
    
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
    PPTipStatSheet.Range("AA4:AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D1TipStatEnd.Count
    D1TipStatEnd.Copy
    PPTipStatSheet.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D2TipStatStart.Count
    D2TipStatStart.Copy
    PPTipStatSheet.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D2TipStartEnd.Count
    D2TipStartEnd.Copy
    PPTipStatSheet.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    Dim PPStatInUse As Range
    Set PPStatInUse = PPTipStatSheet.Range("AA3:AA" & endRow)
    PPStatInUse.Sort Key1:=PPTipStatSheet.Range("AA3"), Order1:=xlAscending, Header:=xlYes
    
    PPTipStatSheet.Range("J2").Value = "TipStation Idle"
    PPTipStatSheet.Range("J3").Value = "Start"
    PPTipStatSheet.Range("K3").Value = "End"
 
    Dim i As Integer, j As Integer
    Dim positivetime_start As Double, positivetime_end As Double
    i = 4
    j = 4
    Do Until i >= endRow + 1
        positivetime_start = PPTipStatSheet.Range("AA" & i).Value
        If positivetime_start >= 0 Then
            If i = 4 Then
                positivetime_end = positivetime_start
                positivetime_start = 0
                i = i + 1
            Else
                positivetime_end = PPTipStatSheet.Range("AA" & i + 1).Value
                i = i + 2
            End If
            PPTipStatSheet.Range("J" & j).Value = positivetime_start
            PPTipStatSheet.Range("K" & j).Value = positivetime_end
            j = j + 1
        Else
            i = i + 1
        End If
    Loop
    PPTipStatSheet.Range("AA4:AA" & endRow).Clear
    PPTipStatSheet.Range("K" & j - 1).Value = 5000
    
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
    PPTipStatSheet.Range("AA1").Value = "PouchLineInUse_Start"
    PPTipStatSheet.Range("AA2:AA" & PchLine_Starts.Count + 1).PasteSpecial xlPasteValues
    PchLine_Ends.Copy
    PPTipStatSheet.Range("AB1").Value = "PouchLineInUse_End"
    PPTipStatSheet.Range("AB2:AB" & PchLine_Ends.Count + 1).PasteSpecial xlPasteValues
    
    PPTipStatSheet.Range("AA1:AB1").Select
    Selection.AutoFilter Field:=1, Criteria1:="<>#N/A", Criteria2:="<> ", Operator:=xlAnd
    PPTipStatSheet.Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    
    PPTipStatSheet.Range("W8").PasteSpecial xlPasteValues
    PPTipStatSheet.Range("AA1:AB1").Select
    Selection.AutoFilter
    PPTipStatSheet.Range("AA:AB").ClearContents

    Dim PchLineIdle_Start As Range, PchLineIdle_End As Range
    Set PchLineIdle_Start = PPTipStatSheet.Range(PPTipStatSheet.Range("X9"), PPTipStatSheet.Range("X9").End(xlDown))
    Set PchLineIdle_End = PPTipStatSheet.Range(PPTipStatSheet.Range("W9"), PPTipStatSheet.Range("W9").End(xlDown))

    PPTipStatSheet.Range("P8").Value = "Start"
    PPTipStatSheet.Range("Q8").Value = "End"
    PPTipStatSheet.Range("R8").Value = "same"
    PPTipStatSheet.Range("P9").Value = 0

    PchLineIdle_End.Copy
    PPTipStatSheet.Range("Q9:Q" & PchLineIdle_End.Count).PasteSpecial xlPasteValues
    PchLineIdle_Start.Copy
    PPTipStatSheet.Range("P10:P" & PchLineIdle_Start.Count).PasteSpecial xlPasteValues
    PPTipStatSheet.Range("Q" & PchLineIdle_Start.Count + 9).Value = wb.Worksheets("Silos").Range("A1").End(xlDown)
    PPTipStatSheet.Range("W:X").ClearContents

    PPTipStatSheet.Range("R9:R" & PchLineIdle_Start.Count + 9).Formula = "=IF(P9=Q9, ""Yes"", ""No"")"
    PPTipStatSheet.Range("P8:R8").Select
    Selection.AutoFilter Field:=3, Criteria1:="No"
    PPTipStatSheet.Range(Selection, Selection.End(xlDown)).Select
    Selection.SpecialCells(xlCellTypeVisible).Select
    Selection.Copy
    PPTipStatSheet.Range("M2").Value = "PouchLine Idle"
    PPTipStatSheet.Range("M3").PasteSpecial xlPasteValues
    PPTipStatSheet.Range("P8:R8").Select
    Selection.AutoFilter
    PPTipStatSheet.Range("O:R").ClearContents

End Sub

Sub findIntersectionsOfIdleTimes(countPouches)
    PPTipStatSheet.Range("P1").Value = "Total Pouch Campaigns: " & countPouches
    PPTipStatSheet.Range("P2").Value = "Both Tip Station & Pouchline Idle"
    PPTipStatSheet.Range("P3").Value = "Potential Slot Point i"
    PPTipStatSheet.Range("Q3").Value = "Start"
    PPTipStatSheet.Range("R3").Value = "End"
    
    Dim TipIdleStart As Double, TipIdleEnd As Double, PchIdleStart_next As Double, PchIdleEnd_next As Double
    Dim i As Integer, j As Integer, k As Integer
    Dim PchLineIdle_Start As Range
    Dim PchIdleStart As Double, PchIdleEnd As Double
    
    Set PchLineIdle_Start = PPTipStatSheet.Range(PPTipStatSheet.Range("N4"), PPTipStatSheet.Range("N4").End(xlDown))
    
    i = 1
    Do Until i > PPTipStatSheet.Range(PPTipStatSheet.Range("K4"), PPTipStatSheet.Range("K4").End(xlDown)).Count
        j = i + 3
        PPTipStatSheet.Range("P" & j).Value = i
    
        TipIdleStart = PPTipStatSheet.Range("J" & j)
        TipIdleEnd = PPTipStatSheet.Range("K" & j)
    
        k = 4
        Do Until k > PchLineIdle_Start.Count + 4
            PchIdleStart = PPTipStatSheet.Range("M" & k)
            PchIdleStart_next = PPTipStatSheet.Range("M" & k + 1)
            PchIdleEnd = PPTipStatSheet.Range("N" & k)
            PchIdleEnd_next = PPTipStatSheet.Range("N" & k + 1)
    
            If TipIdleStart >= PchIdleStart And TipIdleStart < PchIdleStart_next Then
                If TipIdleStart > PchIdleEnd Then
                    PPTipStatSheet.Range("Q" & j).Value = PchIdleStart_next
                    PPTipStatSheet.Range("R" & j).Value = WorksheetFunction.Min(PchIdleEnd_next, TipIdleEnd)
                    Exit Do
                ElseIf TipIdleEnd < PchIdleEnd Then
                    PPTipStatSheet.Range("Q" & j).Value = TipIdleStart
                    PPTipStatSheet.Range("R" & j).Value = TipIdleEnd
                    Exit Do
                ElseIf TipIdleEnd > PchIdleEnd Then
                    PPTipStatSheet.Range("Q" & j).Value = TipIdleStart
                    PPTipStatSheet.Range("R" & j).Value = PchIdleEnd
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
    Set PchIdleRemaining = PPTipStatSheet.Range(PPTipStatSheet.Range("M" & k), PPTipStatSheet.Range("M" & k).End(xlDown))
    Count_PchIdleRemaining = PchIdleRemaining.Count

    Do Until Count_PchIdleRemaining = 0
        PPTipStatSheet.Range("P" & j).Value = i
        PPTipStatSheet.Range("Q" & j).Value = PPTipStatSheet.Range("M" & k)
        PPTipStatSheet.Range("R" & j).Value = PPTipStatSheet.Range("N" & k)

        i = i + 1
        j = j + 1
        k = k + 1
        Count_PchIdleRemaining = Count_PchIdleRemaining - 1
    Loop

End Sub
