Attribute VB_Name = "PPPouch_Initialization"
Option Explicit
Dim wb_PP As Workbook, wb_Main As Workbook
Dim D2Schedule As Worksheet, PPPouchSchedule As Worksheet, PPTipStatSheet As Worksheet, PPRateDSSheet As Worksheet
Dim D1TIPSTAT_PT As PivotTable, D2TIPSTAT_PT As PivotTable
Dim D1TIPSTAT_START As Range, D1TIPSTAT_END As Range, D2TIPSTAT_START As Range, D2TIPSTAT_END As Range
Dim Count_PouchCampaigns As Long

Sub Initialize_PouchInsertion()
    'turn off autosave
    Application.AutoRecover.Enabled = False
    
    initializeWorksheets
    initializePouchWorksheet
    initializePPWorksheet
    
    getTipStatIdleTimes
    getPchLineIdleTimes
    findIntersectionsOfIdleTimes

End Sub

Sub initializeWorksheets()
    'Without Initialising into same workbook
    
    'To adjust to hardcode onto user's path
    'Can also consider moving sheets over to one main workbook
    'Michael: Change reference to an cell value -- solve for this in instructions for documentation -- Lester's Preference KIV
    Set wb_PP = Workbooks.Open("/Users/ben/Desktop/Scope 2/Postponement Creation for Slotting.xlsx")
    Set wb_Main = Workbooks.Open("/Users/ben/Desktop/Model-Testing.xlsm")

    setWorksheet D2Schedule, "D2B1L3B3B4L45T", wb_Main
    setWorksheet PPPouchSchedule, "PP PCH", wb_PP
    setWorksheet PPTipStatSheet, "Testing", wb_Main
    'setWorksheet PPTipStatSheet, "PP", wb_Main
    'To Change to PP Sheet after integration -- Sent to Testing for primarily testing -- can verify once initiation insert is completed (Post Stage 1)
    setWorksheet PPRateDSSheet, "PPRateDS", wb_Main
End Sub

Sub setWorksheet(Worksheet, worksheetName, Workbook)
    On Error GoTo Err
        Set Worksheet = Workbook.Sheets(worksheetName)
    Exit Sub
Err:
    MsgBox worksheetName & " is not in current workbook"
    End
End Sub

Sub initializePouchWorksheet()
    Dim Pouch_OriginalDetails As Range
    Dim lastrow As Long
    
    'Copy and Paste original data to the side
    Count_PouchCampaigns = PPPouchSchedule.Cells(2, 1).End(xlDown).Row
    lastrow = PPPouchSchedule.Cells(2, 19).End(xlDown).Row
    PPPouchSchedule.Range("S2:AF" & lastrow).ClearContents
    Set Pouch_OriginalDetails = PPPouchSchedule.Range("A2:N" & Count_PouchCampaigns)
    Pouch_OriginalDetails.Copy
    PPPouchSchedule.Range("S2:AF" & Count_PouchCampaigns).PasteSpecial xlPasteValues
    
    'Calculate Pouch Fill Times
    Dim effective_fp_tonnes_perhr As Double
    Dim Pouch_Rates As Range
    Set Pouch_Rates = PPRateDSSheet.Range(PPRateDSSheet.Range("D2"), PPRateDSSheet.Range("D2").End(xlDown))
    
    effective_fp_tonnes_perhr = Application.WorksheetFunction.Min(Pouch_Rates)
    PPPouchSchedule.Range("Q1").Value = "Effective FP Tonnes per Hour"
    PPPouchSchedule.Range("Q2:Q" & Count_PouchCampaigns).Formula = "=J2/2.2/1000/" & effective_fp_tonnes_perhr
End Sub

Sub initializePPWorksheet()
'    'D1 - Tip Station (40H Gap)
'    Set D1TIPSTAT_PT = PPTipStatSheet.PivotTables("PivotTable16")
'    'D2 - Tip Station (40H Gap)
'    Set D2TIPSTAT_PT = PPTipStatSheet.PivotTables("PivotTable15")

'    refreshPivotTables
    getPivotTableEntries

End Sub

Sub refreshPivotTables()
    D1TIPSTAT_PT.RefreshTable
    D2TIPSTAT_PT.RefreshTable
End Sub

Sub getPivotTableEntries()
'    Set D1TIPSTAT_START = D1TIPSTAT_PT.PivotFields("Sum of Silo Entry Hr").DataRange
'    Set D1TIPSTAT_END = D1TIPSTAT_PT.PivotFields("Sum of Can After CO Hrs").DataRange
'    Set D2TIPSTATE_START = D2TIPSTAT_PT.PivotFields("Sum of Silo Entry Hr").DataRange
'    Set D2TIPSTAT_END = D2TIPSTAT_PT.PivotFields("Sum of Can After CO Hrs").DataRange

    ''Tipping Station Idle Times Tests
     Set D1TIPSTAT_START = PPTipStatSheet.Range(PPTipStatSheet.Range("A5"), PPTipStatSheet.Range("A5").End(xlDown))
     Set D1TIPSTAT_END = PPTipStatSheet.Range(PPTipStatSheet.Range("B5"), PPTipStatSheet.Range("B5").End(xlDown))
     Set D2TIPSTAT_START = PPTipStatSheet.Range(PPTipStatSheet.Range("D5"), PPTipStatSheet.Range("D5").End(xlDown))
     Set D2TIPSTAT_END = PPTipStatSheet.Range(PPTipStatSheet.Range("E5"), PPTipStatSheet.Range("E5").End(xlDown))
     
End Sub

Sub getTipStatIdleTimes()
    Dim startRow As Integer, endRow As Integer
    startRow = 5 'Insert Checker here -- generalised searcher
    endRow = startRow + D1TIPSTAT_START.Count - 1
    D1TIPSTAT_START.Copy
    PPTipStatSheet.Range("AA4:AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D1TIPSTAT_END.Count
    D1TIPSTAT_END.Copy
    PPTipStatSheet.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D2TIPSTAT_START.Count
    D2TIPSTAT_START.Copy
    PPTipStatSheet.Range("AA" & startRow & ":AA" & endRow).PasteSpecial xlPasteValues
    startRow = endRow
    
    endRow = startRow + D2TIPSTAT_END.Count
    D2TIPSTAT_END.Copy
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
    PPTipStatSheet.Range("Q" & PchLineIdle_Start.Count + 9).Value = wb_Main.Worksheets("Silos").Range("A1").End(xlDown)
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

Sub findIntersectionsOfIdleTimes()
    PPTipStatSheet.Range("P1").Value = "Total Pouch Campaigns: " & Count_PouchCampaigns
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
