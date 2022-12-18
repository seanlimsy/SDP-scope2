Attribute VB_Name = "PPCanStretchInitialization"
Option Explicit
Dim wb As Workbook
Dim D1Schedule As Worksheet, D2Schedule As Worksheet, PPThreshold As Worksheet, PPTippingStation As Worksheet
Dim workingDryerSchedule As Worksheet

Sub stretchingCampaigns()
    Dim PPCanStretching As Range
    Dim D1TipStat_pivotTable As pivotTable, D2TipStat_pivotTable As pivotTable
    Dim D1TipStat_canCOMax As Long, D2TipStat_canCOMax As Long, TipStat_canCOMax As Long
    
    Application.AutoRecover.Enabled = False
    initializeWorksheets
    
    Set PPCanStretching = stretchingInitialisation
    
    Set D1TipStat_pivotTable = pivotFromDryers(1)
    Set D2TipStat_pivotTable = pivotFromDryers(2)
    D1TipStat_pivotTable.RefreshTable
    D2TipStat_pivotTable.RefreshTable
    
    D1TipStat_canCOMax = infoFromDryers(D1TipStat_pivotTable)
    D2TipStat_canCOMax = infoFromDryers(D2TipStat_pivotTable)
        
    
End Sub

Sub initializeWorksheets()
    'Without Initialising into same workbook
    
    'To adjust to hardcode onto user's path
    'Can also consider moving sheets over to one main workbook
    'Michael: Change reference to an cell value -- solve for this in instructions for documentation -- Lester's Preference KIV
    Set wb = ThisWorkbook

    setWorksheet D1Schedule, "D1B1L65T"
    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet PPThreshold, "PP CAN ADDED THRESHOLD"
    setWorksheet PPTippingStation, "PP"
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wb.Sheets(worksheetName)
    Exit Sub
Err:
    MsgBox worksheetName & " is not in current workbook"
    End
End Sub

Function stretchingInitialisation()
    Dim stretchDetails As Range
    
    Set stretchDetails = PPThreshold.Range("A2:N2")
    Set stretchingInitialisation = stretchDetails
    If WorksheetFunction.CountA(stretchDetails) = 0 Then
        MsgBox "Please add PP Stretching Worst Case Campaign -- Missing from PP CAN ADDED THRESHOLD sheet"
        End
    End If
End Function

Function pivotFromDryers(identity)
    If identity = 1 Then
        'D1 - Tip Station (40H Gap)
        Set pivotFromDryers = PPTippingStation.PivotTables("PivotTable16")
    ElseIf identity = 2 Then
        'D2 - Tip Station (40H Gap)
        Set pivotFromDryers = PPTippingStation.PivotTables("PivotTable15")
    End If
End Function

Function infoFromDryers(pivotTable)
    Dim TipStat_canCORange As Range
    Dim TipStat_canCOMax As Long
    
    Set TipStat_canCORange = pivotTable.PivotFields("Sum of Can After CO Hrs").DataRange
    TipStat_canCOMax = Application.WorksheetFunction.Max(TipStat_canCORange)
    
    infoFromDryers = TipStat_canCOMax
End Function


