Option Explicit
Dim wb As Workbook
Dim D1Schedule As Worksheet
Dim D2Schedule As Worksheet
Dim PPThreshold As Worksheet
Dim PPTippingStation As Worksheet
Dim PPRateDSSheet As Worksheet
Dim workingDryerSchedule As Worksheet
Dim Silos As Worksheet

Sub PPCanStretchMain()
    Application.AutoRecover.Enabled = False
    initializeWorksheets
    'runOrDuplicateFile
    initializePPRateDS
    End

    Dim isLogic4Feasible As Boolean
    isLogic4Feasible = logic4()
    If isLogic4Feasible = False Then
        MsgBox "No additional PP Can Campaigns can be inserted by automated process. Terminating Program"
    End If

End Sub

' =============== Setup Logic ===============
Sub initializeWorksheets()
    Set wb = ThisWorkbook

    setWorksheet D1Schedule, "D1B1L65T"
    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
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
    MsgBox worksheetName & " is not in current workbook"
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
    lastRow = PPRateDSSheet.Range("B1").End(xlDown).Row
    canStretchRow = lastRow + 1

    Dim cell As Range
    Dim canRow As Integer
    For Each cell In PPRateDSSheet.Range("A2:A" & lastRow)
        If InStr(cell.Value, "CAN") Then
            canRow = cell.Row
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
        MsgBox "Error in Determining PPRateDS for PP-Can-5 (Stretching add). Check code-base ""initializePPRateDS"". Ending program." 
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

' =============== Main Logic ===============
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
    Dim D1TipStatCanCOMax As Long, D2TipStatCanCOMax As Long, TipStatCanCOMax As Long
    Application.CalculateFull

    Set PPCanStretching = stretchingInitialisation
    Set D1TipStatPivotTable = pivotFromDryers(1)
    Set D2TipStatPivotTable = pivotFromDryers(2)
    D1TipStatPivotTable.RefreshTable
    D2TipStatPivotTable.RefreshTable
    wb.RefreshAll

    D1TipStatCanCOMax = infoFromDryers(D1TipStatPivotTable)
    D2TipStatCanCOMax = infoFromDryers(D2TipStatPivotTable)
    
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
        PPCampaignToInsert = findNextCampaignToInsert(PPCanSchedule)
        
        ' get row of insertion in schedule
        ' -1 if there is no can starve
        Dim D1FirstCanStarveTime As Double
        Dim D2FirstCanStarveTime As Double
        D1FirstCanStarveTime = findFirstCanStarveTime(D1Schedule, d1Skip)
        D2FirstCanStarveTime = findFirstCanStarveTime(D2Schedule, d2Skip)
        
        ' get initial silo constraint violation time
        Dim initialSiloConstraintViolation
        initialSiloConstraintViolation = Silos.Range("K1").Value

        Dim dryerCampaign As Integer
        dryerCampaign = determineDryerCampaign(D1FirstCanStarveTime, D2FirstCanStarveTime, PPCampaignToInsert)
        
        If dryerCampaign = 0 Then
            MsgBox "All can starve slots used. Terminating Program"
            stretchingCampaigns = True
            Exit Function
        ElseIf dryerCampaign = 1 Then 'case: d1 PP campaign
            MsgBox "Add PPCan to Dryer 1"
            d1Skip = addPPCampaign(PPCampaignToInsert, D1Schedule, D1Default, D1FirstCanStarveTime, mainSilo, otherSilo, d1Skip, initialSiloConstraintViolation)
        ElseIf dryerCampaign = 2 Then 'case: d2 PP campaign
            MsgBox "Add PPCan to Dryer 2"
            d2Skip = addPPCampaign(PPCampaignToInsert, D2Schedule, D2Default, D2FirstCanStarveTime, mainSilo, otherSilo, d2Skip, initialSiloConstraintViolation)
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
    stretchingCampaigns = True
End Function

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
    Dim TipStatCanCORange As Range
    Dim TipStatCanCOMax As Long
    
    Set TipStatCanCORange = pivotTable.PivotFields("Sum of Can After CO Hrs").DataRange
    TipStatCanCOMax = Application.WorksheetFunction.Max(TipStatCanCORange)
    
    infoFromDryers = TipStatCanCOMax
End Function

' Incomplete adjustment
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
        Application.CalculateFull

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
    Application.CalculateFull
    
    ' this is to ensure that the pivot table is updated after adding pp campaigns
    wb.RefreshAll
    
    addPPCampaign = dryerSkipArray
End Function
