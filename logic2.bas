Option Explicit
Dim wb As Workbook
Dim D1Schedule As Worksheet, D2Schedule As Worksheet, Silos As Worksheet
Dim PPTippingStation As Worksheet
Dim workingDryerSchedule As Worksheet
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

Sub dryerBlockDelayMain(nextInsertTimeStep As Double)
    ' Print #logic1TextFile, " "
    ' Print #logic1TextFile, "==== Initializing logic 2 ====": Space 0
    
    Dim D1CipHrs As Double, D2CipHrs As Double
    Dim siteCpledCapCurrent As Double
    Dim DiCausingViolationPE As String, DiCausingViolationSG as String, DiCausingViolation as String
    
    Dim exceedTimeStepPE As Double, exceedTimeStepSG as Double, exceedTimeStep as Double
    'Dim nextInsertTimeStep As Double
    Dim dryerBlockBeforeNextInsert_bool As Boolean
    
    Dim idxToDelay As Double
    Dim DiCIPHrs As Double
    
    Application.AutoRecover.Enabled = False
    initializeWorksheetsStage2
    
    'Defining Variables - CIP
    D1CipHrs = wb.Worksheets("Evap DryCIP").Range("T3")
    D2CipHrs = wb.Worksheets("Evap DryCIP").Range("T6")
    
    Dim repeatedSolve As Integer
    repeatedSolve = 0

    Do While True
        If repeatedSolve >= 40 Then 
            ' Print #logic1TextFile, "Issues with resolving dryer blockage at point. Early Termination": Space 0
            reasonForStop = "Unknown effects to delay stage -- Infinite Loop occurred. Restart program"
            ' Print #logic1TextFile, "==== Ending logic 2 ====": Space 0
            End
        Else
            repeatedSolve = repeatedSolve + 1
        End If

        calculateAll
        siteCpledCapCurrent = Round(Silos.Range("R13"), 1)
        DiCausingViolationPE = checkPEDryerResults
        DiCausingViolationSG = checkSGDryerResults

        If DiCausingViolationPE = "None" And DiCausingViolationSG = "None" Then
            ' Print #logic1TextFile, "No more dryer blockages in the system."
            ' Print #logic1TextFile, "==== Ending logic 2 ====": Space 0
            Exit Sub
        Else
            If DiCausingViolationPE <> "None" And DiCausingViolationSG = "None" Then
                exceedTimeStepPE = getExceedTimeStep(DiCausingViolationPE)
                exceedTimeStep = exceedTimeStepPE
                DiCausingViolation = DiCausingViolationPE

            ElseIf DiCausingViolationPE = "None" And DiCausingViolationSG <> "None" Then
                exceedTimeStepSG = getExceedTimeStep(DiCausingViolationSG)
                exceedTimeStep = exceedTimeStepSG
                DiCausingViolation = DiCausingViolationSG

            ElseIf DiCausingViolationPE <> "None" And DiCausingViolationSG <> "None" Then 
                exceedTimeStepPE = getExceedTimeStep(DiCausingViolationPE)
                exceedTimeStepSG = getExceedTimeStep(DiCausingViolationSG)

                If exceedTimeStepPE <= exceedTimeStepSG Then
                    exceedTimeStep = exceedTimeStepPE
                    DiCausingViolation = DiCausingViolationPE
                Else
                    exceedTimeStep = exceedTimeStepSG
                    DiCausingViolation = DiCausingViolationSG
                End If
            End If
        End If
        ' Print #logic1TextFile, "Next possible insert time step: " & nextInsertTimeStep: Space 0
        ' Print #logic1TextFile, "Next dryer exceed time step: " & exceedTimeStep: Space 0
        ' Print #logic1TextFile, "Cause of violation: " & DiCausingViolation: Space 0

        'nextInsertTimeStep = getNextInsertionPointInSchedule(DiCausingViolation)
        dryerBlockBeforeNextInsert_bool = isDryerBlockBeforeNextInsert(exceedTimeStep, nextInsertTimeStep)
        
        If dryerBlockBeforeNextInsert_bool = False Then
            ' Print #logic1TextFile, "Next potential insertion point is before the next time dryer is exceeded. Ending blockage.": Space 0
            ' Print #logic1TextFile, "==== Ending logic 2 ====": Space 0
            Exit Do
        Else
            idxToDelay = getIdxToDelay(exceedTimeStep)
            DiCIPHrs = getCIPHrs(DiCausingViolation, D1CipHrs, D2CipHrs)
            ' Print #logic1TextFile, "Index to Delay: " & idxToDelay: Space 0
            ' Print #logic1TextFile, "Solving Delay...": Space 0
            resolveSiloContraint idxToDelay, DiCIPHrs, exceedTimeStep, DiCausingViolation, siteCpledCapCurrent
            ' Print #logic1TextFile, "Resolved. Moving to next possible TimeStep.": Space 0
            ' Print #logic1TextFile, " "
        End If
    Loop
End Sub

Sub initializeWorksheetsStage2()
    Set wb = ThisWorkbook

    setWorksheet D1Schedule, "D1B1L65T"
    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet Silos, "Silos"
    setWorksheet PPTippingStation, "PP"

    Set D1TipStatPivotTable = PPTippingStation.PivotTables("PivotTableD1")
    Set D2TipStatPivotTable = PPTippingStation.PivotTables("PivotTableD2")    
End Sub

Sub setWorksheet(Worksheet, worksheetName)
    On Error GoTo Err
        Set Worksheet = wb.Sheets(worksheetName)
    Exit Sub
Err:
    MsgBox worksheetName & " is not in current workbook"
    End
End Sub

Function checkPEDryerResults()
    Dim PEPresenceD1 As String, PEPresenceD2 As String

    PEPresenceD1 = Silos.Range("S9")
    PEPresenceD2 = Silos.Range("S10")
    
    If PEPresenceD1 = "YES" Then
        If Silos.Range("R9") = 0 And Silos.Range("R10") = 0 Then
            checkPEDryerResults = "None"
        Else
            checkPEDryerResults = "PED1"
            Set workingDryerSchedule = D1Schedule
        End If
    ElseIf PEPresenceD2 = "YES" Then
        If Silos.Range("R9") = 0 And Silos.Range("R10") = 0 Then
            checkPEDryerResults = "None"
        Else
            checkPEDryerResults = "PED2"
            Set workingDryerSchedule = D2Schedule
        End If
    Else 
        Dim firstViolationTime As Double
        Dim tempD1 As Double, tempD2 As Double

        firstViolationTime = Silos.Range("K1").Value
        tempD1 = Silos.Range("R9")
        tempD2 = Silos.Range("R10")
        
        If tempD1 <= firstViolationTime Then 
            checkPEDryerResults = "PED1"
            Set workingDryerSchedule = D1Schedule
        ElseIf tempD2 <= firstViolationTime Then
            checkPEDryerResults = "PED2"
            Set workingDryerSchedule = D2Schedule
        Else
            checkPEDryerResults = "None"            
        End If
    End If
End Function

Function checkSGDryerResults()
    Dim SGPresenceD1 As String, SGPresenceD2 As String

    SGPresenceD1 = Silos.Range("U9")
    SGPresenceD2 = Silos.Range("U10")
    
    If SGPresenceD1 = "YES" Then
        If Silos.Range("T9") = 0 And Silos.Range("T10") = 0 Then 
            checkSGDryerResults = "None"
        Else
            checkSGDryerResults = "SGD1"
            Set workingDryerSchedule = D1Schedule
        End If
    ElseIf SGPresenceD2 = "YES" Then
        If Silos.Range("T9") = 0 And Silos.Range("T10") = 0 Then 
            checkSGDryerResults = "None"
        Else
            checkSGDryerResults = "SGD2"
            Set workingDryerSchedule = D2Schedule
        End If
    Else
        Dim firstViolationTime As Double
        Dim tempD1 As Double, tempD2 As Double

        firstViolationTime = Silos.Range("K1").Value
        tempD1 = Silos.Range("T9")
        tempD2 = Silos.Range("T10")
        
        If tempD1 <= firstViolationTime Then 
            checkSGDryerResults = "SGD1"
            Set workingDryerSchedule = D1Schedule
        ElseIf tempD2 <= firstViolationTime Then
            checkSGDryerResults = "SGD2"
            Set workingDryerSchedule = D2Schedule
        Else
            checkSGDryerResults = "None"            
        End If
    End If
End Function

Function getExceedTimeStep(DiCause)
    If DiCause = "PED1" Then
        getExceedTimeStep = Silos.Range("R9")
    ElseIf DiCause = "PED2" Then
        getExceedTimeStep = Silos.Range("R10")
    ElseIf DiCause = "SGD1" Then
        getExceedTimeStep = Silos.Range("T9")
    ElseIf DiCause = "SGD2" Then
        getExceedTimeStep = Silos.Range("T10")
    End If
End Function

Function getNextInsertionPointInSchedule(DiCause)
    Dim i As Long
    i = 2
    Do Until workingDryerSchedule.Range("=Ci" & i) <> 0
        i = i + 1
    Loop
    
    getNextInsertionPointInSchedule = workingDryerSchedule.Range("BI" & (i - 1))

End Function

Function isDryerBlockBeforeNextInsert(timeExceed, timeInsert) As Boolean
    If timeExceed <= timeInsert Then
        isDryerBlockBeforeNextInsert = True
    Else
        isDryerBlockBeforeNextInsert = False
    End If
End Function

Function getIdxToDelay(timeExceed)
    getIdxToDelay = Application.WorksheetFunction.Match(timeExceed, workingDryerSchedule.Range("AJ:AJ"), 0)
End Function

Function getCIPHrs(DiCause, D1CIP, D2CIP)
    If DiCause = "PED1" Then
        getCIPHrs = D1CIP
    ElseIf DiCause = "SGD1" Then 
        getCIPHrs = D1CIP
    ElseIf DiCause = "PED2" Then
        getCIPHrs = D2CIP
    ElseIf DiCause = "SGD2" Then
        getCIPHrs = D2CIP
    End If
End Function

Sub resolveSiloContraint(index, CIPHrs, timeExceed, DiCausingViolation, siteCpledCapCurrent)
    Dim timeExceedNext As Double
    Dim existingDryerCIPTimeBase As Double
    Dim siteCpledCapUpdatedCIP As Double
    Dim delayToAdd As Double
    
    existingDryerCIPTimeBase = workingDryerSchedule.Cells(index, 32).Value
    workingDryerSchedule.Cells(index, 32).Value = CIPHrs
    calculateAll
    
    siteCpledCapUpdatedCIP = Round(Silos.Range("R13"), 1)
    delayToAdd = Silos.Range("R7")
    
    timeExceedNext = checkUpdatedCIP(DiCausingViolation)
    ' Print #logic1Textfile, "Current time exceed: " & timeExceed: Space 0
    ' Print #logic1TextFile, "Next time exceed: " & timeExceedNext: Space 0
    checkDryerBlock index, timeExceed, timeExceedNext, existingDryerCIPTimeBase, delayToAdd, CIPHrs, siteCpledCapUpdatedCIP, siteCpledCapCurrent
End Sub

Function checkUpdatedCIP(DiCause)
    If DiCause = "PED1" Then
        checkUpdatedCIP = Silos.Range("R9")
    ElseIf DiCause = "PED2" Then
        checkUpdatedCIP = Silos.Range("R10")
    ElseIf DiCause = "SGD1" Then
        checkUpdatedCIP = Silos.Range("T9")
    ElseIf DiCause = "SGD2" Then
        checkUpdatedCIP = Silos.Range("T10")
    End If
End Function

Sub checkDryerBlock(index, timeExceed, timeExceedNext, currentCIPTimeBase, delay, CIPHrs, cpledCapCIP, currentCpledCap)
    Dim siteCpledCapUpdatedBlock As Double
    
    If timeExceedNext <> timeExceed Then
        ' Print #logic1TextFile, "Next exceeded time step is after current exceed time step. Moving to check improvements.": Space 0
        If cpledCapCIP > currentCpledCap Then
            ' Print #logic1TextFile, "--- Checking effect from adding Blockage Only."
            workingDryerSchedule.Cells(index, 32).Value = currentCIPTimeBase
            workingDryerSchedule.Cells(index, 35).Value = delay
            calculateAll
            
            siteCpledCapUpdatedBlock = Round(Silos.Range("R13"), 1)
            If siteCpledCapUpdatedBlock > cpledCapCIP Then
                ' Print #logic1TextFile, "--- Additional Coupled Capacity incurred by Delay is worse. Reverting back to CIP Only"
                workingDryerSchedule.Cells(index, 35).Value = currentCIPTimeBase
                workingDryerSchedule.Cells(index, 32) = CIPHrs
                calculateAll
            End If
        End If
    Else
        ' Print #logic1TextFile, "Delay at spot still exists. Adding Dryer Delay.": Space 0
        workingDryerSchedule.Cells(index, 35).Value = delay
        calculateAll
    End If
    
End Sub
