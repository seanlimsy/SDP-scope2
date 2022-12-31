Option Explicit
Dim wb As Workbook
Dim D1Schedule As Worksheet, D2Schedule As Worksheet, Silos As Worksheet
Dim workingDryerSchedule As Worksheet

Sub dryerBlockDelayMain(nextInsertTimeStep As Double)
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
    
    Do While True
        siteCpledCapCurrent = Round(Silos.Range("R13"), 1)
        DiCausingViolationPE = checkPEDryerResults
        DiCausingViolationSG = checkSGDryerResults

        If DiCausingViolationPE = "None" And DiCausingViolationSG = "None" Then
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
        
        'nextInsertTimeStep = getNextInsertionPointInSchedule(DiCausingViolation)
        dryerBlockBeforeNextInsert_bool = isDryerBlockBeforeNextInsert(exceedTimeStep, nextInsertTimeStep)
        
        If dryerBlockBeforeNextInsert_bool = False Then
            Exit Do
        Else
            idxToDelay = getIdxToDelay(exceedTimeStep)
            DiCIPHrs = getCIPHrs(DiCausingViolation, D1CipHrs, D2CipHrs)
            resolveSiloContraint idxToDelay, DiCIPHrs, exceedTimeStep, DiCausingViolation, siteCpledCapCurrent

        End If
    Loop
End Sub

Sub initializeWorksheetsStage2()
    Set wb = ThisWorkbook

    setWorksheet D1Schedule, "D1B1L65T"
    setWorksheet D2Schedule, "D2B1L3B3B4L45T"
    setWorksheet Silos, "Silos"
    
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
        checkPEDryerResults = "None"
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
        checkSGDryerResults = "None"
    End If
End Function

Function getExceedTimeStep(DiCause)
    If DiCause = "PED1" Then
        getExceedTimeStep = Silos.Range("R9")
    ElseIf DiCause = "PED2" Then
        getExceedTimeStep = Silos.Range("R10")
    ElseIf DiCause = "SGD1" Then
        getExceedTimeStep = Silos.Range("R9")
    ElseIf DiCause = "SGD2" Then
        getExceedTimeStep = Silos.Range("R10")
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
    Application.Calculate
    
    siteCpledCapUpdatedCIP = Round(Silos.Range("R13"), 1)
    delayToAdd = Silos.Range("R7")
    
    timeExceedNext = checkUpdatedCIP(DiCausingViolation)
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
        If cpledCapCIP > currentCpledCap Then
            workingDryerSchedule.Cells(index, 32).Value = currentCIPTimeBase
            workingDryerSchedule.Cells(index, 35).Value = delay
            Application.Calculate
            
            siteCpledCapUpdatedBlock = Round(Silos.Range("R13"), 1)
            If siteCpledCapUpdatedBlock > cpledCapCIP Then
                workingDryerSchedule.Cells(index, 35).Value = currentCIPTimeBase
                workingDryerSchedule.Cells(index, 32) = CIPHrs
                Application.Calculate
            End If
        End If
    Else
        workingDryerSchedule.Cells(index, 35).Value = delay
        Application.Calculate
    End If
    
End Sub
