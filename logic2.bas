Attribute VB_Name = "Module4"
Option Explicit
Dim wb As Workbook
Dim D1schedule As Worksheet, D2Schedule As Worksheet, Silos As Worksheet
Dim workingDryerSchedule As Worksheet

Sub dryerBlockDelayMain(nextInsertTimeStep As Double)
    Dim D1CipHrs As Double, D2CipHrs As Double
    Dim siteCpledCapCurrent As Double
    Dim DiCausingViolation As String
    
    Dim exceedTimeStep As Double
    'Dim nextInsertTimeStep As Double
    Dim dryerBlockBeforeNextInsert_bool As Boolean
    
    Dim idxToDelay As Double
    Dim DiCIPHrs As Double
    

    Application.AutoRecover.Enabled = False
    initializeWorksheets
    
    'Defining Variables - CIP
    D1CipHrs = wb.Worksheets("Evap DryCIP").Range("T3")
    D2CipHrs = wb.Worksheets("Evap DryCIP").Range("T6")
    
    Do While True
        siteCpledCapCurrent = Round(Silos.Range("R13"), 1)
        DiCausingViolation = checkDryerResults
        exceedTimeStep = getExceedTimeStep(DiCausingViolation)
        
        If DiCausingViolation = "None" Then
            Exit Sub
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

Sub initializeWorksheets()
    'Without Initialising into same workbook
    
    'To adjust to hardcode onto user's path
    'Can also consider moving sheets over to one main workbook
    'Michael: Change reference to an cell value -- solve for this in instructions for documentation -- Lester's Preference KIV
    Set wb = ThisWorkbook

    setWorksheet D1schedule, "D1B1L65T"
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

Function checkDryerResults()
    Dim presenceD1 As String, presenceD2 As String
    presenceD1 = Silos.Range("S9")
    presenceD2 = Silos.Range("S10")
    
    If presenceD1 = "YES" Then
        checkDryerResults = "D1"
        Set workingDryerSchedule = D1schedule
    ElseIf presenceD2 = "YES" Then
        checkDryerResults = "D2"
        Set workingDryerSchedule = D2Schedule
    Else
        checkDryerResults = "None"
    End If
End Function

Function getExceedTimeStep(DiCause)
    If DiCause = "D1" Then
        getExceedTimeStep = Silos.Range("R9")
    ElseIf DiCause = "D2" Then
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
    If DiCause = "D1" Then
        getCIPHrs = D1CIP
    ElseIf DiCause = "D2" Then
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
    If DiCause = "D1" Then
        checkUpdatedCIP = Silos.Range("R9")
    ElseIf DiCause = "D2" Then
        checkUpdatedCIP = Silos.Range("R10")
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


