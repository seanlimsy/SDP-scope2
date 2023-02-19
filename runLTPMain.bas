Option Explicit
Dim wb As Workbook
Public reportWS As Worksheet

' For controlling feasibility of program
Public isLogic1Feasible As Boolean
Public isLogic3Feasible As Boolean
Public isLogic4Feasible As Boolean

' Reason for premature stoppage
Public reasonForStop As String
Public mainSilo
Public otherSilo

' Debugging
' Public logic1File As String
' Public logic1TextFile As String
' Public logic3File As String
' Public logic3TextFile As String
' Public logic4File as String
' Public logic4TextFile As String

Sub runLTP()
    Dim startTime As Double, endTime As Double, totalTime As Double 
    startTime = Timer
    
    initializeRunLTP
    initializeOutputs

    Dim toAttemptStage1 As String
    Dim toAttemptStage3 As String
    Dim toAttemptStage4 As String
    
    toAttemptStage1 = reportWS.Range("C3").Value
    toAttemptStage3 = reportWS.Range("C4").Value
    toAttemptStage4 = reportWS.Range("C5").Value
    
    runStage1 toAttemptStage1
    runStage3 toAttemptStage3
    runStage4 toAttemptStage4

    endTime = Timer 
    totalTime = endTime - startTime
    reportWS.Range("F5").Value = Format(totalTime/3600, "0.00")

End Sub

' ============================================= Setup Logic =============================================
Sub initializeRunLTP()
    Set wb = ThisWorkbook
    Set reportWS = wb.Worksheets("Program Report Page")
    checkWorksheetsRequired
    clearPrevious
End Sub 

Sub clearPrevious()
    reportWS.Range("B3:B5").ClearContents
    reportWS.Range("B7:B8").ClearContents
    reportWS.Range("F3:F4").ClearContents
    reportWS.Range("I3:I4").ClearContents
End Sub

Sub checkWorksheetsRequired()
    checkExists "D1B1L65T"
    checkExists "D1Sched"
    checkExists "D2B1L3B3B4L45T"
    checkExists "D2Sched"
    checkExists "DBSCH Reorder Select"
    checkExists "Silos"
    checkExists "D1Sched (2)"
    checkExists "D2Sched (2)"
    checkExists "PP"
    checkExists "PP CAN"
    checkExists "PP PCH"
    checkExists "PPRateDS"
    checkExists "PP CAN ADDED THRESHOLD"
    checkExists "PP PCH SPACE" 
End Sub

Sub checkExists(checkSheetName)
    Dim sheetName As Worksheet
    For Each sheetName In wb.Worksheets
        If sheetName.Name = checkSheetName Then 
            Exit Sub
        End If
    Next sheetName
    
    MsgBox "Warning! " & checkSheetName & " is not in current workbook. Please check & update the names"
    End
End Sub

Sub initializeOutputs()
    Dim wbPath As String
    wbPath = ThisWorkbook.Path

    ' logic1File = wbPath & "/logic1.txt"
    ' logic1TextFile = FreeFile
    ' Open logic1File For Output as logic1TextFile

    ' logic3File = wbPath & "/logic3.txt"
    ' logic3TextFile = FreeFile
    ' Open logic3File For Output As logic3TextFile 

    ' logic4File = wbPath & "/logic4.txt"
    ' logic4TextFile = FreeFile
    ' Open logic4File For Output As logic4TextFile 

End Sub

' ============================================= Main Logic =============================================
Sub runStage1(toAttemptStage1)
    Dim stage1Progress As Range
    Set stage1Progress = reportWS.Range("B3")
    
    If toAttemptStage1 = "YES" Then
        stage1Progress.Value = "Running"
        programModule1.main        
    Else
        stage1Progress.Value = "Chosen not to attempt"
        End
    End If
    
    If isLogic1Feasible = True Then
        stage1Progress.Value = "Completed"
        reportWS.Range("F3").Value = mainSilo
        reportWS.Range("F4").Value = otherSilo
    Else
        stage1Progress.Value = "Unable to insert PPCAN & 100DB via Program. Terminated here."
        reportWS.Range("B7").Value = reasonForStop
        reportWS.Range("F3").Value = mainSilo
        reportWS.Range("F4").Value = otherSilo
        reportWS.Range("B8").Value = "PPCAN & 100DB INSERT"
        End
    End If  
End Sub

Sub runStage3(toAttemptStage3)
    Dim stage3Progress As Range
    Set stage3Progress = reportWS.Range("B4")
    
    If toAttemptStage3 = "YES" Then
        stage3Progress.Value = "Running"
        programModule3.ppPouchMain
    Else
        stage3Progress.Value = "Chosen not to attempt"
        End
    End If
    
    If isLogic3Feasible = True Then
        stage3Progress.Value = "Completed"
    Else
        stage3Progress.Value = "PPCAN & 100DB Inserted. Unable to insert PPPOUCHES via Program. Terminated here."
        reportWS.Range("B7").Value = reasonForStop
        reportWS.Range("B8").Value = "PPPOUCH INSERT"
        End
    End If
End Sub
    
Sub runStage4(toAttemptStage4)
    Dim stage4Progress As Range
    Set stage4Progress = reportWS.Range("B5")
    Dim sixMonthPPCANStretch As Double
    Dim oneYearPPCANStretch As Double
    
    If toAttemptStage4 = "YES" Then
        stage4Progress.Value = "Running"
        programModule4.PPCanStretchMain
    ElseIf toAttemptStage4 = "NO" Then
        stage4Progress.Value = "Chosen not to attempt"
        End
    End If
    
    If isLogic4Feasible = True Then
        stage4Progress.Value = "Completed"
        reportWS.Range("I3").Formula = "=SUMIFS(D2B1L3B3B4L45T!J:J, D2B1L3B3B4L45T!A:A, ""PP"", D2B1L3B3B4L45T!C:C, 5, D2B1L3B3B4L45T!H:H, ""CAN"")"
        reportWS.Range("I4").Formula = "=SUMIFS(D2B1L3B3B4L45T!J:J, D2B1L3B3B4L45T!A:A, ""PP"", D2B1L3B3B4L45T!C:C, 5, D2B1L3B3B4L45T!H:H, ""CAN"") * 2"
    Else
        stage4Progress.Value = "PPCAN / 100DB / PPPOUCHES Inserted. Unable to insert PPCAN (STRETCH). Terminated Here."
        reportWS.Range("B7").Value = reasonForStop
        reportWS.Range("B8").Value = "PPCAN STRETCHING INSERT"
    End If
End Sub

