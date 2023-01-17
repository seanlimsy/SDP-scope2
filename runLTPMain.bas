Option Explicit
Dim wb As Workbook
Public reportWS As Worksheet

' For controlling feasibility of program and reason for premature stoppage
Public isLogic1Feasible As Boolean
Public isLogic3Feasible As Boolean
Public isLogic4Feasible As Boolean

Public reasonForStop As String
Public mainSilo
Public otherSilo

Sub initializeRunLTP()
    Set wb = ThisWorkbook
    Set reportWS = wb.Worksheets("Program Report Page")
    clearPrevious
End Sub 

Sub clearPrevious()
    reportWS.Range("B3:B5").ClearContents
    reportWS.Range("B7:B8").ClearContents
    reportWS.Range("F3:F4").ClearContents
    reportWS.Range("I3:I4").ClearContents
End Sub

Sub runLTP()
    initializeRunLTP
    Dim toAttemptStage1 As String
    Dim toAttemptStage3 As String
    Dim toAttemptStage4 As String
    
    toAttemptStage1 = reportWS.Range("C3").Value
    toAttemptStage3 = reportWS.Range("C4").Value
    toAttemptStage4 = reportWS.Range("C5").Value

    runStage1 toAttemptStage1
    runStage3 toAttemptStage3
    runStage4 toAttemptStage4

End Sub

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

