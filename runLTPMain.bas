Option Explicit
Dim wb As Workbook
Dim reportWS As Worksheet

' For controlling feasibility of program and reason for premature stoppage
Public isLogic1Feasible As Boolean
Public isLogic3Feasible As Boolean
Public isLogic4Feasible As Boolean
Public reasonForStop As String
Public mainSilo
Public otherSilo

' ============================================= runLTP() Integrated modules logic1 to logic 4 =============================================
Set wb = ThisWorkBook
Set reportWS = wb.Worksheets("Program Report Page")

Dim toAttemptStage1 As String
Dim toAttemptStage3 As String
Dim toAttemptStage4 As String

toAttemptStage1 = reportWS.Range("C2").Value
toAttemptStage3 = reportWS.Range("C3").Value
toAttemptStage4 = reportWS.Range("C4").Value

If IsEmpty(toAttemptStage1) = True OR IsEmpty(toAttemptStage3) = True OR IsEmpty(toAttemptStage4) = True Then 
    MsgBox "Please fill in requirements to running before starting the program. Terminating here."
    End
End If

' ============================================= Logic 1 =============================================
Dim stage1Progress As Range
Set stage1Progress As reportWS.Range("B3")

If toAttemptStage1 = "YES" Then 
    stage1Progress.Value = "Running"
    programModule1.main

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

Else
    stage1Progress.Value = "Chosen not to attempt"
End If

' ============================================= Logic 3 =============================================
Dim stage3Progress As Range
Set stage3Progress As reportWS.Range("B4")

If toAttemptStage3 = "YES" Then 
    stage3Progress.Value = "Running"
    programModule3.ppPouchMain
    If isLogic3Feasible = True Then 
        stage3Progress.Value = "Completed"
    Else
        stage3Progress.Value = "PPCAN & 100DB Inserted. Unable to insert PPPOUCHES via Program. Terminated here."
        reportWS.Range("B7").Value = reasonForStop
        reportWS.Range("B8").Value = "PPPOUCH INSERT"
        End
    End If

Else
    stage3Progress.Value = "Chosen not to attempt"
End If

' ============================================= Logic 4 =============================================
Dim stage4Progress As Range
Set stage4Progress As reportWS.Range("B5")

If toAttemptStage4 = "YES" Then
    stage4Progress.Value = "Running"
    programModule4.PPCanStretchMain
    If isLogic4Feasible = True Then 
        stage4Progress.Value = "Completed"
        reportWS.Range("I3").Formula = "=SUMIFS(D2B1L3B3B4L45T!J:J, D2B1L3B3B4L45T!A:A, ""PP"", D2B1L3B3B4L45TC:C, 5, D2B1L3B3B4L45T!H:H, ""CAN"")"
        reportWS.Range("I4").Formula = "=SUMIFS(D2B1L3B3B4L45T!J:J, D2B1L3B3B4L45T!A:A, ""PP"", D2B1L3B3B4L45TC:C, 5, D2B1L3B3B4L45T!H:H, ""CAN"") * 2" 
    Else
        stage4Progress.Value = "PPCAN / 100DB / PPPOUCHES Inserted. Unable to insert PPCAN (STRETCH). Terminated Here."
        reportWS.Range("B7").Value = reasonForStop
        reportWS.Range("B8").Value = "PPCAN STRETCHING INSERT"
    End If

Else If toAttemptStage4 = "NO" Then
    stage4Progress.Value = "Chosen not to attempt"
    End
End If