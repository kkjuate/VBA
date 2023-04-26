VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Tracker 
   Caption         =   "Tracker"
   ClientHeight    =   8160
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4515
   OleObjectBlob   =   "Tracker.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Tracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub BAcept_Click()
Dim rRow As Long, Pilot As Boolean, i As Long, ExtPiloto As Boolean
Application.Calculation = xlCalculationManual

If Tracker.LLOB.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta designar el LOB afectado"
    Exit Sub
ElseIf Tracker.LAssigned.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta a quien se asignara el ticket"
    Exit Sub
ElseIf Tracker.LOwnership.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta designar el ownership del issue"
    Exit Sub
ElseIf Tracker.LImpact.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta designar el impacto"
    Exit Sub
ElseIf Tracker.LState.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta designar el estado del ticket"
    Exit Sub
ElseIf Tracker.TStartTime.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta la hora de inicio del issue"
    Exit Sub
ElseIf Tracker.TEndTime.Value = "" And Tracker.LState.Value = "Closed" Then
    MsgBox "No se puede registrar el ticket ya que falta la hora de cierre del ticket"
    Exit Sub
ElseIf Tracker.TAffected.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que falta el numero de usuarios afectados"
    Exit Sub
ElseIf Tracker.LSeverity.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que no se asigno la severidad"
    Exit Sub
ElseIf Tracker.LType.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que no se asigno el tipo de issue"
    Exit Sub
ElseIf Tracker.LCategory.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que no se asigno la categoria del issue"
    Exit Sub
ElseIf Tracker.LIssue.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que no se asigno el issue"
    Exit Sub
ElseIf Tracker.LDescription.Value = "" Then
    MsgBox "No se puede registrar el ticket ya que no se asigno la descripcion del issue"
    Exit Sub
End If

If WorksheetFunction.CountIf(WS(1).Columns, TTicket.Value) = 0 Then
rRow = WS(1).Cells(Rows.Count, 1).End(xlUp).Row + 1
ExtPiloto = False
Else
rRow = WorksheetFunction.Match(TTicket.Value, WS(1).Columns(2), 0)
ExtPiloto = True
End If

With WS(1).Cells(rRow, 1)
    .Value = Month(Date)
    .Offset(0, 1) = Tracker.TTicket.Value
    If Not ExtPiloto Then .Offset(0, 2) = Now
    If Not ExtPiloto Then .Offset(0, 3) = WorksheetFunction.VLookup(Environ("Username"), WS(0).Range("M:N"), 2, 0)
    .Offset(0, 4) = LIssue.Value
    .Offset(0, 5) = LType.Value
    .Offset(0, 6) = LCategory.Value
    .Offset(0, 7) = LImpact.Value
    .Offset(0, 8) = LLOB.Value
    .Offset(0, 9) = LOwnership.Value
    .Offset(0, 10) = CDate(TStartTime.Value)
    If LCase$(Tracker.LState.Value) = "closed" Then .Offset(0, 11) = CDate(Tracker.TEndTime.Value) Else .Offset(0, 11).FormulaR1C1 = "=now()"
    .Offset(0, 12).FormulaR1C1 = "=max(0,rc12-rc11)"
    .Offset(0, 13) = Tracker.TAffected.Value
    .Offset(0, 14) = CInt(Left(LSeverity.Value, 1))
    .Offset(0, 15) = Tracker.LDescription.Value
    .Offset(0, 16) = TClientTicket.Value
    .Offset(0, 17) = LAssigned.Value
    .Offset(0, 18) = Tracker.LState.Value
    .Offset(0, 19) = LSummary.Value
    .Offset(0, 20) = LResolution.Value
    
    If Tracker.LState.Value = "Closed" Then
        Select Case CInt(Left$(Tracker.LSeverity.Value, 1))
            Case 1
                If CDate(Tracker.TEndTime.Value) - CDate(Tracker.TStartTime.Value) > 30 / 60 / 24 Then .Offset(0, 22) = 0 Else .Offset(0, 22) = 1
            Case 2
                If CDate(Tracker.TEndTime.Value) - CDate(Tracker.TStartTime.Value) > 60 / 60 / 24 Then .Offset(0, 22) = 0 Else .Offset(0, 22) = 1
            Case 3
                If CDate(Tracker.TEndTime.Value) - CDate(Tracker.TStartTime.Value) > 240 / 60 / 24 Then .Offset(0, 22) = 0 Else .Offset(0, 22) = 1
            Case 4
                If CDate(Tracker.TEndTime.Value) - CDate(Tracker.TStartTime.Value) > 2880 / 60 / 24 Then .Offset(0, 22) = 0 Else .Offset(0, 22) = 1
        End Select
    Else
        .Offset(0, 22).FormulaR1C1 = "=if(or(and(rc15=1,rc12-rc11<30/60/24),and(rc15=2,rc12-rc11<60/60/24),and(rc15=3,rc12-rc11<240/60/24),and(rc15=4,rc12-rc11<2880/60/24)),1,0)"
    End If
End With
Pilot = False
For i = 2 To WS(0).Cells(Rows.Count, 1).End(xlUp).Row
    If LCase$(WS(0).Cells(i, 1)) = LCase$(Tracker.LType.Value) And _
       LCase$(WS(0).Cells(i, 2)) = LCase$(Tracker.LCategory.Value) And _
       LCase$(WS(0).Cells(i, 3)) = LCase$(Tracker.LIssue.Value) Then
       Pilot = True
    End If
Next
If Not Pilot Then
    With WS(0).Cells(Rows.Count, 1).End(xlUp).Offset(1, 0)
        .Value = Tracker.LType.Value
        .Offset(0, 1) = Tracker.LCategory.Value
        .Offset(0, 2) = Tracker.LIssue.Value
    End With
End If

If WorksheetFunction.CountIf(WS(0).Columns(5), Tracker.LSummary.Value) = 0 Then _
        WS(0).Cells(Rows.Count, 5).End(xlUp).Offset(1, 0) = Tracker.LSummary.Value
If WorksheetFunction.CountIf(WS(0).Columns(7), Tracker.LResolution.Value) = 0 Then _
        WS(0).Cells(Rows.Count, 7).End(xlUp).Offset(1, 0) = Tracker.LResolution.Value

If WorksheetFunction.CountIf(WS(0).Columns(16), Tracker.LAssigned.Value) = 0 Then _
        WS(0).Cells(Rows.Count, 16).End(xlUp).Offset(1, 0) = Tracker.LAssigned.Value
If WorksheetFunction.CountIf(WS(0).Columns(11), Tracker.LLOB.Value) = 0 Then _
        WS(0).Cells(Rows.Count, 11).End(xlUp).Offset(1, 0) = Tracker.LLOB.Value


Dim PV As PivotTable
For Each PV In ThisWorkbook.Worksheets("Pivot").PivotTables
    PV.PivotCache.Refresh
Next
ThisWorkbook.Save
Application.Calculation = xlCalculationAutomatic
Unload Tracker
End Sub

Private Sub BCAncel_Click()
Unload Tracker
End Sub

Private Sub CommandButton4_Click()
Lookup.Show False
End Sub

Private Sub Label1_Click()

End Sub

Private Sub Label14_Click()

End Sub

Private Sub Label16_Click()

End Sub

Private Sub Label4_Click()

End Sub

Private Sub Label5_Click()

End Sub

Private Sub LCategory_Change()
Dim i As Integer
LIssue.Clear
For i = 2 To WS(0).Cells(Rows.Count, 2).End(xlUp).Row
    If LCase$(WS(0).Cells(i, 1)) = LCase$(LType.Value) And LCase$(WS(0).Cells(i, 2)) = LCase$(LCategory.Value) Then _
    Tracker.LIssue.AddItem WS(0).Cells(i, 3)
Next
End Sub

Private Sub LDescription_Change()

End Sub

Private Sub LOwnership_Change()
Tracker.TClientTicket.Enabled = (LCase$(LOwnership.Value) = "vendor")
End Sub

Private Sub LResolution_Change()

End Sub

Private Sub LSeverity_Change()

End Sub

Private Sub LState_Change()
If LCase$(LState.Value) = "closed" Then _
Tracker.TEndTime.Value = Format(Now, "mm/dd/yyyy h:mm:ss")

End Sub

Private Sub LType_Change()
Dim i As Integer
LCategory.Clear
For i = 2 To WS(0).Cells(Rows.Count, 2).End(xlUp).Row
    If WorksheetFunction.CountIfs(WS(0).Range("a2", WS(0).Cells(i, 1)), Tracker.LType.Value, WS(0).Range("b2", WS(0).Cells(i, 2)), WS(0).Cells(i, 2)) = 1 Then _
    Tracker.LCategory.AddItem WS(0).Cells(i, 2)
Next
End Sub

Private Sub TEndTime_AfterUpdate()
Dim DDate As Date
On Error Resume Next
    If Tracker.TStartTime.Value <> "" Then DDate = CDate(Tracker.TEndTime.Value)
    If Err <> 0 Then
        MsgBox "La fecha y hora ingresadas no estan en el formato solicitado: MM/DD/YYYY H:MM:SS = 07/01/2020 15:00:00"
        Tracker.TEndTime.Value = ""
    End If
On Error GoTo 0
End Sub


Private Sub TStartTime_AfterUpdate()
Dim DDate As Date
On Error Resume Next
    If Tracker.TStartTime.Value <> "" Then DDate = CDate(Tracker.TStartTime.Value)
    If Err <> 0 Then
        MsgBox "La fecha y hora ingresadas no estan en el formato solicitado: MM/DD/YYYY H:MM:SS = 07/01/2020 15:00:00"
        Tracker.TStartTime.Value = ""
    End If
On Error GoTo 0
End Sub



Private Sub TTicket_Change()
Dim i As Integer

If Not WorksheetFunction.CountIf(WS(1).Columns(2), Tracker.TTicket.Value) = 0 Then
    i = WorksheetFunction.Match(Tracker.TTicket.Value, WS(1).Columns(2), 0)
    With Tracker

        .LIssue.Value = WS(1).Cells(i, 5)
        .LType.Value = WS(1).Cells(i, 6)
        .LCategory.Value = WS(1).Cells(i, 7)
        .LImpact.Value = WS(1).Cells(i, 8)
        .LLOB.Value = WS(1).Cells(i, 9)
        .LOwnership.Value = WS(1).Cells(i, 10)
        .TStartTime.Value = Format(WS(1).Cells(i, 11), "MM/dd/yyyy h:mm:ss")
        If LCase$(WS(1).Cells(i, 19)) = "closed" Then TEndTime.Value = Format(WS(1).Cells(i, 12), "MM/dd/yyyy h:mm:ss")
        .TAffected.Value = WS(1).Cells(i, 14)
        .LSeverity.Value = WS(1).Cells(i, 15)
        .LDescription.Value = WS(1).Cells(i, 16)
        .TClientTicket.Value = WS(1).Cells(i, 17)
        .LAssigned.Value = WS(1).Cells(i, 18)
        .LState.Value = WS(1).Cells(i, 19)
        .LSummary.Value = WS(1).Cells(i, 20)
        .LResolution.Value = WS(1).Cells(i, 21)
        
    End With
End If

End Sub

Private Sub UserForm_Initialize()
Dim LastTicket As Integer, i As Integer, CCases As Integer
Set WS(0) = ThisWorkbook.Worksheets("Routes")
Set WS(1) = ThisWorkbook.Worksheets("Tracker")


If WorksheetFunction.CountIf(WS(0).Columns(13), Environ("Username")) = 0 Then
    With WS(0).Cells(Rows.Count, 13).End(xlUp).Offset(1, 0)
        .Value = Environ("username")
        .Offset(0, 1) = InputBox("Ingrese su nombre", "nombre", Environ("Username"))
    End With
End If
LastTicket = WorksheetFunction.CountIf(WS(1).Columns(2), "*" & WS(1).Range("B1") & "-" & Right(Environ("Username"), 3) & Format(Date, "YYMMDD") & "*")
TTicket.Value = WS(1).Range("B1") & "-" & Right(Environ("Username"), 3) & Format(Date, "YYMMDD") & Format(LastTicket, "000")

For i = 2 To WS(0).Cells(Rows.Count, 11).End(xlUp).Row
    Tracker.LLOB.AddItem WS(0).Cells(i, 11)
Next
For i = 2 To WS(0).Cells(Rows.Count, 16).End(xlUp).Row
    Tracker.LAssigned.AddItem WS(0).Cells(i, 16)
Next
For i = 2 To WS(0).Cells(Rows.Count, 1).End(xlUp).Row
    If WorksheetFunction.CountIf(WS(0).Range("a2", WS(0).Cells(i, 1)), WS(0).Cells(i, 1)) = 1 Then _
    Tracker.LType.AddItem WS(0).Cells(i, 1)
Next
Tracker.LAssigned.Value = WorksheetFunction.VLookup(Environ("Username"), WS(0).Range("M:N"), 2, 0)
Tracker.LState.Value = "Open"
Tracker.LState.AddItem "Open"
Tracker.LState.AddItem "On Revision"
Tracker.LState.AddItem "Closed"
LImpact.Value = "Production"
LImpact.AddItem "Production"
LImpact.AddItem "Non Production"
LOwnership.AddItem "247"
LOwnership.AddItem "Accedo"
LOwnership.AddItem "Vendor"
LOwnership.AddItem "User"
LSeverity.AddItem "4 - Low"
LSeverity.AddItem "3 - Medium"
LSeverity.AddItem "2 - High"
LSeverity.AddItem "1 - Critical"
LSeverity.Value = "3 - Medium"

For i = 2 To WS(0).Cells(Rows.Count, 5).End(xlUp).Row
    Tracker.LSummary.AddItem WS(0).Cells(i, 5)
Next
For i = 2 To WS(0).Cells(Rows.Count, 7).End(xlUp).Row
    Tracker.LResolution.AddItem WS(0).Cells(i, 7)

Next
Tracker.TStartTime.Value = Format(Now, "mm/dd/yyyy h:mm:ss")
For i = 2 To WS(1).Cells(Rows.Count, 1).End(xlUp).Row
    If WorksheetFunction.CountIf(WS(1).Range("P1", WS(1).Cells(i, 16)), WS(1).Cells(i, 16)) = 1 Then _
    Tracker.LDescription.AddItem WS(1).Cells(i, 16)
Next
Tracker.Caption = WorksheetFunction.CountIf(WS(1).Columns(19), "Open") & " Open. " & WorksheetFunction.CountIf(WS(1).Columns(19), "On Revision") & " On revision. " & WorksheetFunction.CountIf(WS(1).Columns(19), "Closed") & " closed."
End Sub
