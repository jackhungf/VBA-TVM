VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   5952
   ClientLeft      =   108
   ClientTop       =   456
   ClientWidth     =   7188
   OleObjectBlob   =   "VBA TVM.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' --- HELPER FUNCTION ---
' Safely converts TextBox strings into numbers. Treats blanks as 0.
Private Function GetVal(ByVal tb As MSForms.TextBox) As Double
    If IsNumeric(tb.Text) Then
        GetVal = CDbl(tb.Text)
    Else
        ' Clear the bad text so the user can re-type it
        tb.Text = ""
        
        ' Safely abort the calculation by throwing a custom error
        Err.Raise Number:=vbObjectError + 1
    End If
End Function

Private Function MasterTVM(N As Double, IY As Double, PV As Double, PMT As Double, FV As Double) As Double
    Dim i As Double
    i = IY / 100
    
    ' Handle zero interest rate scenario
    If Abs(i) < 0.0000001 Then
        MasterTVM = PV + (PMT * N) + FV
    Else
        ' Prevent math errors if (1+i) goes negative during guessing
        If (1 + i) <= 0 Then
            MasterTVM = 1E+15 ' Return a massive error penalty to steer the solver away
            Exit Function
        End If
        ' The core END mode TVM equation (Must equal 0)
        MasterTVM = PV + PMT * ((1 - (1 + i) ^ (-N)) / i) + FV * ((1 + i) ^ (-N))
    End If
End Function

Private Function EvalEq(ByRef Vals() As Double, Unk1Idx As Integer, Unk2Idx As Integer, U1 As Double, U2 As Double) As Double
    Dim V(1 To 5) As Double
    Dim k As Integer
    For k = 1 To 5: V(k) = Vals(k): Next k ' Copy base values
    
    ' Apply the guessed Unkowns (U1 and U2) multiplied by their textbox ratios
    V(Unk1Idx) = U1 * Vals(Unk1Idx)
    V(Unk2Idx) = U2 * Vals(Unk2Idx)
    
    EvalEq = MasterTVM(V(1), V(2), V(3), V(4), V(5))
End Function

Private Sub CBX_simultaneous_Click()
    If CBX_simultaneous.Value Then
        TB_N2.Enabled = Not TB_N2.Enabled
        TB_IY2.Enabled = Not TB_IY2.Enabled
        TB_PV2.Enabled = Not TB_PV2.Enabled
        TB_PMT2.Enabled = Not TB_PMT2.Enabled
        TB_FV2.Enabled = Not TB_FV2.Enabled
        CBX_N.Enabled = Not CBX_N.Enabled
        CBX_IY.Enabled = Not CBX_IY.Enabled
        CBX_PV.Enabled = Not CBX_PV.Enabled
        CBX_PMT.Enabled = Not CBX_PMT.Enabled
        CBX_FV.Enabled = Not CBX_FV.Enabled
        CB_CPT.Enabled = Not CB_CPT.Enabled
        CB_N1.Enabled = Not CB_N1.Enabled
        CB_IY1.Enabled = Not CB_IY1.Enabled
        CB_PV1.Enabled = Not CB_PV1.Enabled
        CB_PMT1.Enabled = Not CB_PMT1.Enabled
        CB_FV1.Enabled = Not CB_FV1.Enabled
    Else
        TB_N2.Enabled = Not TB_N2.Enabled
        TB_IY2.Enabled = Not TB_IY2.Enabled
        TB_PV2.Enabled = Not TB_PV2.Enabled
        TB_PMT2.Enabled = Not TB_PMT2.Enabled
        TB_FV2.Enabled = Not TB_FV2.Enabled
        CBX_N.Enabled = Not CBX_N.Enabled
        CBX_IY.Enabled = Not CBX_IY.Enabled
        CBX_PV.Enabled = Not CBX_PV.Enabled
        CBX_PMT.Enabled = Not CBX_PMT.Enabled
        CBX_FV.Enabled = Not CBX_FV.Enabled
        CB_CPT.Enabled = Not CB_CPT.Enabled
        CB_N1.Enabled = Not CB_N1.Enabled
        CB_IY1.Enabled = Not CB_IY1.Enabled
        CB_PV1.Enabled = Not CB_PV1.Enabled
        CB_PMT1.Enabled = Not CB_PMT1.Enabled
        CB_FV1.Enabled = Not CB_FV1.Enabled
    End If

End Sub

Private Sub AssignResult(SetNum As Integer, Idx As Integer, Result As Double)
    Dim tb As MSForms.TextBox
    If SetNum = 1 Then
        Select Case Idx
            Case 1: Set tb = TB_N1
            Case 2: Set tb = TB_IY1
            Case 3: Set tb = TB_PV1
            Case 4: Set tb = TB_PMT1
            Case 5: Set tb = TB_FV1
        End Select
    Else
        Select Case Idx
            Case 1: Set tb = TB_N2
            Case 2: Set tb = TB_IY2
            Case 3: Set tb = TB_PV2
            Case 4: Set tb = TB_PMT2
            Case 5: Set tb = TB_FV2
        End Select
    End If
    tb.Text = Round(Result, 4)
End Sub

' =========================================================
' 2D NEWTON-RAPHSON SOLVER
' =========================================================
Private Sub CB_CPT_Click()
    On Error GoTo ErrHandler
    Dim Unk1Idx As Integer, Unk2Idx As Integer
    Dim UnkCount As Integer
    
    ' Step 1: Identify which 2 checkboxes are checked
    UnkCount = 0
    If CBX_N.Value Then UnkCount = UnkCount + 1: If UnkCount = 1 Then Unk1Idx = 1 Else Unk2Idx = 1
    If CBX_IY.Value Then UnkCount = UnkCount + 1: If UnkCount = 1 Then Unk1Idx = 2 Else Unk2Idx = 2
    If CBX_PV.Value Then UnkCount = UnkCount + 1: If UnkCount = 1 Then Unk1Idx = 3 Else Unk2Idx = 3
    If CBX_PMT.Value Then UnkCount = UnkCount + 1: If UnkCount = 1 Then Unk1Idx = 4 Else Unk2Idx = 4
    If CBX_FV.Value Then UnkCount = UnkCount + 1: If UnkCount = 1 Then Unk1Idx = 5 Else Unk2Idx = 5
    
    If UnkCount <> 2 Then
        MsgBox "Please check exactly 2 variables to solve for."
        Exit Sub
    End If
    
    ' Step 2: Load the values/multipliers from textboxes into arrays
    Dim Vals1(1 To 5) As Double, Vals2(1 To 5) As Double
    Vals1(1) = GetVal(TB_N1): Vals1(2) = GetVal(TB_IY1): Vals1(3) = GetVal(TB_PV1): Vals1(4) = GetVal(TB_PMT1): Vals1(5) = GetVal(TB_FV1)
    Vals2(1) = GetVal(TB_N2): Vals2(2) = GetVal(TB_IY2): Vals2(3) = GetVal(TB_PV2): Vals2(4) = GetVal(TB_PMT2): Vals2(5) = GetVal(TB_FV2)
    
    ' Step 3: Newton-Raphson Initialization
    Dim U1 As Double, U2 As Double
    Dim E1 As Double, E2 As Double
    Dim a As Double, b As Double, c As Double, d As Double, det As Double
    Dim iter As Integer, h As Double
    
    If Unk1Idx = 1 Then U1 = 20 Else If Unk1Idx = 2 Then U1 = 0.08 Else If Unk1Idx = 3 Then U1 = 1000 Else If Unk1Idx = 4 Then U1 = 200 Else U1 = 1000
    If Unk2Idx = 1 Then U2 = 20 Else If Unk2Idx = 2 Then U2 = 0.08 Else If Unk2Idx = 3 Then U2 = 1000 Else If Unk2Idx = 4 Then U2 = 200 Else U2 = 1000
    
    
    'U1 = 0.01 ' Initial guess for Unknown 1
    U2 = 500 ' Initial guess for Unknown 2
    h = 0.00000001 ' Tiny step size for derivatives
    
    ' Step 4: The Iterative Loop
    For iter = 1 To 500
        ' Calculate current errors
        E1 = EvalEq(Vals1, Unk1Idx, Unk2Idx, U1, U2)
        E2 = EvalEq(Vals2, Unk1Idx, Unk2Idx, U1, U2)
        
        ' Check if both equations equal 0 (Converged successfully!)
        If Abs(E1) < 0.000001 And Abs(E2) < 0.000001 Then Exit For
        
        ' Calculate Jacobian Matrix (Numeric Derivatives)
        a = (EvalEq(Vals1, Unk1Idx, Unk2Idx, U1 + h, U2) - EvalEq(Vals1, Unk1Idx, Unk2Idx, U1 - h, U2)) / (2 * h)
        b = (EvalEq(Vals1, Unk1Idx, Unk2Idx, U1, U2 + h) - EvalEq(Vals1, Unk1Idx, Unk2Idx, U1, U2 - h)) / (2 * h)
        c = (EvalEq(Vals2, Unk1Idx, Unk2Idx, U1 + h, U2) - EvalEq(Vals2, Unk1Idx, Unk2Idx, U1 - h, U2)) / (2 * h)
        d = (EvalEq(Vals2, Unk1Idx, Unk2Idx, U1, U2 + h) - EvalEq(Vals2, Unk1Idx, Unk2Idx, U1, U2 - h)) / (2 * h)
        
        ' Matrix Determinant
        det = a * d - b * c
        If Abs(det) < 0.0000000001 Then
            MsgBox "The numerical approximation failed. Try checking if there is extreme inputs." '& iter
            Exit Sub
        End If
        
        ' Update guesses via 2D Newton-Raphson formula
        U1 = U1 - ((d * E1 - b * E2) / det)
        U2 = U2 - ((-c * E1 + a * E2) / det)
    Next iter
    
    If iter > 150 Then
        MsgBox "Failed to converge in designed iteration."
        Exit Sub
    End If
    
    ' Step 5: Output answers to TextBoxes
    AssignResult 1, Unk1Idx, U1 * Vals1(Unk1Idx)
    AssignResult 1, Unk2Idx, U2 * Vals1(Unk2Idx)
    AssignResult 2, Unk1Idx, U1 * Vals2(Unk1Idx)
    AssignResult 2, Unk2Idx, U2 * Vals2(Unk2Idx)
    
    'MsgBox "Solved! The Base Unknowns are:" & iter & vbCrLf & _
           "U1 = " & Round(U1, 4) & vbCrLf & _
           "U2 = " & Round(U2, 4), vbInformation'
    Exit Sub
ErrHandler:
    MsgBox "Error input"
End Sub





' --- simple TVM ---
' --- COMPUTE N ---
Private Sub CB_N1_Click()
    On Error GoTo ErrHandler
    Dim IY As Double, PV As Double, PMT As Double, FV As Double
    IY = GetVal(TB_IY1) / 100 ' Convert whole % to decimal
    PV = GetVal(TB_PV1)
    PMT = GetVal(TB_PMT1)
    FV = GetVal(TB_FV1)
    
    TB_N1.Text = Round(Application.WorksheetFunction.NPer(IY, PMT, PV, FV), 4)

    Exit Sub
ErrHandler:
    MsgBox "Error input"
End Sub

' --- COMPUTE I/Y ---
Private Sub CB_IY1_Click()
    On Error GoTo ErrHandler
    Dim N As Double, PV As Double, PMT As Double, FV As Double
    N = GetVal(TB_N1)
    PV = GetVal(TB_PV1)
    PMT = GetVal(TB_PMT1)
    FV = GetVal(TB_FV1)
    
    TB_IY1.Text = Round(Application.WorksheetFunction.Rate(N, PMT, PV, FV) * 100, 4)
    Exit Sub
ErrHandler:
    MsgBox "Error input"
End Sub

' --- COMPUTE PV ---
Private Sub CB_PV1_Click()
    On Error GoTo ErrHandler
    Dim N As Double, IY As Double, PMT As Double, FV As Double
    N = GetVal(TB_N1)
    IY = GetVal(TB_IY1) / 100
    PMT = GetVal(TB_PMT1)
    FV = GetVal(TB_FV1)
    
    TB_PV1.Text = Round(Application.WorksheetFunction.PV(IY, N, PMT, FV), 4)
    Exit Sub
ErrHandler:
    MsgBox "Error input"
End Sub

' --- COMPUTE PMT ---
Private Sub CB_PMT1_Click()
    On Error GoTo ErrHandler
    Dim N As Double, IY As Double, PV As Double, FV As Double
    N = GetVal(TB_N1)
    IY = GetVal(TB_IY1) / 100
    PV = GetVal(TB_PV1)
    FV = GetVal(TB_FV1)
    
    TB_PMT1.Text = Round(Application.WorksheetFunction.PMT(IY, N, PV, FV), 4)
    Exit Sub
ErrHandler:
    MsgBox "Error input"
End Sub

' --- COMPUTE FV ---
Private Sub CB_FV1_Click()
    On Error GoTo ErrHandler
    Dim N As Double, IY As Double, PV As Double, PMT As Double
    N = GetVal(TB_N1)
    IY = GetVal(TB_IY1) / 100
    PV = GetVal(TB_PV1)
    PMT = GetVal(TB_PMT1)
    
    TB_FV1.Text = Round(Application.WorksheetFunction.FV(IY, N, PMT, PV), 4)

    Exit Sub
ErrHandler:
    MsgBox "Error input"
End Sub

Private Sub UserForm_Click()

End Sub
