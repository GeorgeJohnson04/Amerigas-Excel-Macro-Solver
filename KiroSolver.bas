Attribute VB_Name = "KiroSolver"
'=====================================================================
' KiroSolver - native Excel min-cost network-flow solver
' Replaces OpenSolver/CBC for the propane supply-planning models.
' Reads the model from the workbook's own Solver names + Period* names,
' solves a min-cost flow (successive shortest paths w/ Dijkstra+potentials),
' writes the decision cells, then verifies against Excel's own formulas.
'
' Lane rule: a BLANK freight cell = closed lane (never shipped).
'            a numeric freight (including 0 = delivered load) = usable.
'
' No OpenSolver, no external tools. Reads dimensions live, so adding
' supply/demand rows does not break it.
'=====================================================================
Option Explicit
Option Base 0

Private Const BIG As Double = 1E+15          ' "infinite" arc capacity (withdraw/inject)
Private Const EPS As Double = 0.000000001    ' 1e-9
Private Const DBGFILE As String = "C:\Users\johnsong\AppData\Local\KiroBuild\kiro_dbg.txt"
Private Const DBGON As Boolean = False

' ---- model dimensions ----
Private mN As Long            ' supply points
Private mM As Long            ' demand points
Private mP As Long            ' periods

' ---- model data (flat arrays) ----
Private mContract() As Double ' [p*N+i]
Private mPrice() As Double    ' [p*N+i]
Private mPercMin() As Double  ' [p*N+i]
Private mPercMax() As Double  ' [p*N+i]
Private mSupMin() As Double   ' [p*N+i]
Private mSupCap() As Double   ' [p*N+i]
Private mDemMin() As Double   ' [p*M+j]
Private mDemCap() As Double   ' [p*M+j]
Private mDemand() As Double   ' [p*M+j]
Private mInitS() As Double    ' [i]
Private mInitD() As Double    ' [j]
Private mTrans() As Double    ' [(p*N+i)*M+j]  cost
Private mClosed() As Boolean  ' [(p*N+i)*M+j]  True = blank freight = closed
Private mShipMin() As Double  ' [(p*N+i)*M+j]
Private mShipMax() As Double  ' [(p*N+i)*M+j]

' ---- adjustable-block addresses (a1, sheet-less) ----
Private mShipA1 As String, mPercA1 As String, mSWdA1 As String
Private mSInjA1 As String, mDWdA1 As String, mDInjA1 As String

' ---- constraint role ranges (a1) ----
Private mDemBalA1 As String, mSupEndA1 As String, mDemEndA1 As String, mSupBalA1 As String

' Contract value as used by the SupplyBalance formula (per p*N+i). May differ from
' mContract (which the objective SUMPRODUCT uses) if a workbook references different
' contract columns in its balance vs its cost. Read empirically via a perc=1 probe.
Private mContractFlow() As Double

' ---- objective cell ----
Private mObjA1 As String

' ---- network (parallel arc arrays; fwd/rev consecutive) ----
Private aTo() As Long
Private aCap() As Double
Private aCost() As Double
Private aHead() As Long       ' head(node) -> first arc id or -1
Private aNext() As Long       ' next(arc) -> next arc from same tail or -1
Private mArc As Long          ' arc count
Private mNode As Long         ' node count
Private mOffset As Double     ' objective offset from lower bounds
Private mB() As Double        ' node supplies

' ---- contradictory-bound detection (min > max => no feasible value exists) ----
Private mBoundConflicts As Long
Private mBoundErr As String

' ---- solution-recovery metadata ----
Private pFwd() As Long, pLo() As Double, pInit() As Double  ' purchase [p*N+i]
Private shFwd() As Long, shLo() As Double, shInit() As Double ' ship [(p*N+i)*M+j]
Private siFwd() As Long, swFwd() As Long                      ' s_inj / s_wd [p*N+i]
Private diFwd() As Long, dwFwd() As Long                      ' d_inj / d_wd [p*M+j]

Private mWs As Worksheet

' ---- network node bases + solve scratch ----
Private mPN As Long, mPMc As Long
Private baseS As Long, baseD As Long, baseSI As Long, baseDI As Long
Private nodeR As Long, nodeSS As Long, nodeTT As Long
Private mNeed As Double
Private ttArc() As Long, ttNode() As Long, mTT As Long
Private hn() As Long, hk() As Double, hsz As Long

'=====================================================================
' ENTRY POINTS
'=====================================================================
Private Sub Dbg(ByVal s As String)
    If Not DBGON Then Exit Sub
    On Error Resume Next
    Dim f As Integer
    f = FreeFile
    Open DBGFILE For Append As #f
    Print #f, Format(Now, "hh:nn:ss") & "  " & s
    Close #f
    On Error GoTo 0
End Sub

Public Sub SolveModel()
    Dim s As String
    s = KiroSolveSilent()
    MsgBox s, vbInformation, "Kiro Solver"
End Sub

Public Sub KiroSolveRibbon(control As Object)
    MsgBox KiroSolveSilent(), vbInformation, "Kiro Solver"
End Sub

' Trivial probe: confirms the module compiles and runs (and that Dbg can write).
Public Function KiroPing() As String
    Dbg "PING ok"
    KiroPing = "PONG"
End Function

' TEMP diagnostic: isolate whether a 2D Double array write lands on the ship block,
' while replicating the KiroSolveSilent state (manual calc, events off, zeroed cells).
' Runs when the workbook opens (macros enabled): gives SolveModel a Ctrl+Shift+K shortcut.
Public Sub Auto_Open()
    On Error Resume Next
    Application.OnKey "^+k", "'" & ThisWorkbook.Name & "'!SolveModel"
End Sub

' Core routine. Returns a one-line status string.
Public Function KiroSolveSilent() As String
    Dim calcSave As XlCalculation, suSave As Boolean, evSave As Boolean
    Dim t0 As Double
    Dim savedVals As Variant   ' array of the 6 adj blocks' original Value2
    Dim okRead As Boolean
    Dim status As String

    On Error GoTo Failed
    On Error Resume Next
    Kill DBGFILE
    On Error GoTo Failed
    Dbg "=== KiroSolveSilent start ==="
    t0 = Timer
    suSave = Application.ScreenUpdating
    evSave = Application.EnableEvents
    calcSave = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    On Error Resume Next
    Application.Calculation = xlCalculationManual
    On Error GoTo Failed

    Set mWs = ThisWorkbook.Worksheets("ModelBase")
    Dbg "got ModelBase; reading model"

    ' --- read the model definition ---
    If Not ReadModel() Then
        status = "KIRO: could not read the model definition (Solver names missing?)."
        GoTo Restore
    End If

    Dbg "model read: N=" & mN & " M=" & mM & " P=" & mP

    ' --- save current adjustable values so we can restore on failure ---
    savedVals = SaveAdjustables()
    Dbg "saved adjustables"

    ' --- zero probe for demand + initial inventory ---
    ZeroAdjustables
    Dbg "zeroed; calculating"
    Application.Calculate
    Dbg "zero-probe calc done"
    If Not ReadDemandInit() Then
        RestoreAdjustables savedVals
        Application.Calculate
        status = "KIRO: could not read demand/inventory (zero probe failed)."
        GoTo Restore
    End If

    Dbg "demand/init read; reading contract-flow (perc=1 probe)"
    ReadContractFlow
    Dbg "contract-flow read; building network"
    ' --- build + solve the network ---
    Dim feasible As Boolean, obj As Double, diag As String
    BuildNetwork
    Dbg "network built: nodes=" & mNode & " arcs=" & mArc & " need=" & Format(mNeed, "0.0")
    If mBoundConflicts > 0 Then
        RestoreAdjustables savedVals
        Application.Calculate
        status = "KIRO: INFEASIBLE - contradictory min/max input bounds in " & mBoundConflicts & " cell(s). " & _
                 IIf(Len(mBoundErr) > 0, mBoundErr & ". ", "") & _
                 "No value can satisfy a minimum that exceeds its maximum - please fix the inputs. (original values restored; nothing written)"
        Dbg "bound conflicts=" & mBoundConflicts & " " & mBoundErr
        GoTo Restore
    End If
    obj = SolveMCF(feasible, diag)
    Dbg "solve done: feasible=" & feasible & " obj=" & Format(obj, "0.00")

    If Not feasible Then
        RestoreAdjustables savedVals
        Application.Calculate
        status = "KIRO: INFEASIBLE. " & diag & _
                 "  (original values restored; nothing written)"
        GoTo Restore
    End If

    ' --- write the solution, recalc, verify ---
    Dbg "writing solution"
    WriteSolution
    Dbg "wrote; recalculating"
    Application.Calculate
    Dbg "recalc done; verifying"

    Dim vmsg As String
    If Not VerifySolution(obj, vmsg) Then
        RestoreAdjustables savedVals
        Application.Calculate
        status = "KIRO: solve produced an invalid result (" & vmsg & _
                 "). Original values restored."
        GoTo Restore
    End If

    status = "KIRO OPTIMAL. Cost = " & Format(obj, "#,##0.00") & _
             "   |  " & Format(Timer - t0, "0.0") & " s  |  verified OK"
    If Len(vmsg) > 0 Then status = status & "  [" & vmsg & "]"

Restore:
    On Error Resume Next
    Application.Calculation = calcSave
    Application.EnableEvents = evSave
    Application.ScreenUpdating = suSave
    On Error GoTo 0
    KiroSolveSilent = status
    Exit Function

Failed:
    status = "KIRO ERROR: " & Err.Number & " - " & Err.Description
    Resume Restore
End Function

'=====================================================================
' MODEL READING
'=====================================================================
Private Function ReadModel() As Boolean
    Dim rt As String, parts() As String, i As Long
    ReadModel = False

    ' --- adjustable blocks from solver_adj (6 comma-separated areas) ---
    rt = NameRefersTo("solver_adj")
    If Len(rt) = 0 Then Exit Function
    rt = StripEq(rt)
    parts = Split(rt, ",")
    Dbg "adj areas=" & (UBound(parts) + 1) & " altShipName=" & (Len(NameRefersTo("Period1_p_AltShip")) > 0)
    If UBound(parts) < 5 Then Exit Function
    mShipA1 = CellPart(parts(0))
    mPercA1 = CellPart(parts(1))
    mSWdA1 = CellPart(parts(2))
    mSInjA1 = CellPart(parts(3))
    mDWdA1 = CellPart(parts(4))
    mDInjA1 = CellPart(parts(5))

    ' dimensions from perc (N x P) and ship (P*N x M)
    Dim pr1 As Long, pc1 As Long, pr2 As Long, pc2 As Long
    ParseA1 mPercA1, pr1, pc1, pr2, pc2
    mN = pr2 - pr1 + 1
    mP = pc2 - pc1 + 1
    Dim sr1 As Long, sc1 As Long, sr2 As Long, sc2 As Long
    ParseA1 mShipA1, sr1, sc1, sr2, sc2
    mM = sc2 - sc1 + 1
    If (sr2 - sr1 + 1) <> mP * mN Then Exit Function

    Dim PN As Long, PM As Long
    PN = mP * mN: PM = mP * mM

    ' --- classify constraints: get bound ranges + demBal/supEnd/demEnd ---
    Dim num As Long
    num = CLng(Val(StripEq(NameRefersTo("solver_num"))))
    If num < 10 Then num = 10

    Dim percMinA1 As String, percMaxA1 As String
    Dim shipMinA1 As String, shipMaxA1 As String
    Dim supMinA1 As String, supMaxA1 As String
    Dim demMinA1 As String, demMaxA1 As String
    Dim k As Long, lhs As String, rel As Long, rhsRT As String, rhsA1 As String
    Dim lr1 As Long, lc1 As Long, lr2 As Long, lc2 As Long
    Dim lcols As Long

    For k = 1 To num
        lhs = CellPart(StripEq(NameRefersTo("solver_lhs" & k)))
        rel = CLng(Val(StripEq(NameRefersTo("solver_rel" & k))))
        rhsRT = NameRefersTo("solver_rhs" & k)
        rhsA1 = ""
        If InStr(rhsRT, "!") > 0 Then rhsA1 = CellPart(StripEq(rhsRT))
        ParseA1 lhs, lr1, lc1, lr2, lc2
        lcols = lc2 - lc1 + 1

        If SameRange(lhs, mPercA1) Then
            If rel = 3 Then
                percMinA1 = rhsA1
            ElseIf rel = 1 Then
                percMaxA1 = rhsA1
            End If
        ElseIf SameRange(lhs, mShipA1) Then
            If rel = 3 Then
                shipMinA1 = rhsA1
            ElseIf rel = 1 Then
                shipMaxA1 = rhsA1
            End If
        ElseIf rel = 2 Then
            If lcols > 1 Then
                mDemBalA1 = lhs                      ' DemBal is M x P
            Else
                mSupBalA1 = lhs                      ' SupBal is P*N x 1
            End If
        Else
            If lcols = 1 Then
                If rel = 1 Then
                    supMaxA1 = rhsA1
                    mSupEndA1 = lhs
                Else
                    supMinA1 = rhsA1
                End If
            Else
                If rel = 1 Then
                    demMaxA1 = rhsA1
                    mDemEndA1 = lhs
                Else
                    demMinA1 = rhsA1
                End If
            End If
        End If
    Next k

    mObjA1 = CellPart(StripEq(NameRefersTo("solver_opt")))
    If Len(mDemBalA1) = 0 Or Len(mSupEndA1) = 0 Or Len(mDemEndA1) = 0 Or Len(mSupBalA1) = 0 Then Exit Function

    ' --- per-period inputs: contract, price, trans (+ blank/closed mask) ---
    ReDim mContract(0 To PN - 1)
    ReDim mPrice(0 To PN - 1)
    ReDim mTrans(0 To PN * mM - 1)
    ReDim mClosed(0 To PN * mM - 1)

    Dim p As Long, ii As Long, jj As Long, baseRow As Long
    Dim cA1 As String, prA1 As String, tA1 As String
    Dim cV As Variant, pV As Variant, tV As Variant
    Dim cColVec As Boolean, prColVec As Boolean, tNormal As Boolean
    Dim cr1 As Long, cc1 As Long, cr2 As Long, cc2 As Long
    Dim tr1 As Long, tc1 As Long, tr2 As Long, tc2 As Long
    Dim tRows As Long, tCols As Long

    For p = 1 To mP
        cA1 = PeriodNameCell("Period" & p & "_p_Contract", "Period1_p_Contract")
        prA1 = PeriodNameCell("Period" & p & "_p_Price", "Period1_p_Price")
        tA1 = PeriodNameCell("Period" & p & "_p_Trans_Cost", "Period1_p_Trans_Cost")
        If Len(cA1) = 0 Or Len(prA1) = 0 Or Len(tA1) = 0 Then Exit Function

        cV = mWs.Range(cA1).Value2
        pV = mWs.Range(prA1).Value2
        tV = mWs.Range(tA1).Value2
        ParseA1 cA1, cr1, cc1, cr2, cc2
        cColVec = ((cr2 - cr1) >= (cc2 - cc1))
        ParseA1 prA1, tr1, tc1, tr2, tc2
        prColVec = ((tr2 - tr1) >= (tc2 - tc1))
        ParseA1 tA1, tr1, tc1, tr2, tc2
        tRows = tr2 - tr1 + 1: tCols = tc2 - tc1 + 1
        tNormal = (tRows = mN)

        baseRow = (p - 1) * mN
        For ii = 0 To mN - 1
            mContract(baseRow + ii) = VecVal(cV, ii, cColVec)
            mPrice(baseRow + ii) = VecVal(pV, ii, prColVec)
            For jj = 0 To mM - 1
                Dim cell As Variant
                If tNormal Then
                    cell = MatCell(tV, ii, jj)
                Else
                    cell = MatCell(tV, jj, ii)
                End If
                Dim idx As Long
                idx = (baseRow + ii) * mM + jj
                If IsNumeric2(cell) Then
                    mTrans(idx) = CDbl(cell)
                    mClosed(idx) = False
                Else
                    mTrans(idx) = 0#
                    mClosed(idx) = True      ' blank / text = closed lane
                End If
            Next jj
        Next ii
    Next p

    ' --- bound ranges ---
    ReadBoundNxP percMinA1, mPercMin, PN
    ReadBoundNxP percMaxA1, mPercMax, PN
    ReadBoundPNx1 supMinA1, mSupMin, PN
    ReadBoundPNx1 supMaxA1, mSupCap, PN
    ReadBoundMxP demMinA1, mDemMin, PM
    ReadBoundMxP demMaxA1, mDemCap, PM
    ReadBoundShip shipMinA1, mShipMin, PN
    ReadBoundShip shipMaxA1, mShipMax, PN

    ReadModel = True
End Function

'=====================================================================
' HELPERS: names, address parsing, cell access
'=====================================================================
Private Function NameRefersTo(ByVal nm As String) As String
    Dim s As String
    s = ""
    On Error Resume Next
    s = mWs.Names(nm).RefersTo
    If Len(s) = 0 Then s = ThisWorkbook.Names(nm).RefersTo
    On Error GoTo 0
    NameRefersTo = s
End Function

Private Function StripEq(ByVal s As String) As String
    If Left$(s, 1) = "=" Then StripEq = Mid$(s, 2) Else StripEq = s
End Function

Private Function CellPart(ByVal s As String) As String
    Dim p As Long
    p = InStr(s, "!")
    If p > 0 Then CellPart = Mid$(s, p + 1) Else CellPart = s
End Function

Private Function ColNum(ByVal letters As String) As Long
    Dim i As Long, n As Long, ch As Long
    n = 0
    letters = UCase$(letters)
    For i = 1 To Len(letters)
        ch = Asc(Mid$(letters, i, 1))
        If ch < 65 Or ch > 90 Then Exit For
        n = n * 26 + (ch - 64)
    Next i
    ColNum = n
End Function

' Parse "$D$259:$M$378" or "$D$259" (sheet already stripped) into 1-based coords.
Private Sub ParseA1(ByVal a1 As String, ByRef r1 As Long, ByRef c1 As Long, _
                    ByRef r2 As Long, ByRef c2 As Long)
    Dim parts() As String
    a1 = Replace$(a1, "$", "")
    parts = Split(a1, ":")
    OneCell parts(0), r1, c1
    If UBound(parts) >= 1 Then
        OneCell parts(1), r2, c2
    Else
        r2 = r1: c2 = c1
    End If
    Dim t As Long
    If r2 < r1 Then t = r1: r1 = r2: r2 = t
    If c2 < c1 Then t = c1: c1 = c2: c2 = t
End Sub

Private Sub OneCell(ByVal s As String, ByRef r As Long, ByRef c As Long)
    Dim i As Long, ch As Long, letters As String, digits As String
    For i = 1 To Len(s)
        ch = Asc(UCase$(Mid$(s, i, 1)))
        If ch >= 65 And ch <= 90 Then
            letters = letters & Mid$(s, i, 1)
        ElseIf ch >= 48 And ch <= 57 Then
            digits = digits & Mid$(s, i, 1)
        End If
    Next i
    c = ColNum(letters)
    r = CLng(Val(digits))
End Sub

Private Function SameRange(ByVal a As String, ByVal b As String) As Boolean
    Dim ar1 As Long, ac1 As Long, ar2 As Long, ac2 As Long
    Dim br1 As Long, bc1 As Long, br2 As Long, bc2 As Long
    ParseA1 a, ar1, ac1, ar2, ac2
    ParseA1 b, br1, bc1, br2, bc2
    SameRange = (ar1 = br1 And ac1 = bc1 And ar2 = br2 And ac2 = bc2)
End Function

Private Function PeriodNameCell(ByVal nm As String, ByVal fallback As String) As String
    Dim rt As String
    rt = NameRefersTo(nm)
    If Len(rt) = 0 Then rt = NameRefersTo(fallback)
    PeriodNameCell = CellPart(StripEq(rt))
End Function

Private Function IsNumeric2(ByVal x As Variant) As Boolean
    Select Case VarType(x)
        Case vbInteger, vbLong, vbSingle, vbDouble, vbCurrency, vbByte
            IsNumeric2 = True
        Case Else
            IsNumeric2 = False
    End Select
End Function

Private Function RawRC(ByVal v As Variant, ByVal r As Long, ByVal c As Long) As Variant
    If IsArray(v) Then RawRC = v(r, c) Else RawRC = v
End Function

Private Function NumRC(ByVal v As Variant, ByVal r As Long, ByVal c As Long) As Double
    Dim x As Variant
    If IsArray(v) Then x = v(r, c) Else x = v
    If IsNumeric2(x) Then NumRC = CDbl(x) Else NumRC = 0#
End Function

Private Function VecVal(ByVal v As Variant, ByVal i0 As Long, ByVal colVec As Boolean) As Double
    If colVec Then
        VecVal = NumRC(v, i0 + 1, 1)
    Else
        VecVal = NumRC(v, 1, i0 + 1)
    End If
End Function

Private Function MatCell(ByVal v As Variant, ByVal r0 As Long, ByVal c0 As Long) As Variant
    MatCell = RawRC(v, r0 + 1, c0 + 1)
End Function

'=====================================================================
' BOUND READERS  (a1 "" -> fill 0)
'=====================================================================
Private Sub ReadBoundNxP(ByVal a1 As String, ByRef arr() As Double, ByVal PN As Long)
    ReDim arr(0 To PN - 1)
    If Len(a1) = 0 Then Exit Sub
    Dim v As Variant: v = mWs.Range(a1).Value2
    Dim p As Long, i As Long
    For p = 0 To mP - 1
        For i = 0 To mN - 1
            arr(p * mN + i) = NumRC(v, i + 1, p + 1)
        Next i
    Next p
End Sub

Private Sub ReadBoundPNx1(ByVal a1 As String, ByRef arr() As Double, ByVal PN As Long)
    ReDim arr(0 To PN - 1)
    If Len(a1) = 0 Then Exit Sub
    Dim v As Variant: v = mWs.Range(a1).Value2
    Dim k As Long
    For k = 0 To PN - 1
        arr(k) = NumRC(v, k + 1, 1)
    Next k
End Sub

Private Sub ReadBoundMxP(ByVal a1 As String, ByRef arr() As Double, ByVal PM As Long)
    ReDim arr(0 To PM - 1)
    If Len(a1) = 0 Then Exit Sub
    Dim v As Variant: v = mWs.Range(a1).Value2
    Dim p As Long, j As Long
    For p = 0 To mP - 1
        For j = 0 To mM - 1
            arr(p * mM + j) = NumRC(v, j + 1, p + 1)
        Next j
    Next p
End Sub

Private Sub ReadBoundShip(ByVal a1 As String, ByRef arr() As Double, ByVal PN As Long)
    ReDim arr(0 To PN * mM - 1)
    If Len(a1) = 0 Then Exit Sub
    Dim v As Variant: v = mWs.Range(a1).Value2
    Dim row As Long, j As Long
    For row = 0 To PN - 1
        For j = 0 To mM - 1
            arr(row * mM + j) = NumRC(v, row + 1, j + 1)
        Next j
    Next row
End Sub

'=====================================================================
' HIDDEN-AWARE BLOCK WRITES
' These models hide unused supply/demand rows (via filter/outline). A bulk
' "Range.Value2 = array/scalar" write SKIPS hidden cells, which both corrupts
' the zero-probe and loses part of the solution. Direct single-cell writes DO
' reach hidden cells, so when a block spans hidden rows/cols we write cell by
' cell. Blocks with no hidden cells use the fast bulk write unchanged.
'=====================================================================
Private Function BlockHidden(ByVal a1 As String) As Boolean
    On Error Resume Next
    Dim rg As Range: Set rg = mWs.Range(a1)
    Dim vis As Long: vis = -1
    vis = rg.SpecialCells(xlCellTypeVisible).Count   ' errors if nothing visible
    On Error GoTo 0
    If vis < 0 Then
        BlockHidden = True
    Else
        BlockHidden = (vis <> rg.Count)
    End If
End Function

Private Sub FillBlock(ByVal a1 As String, ByVal v As Double)
    Dim rg As Range: Set rg = mWs.Range(a1)
    If Not BlockHidden(a1) Then
        rg.Value2 = v
    Else
        Dim r As Long, c As Long
        For r = 1 To rg.Rows.Count
            For c = 1 To rg.Columns.Count
                rg.Cells(r, c).Value2 = v
            Next c
        Next r
    End If
End Sub

Private Sub PutBlockD(ByVal a1 As String, ByRef arr() As Double)
    Dim rg As Range: Set rg = mWs.Range(a1)
    If Not BlockHidden(a1) Then
        rg.Value2 = arr
    Else
        Dim r As Long, c As Long
        For r = 1 To rg.Rows.Count
            For c = 1 To rg.Columns.Count
                rg.Cells(r, c).Value2 = arr(r, c)
            Next c
        Next r
    End If
End Sub

Private Sub PutBlockV(ByVal a1 As String, ByVal arr As Variant)
    Dim rg As Range: Set rg = mWs.Range(a1)
    If Not BlockHidden(a1) Then
        rg.Value2 = arr
    Else
        Dim r As Long, c As Long
        For r = 1 To rg.Rows.Count
            For c = 1 To rg.Columns.Count
                rg.Cells(r, c).Value2 = arr(r, c)
            Next c
        Next r
    End If
End Sub

'=====================================================================
' SAVE / ZERO / RESTORE adjustable blocks; demand+init zero probe
'=====================================================================
Private Function SaveAdjustables() As Variant
    Dim a(0 To 5) As Variant
    a(0) = mWs.Range(mShipA1).Value2
    a(1) = mWs.Range(mPercA1).Value2
    a(2) = mWs.Range(mSWdA1).Value2
    a(3) = mWs.Range(mSInjA1).Value2
    a(4) = mWs.Range(mDWdA1).Value2
    a(5) = mWs.Range(mDInjA1).Value2
    SaveAdjustables = a
End Function

Private Sub ZeroAdjustables()
    FillBlock mShipA1, 0#
    FillBlock mPercA1, 0#
    FillBlock mSWdA1, 0#
    FillBlock mSInjA1, 0#
    FillBlock mDWdA1, 0#
    FillBlock mDInjA1, 0#
End Sub

Private Sub RestoreAdjustables(ByVal a As Variant)
    PutBlockV mShipA1, a(0)
    PutBlockV mPercA1, a(1)
    PutBlockV mSWdA1, a(2)
    PutBlockV mSInjA1, a(3)
    PutBlockV mDWdA1, a(4)
    PutBlockV mDInjA1, a(5)
End Sub

Private Function ReadDemandInit() As Boolean
    On Error GoTo Fail
    Dim PM As Long: PM = mP * mM
    ReDim mDemand(0 To PM - 1)
    ReDim mInitS(0 To mN - 1)
    ReDim mInitD(0 To mM - 1)

    Dim vb As Variant: vb = mWs.Range(mDemBalA1).Value2   ' M x P
    Dim vs As Variant: vs = mWs.Range(mSupEndA1).Value2   ' PN x 1
    Dim vd As Variant: vd = mWs.Range(mDemEndA1).Value2   ' M x P

    Dim p As Long, j As Long, i As Long
    For p = 0 To mP - 1
        For j = 0 To mM - 1
            mDemand(p * mM + j) = -1# * NumRC(vb, j + 1, p + 1)
        Next j
    Next p
    For i = 0 To mN - 1
        mInitS(i) = NumRC(vs, i + 1, 1)      ' period 1 row = i
    Next i
    For j = 0 To mM - 1
        mInitD(j) = NumRC(vd, j + 1, 1)      ' period 1 col = 1
    Next j
    ReadDemandInit = True
    Exit Function
Fail:
    ReadDemandInit = False
End Function

' Perc=1 probe: with all adjustables 0 except perc=1, SupplyBalance = contract (the exact
' contract value the balance uses). Captures contract even when a workbook's balance and
' objective reference different contract columns.
Private Sub ReadContractFlow()
    Dim PN As Long: PN = mP * mN
    ReDim mContractFlow(0 To PN - 1)
    ' Differential probe. SupBal = contractFlow*perc + priorInventory - shipOut - netStore.
    ' At perc=0 (all adjustables still zeroed) SupBal = the constant prior-inventory term.
    ' At perc=1 SupBal = contractFlow + that same constant. Subtracting isolates contractFlow
    ' regardless of any nonzero beginning inventory.
    Dim base As Variant: base = mWs.Range(mSupBalA1).Value2
    FillBlock mPercA1, 1#               ' ship/swd/sinj/dwd/dinj remain 0 from the zero probe
    Application.Calculate
    Dim v As Variant: v = mWs.Range(mSupBalA1).Value2
    Dim k As Long
    For k = 0 To PN - 1
        mContractFlow(k) = NumRC(v, k + 1, 1) - NumRC(base, k + 1, 1)
    Next k
End Sub

'=====================================================================
' NETWORK CONSTRUCTION
'=====================================================================
Private Function maxd(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then maxd = a Else maxd = b
End Function

Private Function AddArc(ByVal u As Long, ByVal v As Long, ByVal cap As Double, ByVal cost As Double) As Long
    aTo(mArc) = v: aCap(mArc) = cap: aCost(mArc) = cost: aNext(mArc) = aHead(u): aHead(u) = mArc
    Dim fwd As Long: fwd = mArc: mArc = mArc + 1
    aTo(mArc) = u: aCap(mArc) = 0#: aCost(mArc) = -cost: aNext(mArc) = aHead(v): aHead(v) = mArc: mArc = mArc + 1
    AddArc = fwd
End Function

' Lower-bounded arc: transforms node balances + offset, adds residual arc. Returns fwd id (-1 if none).
Private Function AddArcLB(ByVal u As Long, ByVal v As Long, ByVal lo As Double, _
                          ByVal hi As Double, ByVal cost As Double) As Long
    Dim capv As Double
    capv = hi - lo
    If capv < -EPS Then mBoundConflicts = mBoundConflicts + 1   ' min > max: no feasible value
    If capv < 0 Then capv = 0#        ' crossed bound -> 0 cap (caught by feasibility/verify)
    mB(u) = mB(u) - lo
    mB(v) = mB(v) + lo
    mOffset = mOffset + lo * cost
    If capv > 0 Then
        AddArcLB = AddArc(u, v, capv, cost)
    Else
        AddArcLB = -1
    End If
End Function

Private Sub BuildNetwork()
    Dim p As Long, i As Long, j As Long, k As Long, nxt As Long
    mPN = mP * mN: mPMc = mP * mM
    baseS = 0: baseD = mPN: baseSI = mPN + mPMc: baseDI = 2 * mPN + mPMc
    nodeR = 2 * mPN + 2 * mPMc: nodeSS = nodeR + 1: nodeTT = nodeR + 2
    mNode = nodeR + 3

    Dim maxArc As Long
    maxArc = 2 * (mPN * mM + 4 * mPN + 3 * mPMc + mNode + 32)
    ReDim aTo(0 To maxArc): ReDim aCap(0 To maxArc)
    ReDim aCost(0 To maxArc): ReDim aNext(0 To maxArc)
    ReDim aHead(0 To mNode - 1)
    For i = 0 To mNode - 1: aHead(i) = -1: Next i
    ReDim mB(0 To mNode - 1)
    mArc = 0: mOffset = 0#
    mBoundConflicts = 0: mBoundErr = ""

    ReDim pFwd(0 To mPN - 1): ReDim pLo(0 To mPN - 1): ReDim pInit(0 To mPN - 1)
    ReDim siFwd(0 To mPN - 1): ReDim swFwd(0 To mPN - 1)
    ReDim shFwd(0 To mPN * mM - 1): ReDim shLo(0 To mPN * mM - 1): ReDim shInit(0 To mPN * mM - 1)
    ReDim diFwd(0 To mPMc - 1): ReDim dwFwd(0 To mPMc - 1)

    ' supply side
    For p = 0 To mP - 1
        For i = 0 To mN - 1
            k = p * mN + i
            Dim sNode As Long, siNode As Long
            sNode = baseS + p * mN + i
            siNode = baseSI + p * mN + i
            Dim cf As Double, cpu As Double, lo As Double, hi As Double
            cf = mContractFlow(k)       ' contract value the SupplyBalance actually uses (physical kgal/%)
            If cf > EPS Then
                ' Physical purchase flow = cf * perc. The objective charges
                ' price * mContract * perc, so cost per physical unit = price*mContract/cf.
                ' When the balance and objective reference the same contract column
                ' (the usual case) cf = mContract and cpu reduces to price.
                cpu = mPrice(k) * mContract(k) / cf
                lo = cf * mPercMin(k): hi = cf * mPercMax(k)
                If mPercMin(k) > mPercMax(k) + 0.000001 And Len(mBoundErr) = 0 Then _
                    mBoundErr = "supply #" & (i + 1) & " period " & (p + 1) & " has minimum take " & _
                                Format(mPercMin(k), "0.0%") & " above maximum take " & Format(mPercMax(k), "0.0%")
                pLo(k) = lo: pInit(k) = maxd(0#, hi - lo)
                pFwd(k) = AddArcLB(nodeR, sNode, lo, hi, cpu)        ' purchase
            Else
                ' Balance contract is 0: perc moves no physical propane, it only
                ' adds price*mContract*perc to the objective, so it pins to its minimum.
                mOffset = mOffset + mPrice(k) * mContract(k) * mPercMin(k)
                pLo(k) = 0#: pInit(k) = 0#
                pFwd(k) = -1
            End If
            siFwd(k) = AddArc(sNode, siNode, BIG, 0#)                ' s_inj
            swFwd(k) = AddArc(siNode, sNode, BIG, 0#)                ' s_wd
            If p < mP - 1 Then nxt = baseSI + (p + 1) * mN + i Else nxt = nodeR
            AddArcLB siNode, nxt, mSupMin(k), mSupCap(k), 0#          ' s_end carry
        Next i
    Next p

    ' demand side
    For p = 0 To mP - 1
        For j = 0 To mM - 1
            k = p * mM + j
            Dim dNode As Long, diNode As Long
            dNode = baseD + p * mM + j
            diNode = baseDI + p * mM + j
            diFwd(k) = AddArc(dNode, diNode, BIG, 0#)                ' d_inj
            dwFwd(k) = AddArc(diNode, dNode, BIG, 0#)                ' d_wd
            If p < mP - 1 Then nxt = baseDI + (p + 1) * mM + j Else nxt = nodeR
            AddArcLB diNode, nxt, mDemMin(k), mDemCap(k), 0#          ' d_end carry
            mB(dNode) = mB(dNode) - mDemand(k)                       ' demand sink
        Next j
    Next p

    ' ship arcs (lane rule: skip closed/blank lanes)
    For p = 0 To mP - 1
        For i = 0 To mN - 1
            For j = 0 To mM - 1
                k = (p * mN + i) * mM + j
                If Not mClosed(k) Then
                    lo = mShipMin(k): hi = mShipMax(k)
                    If Not (hi <= 0 And lo <= 0) Then
                        shLo(k) = lo: shInit(k) = maxd(0#, hi - lo)
                        shFwd(k) = AddArcLB(baseS + p * mN + i, baseD + p * mM + j, lo, hi, mTrans(k))
                    Else
                        shFwd(k) = -1
                    End If
                Else
                    shFwd(k) = -1
                End If
            Next j
        Next i
    Next p

    ' base supplies
    Dim sInitS As Double, sInitD As Double, sDem As Double
    For i = 0 To mN - 1: mB(baseSI + i) = mB(baseSI + i) + mInitS(i): sInitS = sInitS + mInitS(i): Next i
    For j = 0 To mM - 1: mB(baseDI + j) = mB(baseDI + j) + mInitD(j): sInitD = sInitD + mInitD(j): Next j
    For k = 0 To mPMc - 1: sDem = sDem + mDemand(k): Next k
    mB(nodeR) = mB(nodeR) + (sDem - sInitS - sInitD)

    ' super source / sink
    ReDim ttArc(0 To mNode): ReDim ttNode(0 To mNode): mTT = 0
    mNeed = 0#
    Dim v As Long
    For v = 0 To nodeR
        If mB(v) > EPS Then
            AddArc nodeSS, v, mB(v), 0#
            mNeed = mNeed + mB(v)
        ElseIf mB(v) < -EPS Then
            Dim fid As Long
            fid = AddArc(v, nodeTT, -mB(v), 0#)
            ttArc(mTT) = fid: ttNode(mTT) = v: mTT = mTT + 1
        End If
    Next v
End Sub

'=====================================================================
' MIN-COST FLOW  (successive shortest paths, Dijkstra + potentials)
'=====================================================================
Private Sub HPush(ByVal key As Double, ByVal node As Long)
    Dim i As Long, pp As Long, tk As Double, tn As Long
    i = hsz: hk(i) = key: hn(i) = node: hsz = hsz + 1
    Do While i > 0
        pp = (i - 1) \ 2
        If hk(pp) <= hk(i) Then Exit Do
        tk = hk(pp): hk(pp) = hk(i): hk(i) = tk
        tn = hn(pp): hn(pp) = hn(i): hn(i) = tn
        i = pp
    Loop
End Sub

Private Function HPop() As Long
    Dim i As Long, l As Long, r As Long, m As Long, tk As Double, tn As Long
    HPop = hn(0)
    hsz = hsz - 1
    If hsz > 0 Then
        hk(0) = hk(hsz): hn(0) = hn(hsz)
        i = 0
        Do
            l = 2 * i + 1: r = 2 * i + 2: m = i
            If l < hsz Then If hk(l) < hk(m) Then m = l
            If r < hsz Then If hk(r) < hk(m) Then m = r
            If m = i Then Exit Do
            tk = hk(m): hk(m) = hk(i): hk(i) = tk
            tn = hn(m): hn(m) = hn(i): hn(i) = tn
            i = m
        Loop
    End If
End Function

Private Function SolveMCF(ByRef feasible As Boolean, ByRef diag As String) As Double
    Const INF As Double = 1E+300
    Dim s As Long, t As Long
    s = nodeSS: t = nodeTT
    Dim h() As Double, dist() As Double, prevE() As Long
    Dim doneN() As Boolean
    ReDim h(0 To mNode - 1)
    ReDim dist(0 To mNode - 1)
    ReDim prevE(0 To mNode - 1)
    ReDim doneN(0 To mNode - 1)
    ReDim hn(0 To mArc + 16): ReDim hk(0 To mArc + 16)

    Dim totalCost As Double, totalFlow As Double
    Dim u As Long, v As Long, aid As Long, nd As Double, f As Double
    Dim i As Long, augN As Long
    Dbg "SolveMCF begin: s=" & s & " t=" & t

    Do
        For i = 0 To mNode - 1: dist(i) = INF: doneN(i) = False: Next i
        dist(s) = 0#: hsz = 0: HPush 0#, s
        Do While hsz > 0
            u = HPop()
            If Not doneN(u) Then
                doneN(u) = True
                aid = aHead(u)
                Do While aid >= 0
                    If aCap(aid) > EPS Then
                        v = aTo(aid)
                        nd = dist(u) + aCost(aid) + h(u) - h(v)
                        If nd < dist(v) - EPS Then
                            dist(v) = nd: prevE(v) = aid: HPush nd, v
                        End If
                    End If
                    aid = aNext(aid)
                Loop
            End If
        Loop
        If dist(t) >= INF / 2 Then Exit Do
        For i = 0 To mNode - 1
            If dist(i) < INF / 2 Then h(i) = h(i) + dist(i)
        Next i
        ' bottleneck
        f = INF
        v = t
        Do While v <> s
            aid = prevE(v)
            If aCap(aid) < f Then f = aCap(aid)
            v = aTo(aid Xor 1)
        Loop
        v = t
        Do While v <> s
            aid = prevE(v)
            aCap(aid) = aCap(aid) - f
            aCap(aid Xor 1) = aCap(aid Xor 1) + f
            totalCost = totalCost + f * aCost(aid)
            v = aTo(aid Xor 1)
        Loop
        totalFlow = totalFlow + f
        augN = augN + 1
        If augN Mod 200 = 0 Then Dbg "  augment " & augN & " flow=" & Format(totalFlow, "0")
    Loop
    Dbg "  SSP augments=" & augN & " flow=" & Format(totalFlow, "0")

    feasible = (Abs(totalFlow - mNeed) < 0.0001 * (1# + mNeed))
    If Not feasible Then diag = DiagnoseInfeasible(mNeed - totalFlow)
    SolveMCF = mOffset + totalCost
End Function

Private Function DiagnoseInfeasible(ByVal deficit As Double) As String
    Dim t As Long, v As Long, remCap As Double
    Dim dShort As Double, surplus As Double, invShort As Double
    For t = 0 To mTT - 1
        remCap = aCap(ttArc(t))
        If remCap > EPS Then
            v = ttNode(t)
            If v >= baseD And v < baseD + mPMc Then
                dShort = dShort + remCap
            ElseIf v = nodeR Then
                surplus = surplus + remCap
            Else
                invShort = invShort + remCap
            End If
        End If
    Next t
    ' Totals to make the cause unambiguous (and to prove the shortfall is real).
    Dim totDem As Double, totSupMax As Double, totMinLift As Double, totInit As Double, kk As Long
    For kk = 0 To mP * mM - 1: totDem = totDem + mDemand(kk): Next kk
    For kk = 0 To mP * mN - 1
        totSupMax = totSupMax + mContractFlow(kk) * mPercMax(kk)
        totMinLift = totMinLift + mContractFlow(kk) * mPercMin(kk)
    Next kk
    For kk = 0 To mN - 1: totInit = totInit + mInitS(kk): Next kk
    For kk = 0 To mM - 1: totInit = totInit + mInitD(kk): Next kk

    Dim msg As String
    msg = "Short by " & Format(deficit, "#,##0") & " kgal."
    If dShort > EPS Then msg = msg & " Demand unmet " & Format(dShort, "#,##0") & " kgal (not enough open lanes/supply)."
    If surplus > EPS Then msg = msg & " Forced minimum lifts exceed demand by " & Format(surplus, "#,##0") & " kgal with no storage to absorb them (lower the % minimums or open storage)."
    If invShort > EPS Then msg = msg & " Inventory balance infeasible " & Format(invShort, "#,##0") & " kgal."
    msg = msg & " [total demand=" & Format(totDem, "#,##0") & ", max available supply=" & Format(totSupMax + totInit, "#,##0") & _
          " (purchases " & Format(totSupMax, "#,##0") & " + start inv " & Format(totInit, "#,##0") & "), forced min lifts=" & Format(totMinLift, "#,##0") & "]"
    DiagnoseInfeasible = msg
End Function

'=====================================================================
' RECOVER + WRITE SOLUTION
'=====================================================================
Private Sub WriteSolution()
    Dim p As Long, i As Long, j As Long, k As Long
    Dim shipArr() As Double, percArr() As Double
    Dim swdArr() As Double, sinjArr() As Double
    Dim dwdArr() As Double, dinjArr() As Double
    ReDim shipArr(1 To mPN, 1 To mM)
    ReDim percArr(1 To mN, 1 To mP)
    ReDim swdArr(1 To mN, 1 To mP)
    ReDim sinjArr(1 To mN, 1 To mP)
    ReDim dwdArr(1 To mM, 1 To mP)
    ReDim dinjArr(1 To mM, 1 To mP)

    For p = 0 To mP - 1
        For i = 0 To mN - 1
            k = p * mN + i
            Dim x As Double, cf As Double, pc As Double
            If pFwd(k) >= 0 Then
                x = pLo(k) + (pInit(k) - aCap(pFwd(k)))
            Else
                x = pLo(k)
            End If
            cf = mContractFlow(k)
            ' Write the exact flow-implied percentage: contractFlow*perc = x keeps the
            ' supply balance exact. The arc bounds already force x in [cf*percMin, cf*percMax],
            ' so perc lands in [percMin, percMax] (no snapping, which would break balances).
            If cf > EPS Then
                pc = x / cf
            Else
                pc = mPercMin(k)
            End If
            percArr(i + 1, p + 1) = pc
            Dim si As Double, sw As Double
            si = BIG - aCap(siFwd(k)): sw = BIG - aCap(swFwd(k))
            If si < 0.0000001 Then si = 0#
            If sw < 0.0000001 Then sw = 0#
            sinjArr(i + 1, p + 1) = si
            swdArr(i + 1, p + 1) = sw
        Next i
    Next p

    For p = 0 To mP - 1
        For j = 0 To mM - 1
            k = p * mM + j
            Dim di As Double, dw As Double
            di = BIG - aCap(diFwd(k)): dw = BIG - aCap(dwFwd(k))
            If di < 0.0000001 Then di = 0#
            If dw < 0.0000001 Then dw = 0#
            dinjArr(j + 1, p + 1) = di
            dwdArr(j + 1, p + 1) = dw
        Next j
    Next p

    For p = 0 To mP - 1
        For i = 0 To mN - 1
            For j = 0 To mM - 1
                k = (p * mN + i) * mM + j
                Dim sh As Double
                If shFwd(k) >= 0 Then
                    sh = shLo(k) + (shInit(k) - aCap(shFwd(k)))   ' exact flow, already in [shipMin, shipMax]
                Else
                    sh = 0#
                End If
                If sh < 0.0000001 Then sh = 0#
                shipArr((p * mN + i) + 1, j + 1) = sh
            Next j
        Next i
    Next p

    PutBlockD mShipA1, shipArr
    PutBlockD mPercA1, percArr
    PutBlockD mSWdA1, swdArr
    PutBlockD mSInjA1, sinjArr
    PutBlockD mDWdA1, dwdArr
    PutBlockD mDInjA1, dinjArr
End Sub

'=====================================================================
' VERIFY against Excel's own cells
'=====================================================================
Private Function VerifySolution(ByVal obj As Double, ByRef vmsg As String) As Boolean
    VerifySolution = False
    vmsg = ""

    Dim num As Long
    num = CLng(Val(StripEq(NameRefersTo("solver_num"))))
    If num < 10 Then num = 10

    Dim k As Long, lhs As String, rel As Long, rhsRT As String, rhsA1 As String
    Dim isRange As Boolean, rconst As Double
    Dim lhsV As Variant, rhsV As Variant
    Dim r1 As Long, c1 As Long, r2 As Long, c2 As Long, rows As Long, cols As Long
    Dim rr As Long, cc As Long, L As Double, Rv As Double, tol As Double

    For k = 1 To num
        lhs = CellPart(StripEq(NameRefersTo("solver_lhs" & k)))
        If Len(lhs) = 0 Then GoTo nextk
        rel = CLng(Val(StripEq(NameRefersTo("solver_rel" & k))))
        rhsRT = NameRefersTo("solver_rhs" & k)
        isRange = (InStr(rhsRT, "!") > 0)
        lhsV = mWs.Range(lhs).Value2
        If isRange Then
            rhsA1 = CellPart(StripEq(rhsRT))
            rhsV = mWs.Range(rhsA1).Value2
        Else
            rconst = Val(StripEq(rhsRT))
        End If
        ParseA1 lhs, r1, c1, r2, c2
        rows = r2 - r1 + 1: cols = c2 - c1 + 1
        For rr = 1 To rows
            For cc = 1 To cols
                L = NumRC(lhsV, rr, cc)
                If isRange Then Rv = NumRC(rhsV, rr, cc) Else Rv = rconst
                tol = 0.01 + 0.000001 * (Abs(L) + Abs(Rv))
                Dim viol As Double
                Select Case rel
                    Case 1: viol = L - Rv
                    Case 2: viol = Abs(L - Rv)
                    Case 3: viol = Rv - L
                    Case Else: viol = 0#
                End Select
                If viol > tol Then
                    Dim fml As String
                    fml = ""
                    On Error Resume Next
                    fml = mWs.Range(lhs).Cells(rr, cc).Formula
                    On Error GoTo 0
                    vmsg = "c" & k & " rel=" & rel & " at(" & rr & "," & cc & ") L=" & Format(L, "0.00") & " R=" & Format(Rv, "0.00") & " viol=" & Format(viol, "0.000") & " fml=" & Left$(fml, 180)
                    Dbg "VIOL " & vmsg
                    Exit Function
                End If
            Next cc
        Next rr
nextk:
    Next k

    ' --- objective-cell consistency (only enforce if D4 is a live formula) ---
    Dim objVal As Double, hasF As Boolean
    hasF = mWs.Range(mObjA1).HasFormula
    objVal = CDbl(mWs.Range(mObjA1).Value2)
    If hasF Then
        If Abs(objVal - obj) > 1# + 0.000001 * Abs(obj) Then
            vmsg = "objective cell " & Format(objVal, "#,##0.00") & " <> solver " & Format(obj, "#,##0.00")
            Exit Function
        End If
    Else
        vmsg = "NOTE: objective cell is hard-coded (" & Format(objVal, "#,##0.00") & "), not a formula; validated by constraints only"
    End If

    VerifySolution = True
End Function
