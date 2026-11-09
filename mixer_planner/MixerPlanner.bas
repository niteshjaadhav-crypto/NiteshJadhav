Attribute VB_Name = "MixerPlanner"
'==============================================================================
'  MIXER COMPOUND PLANNER  -  automation for Mixing_Production_Plan.xlsm
'
'  Daily flow (buttons on the "Control Panel" sheet):
'    0. Setup_Planner          - run ONCE after importing this module
'    1. Paste SAP LX02 data    - (manual, as today) STOCK updates itself
'    2. Build_Mixer_Plan       - Priority -> mixer allocation -> MX01/MX02/MX03
'                                1440-min plan with line clearance + shifts
'    3. Release POs in SAP, type them in MX01 PO / MX02 PO / MX03 PO and the
'       clean-out sheets (manual, as today)
'    4. Fill_PO_Numbers        - pulls 21/22/29 PO numbers + iterations into
'                                MX01/MX02/MX03
'    5. Build_Tote_Requirement - fills "Tote Requirment" from the plan
'    6. Close_Day              - remembers last iteration used per PO
'
'  Weekly (Thursday): type the planning-team image into STOCK columns D, H, K
'  (K = CONT or the full weekly batch number) and set the week start date in
'  "Planner Settings".  Released POs are deducted automatically.
'==============================================================================
Option Explicit

'---- sheet names (exactly as in the workbook) ---------------------------------
Private Const SH_STOCK As String = "STOCK"
Private Const SH_RAW As String = "Raw Data"
Private Const SH_RECIPE As String = "Recipe Details"
Private Const SH_SET As String = "Planner Settings"
Private Const SH_PRI As String = "Priority"
Private Const SH_COMAP As String = "CO Map"
Private Const SH_ITER As String = "Iteration Log"
Private Const SH_TOTE As String = "Tote Requirment"
Private Const SH_PANEL As String = "Control Panel"

'---- layout constants ---------------------------------------------------------
Private Const PRI_FIRST As Long = 2          ' Priority row r mirrors STOCK row r
Private Const PRI_LAST As Long = 61
Private Const PLAN_FIRST As Long = 4         ' first data row on MX01/MX02/MX03
Private Const PLAN_LAST As Long = 140        ' rows cleared on every rebuild
Private Const HELP_COL As String = "BA"      ' hidden helper columns BA:BE on MX sheets
Private Const TOTE_FIRST As Long = 3
Private Const TOTE_LAST As Long = 32

'---- Planner Settings rows (column B holds the value) -------------------------
Private Const S_PLANDATE As Long = 3
Private Const S_WEEKSTART As Long = 4
Private Const S_CRIT As Long = 5
Private Const S_WATCH As Long = 6
Private Const S_CONTDAYS As Long = 7
Private Const S_INCLWATCH As Long = 8
Private Const S_AUTODEDUCT As Long = 9
Private Const S_DAYMIN As Long = 10
Private Const S_LCMIN As Long = 11
Private Const S_COBEFORE As Long = 12
Private Const S_COAFTER As Long = 13
Private Const S_SHIFTMIN As Long = 14
Private Const S_CYC1 As Long = 15            ' 16 = MX02, 17 = MX03
Private Const S_DEFCO As Long = 18
Private Const S_MAXCOMP As Long = 19
Private Const S_LAST1 As Long = 20           ' 21 = MX02, 22 = MX03

Private Type PItem
    priRow As Long
    code22 As String
    code21 As String
    mat As String
    cover As Double
    req As Double
    exB As Long
    pallet As Double
    wt(1 To 3) As Double
    override As Long
    mixer As Long
    mxB As Long
    remark As String
    done As Boolean
End Type

Private Type PStep
    kind As String          ' "LC" or "PROD"
    item As Long            ' index into items (PROD) / next item (LC)
    mxB As Long
End Type

'==============================================================================
'  0. SETUP  (safe to run again - keeps your settings values)
'==============================================================================
Public Sub Setup_Planner()
    Application.ScreenUpdating = False
    BuildSettingsSheet
    BuildPrioritySheet
    BuildCOMapSheet
    BuildIterationSheet
    BuildControlPanel
    Application.ScreenUpdating = True
    MsgBox "Planner sheets are ready." & vbLf & vbLf & _
           "Please check 'Planner Settings' (yellow cells) - especially the" & vbLf & _
           "mixer cycle time per batch and the week start date.", vbInformation
End Sub

Private Function GetOrAddSheet(nm As String, Optional afterName As String = "") As Worksheet
    On Error Resume Next
    Set GetOrAddSheet = ThisWorkbook.Worksheets(nm)
    On Error GoTo 0
    If GetOrAddSheet Is Nothing Then
        If afterName <> "" Then
            Set GetOrAddSheet = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(afterName))
        Else
            Set GetOrAddSheet = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        End If
        GetOrAddSheet.Name = nm
    End If
End Function

Private Sub SetRow(ws As Worksheet, r As Long, lbl As String, defVal As Variant, note As String)
    ws.Cells(r, 1).Value = lbl
    If IsEmpty(ws.Cells(r, 2).Value) Then ws.Cells(r, 2).Value = defVal
    ws.Cells(r, 2).Interior.Color = RGB(255, 255, 0)
    ws.Cells(r, 2).Font.Color = RGB(0, 0, 255)
    ws.Cells(r, 3).Value = note
End Sub

Private Sub BuildSettingsSheet()
    Dim ws As Worksheet
    Set ws = GetOrAddSheet(SH_SET)
    ws.Range("A1").Value = "PLANNER SETTINGS  (edit yellow cells only)"
    ws.Range("A1").Font.Bold = True: ws.Range("A1").Font.Size = 14
    ws.Range("A2:C2").Value = Array("Setting", "Value", "Note")
    ws.Range("A2:C2").Font.Bold = True
    Dim thu As Date
    thu = Date - ((Weekday(Date, vbMonday) - 4 + 7) Mod 7)
    SetRow ws, S_PLANDATE, "Plan date", Date, "Date written in column A of MX01/MX02/MX03"
    SetRow ws, S_WEEKSTART, "Week start date (Thursday plan received)", thu, "POs released on/after this date are deducted from STOCK limited qty"
    SetRow ws, S_CRIT, "Critical cover (days)", 1.5, "Stock cover below this = press will stop -> MUST plan"
    SetRow ws, S_WATCH, "Watch cover (days)", 2.5, "Below this = plan if mixer has spare time"
    SetRow ws, S_CONTDAYS, "CONT compounds: days of requirement to produce", 3.5, "Your rule: 3 to 4 days once planned"
    SetRow ws, S_INCLWATCH, "Fill spare capacity with WATCH items (YES/NO)", "YES", ""
    SetRow ws, S_AUTODEDUCT, "Auto-deduct released POs from limited qty (YES/NO)", "YES", "YES = type the FULL weekly number in STOCK col K, do not reduce it by hand"
    SetRow ws, S_DAYMIN, "Mixer minutes per day", 1440, "3 shifts x 480"
    SetRow ws, S_LCMIN, "Line clearance total minutes", 80, "Includes all clean-out batches"
    SetRow ws, S_COBEFORE, "Clean-out batches before physical cleaning", 3, ""
    SetRow ws, S_COAFTER, "Clean-out batches after physical cleaning", 1, ""
    SetRow ws, S_SHIFTMIN, "Shift length (minutes)", 480, "Shift A starts at minute 0, then B, then C"
    SetRow ws, S_CYC1, "MX01 cycle time (minutes per batch)", 6.5, "ASSUMPTION - set your real mixing cycle"
    SetRow ws, S_CYC1 + 1, "MX02 cycle time (minutes per batch)", 6.5, "ASSUMPTION - set your real mixing cycle"
    SetRow ws, S_CYC1 + 2, "MX03 cycle time (minutes per batch)", 6.5, "ASSUMPTION - set your real mixing cycle"
    SetRow ws, S_DEFCO, "Default clean-out 29 code", "29000024", "Used when the next compound is not in 'CO Map'"
    SetRow ws, S_MAXCOMP, "Max different compounds (21 codes) per mixer per day", 4, "Limits number of line clearances"
    SetRow ws, S_LAST1, "MX01 compound (21 code) running at end of yesterday", "", "If today's first compound is the same, no LC at start"
    SetRow ws, S_LAST1 + 1, "MX02 compound (21 code) running at end of yesterday", "", ""
    SetRow ws, S_LAST1 + 2, "MX03 compound (21 code) running at end of yesterday", "", ""
    ws.Range("B" & S_PLANDATE & ":B" & S_WEEKSTART).NumberFormat = "dd-mmm-yyyy"
    ws.Range("B" & S_DEFCO).NumberFormat = "@"
    ws.Range("B" & S_LAST1 & ":B" & S_LAST1 + 2).NumberFormat = "@"
    ws.Columns("A").ColumnWidth = 52
    ws.Columns("B").ColumnWidth = 16
    ws.Columns("C").ColumnWidth = 70
    ws.Cells.Font.Name = "Arial"
End Sub

Private Sub BuildPrioritySheet()
    Dim ws As Worksheet, r As Long, s As String, sp As String
    Set ws = GetOrAddSheet(SH_PRI, SH_STOCK)
    sp = "'" & SH_SET & "'!$B$"
    ws.Range("A1:V1").Value = Array("Rank", "22 Code", "Material", "21 Code", "Compound", "Colour", _
        "Stock (EX batches)", "EX batches req / day", "Days cover", "Weekly limit (STOCK K)", _
        "Released this week", "Limit left", "Status", "Plan EX batches", "Pallet kg", _
        "MX01 batch wt", "MX02 batch wt", "MX03 batch wt", "Mixer override (type MX01/MX02/MX03)", _
        "Assigned mixer", "Mixer batches", "Remark")
    For r = PRI_FIRST To PRI_LAST
        s = CStr(r)
        ws.Range("B" & s).Formula = "=IF(TRIM(STOCK!D" & s & "&"""")="""","""",STOCK!D" & s & ")"
        ws.Range("C" & s).Formula = "=IF(B" & s & "="""","""",STOCK!E" & s & ")"
        ws.Range("D" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$C$2:$C$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("E" & s).Formula = "=IF(B" & s & "="""","""",STOCK!F" & s & ")"
        ws.Range("F" & s).Formula = "=IF(B" & s & "="""","""",STOCK!L" & s & ")"
        ws.Range("G" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(N(STOCK!G" & s & "),0))"
        ws.Range("H" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(N(STOCK!H" & s & "),0))"
        ws.Range("I" & s).Formula = "=IF(B" & s & "="""","""",IF(H" & s & ">0,G" & s & "/H" & s & ",99))"
        ws.Range("J" & s).Formula = "=IF(B" & s & "="""","""",IF(TRIM(STOCK!K" & s & "&"""")="""",""CONT"",STOCK!K" & s & "))"
        ws.Range("K" & s).Formula = "=IF(B" & s & "="""","""",SUMIFS('MX01 PO'!$E:$E,'MX01 PO'!$B:$B,B" & s & ",'MX01 PO'!$A:$A,"">=""&" & sp & S_WEEKSTART & ")" & _
            "+SUMIFS('MX02 PO '!$E:$E,'MX02 PO '!$B:$B,B" & s & ",'MX02 PO '!$A:$A,"">=""&" & sp & S_WEEKSTART & ")" & _
            "+SUMIFS('MX03 PO'!$E:$E,'MX03 PO'!$B:$B,B" & s & ",'MX03 PO'!$A:$A,"">=""&" & sp & S_WEEKSTART & "))"
        ws.Range("L" & s).Formula = "=IF(B" & s & "="""","""",IF(UPPER(TRIM(J" & s & "&""""))=""CONT"",""CONT"",MAX(0,IFERROR(VALUE(J" & s & "&""""),0)-IF(" & sp & S_AUTODEDUCT & "=""YES"",K" & s & ",0))))"
        ws.Range("M" & s).Formula = "=IF(B" & s & "="""","""",IF(OR(O" & s & "="""",P" & s & "&Q" & s & "&R" & s & "=""""),""CHECK DATA"",IF(H" & s & "<=0,""NO DEMAND""," & _
            "IF(AND(L" & s & "<>""CONT"",N(L" & s & ")<=0),""LIMIT DONE"",IF(I" & s & "<" & sp & S_CRIT & ",""1-CRITICAL"",IF(I" & s & "<" & sp & S_WATCH & ",""2-WATCH"",""3-OK""))))))"
        ws.Range("N" & s).Formula = "=IF(B" & s & "="""","""",IF(OR(M" & s & "=""1-CRITICAL"",AND(M" & s & "=""2-WATCH""," & sp & S_INCLWATCH & "=""YES""))," & _
            "IF(L" & s & "=""CONT"",MAX(1,ROUNDUP(" & sp & S_CONTDAYS & "*H" & s & ",0)),L" & s & "),0))"
        ws.Range("O" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(1/(1/INDEX('Raw Data'!$H$2:$H$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0))),""""))"
        ws.Range("P" & s).Formula = WtFormula(s, "B", "G", "H")
        ws.Range("Q" & s).Formula = WtFormula(s, "L", "Q", "Q")
        ws.Range("R" & s).Formula = WtFormula(s, "U", "Z", "Z")
        ws.Range("A" & s).Formula = "=IF(N(N" & s & ")>0,COUNTIFS($N$2:$N$61,"">0"",$I$2:$I$61,""<""&I" & s & ")" & _
            "+COUNTIFS($N$2:$N$61,"">0"",$I$2:$I$61,I" & s & ",$H$2:$H$61,"">""&H" & s & ")+1,"""")"
    Next r
    ws.Range("I2:I61").NumberFormat = "0.00"
    ws.Range("O2:R61").NumberFormat = "0.0"
    ws.Range("S2:S61").Interior.Color = RGB(255, 255, 0)
    With ws.Range("A1:V1")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120): .WrapText = True
    End With
    ws.Rows(1).RowHeight = 45
    ws.Cells.Font.Name = "Arial"
    ws.Columns("A:V").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 22: ws.Columns("V").ColumnWidth = 40
    ' status colours
    Dim rng As Range
    Set rng = ws.Range("M2:M61")
    rng.FormatConditions.Delete
    AddCF rng, "1-CRITICAL", RGB(248, 203, 173)
    AddCF rng, "2-WATCH", RGB(255, 230, 153)
    AddCF rng, "3-OK", RGB(198, 239, 206)
    AddCF rng, "CHECK DATA", RGB(189, 215, 238)
    AddCF rng, "LIMIT DONE", RGB(217, 217, 217)
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = False
    ActiveWindow.FreezePanes = True
    If ws.AutoFilterMode = False Then ws.Range("A1:V61").AutoFilter
End Sub

Private Function WtFormula(s As String, codeCol As String, wCol As String, w2Col As String) As String
    Dim m As String, rd As String
    rd = "'Recipe Details'!"
    m = "MATCH(D" & s & "," & rd & "$" & codeCol & "$3:$" & codeCol & "$80,0)"
    WtFormula = "=IF(D" & s & "="""","""",IFERROR(1/(1/INDEX(" & rd & "$" & wCol & "$3:$" & wCol & "$80," & m & "))," & _
                "IFERROR(1/(1/INDEX(" & rd & "$" & w2Col & "$3:$" & w2Col & "$80," & m & ")),"""")))"
End Function

Private Sub AddCF(rng As Range, txt As String, col As Long)
    Dim fc As FormatCondition
    Set fc = rng.FormatConditions.Add(Type:=xlCellValue, Operator:=xlEqual, Formula1:="=""" & txt & """")
    fc.Interior.Color = col
End Sub

Private Sub BuildCOMapSheet()
    Dim ws As Worksheet, rec As Worksheet, r As Long, c As Variant, k As String
    Dim nxt As Long, dict As Object
    Set ws = GetOrAddSheet(SH_COMAP, SH_SET)
    ws.Range("A1:C1").Value = Array("21 Code (next compound)", "Compound", "Clean-out 29 code to use before it")
    ws.Range("A1:C1").Font.Bold = True
    Set dict = CreateObject("Scripting.Dictionary")
    nxt = 2
    Do While Trim(CStr(ws.Cells(nxt, 1).Value)) <> ""
        dict(CodeStr(ws.Cells(nxt, 1).Value)) = True
        nxt = nxt + 1
    Loop
    Set rec = ThisWorkbook.Worksheets(SH_RECIPE)
    For Each c In Array(2, 12, 21)
        For r = 3 To 80
            k = CodeStr(rec.Cells(r, c).Value)
            If Left(k, 2) = "21" And Not dict.Exists(k) Then
                dict(k) = True
                ws.Cells(nxt, 1).NumberFormat = "@"
                ws.Cells(nxt, 1).Value = k
                ws.Cells(nxt, 2).Value = rec.Cells(r, c + 1).Value
                ws.Cells(nxt, 3).NumberFormat = "@"
                ws.Cells(nxt, 3).Value = Cfg(S_DEFCO)
                ws.Cells(nxt, 3).Interior.Color = RGB(255, 255, 0)
                nxt = nxt + 1
            End If
        Next r
    Next c
    ws.Columns("A:C").ColumnWidth = 26
    ws.Cells.Font.Name = "Arial"
End Sub

Private Sub BuildIterationSheet()
    Dim ws As Worksheet
    Set ws = GetOrAddSheet(SH_ITER, SH_COMAP)
    ws.Range("A1:D1").Value = Array("PO number", "Last iteration used", "Updated on", "Source sheet")
    ws.Range("A1:D1").Font.Bold = True
    ws.Columns("A:D").ColumnWidth = 20
    ws.Cells.Font.Name = "Arial"
End Sub

Private Sub BuildControlPanel()
    Dim ws As Worksheet, b As Object, i As Long
    Dim caps As Variant, macs As Variant, notes As Variant
    Set ws = GetOrAddSheet(SH_PANEL)
    ws.Cells.Clear
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    ws.Range("B2").Value = "MIXER COMPOUND PLANNER - daily steps"
    ws.Range("B2").Font.Size = 16: ws.Range("B2").Font.Bold = True
    caps = Array("1. Build Mixer Plan", "2. Fill PO Numbers", "3. Build Tote Requirement", "4. Close Day (log iterations)", "Re-run Setup")
    macs = Array("Build_Mixer_Plan", "Fill_PO_Numbers", "Build_Tote_Requirement", "Close_Day", "Setup_Planner")
    notes = Array("After pasting LX02 (and on Thursday after updating STOCK H/K from the weekly plan). Writes MX01/MX02/MX03.", _
                  "After releasing POs in SAP and typing them in MX01 PO / MX02 PO / MX03 PO and the clean-out sheets.", _
                  "Fills 'Tote Requirment' from today's plan.", _
                  "At end of the day - stores the last iteration of every PO so tomorrow continues from there.", _
                  "Only when sheets are missing. Keeps your settings.")
    For i = 0 To 4
        Set b = ws.Buttons.Add(ws.Range("B" & 4 + i * 3).Left, ws.Range("B" & 4 + i * 3).Top, 220, 30)
        b.Caption = caps(i)
        b.OnAction = macs(i)
        ws.Range("D" & 4 + i * 3).Value = notes(i)
    Next i
    ws.Range("B20").Value = "Weekly (Thursday): type the planning image into STOCK - column D (22 code), H (Extrusion batch/day), K (CONT or FULL weekly batches)."
    ws.Range("B21").Value = "Then set 'Week start date' in Planner Settings. Do not reduce column K by hand - released POs are deducted automatically."
    ws.Range("B22").Value = "Priority sheet: do NOT sort it (rows mirror STOCK). Use the filter on column A (Rank) or M (Status)."
    ws.Columns("B").ColumnWidth = 34
    ws.Cells.Font.Name = "Arial"
End Sub

'==============================================================================
'  1. BUILD MIXER PLAN
'==============================================================================
Public Sub Build_Mixer_Plan()
    Dim wsP As Worksheet, r As Long, n As Long, i As Long, j As Long, m As Long
    Dim it() As PItem, tmp As PItem
    Dim cyc(1 To 3) As Double, load(1 To 3) As Double, ncomp(1 To 3) As Long
    Dim lastRun(1 To 3) As String
    Dim steps(1 To 3, 1 To 200) As PStep, nSteps(1 To 3) As Long
    Dim dayMin As Double, lcMin As Double, maxComp As Long

    If Not SheetExists(SH_PRI) Or Not SheetExists(SH_SET) Then
        MsgBox "Run Setup_Planner first.", vbExclamation: Exit Sub
    End If
    Application.ScreenUpdating = False
    Application.Calculate
    Set wsP = ThisWorkbook.Worksheets(SH_PRI)

    dayMin = GetNum(S_DAYMIN): lcMin = GetNum(S_LCMIN): maxComp = CLng(GetNum(S_MAXCOMP))
    For m = 1 To 3
        cyc(m) = GetNum(S_CYC1 + m - 1)
        If cyc(m) <= 0 Then MsgBox "Set the MX0" & m & " cycle time in Planner Settings.", vbExclamation: GoTo done
        lastRun(m) = CodeStr(Cfg(S_LAST1 + m - 1))
        If lastRun(m) = "" Then lastRun(m) = LastRunning21(m)   ' auto: last compound of yesterday's plan
    Next m

    '--- read items that need planning ----------------------------------------
    ReDim it(1 To PRI_LAST)
    n = 0
    wsP.Range("T2:V" & PRI_LAST).ClearContents
    For r = PRI_FIRST To PRI_LAST
        If IsNumeric(wsP.Cells(r, "N").Value) And CodeStr(wsP.Cells(r, "B").Value) <> "" Then
            If CDbl(wsP.Cells(r, "N").Value) > 0 Then
                n = n + 1
                With it(n)
                    .priRow = r
                    .code22 = CodeStr(wsP.Cells(r, "B").Value)
                    .code21 = CodeStr(wsP.Cells(r, "D").Value)
                    .mat = CStr(wsP.Cells(r, "C").Value)
                    .cover = NumOf(wsP.Cells(r, "I").Value)
                    .req = NumOf(wsP.Cells(r, "H").Value)
                    .exB = CLng(wsP.Cells(r, "N").Value)
                    .pallet = NumOf(wsP.Cells(r, "O").Value)
                    .wt(1) = NumOf(wsP.Cells(r, "P").Value)
                    .wt(2) = NumOf(wsP.Cells(r, "Q").Value)
                    .wt(3) = NumOf(wsP.Cells(r, "R").Value)
                    .override = MixerNo(wsP.Cells(r, "S").Value)
                End With
            End If
        End If
    Next r
    If n = 0 Then MsgBox "Nothing to plan - no CRITICAL / WATCH compounds.", vbInformation: GoTo done

    '--- sort: lowest cover first, then highest consumption -------------------
    For i = 1 To n - 1
        For j = i + 1 To n
            If it(j).cover < it(i).cover Or (it(j).cover = it(i).cover And it(j).req > it(i).req) Then
                tmp = it(i): it(i) = it(j): it(j) = tmp
            End If
        Next j
    Next i

    '--- allocate groups (same 21 code = no LC between them) -------------------
    Dim g As Long, best As Long, bestKey As Double, needMin As Double, lc As Double
    Dim ok As Boolean, key As Double, avail As Double, mxB As Long, fitMx As Long, exFit As Long
    Dim ovr As Long
    For g = 1 To n
        If Not it(g).done Then
            ' override: first item in the group that has one
            ovr = 0
            For j = g To n
                If it(j).code21 = it(g).code21 And it(j).override > 0 Then ovr = it(j).override: Exit For
            Next j
            best = 0: bestKey = 1E+30
            For m = 1 To 3
                ok = (ovr = 0 Or ovr = m) And ncomp(m) < maxComp
                If ok Then
                    needMin = 0
                    For j = g To n
                        If it(j).code21 = it(g).code21 Then
                            If it(j).wt(m) <= 0 Or it(j).pallet <= 0 Then ok = False: Exit For
                            needMin = needMin + MixBatches(it(j).exB, it(j).pallet, it(j).wt(m)) * cyc(m)
                        End If
                    Next j
                End If
                If ok Then
                    lc = LCNeeded(nSteps(m), it(g).code21, lastRun(m), lcMin)
                    key = load(m) + lc + needMin
                    If key > dayMin Then key = key + 1000000#   ' prefer a mixer where it fits fully
                    If key < bestKey Then bestKey = key: best = m
                End If
            Next m

            If best = 0 Then
                For j = g To n
                    If it(j).code21 = it(g).code21 Then
                        it(j).done = True
                        it(j).remark = "Not planned - no eligible mixer (check Recipe Details weights) or compound limit reached"
                    End If
                Next j
            Else
                m = best
                lc = LCNeeded(nSteps(m), it(g).code21, lastRun(m), lcMin)
                If dayMin - load(m) - lc < cyc(m) Then
                    For j = g To n
                        If it(j).code21 = it(g).code21 Then it(j).done = True: it(j).remark = "Not planned - mixer capacity full"
                    Next j
                Else
                    If lc > 0 Then
                        nSteps(m) = nSteps(m) + 1
                        steps(m, nSteps(m)).kind = "LC"
                        steps(m, nSteps(m)).item = g
                    End If
                    load(m) = load(m) + lc
                    ncomp(m) = ncomp(m) + 1
                    For j = g To n
                        If it(j).code21 = it(g).code21 Then
                            it(j).done = True
                            mxB = MixBatches(it(j).exB, it(j).pallet, it(j).wt(m))
                            avail = dayMin - load(m)
                            If mxB * cyc(m) > avail Then
                                fitMx = Int(avail / cyc(m))
                                exFit = Int(fitMx * it(j).wt(m) / it(j).pallet)
                                If exFit < 1 Then
                                    it(j).remark = "Not planned - mixer capacity full"
                                    mxB = 0
                                Else
                                    it(j).remark = "Partial: " & exFit & " of " & it(j).exB & " EX batches (capacity)"
                                    it(j).exB = exFit
                                    mxB = MixBatches(exFit, it(j).pallet, it(j).wt(m))
                                End If
                            End If
                            If mxB > 0 Then
                                it(j).mixer = m
                                it(j).mxB = mxB
                                load(m) = load(m) + mxB * cyc(m)
                                nSteps(m) = nSteps(m) + 1
                                steps(m, nSteps(m)).kind = "PROD"
                                steps(m, nSteps(m)).item = j
                                steps(m, nSteps(m)).mxB = mxB
                            End If
                        End If
                    Next j
                End If
            End If
        End If
    Next g

    '--- write Priority results ------------------------------------------------
    For i = 1 To n
        With it(i)
            If .mixer > 0 Then
                wsP.Cells(.priRow, "T").Value = "MX0" & .mixer
                wsP.Cells(.priRow, "U").Value = .mxB
            End If
            wsP.Cells(.priRow, "V").Value = .remark
        End With
    Next i

    '--- write the three mixer sheets ------------------------------------------
    Dim summary As String
    For m = 1 To 3
        WriteMixerSheet m, it, steps, nSteps(m), cyc(m)
        summary = summary & "MX0" & m & ": " & Format(load(m), "0") & " / " & Format(dayMin, "0") & " min, " & ncomp(m) & " compound(s)" & vbLf
    Next m
    Application.ScreenUpdating = True
    MsgBox "Mixer plan built." & vbLf & vbLf & summary & vbLf & _
           "Next: release POs in SAP, enter them in the PO sheets, then run 'Fill PO Numbers'.", vbInformation
done:
    Application.ScreenUpdating = True
End Sub

Private Function LCNeeded(stepsSoFar As Long, code21 As String, lastRun As String, lcMin As Double) As Double
    If stepsSoFar = 0 And code21 <> "" And code21 = lastRun Then
        LCNeeded = 0
    Else
        LCNeeded = lcMin
    End If
End Function

Private Function MixBatches(exB As Long, pallet As Double, wt As Double) As Long
    ' mixer batches = ROUNDUP(extrusion batches x pallet kg / mixer batch weight)
    Dim x As Double
    x = exB * pallet / wt
    MixBatches = Int(x)
    If x - MixBatches > 0.000001 Then MixBatches = MixBatches + 1
End Function

Private Sub WriteMixerSheet(m As Long, it() As PItem, steps() As PStep, ns As Long, cyc As Double)
    ' steps is (mixer, step) - this sub only reads row m
    Dim ws As Worksheet, r As Long, t As Double, sr As Long, s As Long, k As Long
    Dim shiftLen As Double, coBefore As Long, coAfter As Long, phys As Double
    Dim planDate As Variant, co As String, idx As Long
    Dim n As Long, rel As Long, leftB As Long, take As Long, fit As Long, sIdx As Long
    Dim exPrev As Long, exTo As Long, toB As Long

    Set ws = ThisWorkbook.Worksheets("MX0" & m)
    ' log yesterday's iterations automatically before clearing an older plan
    If IsDate(ws.Range("A" & PLAN_FIRST).Value) Then
        If CDate(ws.Range("A" & PLAN_FIRST).Value) < CDate(Cfg(S_PLANDATE)) Then LogIterations ws
    End If
    ClearPlan ws

    shiftLen = GetNum(S_SHIFTMIN)
    coBefore = CLng(GetNum(S_COBEFORE)): coAfter = CLng(GetNum(S_COAFTER))
    phys = GetNum(S_LCMIN) - (coBefore + coAfter) * cyc
    If phys < 0 Then phys = 0
    planDate = Cfg(S_PLANDATE)
    r = PLAN_FIRST: t = 0: sr = 1

    For s = 1 To ns
        If steps(m, s).kind = "LC" Then
            co = COFor(it(steps(m, s).item).code21)
            For k = 1 To coBefore
                WriteRow ws, r, planDate, ShiftOf(t, shiftLen), sr, co, 0, 0, 0, 0, True
                r = r + 1: sr = sr + 1: t = t + cyc
            Next k
            WriteRow ws, r, planDate, ShiftOf(t, shiftLen), sr, "Line Clearance", 0, 0, 0, 0, False
            r = r + 1: sr = sr + 1: t = t + phys
            For k = 1 To coAfter
                WriteRow ws, r, planDate, ShiftOf(t, shiftLen), sr, co, 0, 0, 0, 0, True
                r = r + 1: sr = sr + 1: t = t + cyc
            Next k
        Else
            idx = steps(m, s).item
            n = steps(m, s).mxB
            rel = 1: leftB = n: exPrev = 0
            Do While leftB > 0
                sIdx = Int(t / shiftLen + 0.000001)
                If sIdx >= 2 Then
                    take = leftB
                Else
                    fit = Int(((sIdx + 1) * shiftLen - t) / cyc + 0.000001)
                    If fit <= 0 Then
                        t = (sIdx + 1) * shiftLen
                        take = 0
                    Else
                        take = IIf(fit < leftB, fit, leftB)
                    End If
                End If
                If take > 0 Then
                    toB = rel + take - 1
                    If toB = n Then exTo = it(idx).exB Else exTo = Int(toB * it(idx).exB / n)
                    If exTo > exPrev Then
                        WriteRow ws, r, planDate, ShiftOf(t, shiftLen), sr, it(idx).code22, rel, toB, exPrev + 1, exTo, False
                        exPrev = exTo
                    Else
                        WriteRow ws, r, planDate, ShiftOf(t, shiftLen), sr, it(idx).code22, rel, toB, 0, 0, False
                    End If
                    r = r + 1: sr = sr + 1
                    t = t + take * cyc: rel = toB + 1: leftB = leftB - take
                End If
            Loop
        End If
        If r > PLAN_LAST Then Exit For
    Next s
End Sub

Private Function ShiftOf(t As Double, shiftLen As Double) As String
    Dim i As Long
    i = Int(t / shiftLen + 0.000001)
    If i > 2 Then i = 2
    ShiftOf = Choose(i + 1, "A", "B", "C")
End Function

Private Sub ClearPlan(ws As Worksheet)
    ws.Range("A" & PLAN_FIRST & ":D" & PLAN_LAST).ClearContents
    ws.Range("F" & PLAN_FIRST & ":H" & PLAN_LAST).ClearContents
    ws.Range("J" & PLAN_FIRST & ":L" & PLAN_LAST).ClearContents
    ws.Range("F" & PLAN_FIRST & ":L" & PLAN_LAST).Interior.ColorIndex = xlNone
    ws.Range(HELP_COL & PLAN_FIRST & ":BE" & PLAN_LAST).ClearContents
    ws.Range(HELP_COL & "3:BE3").Value = Array("Plan 21 rel from", "Plan 21 rel to", "Plan 22 rel from", "Plan 22 rel to", "Row type")
    ws.Columns(HELP_COL & ":BE").Hidden = True
End Sub

Private Sub WriteRow(ws As Worksheet, r As Long, d As Variant, sh As String, sr As Long, code As String, _
                     f21 As Long, t21 As Long, f22 As Long, t22 As Long, isCO As Boolean)
    Dim c As Variant
    ' make sure the row has your lookup formulas (copied from row 4)
    For Each c In Array("E", "I", "M", "O", "P", "Q", "R")
        If Not ws.Range(c & r).HasFormula And ws.Range(c & PLAN_FIRST).HasFormula Then
            ws.Range(c & PLAN_FIRST).Copy ws.Range(c & r)
        End If
    Next c
    ws.Range("A" & r).Value = d
    ws.Range("B" & r).Value = sh
    ws.Range("C" & r).Value = sr
    If IsNumeric(code) Then ws.Range("D" & r).Value = CDbl(code) Else ws.Range("D" & r).Value = code
    If f21 > 0 Then
        ws.Range("G" & r).Value = f21: ws.Range("H" & r).Value = t21
        ws.Cells(r, "BA").Value = f21: ws.Cells(r, "BB").Value = t21
    End If
    If f22 > 0 Then
        ws.Range("K" & r).Value = f22: ws.Range("L" & r).Value = t22
        ws.Cells(r, "BC").Value = f22: ws.Cells(r, "BD").Value = t22
    End If
    If isCO Then
        ws.Cells(r, "BE").Value = "CO"
    ElseIf code = "Line Clearance" Then
        ws.Cells(r, "BE").Value = "LC"
    Else
        ws.Cells(r, "BE").Value = "PROD"
    End If
End Sub

Private Function COFor(code21 As String) As String
    Dim ws As Worksheet, r As Long
    COFor = CodeStr(Cfg(S_DEFCO))
    If Not SheetExists(SH_COMAP) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_COMAP)
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CodeStr(ws.Cells(r, 1).Value) = code21 Then
            If CodeStr(ws.Cells(r, 3).Value) <> "" Then COFor = CodeStr(ws.Cells(r, 3).Value)
            Exit Function
        End If
    Next r
End Function

'==============================================================================
'  2. FILL PO NUMBERS (+ iterations)
'==============================================================================
Public Sub Fill_PO_Numbers()
    Dim m As Long, ws As Worksheet, r As Long, c22 As String, kind As String
    Dim mxPO As String, exPO As String, coPO As String, off As Long
    Dim missing As Long, filled As Long
    Application.ScreenUpdating = False
    For m = 1 To 3
        Set ws = ThisWorkbook.Worksheets("MX0" & m)
        For r = PLAN_FIRST To PLAN_LAST
            kind = CStr(ws.Cells(r, "BE").Value)
            c22 = CodeStr(ws.Range("D" & r).Value)
            If kind = "PROD" Then
                FindPO m, c22, mxPO, exPO
                If mxPO <> "" Then
                    ws.Range("F" & r).Value = CDbl(mxPO)
                    off = LastIteration(mxPO)
                    If NumOf(ws.Cells(r, "BA").Value) > 0 Then
                        ws.Range("G" & r).Value = off + ws.Cells(r, "BA").Value
                        ws.Range("H" & r).Value = off + ws.Cells(r, "BB").Value
                    End If
                    ws.Range("F" & r).Interior.ColorIndex = xlNone
                    filled = filled + 1
                Else
                    ws.Range("F" & r).Interior.Color = RGB(255, 199, 206): missing = missing + 1
                End If
                If exPO <> "" Then
                    ws.Range("J" & r).Value = CDbl(exPO)
                    off = LastIteration(exPO)
                    If NumOf(ws.Cells(r, "BC").Value) > 0 Then
                        ws.Range("K" & r).Value = off + ws.Cells(r, "BC").Value
                        ws.Range("L" & r).Value = off + ws.Cells(r, "BD").Value
                    End If
                    ws.Range("J" & r).Interior.ColorIndex = xlNone
                Else
                    ws.Range("J" & r).Interior.Color = RGB(255, 199, 206): missing = missing + 1
                End If
            ElseIf kind = "CO" Then
                coPO = FindCleanoutPO(m, c22)
                If coPO <> "" Then
                    ws.Range("F" & r).Value = CDbl(coPO): ws.Range("F" & r).Interior.ColorIndex = xlNone
                    filled = filled + 1
                Else
                    ws.Range("F" & r).Interior.Color = RGB(255, 199, 206): missing = missing + 1
                End If
            End If
        Next r
    Next m
    Application.ScreenUpdating = True
    MsgBox filled & " PO cells filled." & IIf(missing > 0, vbLf & missing & " PO(s) not found - marked pink. Enter them in the PO / clean-out sheets and run again.", ""), _
           IIf(missing > 0, vbExclamation, vbInformation)
End Sub

' latest PO lines for a 22 code in the mixer's PO sheet
Private Sub FindPO(m As Long, c22 As String, ByRef mxPO As String, ByRef exPO As String)
    Dim ws As Worksheet, r As Long, firstRow As Long, colMX As Long, lastR As Long
    mxPO = "": exPO = ""
    Select Case m
        Case 1: Set ws = ThisWorkbook.Worksheets("MX01 PO"): firstRow = 2: colMX = 8
        Case 2: Set ws = ThisWorkbook.Worksheets("MX02 PO "): firstRow = 2: colMX = 8
        Case 3: Set ws = ThisWorkbook.Worksheets("MX03 PO"): firstRow = 3: colMX = 9
    End Select
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = lastR To firstRow Step -1
        If CodeStr(ws.Cells(r, 2).Value) = c22 Then
            If mxPO = "" And CodeStr(ws.Cells(r, colMX).Value) <> "" Then mxPO = CodeStr(ws.Cells(r, colMX).Value)
            If exPO = "" And CodeStr(ws.Cells(r, 4).Value) <> "" Then exPO = CodeStr(ws.Cells(r, 4).Value)
            If mxPO <> "" And exPO <> "" Then Exit For
        End If
    Next r
End Sub

' latest clean-out order (New or Reuse) for a 29 code on this mixer
Private Function FindCleanoutPO(m As Long, c29 As String) As String
    Dim ws As Worksheet, r As Long, lastR As Long, bestD As Double, d As Double
    Dim sets As Variant, k As Long
    ' each set = date col, code col, order col
    Select Case m
        Case 1: Set ws = ThisWorkbook.Worksheets("MX1 Clean out"): sets = Array(Array(1, 2, 4), Array(6, 7, 9))
        Case 2: Set ws = ThisWorkbook.Worksheets("MX2 Clean out"): sets = Array(Array(1, 2, 4), Array(6, 7, 9))
        Case 3: Set ws = ThisWorkbook.Worksheets("MX03 Cleanout PO"): sets = Array(Array(1, 2, 5))
    End Select
    bestD = -1
    For k = LBound(sets) To UBound(sets)
        lastR = ws.Cells(ws.Rows.Count, sets(k)(1)).End(xlUp).Row
        For r = 3 To lastR
            If CodeStr(ws.Cells(r, sets(k)(1)).Value) = c29 And CodeStr(ws.Cells(r, sets(k)(2)).Value) <> "" Then
                d = NumOf(ws.Cells(r, sets(k)(0)).Value2)
                If d >= bestD Then bestD = d: FindCleanoutPO = CodeStr(ws.Cells(r, sets(k)(2)).Value)
            End If
        Next r
    Next k
End Function

'==============================================================================
'  3. TOTE REQUIREMENT
'==============================================================================
Public Sub Build_Tote_Requirement()
    Dim ws As Worksheet, mxs As Worksheet, m As Long, r As Long, out As Long, s As String
    Dim dict As Object, key As Variant, c22 As String, exN As Long, rank As Variant
    Set ws = ThisWorkbook.Worksheets(SH_TOTE)
    Set dict = CreateObject("Scripting.Dictionary")
    For m = 1 To 3
        Set mxs = ThisWorkbook.Worksheets("MX0" & m)
        For r = PLAN_FIRST To PLAN_LAST
            If mxs.Cells(r, "BE").Value = "PROD" And NumOf(mxs.Cells(r, "BD").Value) > 0 Then
                c22 = CodeStr(mxs.Range("D" & r).Value)
                exN = mxs.Cells(r, "BD").Value - mxs.Cells(r, "BC").Value + 1
                key = "MX0" & m & "|" & c22
                dict(key) = dict(key) + exN
            End If
        Next r
    Next m
    ws.Range("B" & TOTE_FIRST & ":I" & TOTE_LAST).ClearContents
    out = TOTE_FIRST
    For Each key In dict.Keys
        If out > TOTE_LAST Then Exit For
        s = CStr(out)
        c22 = Split(key, "|")(1)
        ws.Range("A" & s).Value = out - TOTE_FIRST + 1
        ws.Range("B" & s).Value = CDbl(c22)
        ws.Range("D" & s).Value = Split(key, "|")(0)
        ws.Range("E" & s).Value = dict(key)
        rank = Application.Index(ThisWorkbook.Worksheets(SH_PRI).Range("A2:A61"), _
               Application.Match(CDbl(c22), ThisWorkbook.Worksheets(SH_PRI).Range("B2:B61"), 0))
        If Not IsError(rank) Then ws.Range("I" & s).Value = rank
        out = out + 1
    Next key
    ' internal Raw Data lookups (replaces the old link to an external file)
    For r = TOTE_FIRST To TOTE_LAST
        s = CStr(r)
        ws.Range("C" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$E$2:$E$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("F" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$O$2:$O$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("H" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$M$2:$M$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("G" & s).Formula = "=IFERROR(E" & s & "*H" & s & ","""")"
    Next r
    MsgBox "Tote requirement updated (" & dict.Count & " line(s)).", vbInformation
End Sub

'==============================================================================
'  4. CLOSE DAY - remember last iteration per PO
'==============================================================================
Public Sub Close_Day()
    Dim m As Long
    For m = 1 To 3
        LogIterations ThisWorkbook.Worksheets("MX0" & m)
    Next m
    MsgBox "Iterations logged. Tomorrow's plan will continue each PO from its next iteration.", vbInformation
End Sub

Private Sub LogIterations(ws As Worksheet)
    Dim r As Long
    If Not SheetExists(SH_ITER) Then BuildIterationSheet
    For r = PLAN_FIRST To PLAN_LAST
        If CodeStr(ws.Range("F" & r).Value) <> "" And IsNumeric(ws.Range("H" & r).Value) And NumOf(ws.Range("H" & r).Value) > 0 Then
            SetIteration CodeStr(ws.Range("F" & r).Value), CLng(ws.Range("H" & r).Value), ws.Name
        End If
        If CodeStr(ws.Range("J" & r).Value) <> "" And IsNumeric(ws.Range("L" & r).Value) And NumOf(ws.Range("L" & r).Value) > 0 Then
            SetIteration CodeStr(ws.Range("J" & r).Value), CLng(ws.Range("L" & r).Value), ws.Name
        End If
    Next r
End Sub

Private Function LastIteration(po As String) As Long
    Dim ws As Worksheet, f As Range
    If Not SheetExists(SH_ITER) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    Set f = FindPORow(ws, po)
    If Not f Is Nothing Then LastIteration = CLng(NumOf(ws.Cells(f.Row, 2).Value))
End Function

Private Sub SetIteration(po As String, lastIt As Long, src As String)
    Dim ws As Worksheet, f As Range, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    Set f = FindPORow(ws, po)
    If f Is Nothing Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        ws.Cells(r, 1).NumberFormat = "@"
        ws.Cells(r, 1).Value = po
    Else
        r = f.Row
        If NumOf(ws.Cells(r, 2).Value) >= lastIt Then Exit Sub
    End If
    ws.Cells(r, 2).Value = lastIt
    ws.Cells(r, 3).Value = Now
    ws.Cells(r, 4).Value = src
End Sub

Private Function FindPORow(ws As Worksheet, po As String) As Range
    Dim r As Long
    For r = 2 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If CodeStr(ws.Cells(r, 1).Value) = po Then Set FindPORow = ws.Cells(r, 1): Exit Function
    Next r
End Function

'==============================================================================
'  helpers
'==============================================================================
Private Function CodeStr(v As Variant) As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    If IsNumeric(v) And Not VarType(v) = vbString Then
        CodeStr = Format(CDbl(v), "0")
    Else
        CodeStr = Trim(CStr(v))
        If IsNumeric(CodeStr) And CodeStr <> "" Then CodeStr = Format(CDbl(CodeStr), "0")
    End If
End Function

Private Function MixerNo(v As Variant) As Long
    Dim s As String
    s = UCase(Replace(Trim(CStr(v)), " ", ""))
    If s = "MX01" Or s = "MX1" Or s = "1" Then MixerNo = 1
    If s = "MX02" Or s = "MX2" Or s = "2" Then MixerNo = 2
    If s = "MX03" Or s = "MX3" Or s = "3" Then MixerNo = 3
End Function

Private Function Cfg(r As Long) As Variant
    Cfg = ThisWorkbook.Worksheets(SH_SET).Cells(r, 2).Value
End Function

Private Function GetNum(r As Long) As Double
    GetNum = NumOf(Cfg(r))
End Function

Private Function NumOf(v As Variant) As Double
    ' locale-safe number (Val() breaks on comma-decimal settings)
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    If IsNumeric(v) Then NumOf = CDbl(v)
End Function

Private Function Code21Of(c22 As String) As String
    Dim v As Variant
    v = Application.Match(CDbl(c22), ThisWorkbook.Worksheets(SH_RAW).Range("B2:B400"), 0)
    If IsError(v) Then v = Application.Match(c22, ThisWorkbook.Worksheets(SH_RAW).Range("B2:B400"), 0)
    If Not IsError(v) Then Code21Of = CodeStr(ThisWorkbook.Worksheets(SH_RAW).Range("C2:C400").Cells(v, 1).Value)
End Function

' 21 code of the last compound on yesterday's plan (only if the sheet holds an older date)
Private Function LastRunning21(m As Long) As String
    Dim ws As Worksheet, r As Long, c As String
    Set ws = ThisWorkbook.Worksheets("MX0" & m)
    If Not IsDate(ws.Range("A" & PLAN_FIRST).Value) Then Exit Function
    If CDate(ws.Range("A" & PLAN_FIRST).Value) >= CDate(Cfg(S_PLANDATE)) Then Exit Function
    For r = PLAN_LAST To PLAN_FIRST Step -1
        c = CodeStr(ws.Range("D" & r).Value)
        If Left(c, 2) = "22" And Len(c) = 8 Then LastRunning21 = Code21Of(c): Exit Function
    Next r
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nm)
    SheetExists = Not ws Is Nothing
End Function
