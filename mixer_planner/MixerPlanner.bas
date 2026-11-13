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
'  Weekly: Start_New_Week     - snapshots limited-qty compounds into
'                                "Limited Batch Log" (planned / produced / remaining)
'  Mixer availability per weekday (3/2/1 shifts or OFF) is set in Planner Settings.
'  Priority column W lets you change the planned EX batches (0 = do not plan).
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
Private Const SH_LIMLOG As String = "Limited Batch Log"

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
Private Const S_WEEKNO As Long = 23
' mixer availability table on Planner Settings: E2 header, rows 3..9 = Monday..Sunday, F/G/H = MX01/MX02/MX03
Private Const AV_FIRST As Long = 3
Private Const AV_COL As Long = 6

' Iteration Log columns
Private Const IL_PO As Long = 1, IL_TYPE As Long = 2, IL_CODE As Long = 3, IL_MAT As Long = 4, IL_MIX As Long = 5
Private Const IL_DATE As Long = 6, IL_QTY As Long = 7, IL_CLOSED As Long = 8, IL_TFROM As Long = 9, IL_TTO As Long = 10
Private Const IL_LAST As Long = 11, IL_REM As Long = 12, IL_STAT As Long = 13, IL_UPD As Long = 14

Private curPO(1 To 500) As String, curIt(1 To 500) As Long, curBase(1 To 500) As Long, curMix(1 To 500) As String, nCur As Long

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
    manual As Boolean
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
    BuildLimitedLogSheet
    BuildControlPanel
    Application.ScreenUpdating = True
    MsgBox "Planner sheets are ready." & vbLf & vbLf & _
           "Please check 'Planner Settings' (yellow cells) - especially the" & vbLf & _
           "mixer cycle time per batch and the week start date.", vbInformation
End Sub

Private Function GetOrAddSheet(nm As String, afterName As String) As Worksheet
    Dim ws As Worksheet
    If SheetExists(nm) Then
        Set ws = ThisWorkbook.Worksheets(nm)
    Else
        If afterName <> "" Then
            Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(afterName))
        Else
            Set ws = ThisWorkbook.Worksheets.Add(Before:=ThisWorkbook.Worksheets(1))
        End If
        ws.Name = nm
    End If
    Set GetOrAddSheet = ws
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
    Set ws = GetOrAddSheet(SH_SET, "")
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
    SetRow ws, S_WEEKNO, "Week number (as per planning team)", DatePart("ww", thu, vbMonday, vbFirstFourDays), "Written into 'Limited Batch Log' by Start New Week"
    ' --- mixer availability per weekday ---
    Dim d As Long, m As Long
    ws.Range("E2").Value = "MIXER AVAILABILITY (shifts: 3 / 2 / 1 / OFF)"
    ws.Range("F2:H2").Value = Array("MX01", "MX02", "MX03")
    ws.Range("E2:H2").Font.Bold = True
    For d = 0 To 6
        ws.Cells(AV_FIRST + d, 5).Value = Format(DateSerial(2024, 1, 1) + d, "dddd")   ' 1-Jan-2024 = Monday
        For m = 0 To 2
            With ws.Cells(AV_FIRST + d, AV_COL + m)
                If IsEmpty(.Value) Then .Value = 3
                .Interior.Color = RGB(255, 255, 0)
                .Font.Color = RGB(0, 0, 255)
                .HorizontalAlignment = xlCenter
            End With
        Next m
    Next d
    ws.Range("E10").Value = "Example: run only MX01 + MX02 this week -> put OFF for MX03 every day. Sunday limited plan -> 1 for MX01, OFF for others."
    On Error Resume Next             ' drop-down list is a convenience only
    With ws.Range(ws.Cells(AV_FIRST, AV_COL), ws.Cells(AV_FIRST + 6, AV_COL + 2)).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="3,2,1,OFF"
    End With
    On Error GoTo 0
    ws.Columns("E").ColumnWidth = 44
    ws.Columns("F:H").ColumnWidth = 9
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
    ws.Range("A1:Y1").Value = Array("Rank", "22 Code", "Material", "21 Code", "Compound", "Colour", _
        "Stock (EX batches)", "EX batches req / day", "Days cover", "Weekly limit (STOCK K)", _
        "Released this week", "Limit left", "Status", "Plan EX batches", "Pallet kg", _
        "MX01 batch wt", "MX02 batch wt", "MX03 batch wt", "Mixer override (type MX01/MX02/MX03)", _
        "Assigned mixer", "Mixer batches", "Remark", _
        "MANUAL plan EX batches (blank = auto, 0 = do not plan)", "Reason for change (RM / tote / pallet shortage...)", _
        "FINAL plan EX batches")
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
        ws.Range("Y" & s).Formula = "=IF(B" & s & "="""","""",IF(TRIM(W" & s & "&"""")="""",N(N" & s & "),IFERROR(MAX(0,ROUNDDOWN(VALUE(W" & s & "&""""),0)),N(N" & s & "))))"
        ws.Range("A" & s).Formula = "=IF(N(Y" & s & ")>0,COUNTIFS($Y$2:$Y$61,"">0"",$I$2:$I$61,""<""&I" & s & ")" & _
            "+COUNTIFS($Y$2:$Y$61,"">0"",$I$2:$I$61,I" & s & ",$H$2:$H$61,"">""&H" & s & ")+1,"""")"
    Next r
    ws.Range("I2:I61").NumberFormat = "0.00"
    ws.Range("O2:R61").NumberFormat = "0.0"
    ws.Range("S2:S61").Interior.Color = RGB(255, 255, 0)
    ws.Range("W2:X61").Interior.Color = RGB(255, 255, 0)
    ws.Range("W2:W61").Font.Color = RGB(0, 0, 255)
    ws.Range("Y2:Y61").Font.Bold = True
    With ws.Range("A1:Y1")
        .Font.Bold = True: .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 78, 120): .WrapText = True
    End With
    ws.Rows(1).RowHeight = 45
    ws.Cells.Font.Name = "Arial"
    ws.Columns("A:Y").ColumnWidth = 11
    ws.Columns("C").ColumnWidth = 22: ws.Columns("V").ColumnWidth = 40
    ws.Columns("W").ColumnWidth = 18: ws.Columns("X").ColumnWidth = 34
    ' status colours
    Dim rng As Range
    Set rng = ws.Range("M2:M61")
    On Error Resume Next            ' colours are cosmetic - never stop setup for them
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
    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Range("A1:Y61").AutoFilter
    ' STOCK: compound turns green when its weekly limited quantity is finished
    Set rng = ThisWorkbook.Worksheets(SH_STOCK).Range("F2:F61")
    rng.FormatConditions.Add(Type:=xlExpression, Formula1:="=Priority!$M2=""LIMIT DONE""").Interior.Color = RGB(146, 208, 80)
    On Error GoTo 0
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
    Dim nxt As Long, i As Long, seen As Boolean
    Set ws = GetOrAddSheet(SH_COMAP, SH_SET)
    ws.Range("A1:C1").Value = Array("21 Code (next compound)", "Compound", "Clean-out 29 code to use before it")
    ws.Range("A1:C1").Font.Bold = True
    nxt = 2
    Do While Trim(CStr(ws.Cells(nxt, 1).Value)) <> ""
        nxt = nxt + 1
    Loop
    Set rec = ThisWorkbook.Worksheets(SH_RECIPE)
    For Each c In Array(2, 12, 21)
        For r = 3 To 80
            k = CodeStr(rec.Cells(r, c).Value)
            seen = False
            For i = 2 To nxt - 1
                If CodeStr(ws.Cells(i, 1).Value) = k Then seen = True: Exit For
            Next i
            If Left(k, 2) = "21" And Not seen Then
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
    Dim ws As Worksheet, r As Long, lastR As Long, n As Long, i As Long
    Dim oldPO() As String, oldIt() As Double, oldUpd() As Variant, oldSrc() As String
    Set ws = GetOrAddSheet(SH_ITER, SH_SET)
    ' migrate the first version (A=PO, B=Last iteration, C=Updated, D=Sheet)
    If CStr(ws.Range("B1").Value) = "Last iteration used" Then
        lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If lastR >= 2 Then
            ReDim oldPO(2 To lastR): ReDim oldIt(2 To lastR): ReDim oldUpd(2 To lastR): ReDim oldSrc(2 To lastR)
            For r = 2 To lastR
                oldPO(r) = CodeStr(ws.Cells(r, 1).Value): oldIt(r) = NumOf(ws.Cells(r, 2).Value)
                oldUpd(r) = ws.Cells(r, 3).Value: oldSrc(r) = CStr(ws.Cells(r, 4).Value)
            Next r
        End If
        ws.Cells.Clear
        If lastR >= 2 Then
            For r = 2 To lastR
                If oldPO(r) <> "" Then
                    ws.Cells(r, IL_PO).NumberFormat = "@": ws.Cells(r, IL_PO).Value = oldPO(r)
                    ws.Cells(r, IL_MIX).Value = oldSrc(r)
                    ws.Cells(r, IL_CLOSED).Value = oldIt(r)
                    ws.Cells(r, IL_UPD).Value = oldUpd(r)
                    LogRowFormulas ws, r
                End If
            Next r
        End If
    End If
    ws.Range("A1:N1").Value = Array("PO number", "PO type", "Code (22 / 21 / 29)", "Material", "Mixer", _
        "Released on", "Batches released (PO qty)", "Last iteration closed (till yesterday)", _
        "Today's plan: iteration from", "Today's plan: iteration to", "Last iteration used", _
        "Remaining batches", "Status", "Updated on")
    With ws.Range("A1:N1")
        .Font.Bold = True: .WrapText = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(31, 78, 120)
    End With
    ws.Rows(1).RowHeight = 45
    ws.Columns("A:N").ColumnWidth = 14
    ws.Columns("D").ColumnWidth = 24
    ws.Columns("F").NumberFormat = "dd-mmm-yy"
    ws.Columns("N").NumberFormat = "dd-mmm-yy hh:mm"
    ws.Cells.Font.Name = "Arial"
    On Error Resume Next
    ws.Range("M2:M2000").FormatConditions.Delete
    AddCF ws.Range("M2:M2000"), "COMPLETED", RGB(198, 239, 206)
    AddCF ws.Range("M2:M2000"), "RUNNING", RGB(255, 230, 153)
    AddCF ws.Range("M2:M2000"), "OVER PO QTY", RGB(255, 199, 206)
    On Error GoTo 0
End Sub

Private Sub LogRowFormulas(ws As Worksheet, r As Long)
    Dim s As String
    s = CStr(r)
    ws.Cells(r, IL_LAST).Formula = "=IF(A" & s & "="""","""",MAX(N(H" & s & "),N(J" & s & ")))"
    ws.Cells(r, IL_REM).Formula = "=IF(N(G" & s & ")=0,"""",G" & s & "-K" & s & ")"
    ws.Cells(r, IL_STAT).Formula = "=IF(A" & s & "="""","""",IF(N(G" & s & ")=0,"""",IF(K" & s & ">G" & s & ",""OVER PO QTY""," & _
        "IF(K" & s & "=G" & s & ",""COMPLETED"",IF(K" & s & ">0,""RUNNING"",""NOT STARTED"")))))"
End Sub

'==============================================================================
'  LIMITED BATCH LOG  (weekly limited-quantity compounds)
'==============================================================================
Private Sub BuildLimitedLogSheet()
    Dim ws As Worksheet
    Set ws = GetOrAddSheet(SH_LIMLOG, SH_PRI)
    ws.Range("A1:J1").Value = Array("22 Code", "Material", "Compound", "Week number", _
        "Planned batches (start of week)", "Batches produced till date", "Remaining batches for the week", _
        "Status", "Week start date", "Batches released in POs (this week)")
    With ws.Range("A1:J1")
        .Font.Bold = True: .WrapText = True
        .Font.Color = RGB(255, 255, 255): .Interior.Color = RGB(31, 78, 120)
    End With
    ws.Rows(1).RowHeight = 45
    ws.Columns("A:J").ColumnWidth = 13
    ws.Columns("B").ColumnWidth = 24
    ws.Columns("I").NumberFormat = "dd-mmm-yy"
    ws.Cells.Font.Name = "Arial"
    On Error Resume Next
    ws.Range("A2:J1000").FormatConditions.Delete
    ws.Range("A2:J1000").FormatConditions.Add(Type:=xlExpression, Formula1:="=$H2=""PLAN FINISHED""").Interior.Color = RGB(146, 208, 80)
    ws.Range("A2:J1000").FormatConditions.Add(Type:=xlExpression, Formula1:="=$H2=""ALL RELEASED""").Interior.Color = RGB(255, 230, 153)
    On Error GoTo 0
End Sub

' Thursday: copy every limited-quantity compound (number in STOCK col K) into the log
Public Sub Start_New_Week()
    Dim ws As Worksheet, st As Worksheet, r As Long, k As Variant, c22 As String, wk As Long
    Dim ws0 As Date, out As Long, found As Long, i As Long, added As Long, s As String, lastR As Long
    If Not SheetExists(SH_LIMLOG) Then BuildLimitedLogSheet
    Set ws = ThisWorkbook.Worksheets(SH_LIMLOG)
    Set st = ThisWorkbook.Worksheets(SH_STOCK)
    wk = CLng(GetNum(S_WEEKNO))
    ws0 = CDate(Cfg(S_WEEKSTART))
    For r = 2 To 61
        c22 = CodeStr(st.Cells(r, "D").Value)
        k = st.Cells(r, "K").Value
        If c22 <> "" And Not IsError(k) Then
            If UCase(Trim(CStr(k))) <> "CONT" And IsNumeric(k) And Trim(CStr(k)) <> "" Then
                If CDbl(k) > 0 Then
                    found = 0
                    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
                    For i = 2 To lastR
                        If CodeStr(ws.Cells(i, 1).Value) = c22 And NumOf(ws.Cells(i, 4).Value) = wk Then found = i: Exit For
                    Next i
                    If found = 0 Then        ' planned qty is frozen once written; edit E by hand if the plan is revised
                        out = lastR + 1
                        If out < 2 Then out = 2
                        s = CStr(out)
                        ws.Range("A" & s).Value = CDbl(c22)
                        ws.Range("B" & s).Value = st.Cells(r, "E").Value
                        ws.Range("C" & s).Value = st.Cells(r, "F").Value
                        ws.Range("D" & s).Value = wk
                        ws.Range("E" & s).Value = CDbl(k)
                        ws.Range("I" & s).Value = ws0
                        ' produced = extrusion iterations closed in Iteration Log for 22-code POs released this week
                        ws.Range("F" & s).Formula = "=SUMIFS('" & SH_ITER & "'!$H:$H,'" & SH_ITER & "'!$C:$C,A" & s & _
                            ",'" & SH_ITER & "'!$B:$B,""22 code (extruder)"",'" & SH_ITER & "'!$F:$F,"">=""&I" & s & _
                            ",'" & SH_ITER & "'!$F:$F,""<""&(I" & s & "+7))"
                        ws.Range("G" & s).Formula = "=MAX(0,E" & s & "-F" & s & ")"
                        ws.Range("H" & s).Formula = "=IF(F" & s & ">=E" & s & ",""PLAN FINISHED"",IF(J" & s & ">=E" & s & ",""ALL RELEASED"",""IN PROGRESS""))"
                        ws.Range("J" & s).Formula = "=SUMIFS('MX01 PO'!$E:$E,'MX01 PO'!$B:$B,A" & s & ",'MX01 PO'!$A:$A,"">=""&I" & s & ",'MX01 PO'!$A:$A,""<""&(I" & s & "+7))" & _
                            "+SUMIFS('MX02 PO '!$E:$E,'MX02 PO '!$B:$B,A" & s & ",'MX02 PO '!$A:$A,"">=""&I" & s & ",'MX02 PO '!$A:$A,""<""&(I" & s & "+7))" & _
                            "+SUMIFS('MX03 PO'!$E:$E,'MX03 PO'!$B:$B,A" & s & ",'MX03 PO'!$A:$A,"">=""&I" & s & ",'MX03 PO'!$A:$A,""<""&(I" & s & "+7))"
                        added = added + 1
                    End If
                End If
            End If
        End If
    Next r
    MsgBox added & " limited compound(s) added to '" & SH_LIMLOG & "' for week " & wk & ".", vbInformation
End Sub

Private Sub BuildControlPanel()
    Dim ws As Worksheet, b As Object, i As Long
    Dim caps As Variant, macs As Variant, notes As Variant
    Set ws = GetOrAddSheet(SH_PANEL, "")
    ws.Cells.Clear
    On Error Resume Next
    ws.Buttons.Delete
    On Error GoTo 0
    ws.Range("B2").Value = "MIXER COMPOUND PLANNER - daily steps"
    ws.Range("B2").Font.Size = 16: ws.Range("B2").Font.Bold = True
    caps = Array("1. Build Mixer Plan", "2. Fill PO Numbers", "3. Build Tote Requirement", "4. Close Day (log iterations)", _
                 "Thursday: Start New Week", "Update Iteration Log", "Re-run Setup")
    macs = Array("Build_Mixer_Plan", "Fill_PO_Numbers", "Build_Tote_Requirement", "Close_Day", _
                 "Start_New_Week", "Update_Iteration_Log", "Setup_Planner")
    notes = Array("After pasting LX02 (and on Thursday after updating STOCK H/K from the weekly plan). Writes MX01/MX02/MX03.", _
                  "After releasing POs in SAP and typing them in MX01 PO / MX02 PO / MX03 PO and the clean-out sheets.", _
                  "Fills 'Tote Requirment' from today's plan.", _
                  "At end of the day - stores the last iteration of every PO so tomorrow continues from there.", _
                  "After typing the weekly plan into STOCK (D/H/K), setting week start + week number and mixer availability.", _
                  "Adds newly released POs (PO sheets + clean-out sheets) to the Iteration Log. Also runs inside Fill PO / Close Day.", _
                  "Only when sheets are missing. Keeps your settings.")
    For i = 0 To 6
        Set b = ws.Buttons.Add(ws.Range("B" & 4 + i * 3).Left, ws.Range("B" & 4 + i * 3).Top, 220, 30)
        b.Caption = caps(i)
        b.OnAction = macs(i)
        ws.Range("D" & 4 + i * 3).Value = notes(i)
    Next i
    ws.Range("B31").Value = "Change a plan quantity: Priority column W (0 = do not plan) + reason in column X, then press 1 again."
    ws.Range("B28").Value = "Weekly (Thursday): type the planning image into STOCK - column D (22 code), H (Extrusion batch/day), K (CONT or FULL weekly batches)."
    ws.Range("B29").Value = "Then set 'Week start date' in Planner Settings. Do not reduce column K by hand - released POs are deducted automatically."
    ws.Range("B30").Value = "Priority sheet: do NOT sort it (rows mirror STOCK). Use the filter on column A (Rank) or M (Status)."
    ws.Columns("B").ColumnWidth = 34
    ws.Cells.Font.Name = "Arial"
End Sub

'==============================================================================
'  1. BUILD MIXER PLAN
'==============================================================================
Public Sub Build_Mixer_Plan()
    Dim wsP As Worksheet, r As Long, n As Long, i As Long, j As Long, m As Long
    Dim it() As PItem, ord() As Long, tmpI As Long, gi As Long, jj As Long
    Dim cyc(1 To 3) As Double, load(1 To 3) As Double, ncomp(1 To 3) As Long
    Dim lastRun(1 To 3) As String
    Dim steps(1 To 3, 1 To 200) As PStep, nSteps(1 To 3) As Long
    Dim dayMin As Double, lcMin As Double, maxComp As Long, dm(1 To 3) As Double, sh(1 To 3) As Long

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
        sh(m) = ShiftsAvailable(m, CDate(Cfg(S_PLANDATE)))
        dm(m) = sh(m) * GetNum(S_SHIFTMIN)
        If dm(m) > dayMin Then dm(m) = dayMin
        lastRun(m) = CodeStr(Cfg(S_LAST1 + m - 1))
        If lastRun(m) = "" Then lastRun(m) = LastRunning21(m)   ' auto: last compound of yesterday's plan
    Next m

    '--- read items that need planning ----------------------------------------
    ReDim it(1 To PRI_LAST)
    n = 0
    wsP.Range("T2:V" & PRI_LAST).ClearContents
    For r = PRI_FIRST To PRI_LAST
        If IsNumeric(wsP.Cells(r, "Y").Value) And CodeStr(wsP.Cells(r, "B").Value) <> "" Then
            If CDbl(wsP.Cells(r, "Y").Value) > 0 Then
                n = n + 1
                With it(n)
                    .priRow = r
                    .code22 = CodeStr(wsP.Cells(r, "B").Value)
                    .code21 = CodeStr(wsP.Cells(r, "D").Value)
                    .mat = CStr(wsP.Cells(r, "C").Value)
                    .cover = NumOf(wsP.Cells(r, "I").Value)
                    .req = NumOf(wsP.Cells(r, "H").Value)
                    .exB = CLng(wsP.Cells(r, "Y").Value)          ' FINAL plan = manual (col W) or auto (col N)
                    .manual = (Trim(CStr(wsP.Cells(r, "W").Value)) <> "")
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
    '    (an index list is sorted - the records themselves never move)
    ReDim ord(1 To n)
    For i = 1 To n: ord(i) = i: Next i
    For i = 1 To n - 1
        For j = i + 1 To n
            If it(ord(j)).cover < it(ord(i)).cover Or _
               (it(ord(j)).cover = it(ord(i)).cover And it(ord(j)).req > it(ord(i)).req) Or _
               (it(ord(j)).cover = it(ord(i)).cover And it(ord(j)).req = it(ord(i)).req And it(ord(j)).priRow < it(ord(i)).priRow) Then
                tmpI = ord(i): ord(i) = ord(j): ord(j) = tmpI
            End If
        Next j
    Next i

    '--- allocate groups (same 21 code = no LC between them) -------------------
    Dim g As Long, best As Long, bestKey As Double, needMin As Double, lc As Double
    Dim ok As Boolean, key As Double, avail As Double, mxB As Long, fitMx As Long, exFit As Long
    Dim ovr As Long
    For gi = 1 To n
        g = ord(gi)
        If Not it(g).done Then
            ' override: first item in the group that has one
            ovr = 0
            For jj = gi To n
                j = ord(jj)
                If it(j).code21 = it(g).code21 And it(j).override > 0 Then ovr = it(j).override: Exit For
            Next jj
            best = 0: bestKey = 1E+30
            For m = 1 To 3
                ok = (ovr = 0 Or ovr = m) And ncomp(m) < maxComp And dm(m) > 0
                If ok Then
                    needMin = 0
                    For jj = gi To n
                        j = ord(jj)
                        If it(j).code21 = it(g).code21 Then
                            If it(j).wt(m) <= 0 Or it(j).pallet <= 0 Then ok = False: Exit For
                            needMin = needMin + MixBatches(it(j).exB, it(j).pallet, it(j).wt(m)) * cyc(m)
                        End If
                    Next jj
                End If
                If ok Then
                    lc = LCNeeded(nSteps(m), it(g).code21, lastRun(m), lcMin)
                    key = load(m) + lc + needMin
                    If key > dm(m) Then key = key + 1000000#   ' prefer a mixer where it fits fully
                    If key < bestKey Then bestKey = key: best = m
                End If
            Next m

            If best = 0 Then
                For jj = gi To n
                    j = ord(jj)
                    If it(j).code21 = it(g).code21 Then
                        it(j).done = True
                        it(j).remark = "Not planned - no eligible mixer running today (check mixer availability / Recipe Details weights / compound limit)"
                    End If
                Next jj
            Else
                m = best
                lc = LCNeeded(nSteps(m), it(g).code21, lastRun(m), lcMin)
                If dm(m) - load(m) - lc < cyc(m) Then
                    For jj = gi To n
                        j = ord(jj)
                        If it(j).code21 = it(g).code21 Then it(j).done = True: it(j).remark = "Not planned - mixer capacity full"
                    Next jj
                Else
                    If lc > 0 Then
                        nSteps(m) = nSteps(m) + 1
                        steps(m, nSteps(m)).kind = "LC"
                        steps(m, nSteps(m)).item = g
                    End If
                    load(m) = load(m) + lc
                    ncomp(m) = ncomp(m) + 1
                    For jj = gi To n
                        j = ord(jj)
                        If it(j).code21 = it(g).code21 Then
                            it(j).done = True
                            mxB = MixBatches(it(j).exB, it(j).pallet, it(j).wt(m))
                            avail = dm(m) - load(m)
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
                    Next jj
                End If
            End If
        End If
    Next gi

    '--- write Priority results ------------------------------------------------
    For i = 1 To n
        With it(i)
            If .mixer > 0 Then
                wsP.Cells(.priRow, "T").Value = "MX0" & .mixer
                wsP.Cells(.priRow, "U").Value = .mxB
            End If
            If .manual Then
                wsP.Cells(.priRow, "V").Value = Trim("MANUAL qty" & IIf(Trim(CStr(wsP.Cells(.priRow, "X").Value)) <> "", " (" & wsP.Cells(.priRow, "X").Value & ")", "") & ". " & .remark)
            Else
                wsP.Cells(.priRow, "V").Value = .remark
            End If
        End With
    Next i

    '--- write the three mixer sheets ------------------------------------------
    Dim summary As String
    For m = 1 To 3
        WriteMixerSheet m, it, steps, nSteps(m), cyc(m)
        If dm(m) <= 0 Then
            With ThisWorkbook.Worksheets("MX0" & m)
                .Range("A" & PLAN_FIRST).Value = Cfg(S_PLANDATE)
                .Range("D" & PLAN_FIRST).Value = "MIXER OFF today (Planner Settings)"
            End With
            summary = summary & "MX0" & m & ": OFF" & vbLf
        Else
            summary = summary & "MX0" & m & ": " & Format(load(m), "0") & " / " & Format(dm(m), "0") & " min (" & sh(m) & " shift(s)), " & ncomp(m) & " compound(s)" & vbLf
        End If
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
        If CDate(ws.Range("A" & PLAN_FIRST).Value) < CDate(Cfg(S_PLANDATE)) Then
            LogIterations ws                 ' yesterday's plan -> "last iteration closed"
            ClearTodayIterations
        End If
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
    Dim mxPO As String, exPO As String, coPO As String, mxQty As Long, exQty As Long
    Dim n As Long, fromIt As Long, toIt As Long
    Dim missing As Long, filled As Long, over As Long
    Application.ScreenUpdating = False
    SyncPOLog                       ' newly released POs -> Iteration Log
    ClearTodayIterations
    nCur = 0
    For m = 1 To 3
        Set ws = ThisWorkbook.Worksheets("MX0" & m)
        For r = PLAN_FIRST To PLAN_LAST
            kind = CStr(ws.Cells(r, "BE").Value)
            c22 = CodeStr(ws.Range("D" & r).Value)
            If kind = "PROD" Then
                FindPO m, c22, mxPO, mxQty, exPO, exQty
                ' --- 21 code (mixer) PO ---
                ws.Range("F" & r & ":H" & r).Interior.ColorIndex = xlNone
                If mxPO <> "" Then
                    ws.Range("F" & r).Value = CDbl(mxPO)
                    n = CLng(NumOf(ws.Cells(r, "BB").Value) - NumOf(ws.Cells(r, "BA").Value) + 1)
                    If NumOf(ws.Cells(r, "BA").Value) > 0 And n > 0 Then
                        TakeIterations mxPO, n, fromIt, toIt, "MX0" & m   ' continues PO across rows and days
                        ws.Range("G" & r).Value = fromIt
                        ws.Range("H" & r).Value = toIt
                        If mxQty > 0 And toIt > mxQty Then ws.Range("G" & r & ":H" & r).Interior.Color = RGB(255, 192, 0): over = over + 1
                    End If
                    filled = filled + 1
                Else
                    ws.Range("F" & r).Interior.Color = RGB(255, 199, 206): missing = missing + 1
                End If
                ' --- 22 code (extruder) PO ---
                ws.Range("J" & r & ":L" & r).Interior.ColorIndex = xlNone
                If exPO <> "" Then
                    ws.Range("J" & r).Value = CDbl(exPO)
                    n = CLng(NumOf(ws.Cells(r, "BD").Value) - NumOf(ws.Cells(r, "BC").Value) + 1)
                    If NumOf(ws.Cells(r, "BC").Value) > 0 And n > 0 Then
                        TakeIterations exPO, n, fromIt, toIt, "MX0" & m
                        ws.Range("K" & r).Value = fromIt
                        ws.Range("L" & r).Value = toIt
                        If exQty > 0 And toIt > exQty Then ws.Range("K" & r & ":L" & r).Interior.Color = RGB(255, 192, 0): over = over + 1
                    End If
                    filled = filled + 1
                Else
                    ws.Range("J" & r).Interior.Color = RGB(255, 199, 206): missing = missing + 1
                End If
            ElseIf kind = "CO" Then
                coPO = FindCleanoutPO(m, c22)
                If coPO <> "" Then
                    ws.Range("F" & r).Value = CDbl(coPO): ws.Range("F" & r).Interior.ColorIndex = xlNone
                    TakeIterations coPO, 1, fromIt, toIt, "MX0" & m     ' each clean-out row uses 1 batch of the order
                    ws.Range("G" & r).Value = fromIt
                    ws.Range("H" & r).Value = toIt
                    filled = filled + 1
                Else
                    ws.Range("F" & r).Interior.Color = RGB(255, 199, 206): missing = missing + 1
                End If
            End If
        Next r
    Next m
    WriteTodayIterations
    Application.ScreenUpdating = True
    MsgBox filled & " PO cells filled." & _
           IIf(missing > 0, vbLf & missing & " PO(s) not found (pink). Enter them in the PO / clean-out sheets and run again.", "") & _
           IIf(over > 0, vbLf & over & " row(s) need more batches than the PO has (orange). Release more or adjust.", ""), _
           IIf(missing + over > 0, vbExclamation, vbInformation)
End Sub

' running "last iteration" per PO during one Fill run (starts from Iteration Log)
Private Sub TakeIterations(po As String, n As Long, ByRef fromIt As Long, ByRef toIt As Long, mixer As String)
    Dim i As Long
    For i = 1 To nCur
        If curPO(i) = po Then Exit For
    Next i
    If i > nCur Then
        nCur = nCur + 1: i = nCur
        curPO(i) = po
        curIt(i) = LastIteration(po)
        curBase(i) = curIt(i)
        curMix(i) = mixer
    End If
    fromIt = curIt(i) + 1
    toIt = curIt(i) + n
    curIt(i) = toIt
End Sub

' write today's planned iteration range per PO into the Iteration Log
Private Sub WriteTodayIterations()
    Dim ws As Worksheet, i As Long, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    For i = 1 To nCur
        If curIt(i) > curBase(i) Then
            r = EnsureLogRow(curPO(i))
            ws.Cells(r, IL_TFROM).Value = curBase(i) + 1
            ws.Cells(r, IL_TTO).Value = curIt(i)
            If Trim(CStr(ws.Cells(r, IL_MIX).Value)) = "" Then ws.Cells(r, IL_MIX).Value = curMix(i)
            ws.Cells(r, IL_UPD).Value = Now
        End If
    Next i
End Sub

Private Sub ClearTodayIterations()
    Dim ws As Worksheet, lastR As Long
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    lastR = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastR >= 2 Then ws.Range(ws.Cells(2, IL_TFROM), ws.Cells(lastR, IL_TTO)).ClearContents
End Sub

' latest PO lines for a 22 code in the mixer's PO sheet
Private Sub FindPO(m As Long, c22 As String, ByRef mxPO As String, ByRef mxQty As Long, _
                   ByRef exPO As String, ByRef exQty As Long)
    ' latest OPEN PO for the 22 code: released batches > last iteration already used
    Dim ws As Worksheet, r As Long, firstRow As Long, colMX As Long, lastR As Long
    Dim po As String, q As Long
    mxPO = "": exPO = "": mxQty = 0: exQty = 0
    Select Case m
        Case 1: Set ws = ThisWorkbook.Worksheets("MX01 PO"): firstRow = 2: colMX = 8
        Case 2: Set ws = ThisWorkbook.Worksheets("MX02 PO "): firstRow = 2: colMX = 8
        Case 3: Set ws = ThisWorkbook.Worksheets("MX03 PO"): firstRow = 3: colMX = 9
    End Select
    lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
    For r = lastR To firstRow Step -1
        If CodeStr(ws.Cells(r, 2).Value) = c22 Then
            po = CodeStr(ws.Cells(r, colMX).Value)
            If mxPO = "" And po <> "" Then
                q = CLng(NumOf(ws.Cells(r, colMX + 1).Value))
                If q = 0 Or LastIteration(po) < q Then mxPO = po: mxQty = q
            End If
            po = CodeStr(ws.Cells(r, 4).Value)
            If exPO = "" And po <> "" Then
                q = CLng(NumOf(ws.Cells(r, 5).Value))
                If q = 0 Or LastIteration(po) < q Then exPO = po: exQty = q
            End If
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
    Dim c22 As String, exN As Long, rank As Variant, i As Long, found As Boolean
    Dim keyMx(1 To 100) As String, keyCode(1 To 100) As String, qty(1 To 100) As Long, nk As Long
    Set ws = ThisWorkbook.Worksheets(SH_TOTE)
    ' total planned EX batches per mixer + 22 code, in plan order
    For m = 1 To 3
        Set mxs = ThisWorkbook.Worksheets("MX0" & m)
        For r = PLAN_FIRST To PLAN_LAST
            If CStr(mxs.Cells(r, "BE").Value) = "PROD" And NumOf(mxs.Cells(r, "BD").Value) > 0 Then
                c22 = CodeStr(mxs.Range("D" & r).Value)
                exN = CLng(NumOf(mxs.Cells(r, "BD").Value) - NumOf(mxs.Cells(r, "BC").Value) + 1)
                found = False
                For i = 1 To nk
                    If keyMx(i) = "MX0" & m And keyCode(i) = c22 Then qty(i) = qty(i) + exN: found = True: Exit For
                Next i
                If Not found And nk < 100 Then
                    nk = nk + 1: keyMx(nk) = "MX0" & m: keyCode(nk) = c22: qty(nk) = exN
                End If
            End If
        Next r
    Next m
    ws.Range("A" & TOTE_FIRST & ":I" & TOTE_LAST).ClearContents
    out = TOTE_FIRST
    For i = 1 To nk
        If out > TOTE_LAST Then Exit For
        s = CStr(out)
        ws.Range("A" & s).Value = i
        ws.Range("B" & s).Value = CDbl(keyCode(i))
        ws.Range("D" & s).Value = keyMx(i)
        ws.Range("E" & s).Value = qty(i)
        rank = Application.Match(CDbl(keyCode(i)), ThisWorkbook.Worksheets(SH_PRI).Range("B2:B61"), 0)
        If Not IsError(rank) Then ws.Range("I" & s).Value = ThisWorkbook.Worksheets(SH_PRI).Range("A2:A61").Cells(rank, 1).Value
        out = out + 1
    Next i
    ' internal Raw Data lookups (replaces the old link to an external file)
    For r = TOTE_FIRST To TOTE_LAST
        s = CStr(r)
        ws.Range("C" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$E$2:$E$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("F" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$O$2:$O$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("H" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(INDEX('Raw Data'!$M$2:$M$400,MATCH(B" & s & ",'Raw Data'!$B$2:$B$400,0)),""""))"
        ws.Range("G" & s).Formula = "=IF(B" & s & "="""","""",IFERROR(E" & s & "*H" & s & ",""""))"
    Next r
    MsgBox "Tote requirement updated (" & nk & " line(s)).", vbInformation
End Sub

'==============================================================================
'  4. CLOSE DAY - remember last iteration per PO
'==============================================================================
Public Sub Close_Day()
    Dim m As Long
    SyncPOLog
    For m = 1 To 3
        LogIterations ThisWorkbook.Worksheets("MX0" & m)
    Next m
    ClearTodayIterations
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
    If Not f Is Nothing Then LastIteration = CLng(NumOf(ws.Cells(f.Row, IL_CLOSED).Value))
End Function

Private Sub SetIteration(po As String, lastIt As Long, src As String)
    Dim ws As Worksheet, r As Long
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    r = EnsureLogRow(po)
    If Trim(CStr(ws.Cells(r, IL_MIX).Value)) = "" Then ws.Cells(r, IL_MIX).Value = src
    If NumOf(ws.Cells(r, IL_CLOSED).Value) >= lastIt Then Exit Sub
    ws.Cells(r, IL_CLOSED).Value = lastIt
    ws.Cells(r, IL_UPD).Value = Now
End Sub

' row of a PO in the Iteration Log (added if missing)
Private Function EnsureLogRow(po As String) As Long
    Dim ws As Worksheet, f As Range, r As Long
    If Not SheetExists(SH_ITER) Then BuildIterationSheet
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    Set f = FindPORow(ws, po)
    If f Is Nothing Then
        r = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row + 1
        If r < 2 Then r = 2
        ws.Cells(r, IL_PO).NumberFormat = "@"
        ws.Cells(r, IL_PO).Value = po
        LogRowFormulas ws, r
    Else
        r = f.Row
    End If
    EnsureLogRow = r
End Function

' add / update released POs (PO sheets + clean-out sheets) from 7 days before week start onwards
Public Sub Update_Iteration_Log()
    SyncPOLog
    MsgBox "Iteration Log updated with released POs.", vbInformation
End Sub

Private Sub SyncPOLog()
    Dim m As Long, ws As Worksheet, r As Long, lastR As Long, firstRow As Long, colMX As Long
    Dim fromD As Double, d As Double, k As Long, sets As Variant, q As Double
    If Not SheetExists(SH_ITER) Then BuildIterationSheet
    fromD = NumOf(Cfg(S_WEEKSTART)) - 7
    For m = 1 To 3
        Select Case m
            Case 1: Set ws = ThisWorkbook.Worksheets("MX01 PO"): firstRow = 2: colMX = 8
            Case 2: Set ws = ThisWorkbook.Worksheets("MX02 PO "): firstRow = 2: colMX = 8
            Case 3: Set ws = ThisWorkbook.Worksheets("MX03 PO"): firstRow = 3: colMX = 9
        End Select
        lastR = ws.Cells(ws.Rows.Count, 2).End(xlUp).Row
        For r = firstRow To lastR
            d = NumOf(ws.Cells(r, 1).Value2)
            If d >= fromD Then
                ' 22 code (extruder) PO: D = order, E = batches
                PutLog CodeStr(ws.Cells(r, 4).Value), "22 code (extruder)", ws.Cells(r, 2).Value, ws.Cells(r, 3).Value, _
                       "MX0" & m, d, NumOf(ws.Cells(r, 5).Value)
                ' 21 code (mixer) PO
                PutLog CodeStr(ws.Cells(r, colMX).Value), "21 code (mixer)", ws.Cells(r, 7).Value, ws.Cells(r, 3).Value, _
                       "MX0" & m, d, NumOf(ws.Cells(r, colMX + 1).Value)
            End If
        Next r
        ' clean-out orders: each set = date, code, material, order, batches (0 = none)
        Select Case m
            Case 1: Set ws = ThisWorkbook.Worksheets("MX1 Clean out"): sets = Array(Array(1, 2, 3, 4, 5), Array(6, 7, 8, 9, 0))
            Case 2: Set ws = ThisWorkbook.Worksheets("MX2 Clean out"): sets = Array(Array(1, 2, 3, 4, 5), Array(6, 7, 8, 9, 0))
            Case 3: Set ws = ThisWorkbook.Worksheets("MX03 Cleanout PO"): sets = Array(Array(1, 2, 3, 5, 6))
        End Select
        For k = LBound(sets) To UBound(sets)
            lastR = ws.Cells(ws.Rows.Count, sets(k)(3)).End(xlUp).Row
            For r = 3 To lastR
                d = NumOf(ws.Cells(r, sets(k)(0)).Value2)
                If d >= fromD Then
                    q = 0
                    If sets(k)(4) > 0 Then q = NumOf(ws.Cells(r, sets(k)(4)).Value)
                    PutLog CodeStr(ws.Cells(r, sets(k)(3)).Value), "29 code (clean-out)", ws.Cells(r, sets(k)(1)).Value, _
                           ws.Cells(r, sets(k)(2)).Value, "MX0" & m, d, q
                End If
            Next r
        Next k
    Next m
End Sub

Private Sub PutLog(po As String, poType As String, code As Variant, mat As Variant, mixer As String, d As Double, qty As Double)
    Dim ws As Worksheet, r As Long
    If po = "" Or Not IsNumeric(po) Then Exit Sub
    Set ws = ThisWorkbook.Worksheets(SH_ITER)
    r = EnsureLogRow(po)
    ws.Cells(r, IL_TYPE).Value = poType
    If CodeStr(code) <> "" And IsNumeric(CodeStr(code)) Then ws.Cells(r, IL_CODE).Value = CDbl(CodeStr(code))
    If Not IsError(mat) Then
        If Trim(CStr(mat)) <> "" Then ws.Cells(r, IL_MAT).Value = Trim(CStr(mat))
    End If
    ws.Cells(r, IL_MIX).Value = mixer
    If d > 0 Then ws.Cells(r, IL_DATE).Value = CDate(d)
    If qty > 0 Then ws.Cells(r, IL_QTY).Value = qty
End Sub

Private Function FindPORow(ws As Worksheet, po As String) As Range
    Dim r As Long
    Set FindPORow = Nothing
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
    If VarType(v) = vbDate Then NumOf = CDbl(v): Exit Function   ' IsNumeric() is False for dates
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

' shifts a mixer runs on a date (Planner Settings availability table; blank = 3, OFF = 0)
Private Function ShiftsAvailable(m As Long, d As Date) As Long
    Dim v As Variant
    v = ThisWorkbook.Worksheets(SH_SET).Cells(AV_FIRST + Weekday(d, vbMonday) - 1, AV_COL + m - 1).Value
    If IsError(v) Then Exit Function
    If Trim(CStr(v)) = "" Then
        ShiftsAvailable = 3
    ElseIf IsNumeric(v) Then
        ShiftsAvailable = CLng(v)
        If ShiftsAvailable > 3 Then ShiftsAvailable = 3
        If ShiftsAvailable < 0 Then ShiftsAvailable = 0
    Else
        ShiftsAvailable = 0                ' "OFF" or anything else
    End If
End Function

Private Function SheetExists(nm As String) As Boolean
    Dim sh As Object
    For Each sh In ThisWorkbook.Worksheets
        If sh.Name = nm Then SheetExists = True: Exit Function
    Next sh
End Function
