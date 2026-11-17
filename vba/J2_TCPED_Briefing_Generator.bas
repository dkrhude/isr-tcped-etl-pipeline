Attribute VB_Name = "J2_TCPED_Briefing_Generator"
Option Explicit

' ==============================================================================
' Module      : J2_TCPED_Briefing_Generator
' Repository  : isr-tcped-etl-pipeline
' Classification: 100% UNCLASSIFIED // SYNTHETIC EXERCISE DATA
' Description : Ingests normalized ETL output (arcgis_tableau_tcped_export.csv),
'               populates a formatted ListObject data sheet, aggregates Multi-INT
'               fulfillment KPIs using native VBA structs (zero ActiveX/STIG-safe),
'               builds a Priority Tier 1/2 Collection Gap Action Deck, and renders
'               a clustered column briefing chart.
' ==============================================================================

Private Type DisciplineMetric
    DisciplineName As String
    TotalReqs As Long
    SatisfiedCount As Long
    GapCount As Long
End Type

Public Sub GenerateJ2BriefingDeck()
    Dim csvPath As String
    Dim lines As Variant
    Dim headers() As String
    Dim rowFields() As String
    Dim wb As Workbook
    Dim wsData As Worksheet
    Dim wsBrief As Worksheet

    Dim totalTargets As Long
    Dim satisfiedTargets As Long
    Dim totalGaps As Long
    Dim tier1Gaps As Long
    Dim sumLatencySec As Double
    Dim latencySampleCount As Long

    Dim metrics(0 To 3) As DisciplineMetric
    Dim i As Long
    Dim r As Long
    Dim c As Long
    Dim outDataRow As Long
    Dim outGapRow As Long

    csvPath = ResolveCsvPath()
    If Len(csvPath) = 0 Then Exit Sub

    On Error GoTo CleanFail
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    lines = ReadCsvLines(csvPath)
    If UBound(lines) < 1 Then
        MsgBox "Selected CSV file contains no data records.", vbExclamation, "TCPED Briefing Generator"
        GoTo CleanExit
    End If

    ' Initialize canonical TCPED sensor ontology buckets
    metrics(0).DisciplineName = "EO/IR"
    metrics(1).DisciplineName = "SAR"
    metrics(2).DisciplineName = "SIGINT"
    metrics(3).DisciplineName = "LiDAR"

    Set wb = ActiveWorkbook
    Set wsBrief = GetOrResetSheet(wb, "J2_Executive_Brief")
    Set wsData = GetOrResetSheet(wb, "Normalized_TCPED_Deck")

    ' 1. Write Normalized_TCPED_Deck Headers
    headers = ParseCsvLine(CStr(lines(0)))
    For c = 0 To UBound(headers)
        wsData.Cells(1, c + 1).Value = headers(c)
    Next c

    ' Prepare J2_Executive_Brief Action Deck Headers (Row 16)
    BuildBriefStaticHeaders wsBrief

    outDataRow = 2
    outGapRow = 17

    ' 2. Parse Records, Populate Data Tab, Accumulate KPIs & Filter Tier 1/2 Gaps
    For r = 1 To UBound(lines)
        If Len(Trim$(CStr(lines(r)))) > 0 Then
            rowFields = ParseCsvLine(CStr(lines(r)))
            If UBound(rowFields) >= 16 Then
                WriteNormalizedRow wsData, outDataRow, rowFields

                Dim priTier As Long
                Dim sensorType As String
                Dim collStatus As String
                Dim platformName As String
                Dim callsignText As String
                Dim latVal As Double
                Dim lonVal As Double
                Dim latencySec As Double

                priTier = CLng(Val(rowFields(6)))
                sensorType = Trim$(rowFields(7))
                platformName = Trim$(rowFields(8))
                callsignText = Trim$(rowFields(9))
                collStatus = Trim$(rowFields(10))
                latVal = Val(rowFields(3))
                lonVal = Val(rowFields(4))
                latencySec = Val(rowFields(11))

                totalTargets = totalTargets + 1
                If collStatus = "SATISFIED" Then
                    satisfiedTargets = satisfiedTargets + 1
                Else
                    totalGaps = totalGaps + 1
                    If priTier = 1 Then tier1Gaps = tier1Gaps + 1
                End If

                If latencySec > 0 Then
                    sumLatencySec = sumLatencySec + latencySec
                    latencySampleCount = latencySampleCount + 1
                End If

                ' Accumulate Multi-INT Ontology Metrics
                For i = 0 To 3
                    If StrComp(metrics(i).DisciplineName, sensorType, vbTextCompare) = 0 Then
                        metrics(i).TotalReqs = metrics(i).TotalReqs + 1
                        If collStatus = "SATISFIED" Then
                            metrics(i).SatisfiedCount = metrics(i).SatisfiedCount + 1
                        Else
                            metrics(i).GapCount = metrics(i).GapCount + 1
                        End If
                        Exit For
                    End If
                Next i

                ' Populate Priority Tier 1 & 2 Unsatisfied / Degraded Action Deck
                If priTier <= 2 And collStatus <> "SATISFIED" Then
                    wsBrief.Cells(outGapRow, 1).Value = "PRI-" & CStr(priTier)
                    wsBrief.Cells(outGapRow, 2).Value = rowFields(1)
                    wsBrief.Cells(outGapRow, 3).Value = rowFields(2)
                    wsBrief.Cells(outGapRow, 4).Value = sensorType
                    wsBrief.Cells(outGapRow, 5).Value = collStatus
                    wsBrief.Cells(outGapRow, 6).Value = platformName & " (" & callsignText & ")"
                    wsBrief.Cells(outGapRow, 7).Value = Format$(latVal, "0.0000") & ", " & Format$(lonVal, "0.0000")
                    If latencySec > 0 Then
                        wsBrief.Cells(outGapRow, 8).Value = Round(latencySec / 60#, 1)
                        wsBrief.Cells(outGapRow, 8).NumberFormat = "0.0"
                    Else
                        wsBrief.Cells(outGapRow, 8).Value = "N/A"
                    End If
                    wsBrief.Cells(outGapRow, 9).Value = DeriveActionRecommendation(collStatus, sensorType)

                    ApplyStatusCellColor wsBrief.Cells(outGapRow, 5), collStatus
                    outGapRow = outGapRow + 1
                End If

                outDataRow = outDataRow + 1
            End If
        End If
    Next r

    ' 3. Format Normalized Data Sheet as Structured Table
    FormatNormalizedDataSheet wsData, outDataRow - 1, UBound(headers) + 1

    ' 4. Populate KPI Summary, Multi-INT Rollup Table, Borders & Embedded Chart
    Dim meanLatencyMin As Double
    If latencySampleCount > 0 Then
        meanLatencyMin = (sumLatencySec / CDbl(latencySampleCount)) / 60#
    Else
        meanLatencyMin = 0#
    End If

    FinalizeExecutiveBrief wsBrief, totalTargets, satisfiedTargets, totalGaps, _
                           tier1Gaps, meanLatencyMin, metrics, outGapRow - 1

    wsBrief.Activate
    wsBrief.Range("A1").Select

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    Exit Sub

CleanFail:
    Dim errMsg As String
    errMsg = "Runtime Error " & Err.Number & ": " & Err.Description
    Resume CleanErrorPrompt

CleanErrorPrompt:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
    MsgBox errMsg, vbCritical, "TCPED Briefing Generator Error"
End Sub

' ==============================================================================
' HELPER: CSV PATH RESOLUTION & BINARY LINE READER (CRLF / LF SAFE)
' ==============================================================================

Private Function ResolveCsvPath() As String
    Dim candidate1 As String
    Dim candidate2 As String
    Dim picked As Variant

    If Len(ThisWorkbook.Path) > 0 Then
        candidate1 = ThisWorkbook.Path & "\data\processed\arcgis_tableau_tcped_export.csv"
        If Len(Dir$(candidate1)) > 0 Then
            ResolveCsvPath = candidate1
            Exit Function
        End If

        candidate2 = ThisWorkbook.Path & "\..\data\processed\arcgis_tableau_tcped_export.csv"
        If Len(Dir$(candidate2)) > 0 Then
            ResolveCsvPath = candidate2
            Exit Function
        End If
    End If

    picked = Application.GetOpenFilename( _
        FileFilter:="CSV Files (*.csv),*.csv", _
        FilterIndex:=1, _
        Title:="Select arcgis_tableau_tcped_export.csv from data\processed" _
    )

    If VarType(picked) = vbBoolean Then
        ResolveCsvPath = ""
    Else
        ResolveCsvPath = CStr(picked)
    End If
End Function

Private Function ReadCsvLines(ByVal filePath As String) As Variant
    Dim fileNum As Integer
    Dim fileLen As Long
    Dim rawBytes() As Byte
    Dim rawText As String

    fileNum = FreeFile
    Open filePath For Binary Access Read As #fileNum
    fileLen = LOF(fileNum)
    If fileLen = 0 Then
        Close #fileNum
        ReadCsvLines = Array()
        Exit Function
    End If

    ReDim rawBytes(0 To fileLen - 1)
    Get #fileNum, , rawBytes
    Close #fileNum

    rawText = StrConv(rawBytes, vbUnicode)
    If Left$(rawText, 3) = Chr$(239) & Chr$(187) & Chr$(191) Then
        rawText = Mid$(rawText, 4)
    End If

    rawText = Replace(rawText, vbCrLf, vbLf)
    rawText = Replace(rawText, vbCr, vbLf)
    ReadCsvLines = Split(rawText, vbLf)
End Function

' ==============================================================================
' HELPER: DETERMINISTIC RFC-4180 QUOTE-AWARE CSV PARSER
' ==============================================================================

Private Function ParseCsvLine(ByVal lineText As String) As String()
    Dim fields() As String
    Dim fieldCount As Long
    Dim curField As String
    Dim inQuotes As Boolean
    Dim i As Long
    Dim ch As String
    Dim nextCh As String

    ReDim fields(0 To 31)
    fieldCount = 0
    curField = ""
    inQuotes = False

    i = 1
    Do While i <= Len(lineText)
        ch = Mid$(lineText, i, 1)
        If inQuotes Then
            If ch = """" Then
                If i < Len(lineText) Then
                    nextCh = Mid$(lineText, i + 1, 1)
                Else
                    nextCh = ""
                End If
                If nextCh = """" Then
                    curField = curField & """"
                    i = i + 1
                Else
                    inQuotes = False
                End If
            Else
                curField = curField & ch
            End If
        Else
            If ch = """" Then
                inQuotes = True
            ElseIf ch = "," Then
                If fieldCount > UBound(fields) Then
                    ReDim Preserve fields(0 To UBound(fields) * 2)
                End If
                fields(fieldCount) = curField
                fieldCount = fieldCount + 1
                curField = ""
            Else
                curField = curField & ch
            End If
        End If
        i = i + 1
    Loop

    If fieldCount > UBound(fields) Then
        ReDim Preserve fields(0 To fieldCount)
    End If
    fields(fieldCount) = curField
    ReDim Preserve fields(0 To fieldCount)
    ParseCsvLine = fields
End Function

' ==============================================================================
' HELPER: WORKSHEET MANAGEMENT & FORMATTING ROUTINES
' ==============================================================================

Private Function GetOrResetSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
        ws.Name = sheetName
    Else
        Do While ws.ListObjects.Count > 0
            ws.ListObjects(1).Unlist
        Loop
        If ws.ChartObjects.Count > 0 Then
            ws.ChartObjects.Delete
        End If
        ws.Cells.Clear
    End If

    ws.Cells.Font.Name = "Calibri"
    ws.Cells.Font.Size = 10
    Set GetOrResetSheet = ws
End Function

Private Sub WriteNormalizedRow(ByVal ws As Worksheet, ByVal rowIdx As Long, ByRef fields() As String)
    Dim c As Long
    Dim rawVal As String

    For c = 0 To UBound(fields)
        rawVal = Trim$(fields(c))
        Select Case c
            Case 3, 4, 6, 11, 12, 15, 16
                If Len(rawVal) > 0 Then
                    ws.Cells(rowIdx, c + 1).Value = Val(rawVal)
                End If
            Case Else
                ws.Cells(rowIdx, c + 1).Value = rawVal
        End Select
    Next c

    ApplyStatusCellColor ws.Cells(rowIdx, 11), Trim$(fields(10))
End Sub

Private Sub FormatNormalizedDataSheet(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim tblRange As Range
    Dim lo As ListObject

    If lastRow < 2 Then Exit Sub
    Set tblRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
    Set lo = ws.ListObjects.Add(xlSrcRange, tblRange, , xlYes)
    lo.Name = "tbl_TCPED_Normalized"
    lo.TableStyle = "TableStyleMedium2"

    ws.Range(ws.Cells(2, 4), ws.Cells(lastRow, 5)).NumberFormat = "0.00000"
    ws.Range(ws.Cells(2, 12), ws.Cells(lastRow, 12)).NumberFormat = "0"
    ws.Range(ws.Cells(2, 13), ws.Cells(lastRow, 13)).NumberFormat = "0.0"
    ws.Range(ws.Cells(2, 17), ws.Cells(lastRow, 17)).NumberFormat = "0.00"

    ws.Columns.AutoFit
    If ws.Columns(15).ColumnWidth > 48 Then ws.Columns(15).ColumnWidth = 48
End Sub

Private Sub BuildBriefStaticHeaders(ByVal ws As Worksheet)
    ' Classification & Title Banner (CenterAcrossSelection avoids MergeCells hazards)
    ws.Range("A1").Value = "UNCLASSIFIED // SYNTHETIC EXERCISE DATA - NO REAL-WORLD OPSEC"
    With ws.Range("A1:I1")
        .HorizontalAlignment = xlCenterAcrossSelection
        .Interior.Color = RGB(22, 101, 52)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 10
    End With

    ws.Range("A2").Value = "USSTRATCOM / THEATER J2 COLLECTION MANAGEMENT - TCPED GAP BRIEFING (" & _
                           Format$(Now, "yyyy-mm-dd hh:nn") & " LOCAL)"
    With ws.Range("A2:I2")
        .HorizontalAlignment = xlCenterAcrossSelection
        .Interior.Color = RGB(15, 23, 42)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .Font.Size = 12
    End With
    ws.Rows(2).RowHeight = 24

    ' Section 1: KPI Header
    ws.Range("A4").Value = "1. THEATER COLLECTION DECK EXECUTIVE KPIs"
    ws.Range("A4").Font.Bold = True
    ws.Range("A4").Font.Size = 11

    ws.Range("A5:F5").Value = Array( _
        "Validated Targets", "Satisfied Collects", "Total Theater Gaps", _
        "Critical Tier-1 Gaps", "Fulfillment Rate", "Mean PED Latency" _
    )
    StyleTableHeader ws.Range("A5:F5")

    ' Section 2: Multi-INT Rollup Header (Cols A:C contiguous for Chart Source)
    ws.Range("A8").Value = "2. MULTI-INT SENSOR DISCIPLINE FULFILLMENT"
    ws.Range("A8").Font.Bold = True
    ws.Range("A8").Font.Size = 11

    ws.Range("A9:F9").Value = Array( _
        "Sensor Discipline", "Satisfied", "Collection Gaps", _
        "Total Reqs", "Fulfillment %", "Discipline Health" _
    )
    StyleTableHeader ws.Range("A9:F9")

    ' Section 3: Action Deck Header
    ws.Range("A15").Value = "3. PRIORITY TIER 1 & 2 UNSATISFIED / DEGRADED COLLECTION GAPS (J2 ACTION DECK)"
    ws.Range("A15").Font.Bold = True
    ws.Range("A15").Font.Size = 11

    ws.Range("A16:I16").Value = Array( _
        "Priority", "Target ID", "Target Name", "Sensor", "Collection Status", _
        "Assigned Platform", "WGS84 Coords", "PED Latency (m)", "Root Cause & J2 Action" _
    )
    StyleTableHeader ws.Range("A16:I16")
End Sub

Private Sub FinalizeExecutiveBrief( _
    ByVal ws As Worksheet, _
    ByVal totalTargets As Long, _
    ByVal satisfiedTargets As Long, _
    ByVal totalGaps As Long, _
    ByVal tier1Gaps As Long, _
    ByVal meanLatencyMin As Double, _
    ByRef metrics() As DisciplineMetric, _
    ByVal lastGapRow As Long _
)
    Dim i As Long
    Dim rowIdx As Long
    Dim fulfillPct As Double

    ' Populate KPI Values (Row 6)
    ws.Range("A6").Value = totalTargets
    ws.Range("B6").Value = satisfiedTargets
    ws.Range("C6").Value = totalGaps
    ws.Range("D6").Value = tier1Gaps
    If totalTargets > 0 Then
        ws.Range("E6").Value = CDbl(satisfiedTargets) / CDbl(totalTargets)
    Else
        ws.Range("E6").Value = 0#
    End If
    ws.Range("E6").NumberFormat = "0.0%"
    ws.Range("F6").Value = meanLatencyMin
    ws.Range("F6").NumberFormat = "0.0 ""min"""

    With ws.Range("A6:F6")
        .Font.Bold = True
        .Font.Size = 13
        .HorizontalAlignment = xlCenter
        .Interior.Color = RGB(241, 245, 249)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
    If tier1Gaps > 0 Then
        ws.Range("D6").Interior.Color = RGB(254, 202, 202)
        ws.Range("D6").Font.Color = RGB(153, 27, 27)
    End If
    ws.Rows(6).RowHeight = 24

    ' Populate Multi-INT Discipline Table (Rows 10-13)
    For i = 0 To 3
        rowIdx = 10 + i
        ws.Cells(rowIdx, 1).Value = metrics(i).DisciplineName
        ws.Cells(rowIdx, 2).Value = metrics(i).SatisfiedCount
        ws.Cells(rowIdx, 3).Value = metrics(i).GapCount
        ws.Cells(rowIdx, 4).Value = metrics(i).TotalReqs

        If metrics(i).TotalReqs > 0 Then
            fulfillPct = CDbl(metrics(i).SatisfiedCount) / CDbl(metrics(i).TotalReqs)
        Else
            fulfillPct = 0#
        End If
        ws.Cells(rowIdx, 5).Value = fulfillPct
        ws.Cells(rowIdx, 5).NumberFormat = "0.0%"

        If fulfillPct >= 0.8 Then
            ws.Cells(rowIdx, 6).Value = "NOMINAL"
            ApplyStatusCellColor ws.Cells(rowIdx, 6), "SATISFIED"
        ElseIf fulfillPct >= 0.5 Then
            ws.Cells(rowIdx, 6).Value = "AT RISK"
            ApplyStatusCellColor ws.Cells(rowIdx, 6), "DEGRADED_PARTIAL"
        Else
            ws.Cells(rowIdx, 6).Value = "CRITICAL GAP"
            ApplyStatusCellColor ws.Cells(rowIdx, 6), "UNALLOCATED_GAP"
        End If
    Next i

    With ws.Range("A10:F13")
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A10:A13").Font.Bold = True

    If lastGapRow >= 17 Then
        With ws.Range(ws.Cells(17, 1), ws.Cells(lastGapRow, 9))
            .Borders.LineStyle = xlContinuous
            .Borders.Weight = xlThin
        End With
        ws.Range(ws.Cells(17, 1), ws.Cells(lastGapRow, 2)).HorizontalAlignment = xlCenter
        ws.Range(ws.Cells(17, 4), ws.Cells(lastGapRow, 5)).HorizontalAlignment = xlCenter
    End If

    ' AutoFit Columns BEFORE anchoring embedded chart coordinates
    ws.Columns("A:I").AutoFit
    For i = 1 To 6
        If ws.Columns(i).ColumnWidth < 16 Then ws.Columns(i).ColumnWidth = 16
    Next i

    ' Build Embedded Clustered Column Chart beside KPI & Discipline Tables
    Dim chtObj As ChartObject
    Dim cht As Chart
    Set chtObj = ws.ChartObjects.Add( _
        Left:=ws.Range("G4").Left + 8, _
        Top:=ws.Range("G4").Top, _
        Width:=360, _
        Height:=195 _
    )
    chtObj.Name = "cht_MultiIntFulfillment"
    Set cht = chtObj.Chart
    cht.ChartType = xlColumnClustered
    cht.SetSourceData Source:=ws.Range("A9:C13"), PlotBy:=xlColumns
    cht.HasTitle = True
    cht.ChartTitle.Text = "Multi-INT Satisfied vs. Collection Gaps"
    cht.ChartTitle.Font.Size = 10
    cht.ChartTitle.Font.Bold = True
    cht.HasLegend = True
    cht.Legend.Position = xlLegendPositionBottom

    If cht.SeriesCollection.Count >= 2 Then
        cht.SeriesCollection(1).Format.Fill.ForeColor.RGB = RGB(22, 163, 74)
        cht.SeriesCollection(2).Format.Fill.ForeColor.RGB = RGB(220, 38, 38)
    End If
End Sub

Private Sub StyleTableHeader(ByVal hdrRange As Range)
    With hdrRange
        .Interior.Color = RGB(30, 41, 59)
        .Font.Color = RGB(255, 255, 255)
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
    End With
End Sub

Private Sub ApplyStatusCellColor(ByVal targetCell As Range, ByVal statusText As String)
    targetCell.Font.Bold = True
    Select Case statusText
        Case "SATISFIED"
            targetCell.Interior.Color = RGB(220, 252, 231)
            targetCell.Font.Color = RGB(22, 101, 52)
        Case "DEGRADED_PARTIAL"
            targetCell.Interior.Color = RGB(254, 243, 199)
            targetCell.Font.Color = RGB(146, 64, 14)
        Case Else
            targetCell.Interior.Color = RGB(254, 202, 202)
            targetCell.Font.Color = RGB(153, 27, 27)
    End Select
End Sub

Private Function DeriveActionRecommendation(ByVal statusText As String, ByVal sensorType As String) As String
    Select Case statusText
        Case "UNALLOCATED_GAP"
            DeriveActionRecommendation = "CRITICAL: Zero sortie allocated; submit dynamic " & sensorType & " retasking."
        Case "SORTIE_ABORT_GAP"
            DeriveActionRecommendation = "SEVERE: SortieWX/Maint abort; cross-cue theater standby " & sensorType & " asset."
        Case "DEGRADED_PARTIAL"
            DeriveActionRecommendation = "MODERATE: NIIRS/Latency threshold missed; schedule re-collect prior to LTIOV."
        Case Else
            DeriveActionRecommendation = "Monitor collection deck."
    End Select
End Function
