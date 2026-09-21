# Conversation

## User

I am a finance student and I do not know VBA. Using the attached TargetCo workbook and
functional specification, produce \textbf{**one complete VBA standard-module implementation**}
of the Excel Audit Assistant. It must satisfy \textbf{**every requirement**} in the specification. Do not give me pseudocode,
partial procedures, placeholders or alternative versions. Use \texttt{Public Sub AuditOpenTargetCo()} as the macro I will run.
Return, in this order:
(1) one complete VBA code block;
(2) the name of the macro I must run;
(3) a short table stating which workbook the code reads from and which workbook it writes to;
(4) exact installation instructions for desktop Excel.
TargetCo must never be edited, saved, renamed or closed by the macro. All findings must be
written to a new report workbook. Do not choose an EBITDA treatment and do not apply any
correction.
If any requirement cannot be satisfied, tell me \textbf{**before**} giving the code.

## Assistant

```vba
Option Explicit

'=========================================================================
' Excel Audit Assistant
'
' Entry macro:
'   Public Sub AuditOpenTargetCo()
'
' SAFETY DESIGN
' - Reads only from the already-open TargetCo workbook.
' - Never writes to TargetCo.
' - Never saves, renames, deletes, or closes TargetCo.
' - Never calls Shell, a network service, an API, or an external file.
' - Creates a NEW unsaved workbook for the audit report.
' - Does not propose or apply corrections.
' - Does not choose between reported and management-adjusted EBITDA.
'
' The workbook itself declares the following modelling convention:
'   Blue  = hardcoded input
'   Black = same-sheet calculation
'   Green = direct link to another worksheet
'=========================================================================

Private Const FIRST_YEAR_COL As Long = 4     'D
Private Const LAST_YEAR_COL As Long = 7      'G

Private Const RAW_SHEET As String = "01_Raw_Data"
Private Const HIST_SHEET As String = "02_Historical_Analysis"
Private Const ADJ_SHEET As String = "03_Adjustments"
Private Const SUMMARY_SHEET As String = "04_Summary"
Private Const EBITDA_SHEET As String = "05_EBITDA_Bases"
Private Const AUDIT_LOG_SHEET As String = "06_Audit_Log"


'=========================================================================
' MAIN ENTRY POINT
'=========================================================================

Public Sub AuditOpenTargetCo()

    Dim targetWb As Workbook
    Dim reportWb As Workbook
    Dim reportWs As Worksheet

    Dim nextRow As Long
    Dim findingNumber As Long

    On Error GoTo FatalError

    Set targetWb = FindOpenTargetCoWorkbook()

    If targetWb Is Nothing Then
        MsgBox _
            "The open TargetCo workbook could not be identified." & vbCrLf & vbCrLf & _
            "Keep TargetCo_historical_model.xlsx open, together with the separate " & _
            "macro-enabled controller workbook, and run the macro again.", _
            vbExclamation, _
            "Excel Audit Assistant"
        Exit Sub
    End If

    'The report is deliberately a brand-new workbook.
    Set reportWb = Workbooks.Add(xlWBATWorksheet)
    Set reportWs = reportWb.Worksheets(1)

    reportWs.Name = "Audit Findings"

    PrepareReportSheet reportWs, targetWb

    nextRow = 6
    findingNumber = 0

    '1. General formula-pattern and formula-presence review.
    AuditFormulaPatterns targetWb, reportWs, nextRow, findingNumber

    '2. Formula/reference-error review.
    AuditFormulaAndReferenceErrors targetWb, reportWs, nextRow, findingNumber

    '3. Formula-role versus colour-convention review.
    AuditColourMetadata targetWb, reportWs, nextRow, findingNumber

    '4. Independent finance controls.
    AuditFinanceControls targetWb, reportWs, nextRow, findingNumber

    If findingNumber = 0 Then
        reportWs.Cells(nextRow, 1).Value = "No candidate findings were generated."
        reportWs.Cells(nextRow, 1).Font.Bold = True
    End If

    FinaliseReportSheet reportWs, nextRow

    reportWb.Activate
    reportWs.Activate
    reportWs.Range("A1").Select

    MsgBox _
        CStr(findingNumber) & " candidate finding(s) written to a new report workbook." & _
        vbCrLf & vbCrLf & _
        "TargetCo was read only. No correction was applied.", _
        vbInformation, _
        "Excel Audit Assistant"

    Exit Sub

FatalError:

    MsgBox _
        "The audit stopped because of an unexpected error." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
        "No instruction in this module saves, renames, closes, or edits TargetCo.", _
        vbCritical, _
        "Excel Audit Assistant"

End Sub


'=========================================================================
' IDENTIFY THE OPEN TARGET WORKBOOK
'=========================================================================

Private Function FindOpenTargetCoWorkbook() As Workbook

    Dim wb As Workbook
    Dim candidate As Workbook
    Dim candidateCount As Long

    candidateCount = 0

    For Each wb In Application.Workbooks

        If Not wb Is ThisWorkbook Then

            If WorkbookHasTargetCoStructure(wb) Then

                'Prefer a workbook whose filename identifies it as TargetCo.
                If LCase$(wb.Name) Like "targetco_historical_model*.xls*" Then
                    Set FindOpenTargetCoWorkbook = wb
                    Exit Function
                End If

                candidateCount = candidateCount + 1
                Set candidate = wb

            End If

        End If

    Next wb

    'If exactly one other open workbook has the complete required structure,
    'accept it even if the user has a slightly different filename.
    If candidateCount = 1 Then
        Set FindOpenTargetCoWorkbook = candidate
    Else
        Set FindOpenTargetCoWorkbook = Nothing
    End If

End Function


Private Function WorkbookHasTargetCoStructure(ByVal wb As Workbook) As Boolean

    WorkbookHasTargetCoStructure = _
        HasWorksheet(wb, RAW_SHEET) And _
        HasWorksheet(wb, HIST_SHEET) And _
        HasWorksheet(wb, ADJ_SHEET) And _
        HasWorksheet(wb, SUMMARY_SHEET) And _
        HasWorksheet(wb, EBITDA_SHEET) And _
        HasWorksheet(wb, AUDIT_LOG_SHEET)

End Function


Private Function HasWorksheet(ByVal wb As Workbook, ByVal sheetName As String) As Boolean

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = wb.Worksheets(sheetName)
    On Error GoTo 0

    HasWorksheet = Not ws Is Nothing

End Function


'=========================================================================
' REPORT CREATION
'=========================================================================

Private Sub PrepareReportSheet( _
    ByVal ws As Worksheet, _
    ByVal targetWb As Workbook)

    ws.Cells.Clear

    ws.Range("A1").Value = "Excel Audit Assistant — Candidate Findings"
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16

    ws.Range("A2").Value = "Workbook inspected"
    ws.Range("B2").Value = targetWb.Name

    ws.Range("A3").Value = "Audit status"
    ws.Range("B3").Value = _
        "Read-only review. Candidate findings only. No corrections applied. " & _
        "No EBITDA treatment selected."

    ws.Range("A5").Value = "Finding ID"
    ws.Range("B5").Value = "Worksheet"
    ws.Range("C5").Value = "Cell"
    ws.Range("D5").Value = "Detected role"
    ws.Range("E5").Value = "Rule"
    ws.Range("F5").Value = "Formula or value"
    ws.Range("G5").Value = "Reason for flag"
    ws.Range("H5").Value = "Analyst disposition"

    With ws.Range("A5:H5")
        .Font.Bold = True
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With

End Sub


Private Sub AddFinding( _
    ByVal reportWs As Worksheet, _
    ByRef nextRow As Long, _
    ByRef findingNumber As Long, _
    ByVal sourceSheet As String, _
    ByVal sourceCell As String, _
    ByVal detectedRole As String, _
    ByVal ruleName As String, _
    ByVal formulaOrValue As String, _
    ByVal reason As String)

    findingNumber = findingNumber + 1

    reportWs.Cells(nextRow, 1).Value = "F-" & Format$(findingNumber, "000")
    reportWs.Cells(nextRow, 2).Value = sourceSheet
    reportWs.Cells(nextRow, 3).Value = sourceCell
    reportWs.Cells(nextRow, 4).Value = detectedRole
    reportWs.Cells(nextRow, 5).Value = ruleName
    reportWs.Cells(nextRow, 6).Value = formulaOrValue
    reportWs.Cells(nextRow, 7).Value = reason

    'Required to remain blank for analyst review.
    reportWs.Cells(nextRow, 8).Value = vbNullString

    nextRow = nextRow + 1

End Sub


Private Sub FinaliseReportSheet( _
    ByVal ws As Worksheet, _
    ByVal nextRow As Long)

    Dim lastRow As Long

    lastRow = nextRow - 1
    If lastRow < 5 Then lastRow = 5

    With ws.Range("A5:H" & lastRow)
        .VerticalAlignment = xlTop
        .WrapText = True
        .Borders.LineStyle = xlContinuous
    End With

    ws.Columns("A").ColumnWidth = 11
    ws.Columns("B").ColumnWidth = 25
    ws.Columns("C").ColumnWidth = 12
    ws.Columns("D").ColumnWidth = 22
    ws.Columns("E").ColumnWidth = 29
    ws.Columns("F").ColumnWidth = 38
    ws.Columns("G").ColumnWidth = 72
    ws.Columns("H").ColumnWidth = 28

    ws.Rows("1:5").AutoFit

    ws.Range("A5:H" & lastRow).AutoFilter

    ws.Activate
    ActiveWindow.FreezePanes = False
    ws.Range("A6").Select
    ActiveWindow.FreezePanes = True

End Sub


'=========================================================================
' FORMULA-PRESENCE AND STRUCTURAL-PATTERN AUDIT
'=========================================================================

Private Sub AuditFormulaPatterns( _
    ByVal targetWb As Workbook, _
    ByVal reportWs As Worksheet, _
    ByRef nextRow As Long, _
    ByRef findingNumber As Long)

    Dim ws As Worksheet
    Dim r As Long
    Dim c As Long

    Dim formulaCount As Long
    Dim nonBlankConstantCount As Long
    Dim blankCount As Long

    Dim dominantPattern As String
    Dim dominantCount As Long

    Dim cell As Range
    Dim roleName As String

    For Each ws In targetWb.Worksheets

        'Raw data is the declared input sheet. Constants there are expected.
        If ws.Name <> RAW_SHEET And ws.Name <> AUDIT_LOG_SHEET Then

            For r = 1 To ws.UsedRange.Row + ws.UsedRange.Rows.Count - 1

                If Len(Trim$(CStr(ws.Cells(r, 2).Value))) > 0 Then

                    CountYearCellTypes _
                        ws, r, formulaCount, nonBlankConstantCount, blankCount

                    'A row with at least two formula-driven years is treated as
                    'a comparable formula row for presence checking.
                    If formulaCount >= 2 Then

                        For c = FIRST_YEAR_COL To LAST_YEAR_COL

                            Set cell = ws.Cells(r, c)

                            If Not cell.HasFormula Then

                                If Len(CStr(cell.Value2)) > 0 Then

                                    roleName = DetectCellRole(ws, cell)

                                    AddFinding _
                                        reportWs, nextRow, findingNumber, _
                                        ws.Name, cell.Address(False, False), _
                                        roleName, _
                                        "Constant inside formula-driven year row", _
                                        CellFormulaOrValue(cell), _
                                        "Comparable year cells in this row are predominantly formulas, " & _
                                        "but this cell contains a hardcoded value. Review whether the " & _
                                        "pattern break is intentional."

                                ElseIf formulaCount >= 3 Then

                                    AddFinding _
                                        reportWs, nextRow, findingNumber, _
                                        ws.Name, cell.Address(False, False), _
                                        "Blank in formula-driven row", _
                                        "Formula presence across comparable years", _
                                        "(blank)", _
                                        "At least three comparable year cells in this row contain formulas, " & _
                                        "while this year is blank. This may be legitimate where a prior " & _
                                        "period is unavailable, so analyst review is required."

                                End If

                            End If

                        Next c

                    End If

                    'Compare structural FormulaR1C1 patterns.
                    If formulaCount >= 3 Then

                        dominantPattern = GetDominantR1C1Pattern( _
                            ws, r, dominantCount)

                        If Len(dominantPattern) > 0 Then

                            'Only call something an outlier if one pattern has
                            'a strict majority among formula cells.
                            If dominantCount > formulaCount / 2 Then

                                For c = FIRST_YEAR_COL To LAST_YEAR_COL

                                    Set cell = ws.Cells(r, c)

                                    If cell.HasFormula Then

                                        If CStr(cell.FormulaR1C1) <> dominantPattern Then

                                            AddFinding _
                                                reportWs, nextRow, findingNumber, _
                                                ws.Name, cell.Address(False, False), _
                                                DetectCellRole(ws, cell), _
                                                "Structural formula-pattern inconsistency", _
                                                CStr(cell.Formula), _
                                                "The FormulaR1C1 structure differs from the dominant " & _
                                                "pattern across comparable year cells in the same row. " & _
                                                "Review the references and operators."

                                        End If

                                    End If

                                Next c

                            End If

                        End If

                    End If

                End If

            Next r

        End If

    Next ws

End Sub


Private Sub CountYearCellTypes( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByRef formulaCount As Long, _
    ByRef nonBlankConstantCount As Long, _
    ByRef blankCount As Long)

    Dim c As Long
    Dim cell As Range

    formulaCount = 0
    nonBlankConstantCount = 0
    blankCount = 0

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        Set cell = ws.Cells(rowNumber, c)

        If cell.HasFormula Then
            formulaCount = formulaCount + 1
        ElseIf Len(CStr(cell.Value2)) > 0 Then
            nonBlankConstantCount = nonBlankConstantCount + 1
        Else
            blankCount = blankCount + 1
        End If

    Next c

End Sub


Private Function GetDominantR1C1Pattern( _
    ByVal ws As Worksheet, _
    ByVal rowNumber As Long, _
    ByRef bestCount As Long) As String

    Dim c As Long
    Dim testC As Long

    Dim candidate As String
    Dim comparison As String
    Dim candidateCount As Long

    bestCount = 0
    GetDominantR1C1Pattern = vbNullString

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If ws.Cells(rowNumber, c).HasFormula Then

            candidate = CStr(ws.Cells(rowNumber, c).FormulaR1C1)
            candidateCount = 0

            For testC = FIRST_YEAR_COL To LAST_YEAR_COL

                If ws.Cells(rowNumber, testC).HasFormula Then

                    comparison = CStr(ws.Cells(rowNumber, testC).FormulaR1C1)

                    If comparison = candidate Then
                        candidateCount = candidateCount + 1
                    End If

                End If

            Next testC

            If candidateCount > bestCount Then
                bestCount = candidateCount
                GetDominantR1C1Pattern = candidate
            End If

        End If

    Next c

End Function


'=========================================================================
' FORMULA / REFERENCE ERROR AUDIT
'=========================================================================

Private Sub AuditFormulaAndReferenceErrors( _
    ByVal targetWb As Workbook, _
    ByVal reportWs As Worksheet, _
    ByRef nextRow As Long, _
    ByRef findingNumber As Long)

    Dim ws As Worksheet
    Dim cell As Range
    Dim upperFormula As String

    For Each ws In targetWb.Worksheets

        If ws.Name <> AUDIT_LOG_SHEET Then

            For Each cell In ws.UsedRange.Cells

                If cell.HasFormula Then

                    upperFormula = UCase$(CStr(cell.Formula))

                    If FormulaContainsReferenceError(upperFormula) Then

                        AddFinding _
                            reportWs, nextRow, findingNumber, _
                            ws.Name, cell.Address(False, False), _
                            DetectCellRole(ws, cell), _
                            "Formula/reference error", _
                            CStr(cell.Formula), _
                            "The formula text contains an Excel error token or invalid reference."

                    ElseIf IsError(cell.Value) Then

                        AddFinding _
                            reportWs, nextRow, findingNumber, _
                            ws.Name, cell.Address(False, False), _
                            DetectCellRole(ws, cell), _
                            "Formula evaluation error", _
                            CStr(cell.Formula), _
                            "The formula currently evaluates to an Excel error value."

                    End If

                ElseIf IsError(cell.Value) Then

                    AddFinding _
                        reportWs, nextRow, findingNumber, _
                        ws.Name, cell.Address(False, False), _
                        DetectCellRole(ws, cell), _
                        "Cell error value", _
                        SafeCellText(cell), _
                        "The cell contains an Excel error value."

                End If

            Next cell

        End If

    Next ws

End Sub


Private Function FormulaContainsReferenceError( _
    ByVal upperFormula As String) As Boolean

    FormulaContainsReferenceError = _
        (InStr(1, upperFormula, "#REF!", vbTextCompare) > 0) Or _
        (InStr(1, upperFormula, "#DIV/0!", vbTextCompare) > 0) Or _
        (InStr(1, upperFormula, "#VALUE!", vbTextCompare) > 0) Or _
        (InStr(1, upperFormula, "#NAME?", vbTextCompare) > 0) Or _
        (InStr(1, upperFormula, "#NUM!", vbTextCompare) > 0) Or _
        (InStr(1, upperFormula, "#NULL!", vbTextCompare) > 0)

End Function


'=========================================================================
' FORMULA ROLE / COLOUR METADATA AUDIT
'=========================================================================

Private Sub AuditColourMetadata( _
    ByVal targetWb As Workbook, _
    ByVal reportWs As Worksheet, _
    ByRef nextRow As Long, _
    ByRef findingNumber As Long)

    Dim ws As Worksheet
    Dim cell As Range
    Dim expectedColour As Long
    Dim actualColour As Long
    Dim roleName As String

    For Each ws In targetWb.Worksheets

        If ws.Name <> AUDIT_LOG_SHEET Then

            For Each cell In ws.UsedRange.Cells

                roleName = DetectCellRole(ws, cell)

                If roleName = "Declared input" Then

                    expectedColour = RGB(0, 0, 255)
                    actualColour = cell.Font.Color

                    If actualColour <> expectedColour Then

                        AddFinding _
                            reportWs, nextRow, findingNumber, _
                            ws.Name, cell.Address(False, False), _
                            roleName, _
                            "Input-role versus colour metadata", _
                            CellFormulaOrValue(cell), _
                            "This cell is in the declared raw-input area but is not shown " & _
                            "in the workbook's stated blue-input font colour."

                    End If

                ElseIf roleName = "Direct cross-sheet link" Then

                    expectedColour = RGB(0, 122, 51)
                    actualColour = cell.Font.Color

                    If actualColour <> expectedColour Then

                        AddFinding _
                            reportWs, nextRow, findingNumber, _
                            ws.Name, cell.Address(False, False), _
                            roleName, _
                            "Formula-role versus colour metadata", _
                            CellFormulaOrValue(cell), _
                            "This is a direct cross-sheet link but its font colour is not " & _
                            "the workbook's stated green link colour."

                    End If

                ElseIf roleName = "Same-sheet calculation" Then

                    expectedColour = RGB(0, 0, 0)
                    actualColour = cell.Font.Color

                    If actualColour <> expectedColour Then

                        AddFinding _
                            reportWs, nextRow, findingNumber, _
                            ws.Name, cell.Address(False, False), _
                            roleName, _
                            "Formula-role versus colour metadata", _
                            CellFormulaOrValue(cell), _
                            "This is a calculation rather than a direct cross-sheet link, " & _
                            "but its font colour is not the workbook's stated black " & _
                            "calculation colour."

                    End If

                End If

            Next cell

        End If

    Next ws

End Sub


Private Function DetectCellRole( _
    ByVal ws As Worksheet, _
    ByVal cell As Range) As String

    If ws.Name = RAW_SHEET Then

        If cell.Column >= FIRST_YEAR_COL And _
           cell.Column <= LAST_YEAR_COL And _
           Len(CStr(cell.Value2)) > 0 And _
           Not cell.HasFormula Then

            DetectCellRole = "Declared input"
            Exit Function

        End If

    End If

    If cell.HasFormula Then

        If IsDirectCrossSheetLink(CStr(cell.Formula)) Then
            DetectCellRole = "Direct cross-sheet link"
        Else
            DetectCellRole = "Same-sheet calculation"
        End If

    ElseIf Len(CStr(cell.Value2)) = 0 Then

        DetectCellRole = "Blank"

    ElseIf cell.Column >= FIRST_YEAR_COL And _
           cell.Column <= LAST_YEAR_COL Then

        DetectCellRole = "Hardcoded value"

    Else

        DetectCellRole = "Label / metadata"

    End If

End Function


Private Function IsDirectCrossSheetLink(ByVal formulaText As String) As Boolean

    Dim body As String

    body = formulaText

    If Left$(body, 1) = "=" Then
        body = Mid$(body, 2)
    End If

    If InStr(1, body, "!", vbTextCompare) = 0 Then
        IsDirectCrossSheetLink = False
        Exit Function
    End If

    'A pure direct link contains a worksheet reference but no calculation
    'operator. This deliberately treats mixed cross-sheet formulas as
    'calculations rather than simple links.
    If InStr(1, body, "+", vbBinaryCompare) > 0 Or _
       InStr(1, body, "-", vbBinaryCompare) > 0 Or _
       InStr(1, body, "*", vbBinaryCompare) > 0 Or _
       InStr(1, body, "/", vbBinaryCompare) > 0 Or _
       InStr(1, body, "^", vbBinaryCompare) > 0 Or _
       InStr(1, body, "(", vbBinaryCompare) > 0 Or _
       InStr(1, body, ")", vbBinaryCompare) > 0 Then

        IsDirectCrossSheetLink = False

    Else

        IsDirectCrossSheetLink = True

    End If

End Function


'=========================================================================
' INDEPENDENT FINANCE CONTROLS
'=========================================================================

Private Sub AuditFinanceControls( _
    ByVal targetWb As Workbook, _
    ByVal reportWs As Worksheet, _
    ByRef nextRow As Long, _
    ByRef findingNumber As Long)

    Dim rawWs As Worksheet
    Dim histWs As Worksheet
    Dim summaryWs As Worksheet

    Dim c As Long

    Dim revenueCurrent As Double
    Dim revenuePrior As Double
    Dim reportedEBITDA As Double
    Dim depreciation As Double

    Dim receivables As Double
    Dim inventory As Double
    Dim payables As Double
    Dim operatingWC As Double
    Dim priorOperatingWC As Double
    Dim changeOperatingWC As Double

    Dim debt As Double
    Dim cash As Double
    Dim netDebt As Double

    Dim capex As Double
    Dim taxRate As Double

    Dim expectedValue As Double

    Set rawWs = targetWb.Worksheets(RAW_SHEET)
    Set histWs = targetWb.Worksheets(HIST_SHEET)
    Set summaryWs = targetWb.Worksheets(SUMMARY_SHEET)


    '---------------------------------------------------------------------
    ' 1. Revenue growth calibration
    ' FY2022A is intentionally excluded because FY2021A revenue is absent.
    '---------------------------------------------------------------------

    For c = 5 To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(7, c), revenueCurrent) And _
           TryGetNumber(rawWs.Cells(7, c - 1), revenuePrior) Then

            If revenuePrior <> 0 Then

                expectedValue = revenueCurrent / revenuePrior - 1

                CompareExpectedToCell _
                    reportWs, nextRow, findingNumber, _
                    histWs.Cells(7, c), _
                    "Revenue growth", _
                    expectedValue, _
                    "Revenue growth should equal current-period raw revenue divided " & _
                    "by immediately preceding raw revenue, less one."

            End If

        End If

    Next c


    '---------------------------------------------------------------------
    ' 2. Same-period reported EBITDA margin
    '---------------------------------------------------------------------

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(7, c), revenueCurrent) And _
           TryGetNumber(rawWs.Cells(8, c), reportedEBITDA) Then

            If revenueCurrent <> 0 Then

                expectedValue = reportedEBITDA / revenueCurrent

                CompareExpectedToCell _
                    reportWs, nextRow, findingNumber, _
                    histWs.Cells(9, c), _
                    "Same-period reported EBITDA margin", _
                    expectedValue, _
                    "Reported EBITDA margin should use reported EBITDA and revenue " & _
                    "from the same period."

            End If

        End If

    Next c


    '---------------------------------------------------------------------
    ' 3. EBIT = reported EBITDA - D&A
    '---------------------------------------------------------------------

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(8, c), reportedEBITDA) And _
           TryGetNumber(rawWs.Cells(9, c), depreciation) Then

            expectedValue = reportedEBITDA - depreciation

            CompareExpectedToCell _
                reportWs, nextRow, findingNumber, _
                histWs.Cells(11, c), _
                "EBIT bridge", _
                expectedValue, _
                "EBIT should equal reported EBITDA less depreciation and amortisation."

        End If

    Next c


    '---------------------------------------------------------------------
    ' 4. Operating working capital
    '    receivables + inventory - trade payables
    '---------------------------------------------------------------------

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(18, c), receivables) And _
           TryGetNumber(rawWs.Cells(19, c), inventory) And _
           TryGetNumber(rawWs.Cells(20, c), payables) Then

            expectedValue = receivables + inventory - payables

            CompareExpectedToCell _
                reportWs, nextRow, findingNumber, _
                histWs.Cells(24, c), _
                "Operating working capital", _
                expectedValue, _
                "Operating working capital should equal trade receivables plus " & _
                "inventory less trade payables."

        End If

    Next c


    '---------------------------------------------------------------------
    ' 5. Net debt = gross interest-bearing debt - unrestricted cash
    '---------------------------------------------------------------------

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(25, c), debt) And _
           TryGetNumber(rawWs.Cells(26, c), cash) Then

            expectedValue = debt - cash

            CompareExpectedToCell _
                reportWs, nextRow, findingNumber, _
                summaryWs.Cells(24, c), _
                "Net debt", _
                expectedValue, _
                "Net debt should equal gross interest-bearing debt less eligible " & _
                "unrestricted cash."

        End If

    Next c


    '---------------------------------------------------------------------
    ' 6. Net debt / reported EBITDA
    '
    ' This does NOT make an EBITDA judgement. The workbook explicitly names
    ' this metric as net debt / REPORTED EBITDA, so this check uses the
    ' named basis and leaves the management-adjusted treatment untouched.
    '---------------------------------------------------------------------

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(25, c), debt) And _
           TryGetNumber(rawWs.Cells(26, c), cash) And _
           TryGetNumber(rawWs.Cells(8, c), reportedEBITDA) Then

            netDebt = debt - cash

            If reportedEBITDA <> 0 Then

                expectedValue = netDebt / reportedEBITDA

                CompareExpectedToCell _
                    reportWs, nextRow, findingNumber, _
                    summaryWs.Cells(25, c), _
                    "Net debt-to-EBITDA — named reported basis", _
                    expectedValue, _
                    "The workbook names this metric as net debt / reported EBITDA. " & _
                    "The independent check therefore uses raw net debt and raw " & _
                    "reported EBITDA. It does not select or endorse an adjustment."

            End If

        End If

    Next c


    '---------------------------------------------------------------------
    ' 7. Simplified historical FCF control
    '
    ' FCF =
    ' EBIT x (1 - stated tax rate)
    ' + D&A
    ' - CapEx
    ' - increase in operating working capital
    '
    ' FY2022A uses the supplied movement because FY2021A balance-sheet
    ' working capital was not supplied.
    '---------------------------------------------------------------------

    For c = FIRST_YEAR_COL To LAST_YEAR_COL

        If TryGetNumber(rawWs.Cells(8, c), reportedEBITDA) And _
           TryGetNumber(rawWs.Cells(9, c), depreciation) And _
           TryGetNumber(rawWs.Cells(13, c), capex) And _
           TryGetNumber(rawWs.Cells(14, c), taxRate) And _
           TryGetNumber(rawWs.Cells(18, c), receivables) And _
           TryGetNumber(rawWs.Cells(19, c), inventory) And _
           TryGetNumber(rawWs.Cells(20, c), payables) Then

            operatingWC = receivables + inventory - payables

            If c = FIRST_YEAR_COL Then

                If TryGetNumber(rawWs.Cells(21, c), changeOperatingWC) Then

                    expectedValue = _
                        (reportedEBITDA - depreciation) * (1 - taxRate) + _
                        depreciation - _
                        capex - _
                        changeOperatingWC

                    CompareExpectedToCell _
                        reportWs, nextRow, findingNumber, _
                        histWs.Cells(36, c), _
                        "Simplified free-cash-flow control", _
                        expectedValue, _
                        "FY2022A uses the supplied change in operating working " & _
                        "capital because FY2021A working-capital balances are not supplied."

                End If

            Else

                If TryGetNumber(rawWs.Cells(18, c - 1), receivables) And _
                   TryGetNumber(rawWs.Cells(19, c - 1), inventory) And _
                   TryGetNumber(rawWs.Cells(20, c - 1), payables) Then

                    priorOperatingWC = receivables + inventory - payables

                    'Re-read current-period values because the variables above
                    'were reused to construct prior-period working capital.
                    Call TryGetNumber(rawWs.Cells(18, c), receivables)
                    Call TryGetNumber(rawWs.Cells(19, c), inventory)
                    Call TryGetNumber(rawWs.Cells(20, c), payables)

                    operatingWC = receivables + inventory - payables
                    changeOperatingWC = operatingWC - priorOperatingWC

                    expectedValue = _
                        (reportedEBITDA - depreciation) * (1 - taxRate) + _
                        depreciation - _
                        capex - _
                        changeOperatingWC

                    CompareExpectedToCell _
                        reportWs, nextRow, findingNumber, _
                        histWs.Cells(36, c), _
                        "Simplified free-cash-flow control", _
                        expectedValue, _
                        "The simplified control should equal tax-effected EBIT plus " & _
                        "D&A less CapEx less the increase in operating working capital."

                End If

            End If

        End If

    Next c

End Sub


Private Sub CompareExpectedToCell( _
    ByVal reportWs As Worksheet, _
    ByRef nextRow As Long, _
    ByRef findingNumber As Long, _
    ByVal testedCell As Range, _
    ByVal ruleName As String, _
    ByVal expectedValue As Double, _
    ByVal ruleExplanation As String)

    Dim actualValue As Double
    Dim difference As Double
    Dim tolerance As Double

    If Not TryGetNumber(testedCell, actualValue) Then

        AddFinding _
            reportWs, nextRow, findingNumber, _
            testedCell.Worksheet.Name, _
            testedCell.Address(False, False), _
            DetectCellRole(testedCell.Worksheet, testedCell), _
            ruleName, _
            CellFormulaOrValue(testedCell), _
            ruleExplanation & _
            " The displayed cell does not currently contain a usable numeric result."

        Exit Sub

    End If

    difference = Abs(actualValue - expectedValue)

    tolerance = Abs(expectedValue) * 0.00000001
    If tolerance < 0.00000001 Then tolerance = 0.00000001

    If difference > tolerance Then

        AddFinding _
            reportWs, nextRow, findingNumber, _
            testedCell.Worksheet.Name, _
            testedCell.Address(False, False), _
            DetectCellRole(testedCell.Worksheet, testedCell), _
            ruleName, _
            CellFormulaOrValue(testedCell), _
            ruleExplanation & _
            " Independent recomputation gives " & _
            FormatNumberForReport(expectedValue) & _
            ", while the workbook currently gives " & _
            FormatNumberForReport(actualValue) & _
            ". Candidate issue only; no correction has been applied."

    End If

End Sub


'=========================================================================
' VALUE / DISPLAY HELPERS
'=========================================================================

Private Function TryGetNumber( _
    ByVal cell As Range, _
    ByRef result As Double) As Boolean

    If IsError(cell.Value) Then
        TryGetNumber = False
        Exit Function
    End If

    If IsNumeric(cell.Value2) And Len(CStr(cell.Value2)) > 0 Then
        result = CDbl(cell.Value2)
        TryGetNumber = True
    Else
        TryGetNumber = False
    End If

End Function


Private Function CellFormulaOrValue(ByVal cell As Range) As String

    If cell.HasFormula Then
        CellFormulaOrValue = CStr(cell.Formula)
    ElseIf IsError(cell.Value) Then
        CellFormulaOrValue = SafeCellText(cell)
    ElseIf Len(CStr(cell.Value2)) = 0 Then
        CellFormulaOrValue = "(blank)"
    Else
        CellFormulaOrValue = CStr(cell.Value2)
    End If

End Function


Private Function SafeCellText(ByVal cell As Range) As String

    On Error Resume Next
    SafeCellText = CStr(cell.Text)

    If Err.Number <> 0 Then
        Err.Clear
        SafeCellText = "(Excel error value)"
    End If

    On Error GoTo 0

End Function


Private Function FormatNumberForReport(ByVal value As Double) As String

    FormatNumberForReport = Format$(value, "0.000000")

End Function
```

## (2) Macro to run

**`AuditOpenTargetCo`**

The implementation follows the specification's read-only design: TargetCo is inspected, findings are written separately, analyst disposition remains for the analyst, and no correction or EBITDA-treatment decision is made.

## (3) What it reads and writes

| Workbook | What the macro does |
|---|---|
| **TargetCo workbook** | **Reads only.** It inspects cells, formulas, `FormulaR1C1`, calculated values, and font-colour metadata. |
| **Your `.xlsm` controller** | Contains the VBA module. The macro does not use it as the audit target. |
| **New report workbook** | **Writes all candidate findings here.** It creates a new workbook with an `Audit Findings` sheet. |

There is no `Save`, `SaveAs`, `Close`, `Delete`, `Shell`, network/API call, or write operation against TargetCo in the module. That implements the specification's safety restrictions.

## (4) Exact desktop Excel installation instructions

1. Open **TargetCo** in desktop Excel and leave it open.
2. In Excel, create a **second, completely blank workbook**.
3. Save that second workbook as an **Excel Macro-Enabled Workbook (`.xlsm`)**. For example, `Excel_Audit_Assistant.xlsm`.
4. Keep **both workbooks open**.
5. Open the VBA editor:
   - **Windows:** `Developer` → `Visual Basic`.
   - **Mac:** `Tools` → `Macro` → `Visual Basic Editor`; alternatively use `Developer` → `Visual Basic` if the Developer tab is enabled.
6. In the VBA editor, look at the **Project Explorer** on the left. Select the project corresponding to your new `.xlsm` controller workbook—not TargetCo.
7. Choose **Insert → Module**.
8. A new standard module such as `Module1` will appear.
9. Copy the **entire VBA code block above**, from `Option Explicit` through the final `End Function`.
10. Paste it into that module.
11. Do **not** paste anything into `ThisWorkbook`, a worksheet object, or the TargetCo VBA project.
12. Return to Excel while keeping TargetCo open.
13. Open the Macro dialog:
    - **Windows:** `Developer` → `Macros`, or press `Alt+F8`.
    - **Mac:** `Tools` → `Macro` → `Macros`.
14. Select **`AuditOpenTargetCo`**.
15. Click **Run**.
16. The macro should create a **third workbook**, containing a worksheet named **`Audit Findings`**.
17. Review that report. The `Analyst disposition` column is deliberately blank for you to complete.
18. Do **not** interpret every flag as a confirmed error. The specification explicitly says candidate flags require analyst attention and independent checking.
19. You may save the **report workbook** yourself if you want to retain it. The macro itself does not save anything.
20. Verify that TargetCo remains open and unchanged. The specification likewise instructs you to confirm that the separate report was created without editing or saving TargetCo.
