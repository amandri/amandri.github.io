Attribute VB_Name = "TargetCo_AuditAssistant"
Option Explicit

Private Const TARGET_NAME As String = "TargetCo_historical_model.xlsx"

Public Sub AuditOpenTargetCo()
    Dim targetWb As Workbook
    Dim reportWb As Workbook
    Dim reportWs As Worksheet
    Dim nextRow As Long

    Set targetWb = FindOpenWorkbook(TARGET_NAME)
    If targetWb Is Nothing Then
        MsgBox "Open " & TARGET_NAME & " before running the audit.", vbExclamation
        Exit Sub
    End If

    Set reportWb = Application.Workbooks.Add(xlWBATWorksheet)
    Set reportWs = reportWb.Worksheets(1)
    reportWs.Name = "Audit Report"
    PrepareReport reportWs
    nextRow = 2

    AddFinding reportWs, nextRow, "01_Raw_Data", "D:G", "Declared input area", _
        "DECLARED_INPUT_AREA", "", "Expected hardcoded source area. Excluded from constant-in-formula-row flags.", "Context"

    ScanComparableYearCells targetWb, reportWs, nextRow
    CheckTargetCoFinanceRules targetWb, reportWs, nextRow

    reportWs.Columns("A:H").EntireColumn.AutoFit
    reportWs.Columns("F:G").ColumnWidth = 42
    reportWs.Columns("F:G").WrapText = True
    reportWs.Activate
    MsgBox CStr(nextRow - 2) & " report records created. Review each candidate in Excel.", vbInformation
End Sub

Private Function FindOpenWorkbook(ByVal requiredName As String) As Workbook
    Dim wb As Workbook
    For Each wb In Application.Workbooks
        If StrComp(wb.Name, requiredName, vbTextCompare) = 0 Then
            Set FindOpenWorkbook = wb
            Exit Function
        End If
    Next wb
End Function

Private Sub PrepareReport(ByVal reportWs As Worksheet)
    Dim headers As Variant
    Dim index As Long
    headers = Array("Worksheet", "Cell", "Detected role", "Rule", "Formula or value", "Reason for flag", "Analyst disposition", "Review status")
    For index = LBound(headers) To UBound(headers)
        reportWs.Cells(1, index + 1).Value = headers(index)
    Next index
    reportWs.Rows(1).Font.Bold = True
    reportWs.Rows(1).Interior.Color = RGB(31, 56, 100)
    reportWs.Rows(1).Font.Color = RGB(255, 255, 255)
End Sub

Private Sub ScanComparableYearCells(ByVal targetWb As Workbook, ByVal reportWs As Worksheet, ByRef nextRow As Long)
    Dim sourceWs As Worksheet
    Dim rowNumber As Long
    Dim columnNumber As Long
    Dim formulaCount As Long
    Dim nonBlankCount As Long
    Dim targetCell As Range
    Dim majorityPattern As String

    For Each sourceWs In targetWb.Worksheets
        If sourceWs.Name <> "06_Audit_Log" Then
            For rowNumber = 1 To sourceWs.UsedRange.Rows(sourceWs.UsedRange.Rows.Count).Row
                formulaCount = 0
                nonBlankCount = 0
                For columnNumber = 4 To 7
                    Set targetCell = sourceWs.Cells(rowNumber, columnNumber)
                    If Len(CStr(targetCell.Value2)) > 0 Or targetCell.HasFormula Then nonBlankCount = nonBlankCount + 1
                    If targetCell.HasFormula Then formulaCount = formulaCount + 1
                Next columnNumber

                If formulaCount >= 2 Then
                    majorityPattern = MostCommonR1C1(sourceWs, rowNumber)
                    For columnNumber = 4 To 7
                        Set targetCell = sourceWs.Cells(rowNumber, columnNumber)
                        If IsDeclaredInput(sourceWs.Name, rowNumber, columnNumber) Then
                            ' Declared inputs are not calculation anomalies.
                        ElseIf targetCell.HasFormula Then
                            If FormulaHasError(targetCell) Then
                                AddFinding reportWs, nextRow, sourceWs.Name, targetCell.Address(False, False), _
                                    "Formula", "FORMULA_OR_REFERENCE_ERROR", targetCell.Formula, _
                                    "The formula or its calculated value contains an Excel error."
                            End If
                            If Len(majorityPattern) > 0 And targetCell.FormulaR1C1 <> majorityPattern Then
                                AddFinding reportWs, nextRow, sourceWs.Name, targetCell.Address(False, False), _
                                    "Formula", "FORMULA_R1C1_DEVIATION", targetCell.Formula, _
                                    "FormulaR1C1 differs from the most common pattern in comparable year cells."
                            End If
                            If targetCell.Font.ColorIndex = 5 Then
                                AddFinding reportWs, nextRow, sourceWs.Name, targetCell.Address(False, False), _
                                    "Formula", "ROLE_COLOUR_MISMATCH", targetCell.Formula, _
                                    "A formula uses the workbook's blue input colour metadata."
                            End If
                        ElseIf Len(CStr(targetCell.Value2)) > 0 And nonBlankCount >= 3 Then
                            AddFinding reportWs, nextRow, sourceWs.Name, targetCell.Address(False, False), _
                                "Constant in calculation row", "CONSTANT_IN_FORMULA_ROW", CStr(targetCell.Value2), _
                                "A constant appears inside a row otherwise driven by formulas."
                        End If
                    Next columnNumber
                End If
            Next rowNumber
        End If
    Next sourceWs
End Sub

Private Function MostCommonR1C1(ByVal sourceWs As Worksheet, ByVal rowNumber As Long) As String
    Dim patterns(1 To 4) As String
    Dim counts(1 To 4) As Long
    Dim leftIndex As Long
    Dim rightIndex As Long
    Dim bestIndex As Long
    Dim targetCell As Range

    For leftIndex = 1 To 4
        Set targetCell = sourceWs.Cells(rowNumber, leftIndex + 3)
        If targetCell.HasFormula Then patterns(leftIndex) = targetCell.FormulaR1C1
    Next leftIndex

    For leftIndex = 1 To 4
        If Len(patterns(leftIndex)) > 0 Then
            For rightIndex = 1 To 4
                If patterns(rightIndex) = patterns(leftIndex) Then counts(leftIndex) = counts(leftIndex) + 1
            Next rightIndex
            If bestIndex = 0 Then
                bestIndex = leftIndex
            ElseIf counts(leftIndex) > counts(bestIndex) Then
                bestIndex = leftIndex
            End If
        End If
    Next leftIndex

    If bestIndex > 0 And counts(bestIndex) >= 2 Then MostCommonR1C1 = patterns(bestIndex)
End Function

Private Function FormulaHasError(ByVal targetCell As Range) As Boolean
    FormulaHasError = InStr(1, targetCell.Formula, "#REF!", vbTextCompare) > 0 _
        Or InStr(1, targetCell.Formula, "#NAME?", vbTextCompare) > 0 _
        Or IsError(targetCell.Value)
End Function

Private Function IsDeclaredInput(ByVal sheetName As String, ByVal rowNumber As Long, ByVal columnNumber As Long) As Boolean
    If sheetName = "01_Raw_Data" And columnNumber >= 4 And columnNumber <= 7 Then
        Select Case rowNumber
            Case 7, 8, 9, 13, 14, 18, 19, 20, 21, 25, 26, 30, 34
                IsDeclaredInput = True
                Exit Function
        End Select
    End If

    If sheetName = "03_Adjustments" And columnNumber >= 6 And columnNumber <= 7 Then
        If rowNumber = 16 Or rowNumber = 17 Or rowNumber = 23 Or rowNumber = 24 Then IsDeclaredInput = True
    End If
End Function

Private Sub CheckTargetCoFinanceRules(ByVal targetWb As Workbook, ByVal reportWs As Worksheet, ByRef nextRow As Long)
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "F7", "=F6/E6-1", "GROWTH_PRECEDING_PERIOD", "Growth must use the immediately preceding period."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "G9", "=G8/G6", "MARGIN_MATCHING_PERIOD", "Reported EBITDA margin must use same-period revenue."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "F11", "=F8-F10", "EBIT_DEPENDENCY", "EBIT must retain the dependency on reported EBITDA and D&A."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "G24", "=G21+G22-G23", "OPERATING_NWC_DEFINITION", "Operating NWC must add receivables and inventory and deduct payables."
    CheckExpectedFormula targetWb, reportWs, nextRow, "04_Summary", "G24", "=G22-G23", "NET_DEBT_DEFINITION", "Net debt must deduct eligible unrestricted cash from gross debt."
    CheckExpectedFormula targetWb, reportWs, nextRow, "04_Summary", "D25", "=D24/D16", "NAMED_EBITDA_DENOMINATOR", "The row label names reported EBITDA, which is held on row 16."
    CheckExpectedFormula targetWb, reportWs, nextRow, "04_Summary", "E25", "=E24/E16", "NAMED_EBITDA_DENOMINATOR", "The row label names reported EBITDA, which is held on row 16."
    CheckExpectedFormula targetWb, reportWs, nextRow, "04_Summary", "F25", "=F24/F16", "NAMED_EBITDA_DENOMINATOR", "The row label names reported EBITDA, which is held on row 16."
    CheckExpectedFormula targetWb, reportWs, nextRow, "04_Summary", "G25", "=G24/G16", "NAMED_EBITDA_DENOMINATOR", "The row label names reported EBITDA, which is held on row 16."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "D36", "=SUM(D32:D35)", "HISTORICAL_FCF_CONTROL", "Historical FCF must reconcile NOPAT, D&A, CapEx and the change in operating NWC."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "E36", "=SUM(E32:E35)", "HISTORICAL_FCF_CONTROL", "Historical FCF must reconcile NOPAT, D&A, CapEx and the change in operating NWC."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "F36", "=SUM(F32:F35)", "HISTORICAL_FCF_CONTROL", "Historical FCF must reconcile NOPAT, D&A, CapEx and the change in operating NWC."
    CheckExpectedFormula targetWb, reportWs, nextRow, "02_Historical_Analysis", "G36", "=SUM(G32:G35)", "HISTORICAL_FCF_CONTROL", "Historical FCF must reconcile NOPAT, D&A, CapEx and the change in operating NWC."
End Sub

Private Sub CheckExpectedFormula(ByVal targetWb As Workbook, ByVal reportWs As Worksheet, ByRef nextRow As Long, _
    ByVal sheetName As String, ByVal cellAddress As String, ByVal expectedFormula As String, _
    ByVal ruleName As String, ByVal reason As String)

    Dim targetCell As Range
    Set targetCell = targetWb.Worksheets(sheetName).Range(cellAddress)
    If Not targetCell.HasFormula Or StrComp(targetCell.Formula, expectedFormula, vbTextCompare) <> 0 Then
        AddFinding reportWs, nextRow, sheetName, cellAddress, IIf(targetCell.HasFormula, "Formula", "Constant"), _
            ruleName, FormulaOrValue(targetCell), reason
    End If
End Sub

Private Function FormulaOrValue(ByVal targetCell As Range) As String
    If targetCell.HasFormula Then
        FormulaOrValue = targetCell.Formula
    Else
        FormulaOrValue = CStr(targetCell.Value2)
    End If
End Function

Private Sub AddFinding(ByVal reportWs As Worksheet, ByRef nextRow As Long, ByVal sheetName As String, _
    ByVal cellAddress As String, ByVal detectedRole As String, ByVal ruleName As String, _
    ByVal formulaOrValueText As String, ByVal reason As String, _
    Optional ByVal reviewStatus As String = "Candidate")

    reportWs.Cells(nextRow, 1).Value = sheetName
    reportWs.Cells(nextRow, 2).Value = cellAddress
    reportWs.Cells(nextRow, 3).Value = detectedRole
    reportWs.Cells(nextRow, 4).Value = ruleName
    reportWs.Cells(nextRow, 5).NumberFormat = "@"
    reportWs.Cells(nextRow, 5).Value = formulaOrValueText
    reportWs.Cells(nextRow, 6).Value = reason
    reportWs.Cells(nextRow, 7).Value = ""
    reportWs.Cells(nextRow, 8).Value = reviewStatus
    nextRow = nextRow + 1
End Sub
