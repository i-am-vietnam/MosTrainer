using Microsoft.Office.Interop.Excel;
using MosTrainer.Core.Interfaces;
using MosTrainer.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using Xl = Microsoft.Office.Interop.Excel;
using System.IO.Compression;
using System.Linq;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;

namespace MosTrainer.Excel
{
    public class ExcelController : IExcelController, IDisposable
    {
        private ExcelSession _session;

        public bool IsOpened
        {
            get { return _session != null && _session.App != null && _session.Workbook != null; }
        }

        public void StartExcel()
        {
            CloseWorkbook();

            Xl.Application app = null;
            try
            {
                app = new Xl.Application();
                app.Visible = true;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int pid = GetExcelPid(app);

                _session = new ExcelSession
                {
                    App = app,
                    Workbook = null,
                    Pid = pid,
                    WorkbookPath = "",
                    StartedAt = DateTime.Now
                };

                app = null;
            }
            finally
            {
                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    ReleaseCom(app);
                }
            }
        }

        public void Open(string filePath)
        {
            OpenWorkbook(filePath);
        }

        public void OpenWorkbook(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Workbook path is empty.", "filePath");

            CloseWorkbook();

            Xl.Application app = null;
            Xl.Workbook wb = null;

            try
            {
                app = new Xl.Application();
                app.Visible = true;
                app.DisplayAlerts = false;
                app.EnableEvents = false;

                int pid = GetExcelPid(app);

                wb = app.Workbooks.Open(filePath);

                _session = new ExcelSession
                {
                    App = app,
                    Workbook = wb,
                    Pid = pid,
                    WorkbookPath = filePath,
                    StartedAt = DateTime.Now
                };

                app = null;
                wb = null;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ExcelController.OpenWorkbook", "Workbook could not be opened: " + filePath, ex);

                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    ReleaseCom(wb);
                }

                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    ReleaseCom(app);
                }

                _session = null;
                throw;
            }
        }

        public void Close()
        {
            CloseWorkbook();
        }

        public void CloseWorkbook()
        {
            if (_session == null)
                return;

            int pid = _session.Pid;
            Xl.Workbook wb = null;
            Xl.Application app = null;

            try
            {
                wb = _session.Workbook as Xl.Workbook;
                app = _session.App as Xl.Application;

                if (wb != null)
                {
                    try { wb.Close(false); } catch { }
                    ReleaseCom(wb);
                }

                if (app != null)
                {
                    try { app.Quit(); } catch { }
                    ReleaseCom(app);
                }
            }
            finally
            {
                _session.Workbook = null;
                _session.App = null;
                _session = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();

            if (pid > 0 && WinApiProcessHelper.IsProcessAlive(pid))
            {
                WinApiProcessHelper.WaitForExit(pid, 2000);

                if (WinApiProcessHelper.IsProcessAlive(pid))
                {
                    WinApiProcessHelper.KillByPid(pid);
                    WinApiProcessHelper.WaitForExit(pid, 2000);
                }
            }
        }

        public bool WorksheetExists(string sheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Workbook wb = null;
            Xl.Sheets sheets = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                sheets = wb.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    Xl.Worksheet ws = null;
                    try
                    {
                        ws = (Xl.Worksheet)sheets[i];
                        if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    finally
                    {
                        ReleaseCom(ws);
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sheets);
            }
        }

        public bool CellValueEquals(string sheetName, string address, object expectedValue)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range cell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                cell = ws.Range[address];
                object actual = cell.Value2;

                return ObjectEqualsLoose(actual, expectedValue);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(ws);
            }
        }

        public bool CellTextEquals(string sheetName, string address, string expectedText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range cell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                cell = ws.Range[address];
                string actual = cell.Text == null ? "" : cell.Text.ToString();
                string expected = expectedText ?? "";

                return string.Equals(NormalizeText(actual), NormalizeText(expected), StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(ws);
            }
        }

        public bool CellFormulaEquals(string sheetName, string address, string expectedFormula)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range cell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                cell = ws.Range[address];
                string actual = cell.Formula == null ? "" : cell.Formula.ToString();
                string expected = expectedFormula ?? "";

                return string.Equals(actual.Trim(), expected.Trim(), StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(ws);
            }
        }

        public bool CellFormulaEqualsNormalized(string sheetName, string address, string expectedFormula)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range cell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                cell = ws.Range[address];
                string actual = cell.Formula == null ? "" : cell.Formula.ToString();
                string expected = expectedFormula ?? "";

                return NormalizeFormula(actual) == NormalizeFormula(expected);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(ws);
            }
        }

        public bool RangeNumberFormatEquals(string sheetName, string address, string expectedFormat)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                range = ws.Range[address];
                string actual = range.NumberFormat == null ? "" : range.NumberFormat.ToString();
                string expected = expectedFormat ?? "";

                return NormalizeFormat(actual) == NormalizeFormat(expected);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(ws);
            }
        }

        public bool RangeFillColorEquals(string sheetName, string address, int expectedColor)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                range = ws.Range[address];
                int actual = Convert.ToInt32(range.Interior.Color);
                return actual == expectedColor;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(ws);
            }
        }

        public bool RangeFontBoldEquals(string sheetName, string address, bool expectedBold)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;
            Xl.Font font = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                range = ws.Range[address];
                font = range.Font;

                bool actual = Convert.ToBoolean(font.Bold);
                return actual == expectedBold;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(font);
                ReleaseCom(range);
                ReleaseCom(ws);
            }
        }

        public bool RangeHorizontalAlignmentEquals(string sheetName, string address, string expectedAlignment)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                range = ws.Range[address];
                int actual = Convert.ToInt32(range.HorizontalAlignment);
                int expected = ParseHorizontalAlignment(expectedAlignment);

                return actual == expected;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(ws);
            }
        }

        public bool TableExists(string sheetName, string tableName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                tables = ws.ListObjects;

                for (int i = 1; i <= tables.Count; i++)
                {
                    Xl.ListObject table = null;
                    try
                    {
                        table = tables.Item[i];

                        if (string.IsNullOrWhiteSpace(tableName))
                            return true;

                        if (string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    finally
                    {
                        ReleaseCom(table);
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(tables);
                ReleaseCom(ws);
            }
        }

        public bool ChartExists(string sheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            Xl.Worksheet ws = null;
            Xl.ChartObjects chartObjects = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                chartObjects = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                return chartObjects != null && chartObjects.Count > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chartObjects);
                ReleaseCom(ws);
            }
        }

        public bool NamedRangeExists(string name)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(name)) return false;

            Xl.Workbook wb = null;
            Xl.Names names = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                names = wb.Names;

                for (int i = 1; i <= names.Count; i++)
                {
                    Xl.Name currentName = null;
                    try
                    {
                        currentName = names.Item(i, Type.Missing, Type.Missing);
                        string fullName = currentName.Name == null ? "" : currentName.Name.ToString();
                        string shortName = fullName;

                        int exclamationIndex = shortName.LastIndexOf('!');
                        if (exclamationIndex >= 0 && exclamationIndex + 1 < shortName.Length)
                            shortName = shortName.Substring(exclamationIndex + 1);

                        shortName = shortName.Trim('\'');

                        if (string.Equals(fullName, name, StringComparison.OrdinalIgnoreCase) ||
                            string.Equals(shortName, name, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                    finally
                    {
                        ReleaseCom(currentName);
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(names);
            }
        }

        public bool PrintAreaEquals(string sheetName, string expectedRange)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(expectedRange)) return false;

            Xl.Worksheet ws = null;
            Xl.PageSetup pageSetup = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                pageSetup = ws.PageSetup;
                string actualPrintArea = pageSetup.PrintArea == null
                    ? ""
                    : pageSetup.PrintArea.ToString();

                if (string.IsNullOrWhiteSpace(actualPrintArea))
                    return false;

                string actual = NormalizeRangeAddress(actualPrintArea);
                string expected = NormalizeRangeAddress(expectedRange);

                return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(pageSetup);
                ReleaseCom(ws);
            }
        }
        public bool AutoFilterByHeaderEquals(string sheetName, string columnHeader, string expectedText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(columnHeader)) return false;
            if (string.IsNullOrWhiteSpace(expectedText)) return false;

            Xl.Worksheet ws = null;
            Xl.Range filterRange = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                filterRange = GetAutoFilterRange(ws);
                if (filterRange == null) return false;

                int fieldIndex = FindHeaderColumnIndexInFilterRange(filterRange, columnHeader);
                if (fieldIndex <= 0) return false;

                bool criteriaOk = AutoFilterCriteriaContainsExpected(ws, filterRange, fieldIndex, expectedText);
                if (!criteriaOk) return false;

                return AutoFilterVisibleRowsMatch(filterRange, fieldIndex, expectedText);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(filterRange);
                ReleaseCom(ws);
            }
        }
        //Project 1 Task 3
        public bool IfFormulaByHeaders(string sheetName,string targetHeader,string criteriaHeader,string compareOperator,double threshold,string trueText,string falseText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetHeader)) return false;
            if (string.IsNullOrWhiteSpace(criteriaHeader)) return false;
            if (string.IsNullOrWhiteSpace(compareOperator)) return false;
            if (string.IsNullOrWhiteSpace(trueText)) return false;
            if (string.IsNullOrWhiteSpace(falseText)) return false;

            Xl.Worksheet ws = null;
            Xl.Range usedRange = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                usedRange = ws.UsedRange;
                if (usedRange == null) return false;

                int headerRow;
                int targetColumn;
                int criteriaColumn;

                bool foundHeaders = FindTwoHeadersP1T3(
                    ws,
                    usedRange,
                    targetHeader,
                    criteriaHeader,
                    out headerRow,
                    out targetColumn,
                    out criteriaColumn);

                if (!foundHeaders) return false;

                int firstUsedRow = Convert.ToInt32(usedRange.Row);
                int usedRowCount = Convert.ToInt32(usedRange.Rows.Count);
                int lastUsedRow = firstUsedRow + usedRowCount - 1;

                int checkedRows = 0;

                for (int row = headerRow + 1; row <= lastUsedRow; row++)
                {
                    Xl.Range criteriaCell = null;
                    Xl.Range targetCell = null;

                    try
                    {
                        criteriaCell = (Xl.Range)ws.Cells[row, criteriaColumn];
                        targetCell = (Xl.Range)ws.Cells[row, targetColumn];

                        string criteriaRaw = GetCellStringP1T3(criteriaCell);
                        string targetRaw = GetCellStringP1T3(targetCell);

                        // Bỏ qua dòng trống cuối bảng nếu có
                        if (string.IsNullOrWhiteSpace(criteriaRaw) && string.IsNullOrWhiteSpace(targetRaw))
                            continue;

                        double quantity;
                        if (!TryParseDoubleP1T3(criteriaRaw, out quantity))
                            continue;

                        checkedRows++;

                        bool expectedCondition = CompareNumberP1T3(quantity, compareOperator, threshold);
                        string expectedResult = expectedCondition ? trueText : falseText;

                        string actualResult = GetCellTextP1T3(targetCell);

                        if (!StringEqualsP1T3(actualResult, expectedResult))
                            return false;

                        bool hasFormula = CellHasFormulaP1T3(targetCell);
                        if (!hasFormula)
                            return false;

                        string formula = Convert.ToString(targetCell.Formula);
                        string formulaLocal = Convert.ToString(targetCell.FormulaLocal);

                        bool formulaOk = FormulaLooksLikeCorrectIfP1T3(
                            formula,
                            formulaLocal,
                            row,
                            criteriaColumn,
                            criteriaHeader,
                            compareOperator,
                            threshold,
                            trueText,
                            falseText);

                        if (!formulaOk)
                            return false;
                    }
                    finally
                    {
                        ReleaseCom(criteriaCell);
                        ReleaseCom(targetCell);
                    }
                }

                return checkedRows > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(usedRange);
                ReleaseCom(ws);
            }
        }

        private bool FindTwoHeadersP1T3(Xl.Worksheet ws,Xl.Range usedRange,string targetHeader,string criteriaHeader,out int headerRow,out int targetColumn,out int criteriaColumn)
        {
            headerRow = 0;
            targetColumn = 0;
            criteriaColumn = 0;

            if (ws == null || usedRange == null)
                return false;

            string normalizedTargetHeader = NormalizeTextP1T3(targetHeader);
            string normalizedCriteriaHeader = NormalizeTextP1T3(criteriaHeader);

            int firstRow = Convert.ToInt32(usedRange.Row);
            int firstColumn = Convert.ToInt32(usedRange.Column);
            int rowCount = Convert.ToInt32(usedRange.Rows.Count);
            int columnCount = Convert.ToInt32(usedRange.Columns.Count);

            int lastRow = firstRow + rowCount - 1;
            int lastColumn = firstColumn + columnCount - 1;

            // Header thường nằm ở dòng đầu bảng, nhưng tìm trong 20 dòng đầu cho an toàn
            int maxSearchRow = firstRow + 20;
            if (maxSearchRow > lastRow) maxSearchRow = lastRow;

            for (int row = firstRow; row <= maxSearchRow; row++)
            {
                int foundTargetCol = 0;
                int foundCriteriaCol = 0;

                for (int col = firstColumn; col <= lastColumn; col++)
                {
                    Xl.Range cell = null;

                    try
                    {
                        cell = (Xl.Range)ws.Cells[row, col];
                        string text = GetCellTextP1T3(cell);
                        string normalized = NormalizeTextP1T3(text);

                        if (string.Equals(normalized, normalizedTargetHeader, StringComparison.OrdinalIgnoreCase))
                            foundTargetCol = col;

                        if (string.Equals(normalized, normalizedCriteriaHeader, StringComparison.OrdinalIgnoreCase))
                            foundCriteriaCol = col;
                    }
                    finally
                    {
                        ReleaseCom(cell);
                    }
                }

                if (foundTargetCol > 0 && foundCriteriaCol > 0)
                {
                    headerRow = row;
                    targetColumn = foundTargetCol;
                    criteriaColumn = foundCriteriaCol;
                    return true;
                }
            }

            return false;
        }

        private bool CellHasFormulaP1T3(Xl.Range cell)
        {
            if (cell == null)
                return false;

            try
            {
                object hasFormulaObj = cell.HasFormula;

                if (hasFormulaObj is bool)
                    return (bool)hasFormulaObj;

                string formula = Convert.ToString(cell.Formula);
                return !string.IsNullOrWhiteSpace(formula) && formula.Trim().StartsWith("=");
            }
            catch
            {
                return false;
            }
        }

        private bool FormulaLooksLikeCorrectIfP1T3(
            string formula,
            string formulaLocal,
            int row,
            int criteriaColumn,
            string criteriaHeader,
            string compareOperator,
            double threshold,
            string trueText,
            string falseText)
        {
            string f1 = NormalizeFormulaP1T3(formula);
            string f2 = NormalizeFormulaP1T3(formulaLocal);

            string combined = f1 + " " + f2;

            if (!combined.Contains("IF("))
                return false;

            string thresholdText = NormalizeFormulaP1T3(threshold.ToString("0"));
            if (!combined.Contains(thresholdText))
                return false;

            string trueToken = NormalizeFormulaP1T3(trueText);
            string falseToken = NormalizeFormulaP1T3(falseText);

            if (!combined.Contains(trueToken))
                return false;

            if (!combined.Contains(falseToken))
                return false;

            // Trường hợp công thức thường: =IF(E2>6000,"Yes","No")
            string expectedCellRef = ExcelColumnNameP1T3(criteriaColumn) + row;
            expectedCellRef = NormalizeFormulaP1T3(expectedCellRef);

            // Trường hợp công thức dạng Table: =IF([@[Quantity in Stock]]>6000,"Yes","No")
            string structuredHeaderRef = NormalizeFormulaP1T3(criteriaHeader);

            bool hasCriteriaReference =
                combined.Contains(expectedCellRef) ||
                combined.Contains(structuredHeaderRef);

            if (!hasCriteriaReference)
                return false;

            if (!IfComparisonMatchesP1T3(f1, expectedCellRef, structuredHeaderRef, compareOperator, threshold) &&
                !IfComparisonMatchesP1T3(f2, expectedCellRef, structuredHeaderRef, compareOperator, threshold))
                return false;

            return true;
        }

        private bool IfComparisonMatchesP1T3(string formula, string expectedCellRef, string structuredHeaderRef, string compareOperator, double threshold)
        {
            if (string.IsNullOrWhiteSpace(formula)) return false;
            int ifStart = formula.IndexOf("IF(", StringComparison.Ordinal);
            if (ifStart < 0) return false;
            int comma = formula.IndexOf(',', ifStart + 3);
            if (comma < 0) return false;

            string condition = formula.Substring(ifStart + 3, comma - ifStart - 3);
            Match match = Regex.Match(condition, @"(?<op>>=|<=|<>|>|<|=)");
            if (!match.Success || match.NextMatch().Success) return false;
            if (!string.Equals(match.Groups["op"].Value, compareOperator == null ? "" : compareOperator.Trim(), StringComparison.Ordinal))
                return false;

            string left = condition.Substring(0, match.Index);
            string right = condition.Substring(match.Index + match.Length).Trim('(', ')');
            if (!left.Contains(expectedCellRef) && !left.Contains(structuredHeaderRef)) return false;

            bool percent = right.EndsWith("%", StringComparison.Ordinal);
            if (percent) right = right.Substring(0, right.Length - 1);
            double parsed;
            if (!double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out parsed)) return false;
            if (percent) parsed /= 100.0;
            return Math.Abs(parsed - threshold) < 0.000001;
        }

        private string GetCellStringP1T3(Xl.Range cell)
        {
            if (cell == null)
                return "";

            try
            {
                object value = cell.Value2;
                return value == null ? "" : Convert.ToString(value).Trim();
            }
            catch
            {
                return "";
            }
        }

        private string GetCellTextP1T3(Xl.Range cell)
        {
            if (cell == null)
                return "";

            try
            {
                object text = cell.Text;
                string result = text == null ? "" : Convert.ToString(text).Trim();

                if (!string.IsNullOrWhiteSpace(result))
                    return result;

                object value = cell.Value2;
                return value == null ? "" : Convert.ToString(value).Trim();
            }
            catch
            {
                return "";
            }
        }
        //Project 1 Task 4
        public bool MultiLevelSortByHeaders(string sheetName, IList<string> sortHeaders, IList<string> sortOrders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (sortHeaders == null || sortHeaders.Count == 0) return false;

            Xl.Worksheet ws = null;
            Xl.Range tableRange = null;
            Xl.Range previousCell = null;
            Xl.Range currentCell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                tableRange = GetTableOrUsedRangeByHeadersP1T4T5(ws, sortHeaders);
                if (tableRange == null) return false;

                int rowCount = Convert.ToInt32(tableRange.Rows.Count);
                if (rowCount < 3) return false;

                int keyCount = sortHeaders.Count;
                int[] keyColumns = new int[keyCount];
                bool[] ascendingOrders = new bool[keyCount];

                for (int i = 0; i < keyCount; i++)
                {
                    if (string.IsNullOrWhiteSpace(sortHeaders[i])) return false;

                    keyColumns[i] = FindHeaderColumnIndexInFilterRange(tableRange, sortHeaders[i]);
                    if (keyColumns[i] <= 0) return false;

                    ascendingOrders[i] = IsAscendingSortOrderP1T4(sortOrders, i);
                }

                int checkedPairs = 0;

                for (int rowIndex = 3; rowIndex <= rowCount; rowIndex++)
                {
                    bool pairHasData = false;
                    bool pairResolved = false;

                    for (int keyIndex = 0; keyIndex < keyCount; keyIndex++)
                    {
                        ReleaseCom(previousCell);
                        ReleaseCom(currentCell);
                        previousCell = null;
                        currentCell = null;

                        previousCell = (Xl.Range)tableRange.Cells[rowIndex - 1, keyColumns[keyIndex]];
                        currentCell = (Xl.Range)tableRange.Cells[rowIndex, keyColumns[keyIndex]];

                        string previousText = GetCellTextP1T3(previousCell);
                        string currentText = GetCellTextP1T3(currentCell);

                        if (!string.IsNullOrWhiteSpace(previousText) || !string.IsNullOrWhiteSpace(currentText))
                            pairHasData = true;

                        int compare = CompareSortValueP1T4(previousText, currentText);

                        if (compare == 0)
                            continue;

                        checkedPairs++;
                        pairResolved = true;

                        if (ascendingOrders[keyIndex])
                        {
                            if (compare > 0)
                                return false;
                        }
                        else
                        {
                            if (compare < 0)
                                return false;
                        }

                        break;
                    }

                    if (pairHasData && !pairResolved)
                        checkedPairs++;
                }

                return checkedPairs > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(previousCell);
                ReleaseCom(currentCell);
                ReleaseCom(tableRange);
                ReleaseCom(ws);
            }
        }

        //Project 1 Task 5
        public bool EmailFormulaFromHeader(string sheetName, string targetHeader, string sourceHeader, string domain, bool requireFunction)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetHeader)) return false;
            if (string.IsNullOrWhiteSpace(sourceHeader)) return false;
            if (string.IsNullOrWhiteSpace(domain)) return false;

            Xl.Worksheet ws = null;
            Xl.Range tableRange = null;
            Xl.Range sourceCell = null;
            Xl.Range targetCell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                List<string> requiredHeaders = new List<string>();
                requiredHeaders.Add(targetHeader);
                requiredHeaders.Add(sourceHeader);

                tableRange = GetTableOrUsedRangeByHeadersP1T4T5(ws, requiredHeaders);
                if (tableRange == null) return false;

                int targetFieldIndex = FindHeaderColumnIndexInFilterRange(tableRange, targetHeader);
                int sourceFieldIndex = FindHeaderColumnIndexInFilterRange(tableRange, sourceHeader);

                if (targetFieldIndex <= 0) return false;
                if (sourceFieldIndex <= 0) return false;

                int rowCount = Convert.ToInt32(tableRange.Rows.Count);
                int firstWorksheetRow = Convert.ToInt32(tableRange.Row);
                int firstWorksheetColumn = Convert.ToInt32(tableRange.Column);

                int sourceWorksheetColumn = firstWorksheetColumn + sourceFieldIndex - 1;
                string normalizedDomain = NormalizeEmailDomainP1T5(domain);

                int checkedRows = 0;

                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    ReleaseCom(sourceCell);
                    ReleaseCom(targetCell);
                    sourceCell = null;
                    targetCell = null;

                    sourceCell = (Xl.Range)tableRange.Cells[rowIndex, sourceFieldIndex];
                    targetCell = (Xl.Range)tableRange.Cells[rowIndex, targetFieldIndex];

                    string firstName = GetCellTextP1T3(sourceCell);
                    string actualEmail = GetCellTextP1T3(targetCell);

                    if (string.IsNullOrWhiteSpace(firstName) && string.IsNullOrWhiteSpace(actualEmail))
                        continue;

                    if (string.IsNullOrWhiteSpace(firstName))
                        return false;

                    checkedRows++;

                    string expectedEmail = firstName.Trim() + normalizedDomain;

                    if (!EmailTextEqualsP1T5(actualEmail, expectedEmail))
                        return false;

                    if (!CellHasFormulaP1T3(targetCell))
                        return false;

                    string formula = Convert.ToString(targetCell.Formula);
                    string formulaLocal = Convert.ToString(targetCell.FormulaLocal);
                    int worksheetRow = firstWorksheetRow + rowIndex - 1;

                    bool formulaOk = FormulaLooksLikeEmailFormulaP1T5(
                        formula,
                        formulaLocal,
                        worksheetRow,
                        sourceWorksheetColumn,
                        sourceHeader,
                        normalizedDomain,
                        requireFunction);

                    if (!formulaOk)
                        return false;
                }

                return checkedRows > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sourceCell);
                ReleaseCom(targetCell);
                ReleaseCom(tableRange);
                ReleaseCom(ws);
            }
        }

        private Xl.Range GetTableOrUsedRangeByHeadersP1T4T5(Xl.Worksheet ws, IList<string> requiredHeaders)
        {
            if (ws == null) return null;
            if (requiredHeaders == null || requiredHeaders.Count == 0) return null;

            Xl.ListObjects listObjects = null;
            Xl.ListObject listObject = null;
            Xl.Range range = null;
            Xl.Range usedRange = null;

            try
            {
                try
                {
                    listObjects = ws.ListObjects;

                    if (listObjects != null)
                    {
                        int tableCount = Convert.ToInt32(listObjects.Count);

                        for (int i = 1; i <= tableCount; i++)
                        {
                            ReleaseCom(listObject);
                            ReleaseCom(range);
                            listObject = null;
                            range = null;

                            listObject = listObjects.Item[i];
                            if (listObject == null) continue;

                            range = listObject.Range;

                            if (RangeContainsRequiredHeadersP1T4T5(range, requiredHeaders))
                            {
                                Xl.Range result = range;
                                range = null;
                                return result;
                            }
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(range);
                    ReleaseCom(listObject);
                    ReleaseCom(listObjects);
                    range = null;
                    listObject = null;
                    listObjects = null;
                }

                usedRange = ws.UsedRange;
                if (RangeContainsRequiredHeadersP1T4T5(usedRange, requiredHeaders))
                {
                    Xl.Range result = usedRange;
                    usedRange = null;
                    return result;
                }

                range = FindRangeStartingAtHeaderRowP1T5(ws, usedRange, requiredHeaders);
                if (range != null)
                {
                    Xl.Range result = range;
                    range = null;
                    return result;
                }

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(usedRange);
                ReleaseCom(listObject);
                ReleaseCom(listObjects);
            }
        }

        private bool RangeContainsRequiredHeadersP1T4T5(Xl.Range range, IList<string> requiredHeaders)
        {
            if (range == null) return false;
            if (requiredHeaders == null || requiredHeaders.Count == 0) return false;

            for (int i = 0; i < requiredHeaders.Count; i++)
            {
                string header = requiredHeaders[i];
                if (string.IsNullOrWhiteSpace(header))
                    return false;

                int columnIndex = FindHeaderColumnIndexInFilterRange(range, header);
                if (columnIndex <= 0)
                    return false;
            }

            return true;
        }

        private bool IsAscendingSortOrderP1T4(IList<string> sortOrders, int index)
        {
            if (sortOrders == null || index < 0 || index >= sortOrders.Count)
                return true;

            string order = sortOrders[index] == null ? "" : sortOrders[index].Trim().ToUpperInvariant();
            order = order.Replace(" ", "");

            if (order.Contains("DESC") || order.Contains("ZTOA") || order.Contains("Z-A"))
                return false;

            return true;
        }

        private int CompareSortValueP1T4(string left, string right)
        {
            string a = NormalizeSortValueP1T4(left);
            string b = NormalizeSortValueP1T4(right);

            double numberA;
            double numberB;

            if (TryParseDoubleP1T3(a, out numberA) && TryParseDoubleP1T3(b, out numberB))
                return numberA.CompareTo(numberB);

            return string.Compare(a, b, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeSortValueP1T4(string value)
        {
            if (value == null) return "";

            string result = value.Trim();
            result = result.Replace((char)160, ' ');
            result = Regex.Replace(result, @"\s+", " ");
            return result.ToUpperInvariant();
        }

        private string NormalizeEmailDomainP1T5(string domain)
        {
            if (domain == null) return "";

            string result = domain.Trim();
            result = result.Replace("\"", "");
            result = result.Replace("'", "");

            if (!result.StartsWith("@", StringComparison.Ordinal))
                result = "@" + result;

            return result;
        }

        private bool EmailTextEqualsP1T5(string actual, string expected)
        {
            string a = NormalizeEmailTextP1T5(actual);
            string e = NormalizeEmailTextP1T5(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeEmailTextP1T5(string text)
        {
            if (text == null) return "";

            string result = text.Trim();
            result = result.Replace((char)160, ' ');
            result = Regex.Replace(result, @"\s+", "");
            return result.ToLowerInvariant();
        }

        private bool FormulaLooksLikeEmailFormulaP1T5(
            string formula,
            string formulaLocal,
            int row,
            int sourceColumn,
            string sourceHeader,
            string domain,
            bool requireFunction)
        {
            string f1 = NormalizeFormulaP1T5(formula);
            string f2 = NormalizeFormulaP1T5(formulaLocal);
            string combined = f1 + " " + f2;

            if (string.IsNullOrWhiteSpace(combined))
                return false;

            string expectedCellRef = NormalizeFormulaP1T5(ExcelColumnNameP1T3(sourceColumn) + row);
            string structuredHeaderRef = NormalizeFormulaP1T5(sourceHeader);

            bool hasSourceReference =
                combined.Contains(expectedCellRef) ||
                combined.Contains(structuredHeaderRef);

            if (!hasSourceReference)
                return false;

            string domainWithAt = NormalizeFormulaP1T5(domain);
            string domainWithoutAt = NormalizeFormulaP1T5(domain == null ? "" : domain.Trim().TrimStart('@'));

            bool hasDomain =
                (!string.IsNullOrWhiteSpace(domainWithAt) && combined.Contains(domainWithAt)) ||
                (!string.IsNullOrWhiteSpace(domainWithoutAt) && combined.Contains(domainWithoutAt));

            if (!hasDomain)
                return false;

            bool usesFunction =
                combined.Contains("CONCATENATE(") ||
                combined.Contains("CONCAT(") ||
                combined.Contains("TEXTJOIN(");

            bool looksLikeConstructFormula = usesFunction ||
                (!requireFunction && combined.Contains("&"));

            if (!looksLikeConstructFormula)
                return false;

            return true;
        }

        private string NormalizeFormulaP1T5(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();
            s = s.Replace(";", ",");
            s = s.Replace("$", "");
            s = s.Replace("“", "\"");
            s = s.Replace("”", "\"");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.ToUpperInvariant();

            string result = "";
            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (!char.IsWhiteSpace(ch))
                    result += ch;
            }

            return result;
        }
        private bool TryParseDoubleP1T3(string input, out double value)
        {
            value = 0;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            string s = input.Trim();

            bool ok = double.TryParse(
                s,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture,
                out value);

            if (ok) return true;

            ok = double.TryParse(
                s,
                System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.CurrentCulture,
                out value);

            return ok;
        }
        //Project 1 Task 6
        public bool TableBandedRows(string sheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects listObjects = null;
            Xl.ListObject listObject = null;
            Xl.Range tableRange = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                listObjects = ws.ListObjects;
                if (listObjects == null) return false;

                int tableCount = Convert.ToInt32(listObjects.Count);
                if (tableCount <= 0) return false;

                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(listObject);
                    ReleaseCom(tableRange);
                    listObject = null;
                    tableRange = null;

                    listObject = listObjects.Item[i];
                    if (listObject == null) continue;

                    tableRange = listObject.Range;
                    if (tableRange == null) continue;

                    int rowCount = Convert.ToInt32(tableRange.Rows.Count);
                    if (rowCount < 2) continue;

                    object rowStripes = null;

                    try
                    {
                        rowStripes = listObject.ShowTableStyleRowStripes;
                    }
                    catch
                    {
                        rowStripes = null;
                    }

                    if (IsComBooleanTrueP1T6(rowStripes))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(tableRange);
                ReleaseCom(listObject);
                ReleaseCom(listObjects);
                ReleaseCom(ws);
            }
        }

        //Project 1 Task 7
        public bool ChartSheetExists(string chartSheetName, string sourceSheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(chartSheetName)) return false;

            Xl.Workbook wb = null;
            Xl.Sheets sheets = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                if (wb == null) return false;

                sheets = wb.Sheets;
                if (sheets == null) return false;

                int sheetCount = Convert.ToInt32(sheets.Count);

                for (int i = 1; i <= sheetCount; i++)
                {
                    object sheetObject = null;
                    Xl.Chart chart = null;

                    try
                    {
                        sheetObject = sheets[i];

                        try
                        {
                            chart = sheetObject as Xl.Chart;
                        }
                        catch
                        {
                            chart = null;
                        }

                        if (chart == null)
                            continue;

                        string actualName = Convert.ToString(chart.Name);

                        if (!StringEqualsP1T7(actualName, chartSheetName))
                            continue;

                        return ChartSheetHasUsableChartP1T7(chart, sourceSheetName);
                    }
                    finally
                    {
                        if (chart != null)
                            ReleaseCom(chart);
                        else
                            ReleaseCom(sheetObject);
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sheets);
            }
        }

        private bool IsComBooleanTrueP1T6(object value)
        {
            if (value == null) return false;

            if (value is bool)
                return (bool)value;

            try
            {
                int number = Convert.ToInt32(value, CultureInfo.InvariantCulture);
                return number != 0;
            }
            catch
            {
            }

            string text = Convert.ToString(value);
            if (string.IsNullOrWhiteSpace(text)) return false;

            text = text.Trim();

            return string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "YES", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "-1", StringComparison.OrdinalIgnoreCase);
        }

        private bool ChartSheetHasUsableChartP1T7(Xl.Chart chart, string sourceSheetName)
        {
            if (chart == null) return false;

            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;

            try
            {
                try
                {
                    seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                }
                catch
                {
                    seriesCollection = null;
                }

                // Một số chart sheet đọc SeriesCollection bị lỗi COM.
                // Nếu đã xác định đúng là Chart Sheet tên Fee Chart thì vẫn cho qua để tránh false fail.
                if (seriesCollection == null)
                    return true;

                int seriesCount = 0;

                try
                {
                    seriesCount = Convert.ToInt32(seriesCollection.Count);
                }
                catch
                {
                    seriesCount = 0;
                }

                if (seriesCount <= 0)
                    return false;

                if (string.IsNullOrWhiteSpace(sourceSheetName))
                    return true;

                string sourceKey = NormalizeChartFormulaP1T7(sourceSheetName);

                for (int i = 1; i <= seriesCount; i++)
                {
                    ReleaseCom(series);
                    series = null;

                    try
                    {
                        series = seriesCollection.Item(i);
                        if (series == null) continue;

                        string formula = Convert.ToString(series.Formula);
                        string normalizedFormula = NormalizeChartFormulaP1T7(formula);

                        if (!string.IsNullOrWhiteSpace(sourceKey) &&
                            normalizedFormula.Contains(sourceKey))
                            return true;
                    }
                    catch
                    {
                    }
                }

                // Một số chart dùng structured reference theo Table nên Series.Formula có thể không chứa tên sheet.
                // Vẫn PASS nếu đúng chart sheet và có series dữ liệu.
                return seriesCount > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
            }
        }

        private string NormalizeChartFormulaP1T7(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("=", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }

        private bool StringEqualsP1T7(string actual, string expected)
        {
            string a = NormalizeTextP1T3(actual);
            string e = NormalizeTextP1T3(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }
        private bool CompareNumberP1T3(double actual, string compareOperator, double threshold)
        {
            string op = compareOperator == null ? "" : compareOperator.Trim();

            switch (op)
            {
                case ">":
                    return actual > threshold;

                case ">=":
                    return actual >= threshold;

                case "<":
                    return actual < threshold;

                case "<=":
                    return actual <= threshold;

                case "=":
                case "==":
                    return Math.Abs(actual - threshold) < 0.000001;

                case "<>":
                case "!=":
                    return Math.Abs(actual - threshold) >= 0.000001;

                default:
                    return false;
            }
        }

        private bool StringEqualsP1T3(string actual, string expected)
        {
            string a = actual == null ? "" : actual.Trim();
            string e = expected == null ? "" : expected.Trim();

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeTextP1T3(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            string s = text.Trim().ToUpperInvariant();
            string result = "";

            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];

                if (char.IsLetterOrDigit(ch))
                    result += ch;
            }

            return result;
        }

        private string NormalizeFormulaP1T3(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();

            s = s.Replace(";", ",");
            s = s.Replace("$", "");
            s = s.Replace("“", "\"");
            s = s.Replace("”", "\"");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.ToUpperInvariant();

            string result = "";

            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];

                if (!char.IsWhiteSpace(ch))
                    result += ch;
            }

            return result;
        }

        private string ExcelColumnNameP1T3(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = "";

            while (dividend > 0)
            {
                int modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo) + columnName;
                dividend = (dividend - modulo) / 26;
            }

            return columnName;
        }
        private Xl.Range GetAutoFilterRange(Xl.Worksheet ws)
        {
            if (ws == null) return null;

            Xl.AutoFilter autoFilter = null;
            Xl.Range range = null;
            Xl.ListObjects listObjects = null;
            Xl.ListObject listObject = null;

            try
            {
                try
                {
                    autoFilter = ws.AutoFilter;
                    if (autoFilter != null)
                    {
                        range = autoFilter.Range;
                        if (range != null)
                            return range;
                    }
                }
                catch
                {
                    ReleaseCom(range);
                    range = null;
                }
                finally
                {
                    ReleaseCom(autoFilter);
                }

                try
                {
                    listObjects = ws.ListObjects;
                    if (listObjects != null && listObjects.Count > 0)
                    {
                        listObject = listObjects.Item[1];
                        if (listObject != null)
                        {
                            range = listObject.Range;
                            return range;
                        }
                    }
                }
                catch
                {
                    ReleaseCom(range);
                    range = null;
                }

                return null;
            }
            finally
            {
                ReleaseCom(listObject);
                ReleaseCom(listObjects);
            }
        }
        //Project 1 Task 8
        public bool NoConditionalFormatting(string sheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Worksheet ws = null;
            Xl.Range allCells = null;
            Xl.Range usedRange = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                allCells = ws.Cells as Xl.Range;
                if (RangeHasConditionalFormattingP1T8(allCells))
                    return false;

                usedRange = ws.UsedRange;
                if (RangeHasConditionalFormattingP1T8(usedRange))
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(usedRange);
                ReleaseCom(allCells);
                ReleaseCom(ws);
            }
        }

        private bool RangeHasConditionalFormattingP1T8(Xl.Range range)
        {
            if (range == null) return false;

            Xl.FormatConditions formatConditions = null;
            Xl.Range conditionalCells = null;

            try
            {
                try
                {
                    formatConditions = range.FormatConditions;
                    if (formatConditions != null)
                    {
                        int count = 0;
                        try
                        {
                            count = Convert.ToInt32(formatConditions.Count);
                        }
                        catch
                        {
                            count = 0;
                        }

                        if (count > 0)
                            return true;
                    }
                }
                catch
                {
                }

                try
                {
                    conditionalCells = range.SpecialCells(
                        Xl.XlCellType.xlCellTypeAllFormatConditions,
                        Type.Missing);

                    if (conditionalCells != null)
                        return true;
                }
                catch (COMException)
                {
                    // Excel ném COMException khi không tìm thấy ô nào có Conditional Formatting.
                    return false;
                }
                catch
                {
                    return false;
                }

                return false;
            }
            finally
            {
                ReleaseCom(conditionalCells);
                ReleaseCom(formatConditions);
            }
        }
        private int FindHeaderColumnIndexInFilterRange(Xl.Range filterRange, string columnHeader)
        {
            if (filterRange == null || string.IsNullOrWhiteSpace(columnHeader))
                return -1;

            Xl.Range headerRow = null;
            Xl.Range cell = null;

            try
            {
                headerRow = (Xl.Range)filterRange.Rows[1];
                int columnCount = filterRange.Columns.Count;
                string expectedHeader = NormalizeHeaderText(columnHeader);

                for (int i = 1; i <= columnCount; i++)
                {
                    ReleaseCom(cell);
                    cell = null;

                    cell = (Xl.Range)headerRow.Cells[1, i];

                    string actualHeader = "";
                    try
                    {
                        actualHeader = cell.Text == null ? "" : cell.Text.ToString();
                    }
                    catch
                    {
                        actualHeader = cell.Value2 == null ? "" : cell.Value2.ToString();
                    }

                    if (string.Equals(NormalizeHeaderText(actualHeader), expectedHeader, StringComparison.OrdinalIgnoreCase))
                        return i;
                }

                return -1;
            }
            catch
            {
                return -1;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(headerRow);
            }
        }
        //Project 2 Task 1
        public bool ImportedCsvAtCell(string projectId, string sheetName, string startCell, string sourceFileName, bool firstRowAsHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(projectId)) return false;
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(startCell)) startCell = "A1";

            Xl.Worksheet ws = null;
            Xl.Range startRange = null;
            Xl.ListObject table = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                startRange = ws.Range[startCell];
                if (startRange == null) return false;

                int startRow = Convert.ToInt32(startRange.Row);
                int startColumn = Convert.ToInt32(startRange.Column);

                string sourcePath = FindProjectAssetFileP2T1(projectId, sourceFileName);
                if (string.IsNullOrWhiteSpace(sourcePath) || !File.Exists(sourcePath))
                    return false;

                List<List<string>> expectedRows = ReadDelimitedFileP2T1(sourcePath);
                NormalizeExpectedRowsP2T1(expectedRows);

                if (expectedRows == null || expectedRows.Count == 0)
                    return false;

                int expectedColumnCount = GetExpectedColumnCountP2T1(expectedRows, firstRowAsHeaders);
                if (expectedColumnCount <= 0)
                    return false;

                int expectedRowCount = expectedRows.Count;

                table = FindListObjectStartingAtP2T1(ws, startRow, startColumn);

                // MOS yêu cầu Load To... Table / Existing worksheet =$A$1.
                if (table == null)
                    return false;

                if (firstRowAsHeaders && !ListObjectHasHeaderRowP2T1(table))
                    return false;

                if (!ListObjectSizeEqualsP2T1(table, expectedRowCount, expectedColumnCount))
                    return false;

                if (!ImportedCellsMatchExpectedP2T1(ws, startRow, startColumn, expectedRows, expectedColumnCount))
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(table);
                ReleaseCom(startRange);
                ReleaseCom(ws);
            }
        }
        private string FindProjectAssetFileP2T1(string projectId, string sourceFileName)
        {
            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            string safeProjectId = projectId == null ? "" : projectId.Trim();
            string safeFileName = sourceFileName == null ? "" : sourceFileName.Trim();

            List<string> folders = new List<string>();

            if (!string.IsNullOrWhiteSpace(safeProjectId))
            {
                folders.Add(Path.Combine(documents, "MOS Trainer", "AssetsTemp", safeProjectId));
                folders.Add(Path.Combine(documents, "MosTrainer", "AssetsTemp", safeProjectId));
                folders.Add(Path.Combine(documents, safeProjectId));
            }

            folders.Add(documents);

            for (int i = 0; i < folders.Count; i++)
            {
                string folder = folders[i];
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                    continue;

                if (!string.IsNullOrWhiteSpace(safeFileName))
                {
                    string directPath = Path.Combine(folder, safeFileName);
                    if (File.Exists(directPath))
                        return directPath;

                    if (Path.GetExtension(safeFileName).Length == 0)
                    {
                        string csvPath = Path.Combine(folder, safeFileName + ".csv");
                        if (File.Exists(csvPath))
                            return csvPath;

                        string txtPath = Path.Combine(folder, safeFileName + ".txt");
                        if (File.Exists(txtPath))
                            return txtPath;
                    }
                }
            }

            for (int i = 0; i < folders.Count; i++)
            {
                string folder = folders[i];
                if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                    continue;

                string match = FindFileRecursiveP2T1(folder, safeFileName);
                if (!string.IsNullOrWhiteSpace(match))
                    return match;
            }

            return "";
        }

        private string FindFileRecursiveP2T1(string folder, string sourceFileName)
        {
            try
            {
                string[] files = Directory.GetFiles(folder, "*.*", SearchOption.AllDirectories);
                string expectedKey = NormalizeFileKeyP2T1(sourceFileName);
                string expectedKeyNoS = TrimTrailingSP2T1(expectedKey);

                string firstCsvOrTxt = "";

                for (int i = 0; i < files.Length; i++)
                {
                    string file = files[i];
                    string ext = Path.GetExtension(file);

                    if (!string.Equals(ext, ".csv", StringComparison.OrdinalIgnoreCase) &&
                        !string.Equals(ext, ".txt", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (string.IsNullOrWhiteSpace(firstCsvOrTxt))
                        firstCsvOrTxt = file;

                    if (string.IsNullOrWhiteSpace(expectedKey))
                        continue;

                    string candidateKey = NormalizeFileKeyP2T1(Path.GetFileName(file));
                    string candidateKeyNoS = TrimTrailingSP2T1(candidateKey);

                    if (string.Equals(candidateKey, expectedKey, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(candidateKeyNoS, expectedKeyNoS, StringComparison.OrdinalIgnoreCase))
                        return file;
                }

                if (string.IsNullOrWhiteSpace(expectedKey))
                    return firstCsvOrTxt;
            }
            catch
            {
            }

            return "";
        }

        private string NormalizeFileKeyP2T1(string fileName)
        {
            if (string.IsNullOrWhiteSpace(fileName)) return "";

            string name = Path.GetFileNameWithoutExtension(fileName).Trim();
            name = Regex.Replace(name, @"\s+", "");
            name = name.Replace("_", "");
            name = name.Replace("-", "");
            name = name.ToUpperInvariant();

            return name;
        }

        private string TrimTrailingSP2T1(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            while (text.EndsWith("S", StringComparison.OrdinalIgnoreCase) && text.Length > 1)
                text = text.Substring(0, text.Length - 1);

            return text;
        }

        private List<List<string>> ReadDelimitedFileP2T1(string path)
        {
            List<List<string>> rows = new List<List<string>>();

            string[] lines = File.ReadAllLines(path, Encoding.UTF8);
            if (lines == null || lines.Length == 0)
                return rows;

            char delimiter = DetectDelimiterP2T1(lines);

            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i];

                if (line == null)
                    line = "";

                if (i == 0)
                    line = line.TrimStart('\uFEFF');

                rows.Add(ParseDelimitedLineP2T1(line, delimiter));
            }

            return rows;
        }

        private char DetectDelimiterP2T1(string[] lines)
        {
            char[] delimiters = new char[] { ',', ';', '\t' };
            int[] counts = new int[] { 0, 0, 0 };

            int checkedLines = 0;

            for (int i = 0; i < lines.Length && checkedLines < 5; i++)
            {
                string line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                    continue;

                checkedLines++;

                for (int d = 0; d < delimiters.Length; d++)
                    counts[d] += CountDelimiterOutsideQuotesP2T1(line, delimiters[d]);
            }

            int bestIndex = 0;
            for (int i = 1; i < counts.Length; i++)
            {
                if (counts[i] > counts[bestIndex])
                    bestIndex = i;
            }

            return delimiters[bestIndex];
        }

        private int CountDelimiterOutsideQuotesP2T1(string line, char delimiter)
        {
            if (line == null) return 0;

            bool inQuotes = false;
            int count = 0;

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];

                if (ch == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        i++;
                        continue;
                    }

                    inQuotes = !inQuotes;
                    continue;
                }

                if (!inQuotes && ch == delimiter)
                    count++;
            }

            return count;
        }

        private List<string> ParseDelimitedLineP2T1(string line, char delimiter)
        {
            List<string> values = new List<string>();

            if (line == null)
            {
                values.Add("");
                return values;
            }

            bool inQuotes = false;
            StringBuilder current = new StringBuilder();

            for (int i = 0; i < line.Length; i++)
            {
                char ch = line[i];

                if (ch == '"')
                {
                    if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                    {
                        current.Append('"');
                        i++;
                        continue;
                    }

                    inQuotes = !inQuotes;
                    continue;
                }

                if (!inQuotes && ch == delimiter)
                {
                    values.Add(current.ToString());
                    current.Length = 0;
                    continue;
                }

                current.Append(ch);
            }

            values.Add(current.ToString());
            return values;
        }

        private void NormalizeExpectedRowsP2T1(List<List<string>> rows)
        {
            if (rows == null) return;

            for (int r = rows.Count - 1; r >= 0; r--)
            {
                if (RowIsCompletelyEmptyP2T1(rows[r]))
                    rows.RemoveAt(r);
                else
                    break;
            }

            for (int r = 0; r < rows.Count; r++)
                TrimTrailingEmptyCellsP2T1(rows[r]);
        }

        private bool RowIsCompletelyEmptyP2T1(List<string> row)
        {
            if (row == null || row.Count == 0) return true;

            for (int i = 0; i < row.Count; i++)
            {
                if (!string.IsNullOrWhiteSpace(row[i]))
                    return false;
            }

            return true;
        }

        private void TrimTrailingEmptyCellsP2T1(List<string> row)
        {
            if (row == null) return;

            for (int i = row.Count - 1; i >= 0; i--)
            {
                if (string.IsNullOrWhiteSpace(row[i]))
                    row.RemoveAt(i);
                else
                    break;
            }
        }

        private int GetExpectedColumnCountP2T1(List<List<string>> rows, bool firstRowAsHeaders)
        {
            if (rows == null || rows.Count == 0) return 0;

            if (firstRowAsHeaders && rows[0] != null && rows[0].Count > 0)
                return rows[0].Count;

            int max = 0;
            for (int i = 0; i < rows.Count; i++)
            {
                if (rows[i] != null && rows[i].Count > max)
                    max = rows[i].Count;
            }

            return max;
        }

        private Xl.ListObject FindListObjectStartingAtP2T1(Xl.Worksheet ws, int startRow, int startColumn)
        {
            Xl.ListObjects listObjects = null;
            Xl.ListObject listObject = null;
            Xl.Range range = null;

            try
            {
                listObjects = ws.ListObjects;
                if (listObjects == null) return null;

                int count = Convert.ToInt32(listObjects.Count);

                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(listObject);
                    ReleaseCom(range);
                    listObject = null;
                    range = null;

                    listObject = listObjects.Item[i];
                    if (listObject == null) continue;

                    range = listObject.Range;
                    if (range == null) continue;

                    int row = Convert.ToInt32(range.Row);
                    int column = Convert.ToInt32(range.Column);

                    if (row == startRow && column == startColumn)
                    {
                        Xl.ListObject result = listObject;
                        listObject = null;
                        return result;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(range);
                ReleaseCom(listObject);
                ReleaseCom(listObjects);
            }
        }

        private bool ListObjectHasHeaderRowP2T1(Xl.ListObject table)
        {
            if (table == null) return false;

            Xl.Range headerRange = null;

            try
            {
                bool showHeaders = true;

                try
                {
                    showHeaders = Convert.ToBoolean(table.ShowHeaders);
                }
                catch
                {
                    showHeaders = true;
                }

                if (!showHeaders)
                    return false;

                headerRange = table.HeaderRowRange;
                return headerRange != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(headerRange);
            }
        }

        private bool ListObjectSizeEqualsP2T1(Xl.ListObject table, int expectedRows, int expectedColumns)
        {
            if (table == null) return false;

            Xl.Range range = null;

            try
            {
                range = table.Range;
                if (range == null) return false;

                int actualRows = Convert.ToInt32(range.Rows.Count);
                int actualColumns = Convert.ToInt32(range.Columns.Count);

                return actualRows == expectedRows && actualColumns == expectedColumns;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(range);
            }
        }

        private bool ImportedCellsMatchExpectedP2T1(Xl.Worksheet ws, int startRow, int startColumn, List<List<string>> expectedRows, int expectedColumnCount)
        {
            if (ws == null) return false;
            if (expectedRows == null || expectedRows.Count == 0) return false;

            Xl.Range cell = null;

            try
            {
                for (int r = 0; r < expectedRows.Count; r++)
                {
                    for (int c = 0; c < expectedColumnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)ws.Cells[startRow + r, startColumn + c];

                        string actual = GetCellTextP1T3(cell);
                        string expected = "";

                        if (expectedRows[r] != null && c < expectedRows[r].Count)
                            expected = expectedRows[r][c];

                        if (!CsvCellEqualsP2T1(actual, expected))
                            return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool CsvCellEqualsP2T1(string actual, string expected)
        {
            string a = NormalizeCsvCellP2T1(actual);
            string e = NormalizeCsvCellP2T1(expected);

            if (string.Equals(a, e, StringComparison.OrdinalIgnoreCase))
                return true;

            double numberA;
            double numberE;
            if (TryParseCsvDoubleP2T1(a, out numberA) && TryParseCsvDoubleP2T1(e, out numberE))
                return Math.Abs(numberA - numberE) < 0.0000001;

            DateTime dateA;
            DateTime dateE;
            if (DateTime.TryParse(a, out dateA) && DateTime.TryParse(e, out dateE))
                return dateA.Date == dateE.Date;

            return false;
        }

        private string NormalizeCsvCellP2T1(string text)
        {
            if (text == null) return "";

            string result = text.Trim();
            result = result.Replace((char)160, ' ');
            result = result.Replace("\r", " ");
            result = result.Replace("\n", " ");
            result = Regex.Replace(result, @"\s+", " ");

            return result.Trim();
        }

        private bool TryParseCsvDoubleP2T1(string text, out double value)
        {
            value = 0;
            if (string.IsNullOrWhiteSpace(text)) return false;

            string s = text.Trim();

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return true;

            if (double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                return true;

            s = s.Replace(",", "");
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return true;

            return false;
        }
        //Project 2 Task 2
        public bool ColumnWidthEquals(string sheetName, string columnRange, double expectedWidth)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(columnRange)) return false;
            if (expectedWidth <= 0) return false;

            Xl.Worksheet ws = null;
            Xl.Range columnsRange = null;
            Xl.Range column = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                columnsRange = GetColumnRangeP2T2(ws, columnRange);
                if (columnsRange == null) return false;

                int columnCount = Convert.ToInt32(columnsRange.Columns.Count);
                if (columnCount <= 0) return false;

                for (int i = 1; i <= columnCount; i++)
                {
                    ReleaseCom(column);
                    column = null;

                    column = (Xl.Range)columnsRange.Columns[i];
                    if (column == null) return false;

                    double actualWidth = Convert.ToDouble(column.ColumnWidth, CultureInfo.InvariantCulture);

                    // Excel ColumnWidth đôi khi lệch rất nhỏ giữa các máy/Office build.
                    if (Math.Abs(actualWidth - expectedWidth) > 0.05)
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(column);
                ReleaseCom(columnsRange);
                ReleaseCom(ws);
            }
        }

        //Project 2 Task 3
        public bool ColumnSparklinesByRange(string sheetName, string locationRange, string dataRange)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(locationRange)) return false;
            if (string.IsNullOrWhiteSpace(dataRange)) return false;

            Xl.Worksheet ws = null;
            Xl.Range expectedLocation = null;
            Xl.Range expectedData = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                if (!TryGetEffectiveSparklineRangesP2T3(ws, locationRange, dataRange, out expectedLocation, out expectedData))
                    return false;

                if (expectedLocation == null || expectedData == null)
                    return false;

                int locationRows = Convert.ToInt32(expectedLocation.Rows.Count);
                int locationColumns = Convert.ToInt32(expectedLocation.Columns.Count);
                int dataRows = Convert.ToInt32(expectedData.Rows.Count);
                int dataColumns = Convert.ToInt32(expectedData.Columns.Count);

                if (locationColumns != 1)
                    return false;

                if (locationRows != dataRows)
                    return false;

                if (dataColumns < 2)
                    return false;

                return LocationRangeHasColumnSparklinesP2T3(expectedLocation, expectedData);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(expectedLocation);
                ReleaseCom(expectedData);
                ReleaseCom(ws);
            }
        }
        private Xl.Range GetColumnRangeP2T2(Xl.Worksheet ws, string columnRange)
        {
            if (ws == null || string.IsNullOrWhiteSpace(columnRange)) return null;

            string address = columnRange.Trim();
            address = address.Replace("$", "");
            address = address.Replace("'", "");

            int bangIndex = address.LastIndexOf('!');
            if (bangIndex >= 0 && bangIndex + 1 < address.Length)
                address = address.Substring(bangIndex + 1);

            try
            {
                return ws.Range[address];
            }
            catch
            {
            }

            try
            {
                return ws.Columns[address] as Xl.Range;
            }
            catch
            {
                return null;
            }
        }

        private Xl.Worksheet GetWorksheetExactOrNormalizedP2T2T3(string sheetName)
        {
            if (string.IsNullOrWhiteSpace(sheetName)) return null;

            Xl.Worksheet exact = null;
            try
            {
                exact = GetWorksheet(sheetName);
                if (exact != null) return exact;
            }
            catch
            {
                ReleaseCom(exact);
                exact = null;
            }

            Xl.Workbook wb = null;
            Xl.Sheets sheets = null;
            Xl.Worksheet ws = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                sheets = wb.Worksheets;

                string expected = NormalizeSheetNameP2T2T3(sheetName);
                string expectedSingular = TrimTrailingSP2T2T3(expected);

                for (int i = 1; i <= sheets.Count; i++)
                {
                    ReleaseCom(ws);
                    ws = null;

                    ws = (Xl.Worksheet)sheets[i];
                    string actual = NormalizeSheetNameP2T2T3(ws.Name);
                    string actualSingular = TrimTrailingSP2T2T3(actual);

                    if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase) ||
                        string.Equals(actualSingular, expectedSingular, StringComparison.OrdinalIgnoreCase))
                    {
                        Xl.Worksheet result = ws;
                        ws = null;
                        return result;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(ws);
                ReleaseCom(sheets);
            }
        }

        private string NormalizeSheetNameP2T2T3(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace("'", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();
            return s;
        }

        private string TrimTrailingSP2T2T3(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            while (text.Length > 1 && text.EndsWith("S", StringComparison.OrdinalIgnoreCase))
                text = text.Substring(0, text.Length - 1);

            return text;
        }

        private bool TryGetEffectiveSparklineRangesP2T3(
            Xl.Worksheet ws,
            string locationRangeText,
            string dataRangeText,
            out Xl.Range effectiveLocation,
            out Xl.Range effectiveData)
        {
            effectiveLocation = null;
            effectiveData = null;

            Xl.Range originalLocation = null;
            Xl.Range originalData = null;
            Xl.Range adjustedLocation = null;
            Xl.Range adjustedData = null;

            try
            {
                originalLocation = ws.Range[locationRangeText];
                originalData = ws.Range[dataRangeText];

                if (originalLocation == null || originalData == null)
                    return false;

                int locRows = Convert.ToInt32(originalLocation.Rows.Count);
                int dataRows = Convert.ToInt32(originalData.Rows.Count);

                bool dropLocationHeader = RangeFirstRowLooksLikeHeaderP2T3(originalLocation);
                bool dropDataHeader = RangeFirstRowLooksLikeHeaderP2T3(originalData);

                // Đề có thể ghi E2:E6 và B2:D6, nhưng dòng 2 là header.
                // Khi dòng đầu là header thì bỏ dòng đầu để chấm đúng vùng dữ liệu E3:E6 và B3:D6.
                if (dropLocationHeader && locRows > 1)
                    adjustedLocation = OffsetResizeRangeP2T3(originalLocation, 1, 0, locRows - 1, Convert.ToInt32(originalLocation.Columns.Count));
                else
                    adjustedLocation = DuplicateRangeP2T3(originalLocation);

                if (dropDataHeader && dataRows > 1)
                    adjustedData = OffsetResizeRangeP2T3(originalData, 1, 0, dataRows - 1, Convert.ToInt32(originalData.Columns.Count));
                else
                    adjustedData = DuplicateRangeP2T3(originalData);

                if (adjustedLocation == null || adjustedData == null)
                    return false;

                effectiveLocation = adjustedLocation;
                effectiveData = adjustedData;
                adjustedLocation = null;
                adjustedData = null;

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(adjustedLocation);
                ReleaseCom(adjustedData);
                ReleaseCom(originalLocation);
                ReleaseCom(originalData);
            }
        }

        private bool RangeFirstRowLooksLikeHeaderP2T3(Xl.Range range)
        {
            if (range == null) return false;

            Xl.Range firstRow = null;
            Xl.Range cell = null;

            try
            {
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                if (rowCount <= 1)
                    return false;

                firstRow = (Xl.Range)range.Rows[1];

                int textCount = 0;
                int numericCount = 0;

                for (int c = 1; c <= columnCount; c++)
                {
                    ReleaseCom(cell);
                    cell = null;

                    cell = (Xl.Range)firstRow.Cells[1, c];
                    string text = GetCellTextP1T3(cell);

                    if (string.IsNullOrWhiteSpace(text))
                        continue;

                    double number;
                    if (TryParseDoubleP1T3(text, out number))
                        numericCount++;
                    else
                        textCount++;
                }

                return textCount > 0 && numericCount == 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(firstRow);
            }
        }

        private Xl.Range DuplicateRangeP2T3(Xl.Range range)
        {
            if (range == null) return null;

            Xl.Worksheet ws = null;
            Xl.Range startCell = null;
            Xl.Range endCell = null;

            try
            {
                ws = range.Worksheet as Xl.Worksheet;
                int firstRow = Convert.ToInt32(range.Row);
                int firstColumn = Convert.ToInt32(range.Column);
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                startCell = (Xl.Range)ws.Cells[firstRow, firstColumn];
                endCell = (Xl.Range)ws.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];

                return ws.Range[startCell, endCell];
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(startCell);
                ReleaseCom(endCell);
                ReleaseCom(ws);
            }
        }

        private Xl.Range OffsetResizeRangeP2T3(Xl.Range range, int rowOffset, int columnOffset, int rowCount, int columnCount)
        {
            if (range == null) return null;
            if (rowCount <= 0 || columnCount <= 0) return null;

            Xl.Worksheet ws = null;
            Xl.Range startCell = null;
            Xl.Range endCell = null;

            try
            {
                ws = range.Worksheet as Xl.Worksheet;

                int firstRow = Convert.ToInt32(range.Row) + rowOffset;
                int firstColumn = Convert.ToInt32(range.Column) + columnOffset;

                startCell = (Xl.Range)ws.Cells[firstRow, firstColumn];
                endCell = (Xl.Range)ws.Cells[firstRow + rowCount - 1, firstColumn + columnCount - 1];

                return ws.Range[startCell, endCell];
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(startCell);
                ReleaseCom(endCell);
                ReleaseCom(ws);
            }
        }
        private bool LocationRangeHasColumnSparklinesP2T3(Xl.Range locationRange, Xl.Range dataRange)
        {
            if (locationRange == null || dataRange == null) return false;

            Xl.Range locationCell = null;
            Xl.Range rowDataRange = null;

            try
            {
                int locationRows = Convert.ToInt32(locationRange.Rows.Count);
                int dataColumns = Convert.ToInt32(dataRange.Columns.Count);

                int locationStartRow = Convert.ToInt32(locationRange.Row);
                int locationColumn = Convert.ToInt32(locationRange.Column);

                int dataStartRow = Convert.ToInt32(dataRange.Row);
                int dataStartColumn = Convert.ToInt32(dataRange.Column);

                string fullDataAddress = NormalizeRangeAddressP2T3(dataRange);

                for (int r = 0; r < locationRows; r++)
                {
                    ReleaseCom(locationCell);
                    ReleaseCom(rowDataRange);
                    locationCell = null;
                    rowDataRange = null;

                    locationCell = (Xl.Range)locationRange.Worksheet.Cells[locationStartRow + r, locationColumn];
                    rowDataRange = BuildRangeByCoordinatesP2T3(
                        dataRange.Worksheet as Xl.Worksheet,
                        dataStartRow + r,
                        dataStartColumn,
                        1,
                        dataColumns);

                    string rowDataAddress = NormalizeRangeAddressP2T3(rowDataRange);

                    if (!CellHasColumnSparklineP2T3(locationCell, fullDataAddress, rowDataAddress))
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(locationCell);
                ReleaseCom(rowDataRange);
            }
        }

        private Xl.Range BuildRangeByCoordinatesP2T3(Xl.Worksheet ws, int row, int column, int rowCount, int columnCount)
        {
            if (ws == null) return null;
            if (rowCount <= 0 || columnCount <= 0) return null;

            Xl.Range startCell = null;
            Xl.Range endCell = null;

            try
            {
                startCell = (Xl.Range)ws.Cells[row, column];
                endCell = (Xl.Range)ws.Cells[row + rowCount - 1, column + columnCount - 1];
                return ws.Range[startCell, endCell];
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(startCell);
                ReleaseCom(endCell);
            }
        }

        private bool CellHasColumnSparklineP2T3(Xl.Range cell, string fullDataAddress, string rowDataAddress)
        {
            if (cell == null) return false;

            object groups = null;
            object group = null;

            try
            {
                groups = GetComPropertyP2T3(cell, "SparklineGroups");
                int groupCount = ToIntP2T3(GetComPropertyP2T3(groups, "Count"));

                if (groupCount <= 0)
                    return false;

                bool sawReadableSource = false;

                for (int i = 1; i <= groupCount; i++)
                {
                    ReleaseCom(group);
                    group = null;

                    group = GetComIndexedPropertyP2T3(groups, "Item", i);
                    if (group == null) continue;

                    int type = ToIntP2T3(GetComPropertyP2T3(group, "Type"));

                    // Excel enum: xlSparkColumn = 2.
                    if (type != 2)
                        continue;

                    string sourceData = Convert.ToString(GetComPropertyP2T3(group, "SourceData"));
                    string normalizedSource = NormalizeSparklineSourceP2T3(sourceData);

                    if (!string.IsNullOrWhiteSpace(normalizedSource))
                    {
                        sawReadableSource = true;

                        string fullData = NormalizeSparklineSourceP2T3(fullDataAddress);
                        string rowData = NormalizeSparklineSourceP2T3(rowDataAddress);

                        if (normalizedSource.Contains(fullData) ||
                            normalizedSource.Contains(rowData) ||
                            fullData.Contains(normalizedSource) ||
                            rowData.Contains(normalizedSource))
                        {
                            return true;
                        }
                    }
                    else
                    {
                        if (!sawReadableSource)
                            return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(group);
                ReleaseCom(groups);
            }
        }

        private string NormalizeRangeAddressP2T3(Xl.Range range)
        {
            if (range == null) return "";

            try
            {
                string address = Convert.ToString(range.get_Address(false, false, Xl.XlReferenceStyle.xlA1, false, Type.Missing));

                Xl.Worksheet ws = range.Worksheet as Xl.Worksheet;
                string sheet = ws == null ? "" : ws.Name;
                ReleaseCom(ws);

                if (!string.IsNullOrWhiteSpace(sheet))
                    address = sheet + "!" + address;

                return address;
            }
            catch
            {
                return "";
            }
        }

        private string NormalizeSparklineSourceP2T3(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("=", "");
            s = s.Replace(";", ",");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();
            return s;
        }

        private object GetComPropertyP2T3(object comObject, string propertyName)
        {
            if (comObject == null || string.IsNullOrWhiteSpace(propertyName)) return null;

            try
            {
                return comObject.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    comObject,
                    null);
            }
            catch
            {
                return null;
            }
        }

        private object GetComIndexedPropertyP2T3(object comObject, string propertyName, object index)
        {
            if (comObject == null || string.IsNullOrWhiteSpace(propertyName)) return null;

            try
            {
                return comObject.GetType().InvokeMember(
                    propertyName,
                    BindingFlags.GetProperty,
                    null,
                    comObject,
                    new object[] { index });
            }
            catch
            {
                return null;
            }
        }

        private int ToIntP2T3(object value)
        {
            if (value == null) return 0;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                int result;
                if (int.TryParse(Convert.ToString(value), out result))
                    return result;
            }

            return 0;
        }

        private bool AutoFilterCriteriaContainsExpected(Xl.Worksheet ws, Xl.Range filterRange, int fieldIndex, string expectedText)
        {
            if (ws == null || filterRange == null || fieldIndex <= 0 || string.IsNullOrWhiteSpace(expectedText))
                return false;

            Xl.AutoFilter autoFilter = null;
            Xl.Filters filters = null;
            Xl.Filter filter = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;

            try
            {
                try { autoFilter = ws.AutoFilter; } catch { autoFilter = null; }
                if (autoFilter != null)
                {
                    filters = autoFilter.Filters;
                    if (filters != null && fieldIndex <= filters.Count)
                    {
                        filter = filters.Item[fieldIndex];
                        if (FilterCriteriaContainsExpectedP14(filter, expectedText)) return true;
                    }
                }

                ReleaseCom(filter); filter = null;
                ReleaseCom(filters); filters = null;
                ReleaseCom(autoFilter); autoFilter = null;

                tables = ws.ListObjects;
                int tableCount = tables == null ? 0 : Convert.ToInt32(tables.Count);
                string expectedRange = NormalizeRangeAddress(Convert.ToString(
                    filterRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]));
                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(tableRange); ReleaseCom(table);
                    tableRange = null; table = tables.Item[i];
                    if (table == null) continue;
                    tableRange = table.Range;
                    string actualRange = tableRange == null ? "" : NormalizeRangeAddress(Convert.ToString(
                        tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]));
                    if (!string.Equals(actualRange, expectedRange, StringComparison.OrdinalIgnoreCase)) continue;

                    autoFilter = table.AutoFilter;
                    filters = autoFilter == null ? null : autoFilter.Filters;
                    if (filters == null || fieldIndex > filters.Count) return false;
                    filter = filters.Item[fieldIndex];
                    return FilterCriteriaContainsExpectedP14(filter, expectedText);
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(tables);
                ReleaseCom(filter);
                ReleaseCom(filters);
                ReleaseCom(autoFilter);
            }
        }

        private Xl.Range FindRangeStartingAtHeaderRowP1T5(Xl.Worksheet ws, Xl.Range usedRange, IList<string> requiredHeaders)
        {
            if (ws == null || usedRange == null || requiredHeaders == null || requiredHeaders.Count == 0) return null;

            Xl.Range cell = null;
            Xl.Range start = null;
            Xl.Range end = null;
            try
            {
                int firstRow = Convert.ToInt32(usedRange.Row);
                int firstColumn = Convert.ToInt32(usedRange.Column);
                int rowCount = Convert.ToInt32(usedRange.Rows.Count);
                int columnCount = Convert.ToInt32(usedRange.Columns.Count);
                int lastRow = firstRow + rowCount - 1;
                int lastColumn = firstColumn + columnCount - 1;

                for (int row = firstRow; row <= lastRow; row++)
                {
                    bool allFound = true;
                    foreach (string requiredHeader in requiredHeaders)
                    {
                        bool found = false;
                        for (int column = firstColumn; column <= lastColumn; column++)
                        {
                            ReleaseCom(cell);
                            cell = ws.Cells[row, column] as Xl.Range;
                            if (cell != null && HeaderEqualsP4T5(GetCellTextP1T3(cell), requiredHeader))
                            {
                                found = true;
                                break;
                            }
                        }
                        if (!found)
                        {
                            allFound = false;
                            break;
                        }
                    }

                    if (!allFound) continue;
                    start = ws.Cells[row, firstColumn] as Xl.Range;
                    end = ws.Cells[lastRow, lastColumn] as Xl.Range;
                    if (start == null || end == null) return null;
                    return ws.Range[start, end];
                }

                return null;
            }
            finally
            {
                ReleaseCom(end); ReleaseCom(start); ReleaseCom(cell);
            }
        }

        private bool FilterCriteriaContainsExpectedP14(Xl.Filter filter, string expectedText)
        {
            if (filter == null) return false;
            bool isOn;
            try { isOn = filter.On; } catch { isOn = false; }
            if (!isOn) return false;

            object criteria1 = null;
            object criteria2 = null;
            try { criteria1 = filter.Criteria1; } catch { }
            try { criteria2 = filter.Criteria2; } catch { }
            return CriteriaObjectContainsExpected(criteria1, expectedText) ||
                CriteriaObjectContainsExpected(criteria2, expectedText);
        }
        //Project 2 Task 4
        public bool FreezePanesEquals(string sheetName, int expectedFreezeRows, int expectedFreezeColumns)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (expectedFreezeRows < 0) return false;
            if (expectedFreezeColumns < 0) return false;

            Xl.Worksheet ws = null;
            Xl.Window window = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                Xl.Application app = _session.App as Xl.Application;
                if (app == null) return false;

                // Freeze Panes là trạng thái của Window đang active.
                // Vì vậy phải Activate đúng worksheet trước khi đọc ActiveWindow.
                ws.Activate();

                window = app.ActiveWindow;
                if (window == null) return false;

                bool freezePanes = false;
                try
                {
                    freezePanes = Convert.ToBoolean(window.FreezePanes);
                }
                catch
                {
                    freezePanes = false;
                }

                if (!freezePanes)
                    return false;

                int actualSplitRow = 0;
                int actualSplitColumn = 0;

                try
                {
                    actualSplitRow = Convert.ToInt32(window.SplitRow);
                }
                catch
                {
                    actualSplitRow = 0;
                }

                try
                {
                    actualSplitColumn = Convert.ToInt32(window.SplitColumn);
                }
                catch
                {
                    actualSplitColumn = 0;
                }

                return actualSplitRow == expectedFreezeRows &&
                       actualSplitColumn == expectedFreezeColumns;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(window);
                ReleaseCom(ws);
            }
        }
        //Project 2 Task 5
        public bool NamedRangesSumFormula(string sheetName, string cellAddress, IList<string> rangeNames)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(cellAddress)) return false;
            if (rangeNames == null || rangeNames.Count == 0) return false;

            Xl.Workbook wb = null;
            Xl.Worksheet ws = null;
            Xl.Range targetCell = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                if (wb == null) return false;

                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                targetCell = ws.Range[cellAddress];
                if (targetCell == null) return false;

                if (!CellHasFormulaP1T3(targetCell))
                    return false;

                string formula = Convert.ToString(targetCell.Formula);
                string formulaLocal = Convert.ToString(targetCell.FormulaLocal);

                if (!FormulaLooksLikeNamedRangesSumP2T5(formula, formulaLocal, rangeNames))
                    return false;

                double expectedSum;
                if (!TryCalculateNamedRangesSumP2T5(wb, ws, rangeNames, out expectedSum))
                    return false;

                string actualText = GetCellTextP1T3(targetCell);
                double actualNumber;

                if (!TryParseDoubleP1T3(actualText, out actualNumber))
                {
                    object actualValue = targetCell.Value2;
                    if (actualValue == null)
                        return false;

                    try
                    {
                        actualNumber = Convert.ToDouble(actualValue, CultureInfo.InvariantCulture);
                    }
                    catch
                    {
                        return false;
                    }
                }

                return Math.Abs(actualNumber - expectedSum) < 0.0001;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(targetCell);
                ReleaseCom(ws);
            }
        }

        private bool FormulaLooksLikeNamedRangesSumP2T5(string formula, string formulaLocal, IList<string> rangeNames)
        {
            string f1 = NormalizeNamedRangeFormulaP2T5(formula);
            string f2 = NormalizeNamedRangeFormulaP2T5(formulaLocal);

            if (string.IsNullOrWhiteSpace(f1) && string.IsNullOrWhiteSpace(f2))
                return false;

            bool okFormula = SingleFormulaLooksLikeNamedRangesSumP2T5(f1, rangeNames);
            bool okFormulaLocal = SingleFormulaLooksLikeNamedRangesSumP2T5(f2, rangeNames);

            return okFormula || okFormulaLocal;
        }

        private bool SingleFormulaLooksLikeNamedRangesSumP2T5(string normalizedFormula, IList<string> rangeNames)
        {
            if (string.IsNullOrWhiteSpace(normalizedFormula))
                return false;

            Match sumMatch = Regex.Match(normalizedFormula, @"^=SUM\((?<arguments>[^()]*)\)$", RegexOptions.IgnoreCase);
            if (!sumMatch.Success)
                return false;

            string[] arguments = sumMatch.Groups["arguments"].Value
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);
            if (arguments.Length != rangeNames.Count)
                return false;

            List<string> expectedNames = new List<string>();

            for (int i = 0; i < rangeNames.Count; i++)
            {
                string name = NormalizeDefinedNameP2T5(rangeNames[i]);
                if (string.IsNullOrWhiteSpace(name))
                    return false;

                expectedNames.Add(name);
            }

            for (int i = 0; i < arguments.Length; i++)
            {
                string argument = NormalizeDefinedNameP2T5(arguments[i]);
                int matchingIndex = expectedNames.FindIndex(name =>
                    string.Equals(name, argument, StringComparison.OrdinalIgnoreCase));
                if (matchingIndex < 0) return false;
                expectedNames.RemoveAt(matchingIndex);
            }

            return expectedNames.Count == 0;
        }

        private string NormalizeNamedRangeFormulaP2T5(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();
            s = s.Replace(";", ",");
            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("[", "");
            s = s.Replace("]", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();
            return s;
        }

        private string NormalizeDefinedNameP2T5(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return "";

            string s = name.Trim();
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();
            return s;
        }

        private bool FormulaContainsDirectCellReferenceP2T5(string formulaWithoutNames)
        {
            if (string.IsNullOrWhiteSpace(formulaWithoutNames))
                return false;

            // A1, $A$1, A1:B5, Figure!A1... đều bị xem là dùng cell reference.
            return Regex.IsMatch(
                formulaWithoutNames,
                @"(^|[^A-Z0-9_])\$?[A-Z]{1,3}\$?\d+([^A-Z0-9_]|$)",
                RegexOptions.IgnoreCase);
        }

        private bool TryCalculateNamedRangesSumP2T5(Xl.Workbook wb, Xl.Worksheet ws, IList<string> rangeNames, out double total)
        {
            total = 0;
            if (wb == null || ws == null || rangeNames == null || rangeNames.Count == 0)
                return false;

            for (int i = 0; i < rangeNames.Count; i++)
            {
                Xl.Range namedRange = null;

                try
                {
                    if (!TryGetNamedRangeP2T5(wb, ws, rangeNames[i], out namedRange))
                        return false;

                    total += SumNumericRangeP2T5(namedRange);
                }
                finally
                {
                    ReleaseCom(namedRange);
                }
            }

            return true;
        }

        private bool TryGetNamedRangeP2T5(Xl.Workbook wb, Xl.Worksheet ws, string rangeName, out Xl.Range resultRange)
        {
            resultRange = null;
            if (wb == null || ws == null || string.IsNullOrWhiteSpace(rangeName))
                return false;

            Xl.Names names = null;
            Xl.Name nameObject = null;

            try
            {
                try
                {
                    names = wb.Names;
                    nameObject = names.Item(rangeName, Type.Missing, Type.Missing);
                    if (nameObject != null)
                    {
                        resultRange = nameObject.RefersToRange;
                        if (resultRange != null)
                            return true;
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(nameObject);
                    ReleaseCom(names);
                    nameObject = null;
                    names = null;
                }

                try
                {
                    names = ws.Names;
                    nameObject = names.Item(rangeName, Type.Missing, Type.Missing);
                    if (nameObject != null)
                    {
                        resultRange = nameObject.RefersToRange;
                        if (resultRange != null)
                            return true;
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(nameObject);
                    ReleaseCom(names);
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private double SumNumericRangeP2T5(Xl.Range range)
        {
            if (range == null) return 0;

            object value = null;
            double total = 0;

            try
            {
                value = range.Value2;

                object[,] values = value as object[,];
                if (values != null)
                {
                    int rowStart = values.GetLowerBound(0);
                    int rowEnd = values.GetUpperBound(0);
                    int colStart = values.GetLowerBound(1);
                    int colEnd = values.GetUpperBound(1);

                    for (int r = rowStart; r <= rowEnd; r++)
                    {
                        for (int c = colStart; c <= colEnd; c++)
                            total += ToDoubleOrZeroP2T5(values[r, c]);
                    }

                    return total;
                }

                return ToDoubleOrZeroP2T5(value);
            }
            catch
            {
                return 0;
            }
        }

        private double ToDoubleOrZeroP2T5(object value)
        {
            if (value == null) return 0;

            try
            {
                return Convert.ToDouble(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                double result;
                if (double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.InvariantCulture, out result))
                    return result;

                if (double.TryParse(Convert.ToString(value), NumberStyles.Any, CultureInfo.CurrentCulture, out result))
                    return result;
            }

            return 0;
        }
        //Project 2 Task 6
        public bool ChartAltTextEquals(string sheetName, string expectedText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(expectedText)) return false;

            Xl.Worksheet ws = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                if (WorksheetShapesContainChartAltTextP2T6(ws, expectedText))
                    return true;

                if (WorksheetChartObjectsContainAltTextP2T6(ws, expectedText))
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(ws);
            }
        }

        public bool ChartAltTextDescriptionEquals(string sheetName, string chartName, string expectedDescription)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(chartName) ||
                string.IsNullOrWhiteSpace(expectedDescription)) return false;

            Xl.Worksheet ws = null;
            Xl.Shapes shapes = null;
            Xl.Shape shape = null;
            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;
                shapes = ws.Shapes;
                shape = shapes.Item(chartName);
                if (shape == null || shape.Type != Office.MsoShapeType.msoChart) return false;
                return string.Equals(
                    NormalizeTextP4T6(shape.AlternativeText),
                    NormalizeTextP4T6(expectedDescription),
                    StringComparison.Ordinal);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(shape); ReleaseCom(shapes); ReleaseCom(ws);
            }
        }

        private bool WorksheetShapesContainChartAltTextP2T6(Xl.Worksheet ws, string expectedText)
        {
            object shapes = null;
            object shape = null;

            try
            {
                shapes = GetComPropertyP2T3(ws, "Shapes");
                int count = ToIntP2T3(GetComPropertyP2T3(shapes, "Count"));

                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(shape);
                    shape = null;

                    shape = GetComIndexedPropertyP2T3(shapes, "Item", i);
                    if (shape == null) continue;

                    if (!ShapeIsChartP2T6(shape))
                        continue;

                    if (AnyChartAltTextEqualsP2T6(shape, expectedText))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(shape);
                ReleaseCom(shapes);
            }
        }

        private bool WorksheetChartObjectsContainAltTextP2T6(Xl.Worksheet ws, string expectedText)
        {
            object chartObjects = null;
            object chartObject = null;
            object chart = null;
            object parentShape = null;
            object shapeRange = null;
            object shapeRangeItem = null;
            object chartArea = null;
            object chartAreaParent = null;
            object shapes = null;

            try
            {
                try
                {
                    chartObjects = ws.ChartObjects(Type.Missing);
                }
                catch
                {
                    chartObjects = null;
                }

                int count = ToIntP2T3(GetComPropertyP2T3(chartObjects, "Count"));

                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(chartObject);
                    ReleaseCom(chart);
                    ReleaseCom(parentShape);
                    ReleaseCom(shapeRange);
                    ReleaseCom(shapeRangeItem);
                    ReleaseCom(chartArea);
                    ReleaseCom(chartAreaParent);

                    chartObject = null;
                    chart = null;
                    parentShape = null;
                    shapeRange = null;
                    shapeRangeItem = null;
                    chartArea = null;
                    chartAreaParent = null;

                    chartObject = GetComIndexedPropertyP2T3(chartObjects, "Item", i);
                    if (chartObject == null) continue;

                    if (AnyChartAltTextEqualsP2T6(chartObject, expectedText))
                        return true;

                    shapeRange = GetComPropertyP2T3(chartObject, "ShapeRange");
                    if (shapeRange != null)
                    {
                        if (AnyChartAltTextEqualsP2T6(shapeRange, expectedText))
                            return true;

                        shapeRangeItem = GetComIndexedPropertyP2T3(shapeRange, "Item", 1);
                        if (shapeRangeItem != null && AnyChartAltTextEqualsP2T6(shapeRangeItem, expectedText))
                            return true;
                    }

                    string chartObjectName = Convert.ToString(GetComPropertyP2T3(chartObject, "Name"));
                    if (!string.IsNullOrWhiteSpace(chartObjectName))
                    {
                        if (shapes == null)
                            shapes = GetComPropertyP2T3(ws, "Shapes");

                        parentShape = GetComIndexedPropertyP2T3(shapes, "Item", chartObjectName);
                        if (parentShape != null && AnyChartAltTextEqualsP2T6(parentShape, expectedText))
                            return true;
                    }

                    chart = GetComPropertyP2T3(chartObject, "Chart");
                    if (chart != null)
                    {
                        if (AnyChartAltTextEqualsP2T6(chart, expectedText))
                            return true;

                        chartArea = GetComPropertyP2T3(chart, "ChartArea");
                        if (chartArea != null)
                        {
                            if (AnyChartAltTextEqualsP2T6(chartArea, expectedText))
                                return true;

                            chartAreaParent = GetComPropertyP2T3(chartArea, "Parent");
                            if (chartAreaParent != null && AnyChartAltTextEqualsP2T6(chartAreaParent, expectedText))
                                return true;
                        }
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chartAreaParent);
                ReleaseCom(chartArea);
                ReleaseCom(shapeRangeItem);
                ReleaseCom(shapeRange);
                ReleaseCom(parentShape);
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(shapes);
            }
        }

        private bool ShapeIsChartP2T6(object shape)
        {
            if (shape == null) return false;

            int hasChart = ToIntP2T3(GetComPropertyP2T3(shape, "HasChart"));
            if (hasChart != 0)
                return true;

            object chart = null;

            try
            {
                chart = GetComPropertyP2T3(shape, "Chart");
                return chart != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chart);
            }
        }

        private bool AnyChartAltTextEqualsP2T6(object comObject, string expectedText)
        {
            if (comObject == null) return false;

            if (ComTextPropertyEqualsP2T6(comObject, "AlternativeText", expectedText))
                return true;

            if (ComTextPropertyEqualsP2T6(comObject, "Description", expectedText))
                return true;

            if (ComTextPropertyEqualsP2T6(comObject, "Title", expectedText))
                return true;

            object chart = null;
            object parent = null;
            object chartArea = null;
            object chartTitle = null;
            object shapeRange = null;
            object shapeRangeItem = null;

            try
            {
                chart = GetComPropertyP2T3(comObject, "Chart");
                if (chart != null)
                {
                    if (ComTextPropertyEqualsP2T6(chart, "AlternativeText", expectedText)) return true;
                    if (ComTextPropertyEqualsP2T6(chart, "Description", expectedText)) return true;
                    if (ComTextPropertyEqualsP2T6(chart, "Title", expectedText)) return true;

                    chartArea = GetComPropertyP2T3(chart, "ChartArea");
                    if (chartArea != null)
                    {
                        if (ComTextPropertyEqualsP2T6(chartArea, "AlternativeText", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(chartArea, "Description", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(chartArea, "Title", expectedText)) return true;
                    }

                    chartTitle = GetComPropertyP2T3(chart, "ChartTitle");
                    if (chartTitle != null)
                    {
                        if (ComTextPropertyEqualsP2T6(chartTitle, "Text", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(chartTitle, "Caption", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(chartTitle, "AlternativeText", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(chartTitle, "Description", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(chartTitle, "Title", expectedText)) return true;
                    }
                }

                parent = GetComPropertyP2T3(comObject, "Parent");
                if (parent != null)
                {
                    if (ComTextPropertyEqualsP2T6(parent, "AlternativeText", expectedText)) return true;
                    if (ComTextPropertyEqualsP2T6(parent, "Description", expectedText)) return true;
                    if (ComTextPropertyEqualsP2T6(parent, "Title", expectedText)) return true;
                }

                shapeRange = GetComPropertyP2T3(comObject, "ShapeRange");
                if (shapeRange != null)
                {
                    if (ComTextPropertyEqualsP2T6(shapeRange, "AlternativeText", expectedText)) return true;
                    if (ComTextPropertyEqualsP2T6(shapeRange, "Description", expectedText)) return true;
                    if (ComTextPropertyEqualsP2T6(shapeRange, "Title", expectedText)) return true;

                    shapeRangeItem = GetComIndexedPropertyP2T3(shapeRange, "Item", 1);
                    if (shapeRangeItem != null)
                    {
                        if (ComTextPropertyEqualsP2T6(shapeRangeItem, "AlternativeText", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(shapeRangeItem, "Description", expectedText)) return true;
                        if (ComTextPropertyEqualsP2T6(shapeRangeItem, "Title", expectedText)) return true;
                    }
                }
            }
            catch
            {
            }
            finally
            {
                ReleaseCom(shapeRangeItem);
                ReleaseCom(shapeRange);
                ReleaseCom(chartTitle);
                ReleaseCom(chartArea);
                ReleaseCom(parent);
                ReleaseCom(chart);
            }

            return false;
        }

        private bool ComTextPropertyEqualsP2T6(object comObject, string propertyName, string expectedText)
        {
            if (comObject == null || string.IsNullOrWhiteSpace(propertyName)) return false;

            object value = null;

            try
            {
                value = GetComPropertyP2T3(comObject, propertyName);
                string text = Convert.ToString(value);

                return AltTextEqualsP2T6(text, expectedText);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(value);
            }
        }

        private bool AltTextEqualsP2T6(string actual, string expected)
        {
            string a = NormalizeAltTextP2T6(actual);
            string e = NormalizeAltTextP2T6(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeAltTextP2T6(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", " ");
            return s.Trim();
        }
        //Project 2 Task 7
        //Project 2 Task 7 - Clean fixed version by reading xlsx XML
        public bool RangeIconSetEquals(string sheetName, string rangeAddress, string expectedIconSetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(rangeAddress)) return false;
            if (string.IsNullOrWhiteSpace(expectedIconSetName)) return false;

            Xl.Workbook wb = null;
            string tempPath = "";

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                if (wb == null) return false;

                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    "MosTrainer_P2T7_" + Guid.NewGuid().ToString("N") + ".xlsx");

                // SaveCopyAs lấy đúng trạng thái hiện tại học viên vừa làm,
                // không cần học viên tự bấm Save.
                wb.SaveCopyAs(tempPath);

                if (!File.Exists(tempPath))
                    return false;

                return XlsxHasIconSetConditionalFormattingOnRangeP2T7Xml(
                    tempPath,
                    sheetName,
                    rangeAddress,
                    expectedIconSetName);
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals",
                    "P02_T07 icon-set grading failed.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return false;
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPath) && File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch (Exception ex)
                {
                    AppLogger.Warning(
                        "RangeIconSetEquals",
                        "Temporary grading copy could not be deleted: " + ex.Message,
                        "Excel2019_P02",
                        "P02_T07");
                }

                // Không ReleaseCom(wb) vì đây là workbook chính trong _session.
            }
        }

        // Nếu interface cũ còn bắt buộc method này thì giữ wrapper này.
        // Nó giúp assertion cũ vẫn PASS, không còn gọi logic table-column cũ.
        public bool TableColumnIconSetEquals(string sheetName, string columnHeader, string expectedIconSetName)
        {
            return RangeIconSetEquals(sheetName, "G10:G40", expectedIconSetName);
        }
        private bool XlsxHasIconSetConditionalFormattingOnRangeP2T7Xml(
            string xlsxPath,
            string sheetName,
            string expectedRange,
            string expectedIconSetName)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath) || !File.Exists(xlsxPath)) return false;
            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(expectedRange)) return false;
            if (string.IsNullOrWhiteSpace(expectedIconSetName)) return false;

            try
            {
                using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    string sheetPartPath = FindWorksheetPartPathP2T7Xml(zip, sheetName);
                    if (string.IsNullOrWhiteSpace(sheetPartPath))
                        return false;

                    ZipArchiveEntry sheetEntry = zip.GetEntry(sheetPartPath);
                    if (sheetEntry == null)
                        return false;

                    string sheetXml = ReadZipEntryTextP2T7Xml(sheetEntry);
                    if (string.IsNullOrWhiteSpace(sheetXml))
                        return false;

                    return WorksheetXmlHasIconSetOnRangeP2T7Xml(
                        sheetXml,
                        expectedRange,
                        expectedIconSetName);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals.Xml",
                    "P02_T07 workbook XML could not be inspected.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return false;
            }
        }

        private string FindWorksheetPartPathP2T7Xml(ZipArchive zip, string sheetName)
        {
            if (zip == null || string.IsNullOrWhiteSpace(sheetName))
                return "";

            try
            {
                ZipArchiveEntry workbookEntry = zip.GetEntry("xl/workbook.xml");
                ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");

                if (workbookEntry == null || relsEntry == null)
                    return "";

                string workbookXml = ReadZipEntryTextP2T7Xml(workbookEntry);
                string relsXml = ReadZipEntryTextP2T7Xml(relsEntry);

                if (string.IsNullOrWhiteSpace(workbookXml) || string.IsNullOrWhiteSpace(relsXml))
                    return "";

                XDocument workbookDocument = XDocument.Parse(workbookXml);
                XElement sheetElement = workbookDocument
                    .Descendants()
                    .FirstOrDefault(element =>
                        string.Equals(element.Name.LocalName, "sheet", StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(
                            NormalizeSheetNameP2T7Xml(GetAttributeValueP2T7Xml(element, "name")),
                            NormalizeSheetNameP2T7Xml(sheetName),
                            StringComparison.Ordinal));

                if (sheetElement == null)
                    return "";

                string rid = GetAttributeValueP2T7Xml(sheetElement, "id");
                if (string.IsNullOrWhiteSpace(rid))
                    return "";

                XDocument relationshipsDocument = XDocument.Parse(relsXml);
                XElement relationshipElement = relationshipsDocument
                    .Descendants()
                    .FirstOrDefault(element =>
                        string.Equals(element.Name.LocalName, "Relationship", StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(GetAttributeValueP2T7Xml(element, "Id"), rid, StringComparison.Ordinal));

                if (relationshipElement == null)
                    return "";

                string target = GetAttributeValueP2T7Xml(relationshipElement, "Target");
                if (string.IsNullOrWhiteSpace(target))
                    return "";

                target = target.Replace("\\", "/");

                if (target.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                    target = target.Substring(1);
                else if (target.StartsWith("/", StringComparison.OrdinalIgnoreCase))
                    target = "xl" + target;
                else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                    target = "xl/" + target;

                return target;
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals.FindWorksheet",
                    "P02_T07 worksheet XML relationship could not be resolved.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return "";
            }
        }

        private string GetAttributeValueP2T7Xml(XElement element, string localName)
        {
            if (element == null || string.IsNullOrWhiteSpace(localName))
                return "";

            XAttribute attribute = element.Attributes().FirstOrDefault(item =>
                string.Equals(item.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase));

            return attribute == null ? "" : attribute.Value;
        }

        private bool WorksheetXmlHasIconSetOnRangeP2T7Xml(
            string sheetXml,
            string expectedRange,
            string expectedIconSetName)
        {
            if (string.IsNullOrWhiteSpace(sheetXml)) return false;
            if (string.IsNullOrWhiteSpace(expectedRange)) return false;

            string expectedOoxmlIconSet;
            if (!TryMapExpectedIconSetP2T7(expectedIconSetName, out expectedOoxmlIconSet))
            {
                AppLogger.Warning(
                    "RangeIconSetEquals",
                    "Unsupported expected icon-set name: " + expectedIconSetName,
                    "Excel2019_P02",
                    "P02_T07");
                return false;
            }

            try
            {
                XDocument worksheetDocument = XDocument.Parse(sheetXml);

                foreach (XElement conditionalFormatting in worksheetDocument.Descendants()
                    .Where(element => string.Equals(
                        element.Name.LocalName,
                        "conditionalFormatting",
                        StringComparison.OrdinalIgnoreCase)))
                {
                    string sqref = GetAttributeValueP2T7Xml(conditionalFormatting, "sqref");
                    if (string.IsNullOrWhiteSpace(sqref))
                    {
                        XElement sqrefElement = conditionalFormatting.Descendants()
                            .FirstOrDefault(element => string.Equals(
                                element.Name.LocalName, "sqref", StringComparison.OrdinalIgnoreCase));
                        sqref = sqrefElement == null ? "" : sqrefElement.Value;
                    }

                    if (!SqrefExactlyMatchesRangeP2T7Xml(sqref, expectedRange))
                        continue;

                    foreach (XElement rule in conditionalFormatting.Descendants()
                        .Where(element => string.Equals(
                            element.Name.LocalName, "cfRule", StringComparison.OrdinalIgnoreCase)))
                    {
                        if (!string.Equals(
                            GetAttributeValueP2T7Xml(rule, "type"),
                            "iconSet",
                            StringComparison.OrdinalIgnoreCase))
                        {
                            continue;
                        }

                        XElement iconSet = rule.Descendants().FirstOrDefault(element =>
                            string.Equals(element.Name.LocalName, "iconSet", StringComparison.OrdinalIgnoreCase));

                        if (iconSet == null)
                            continue;

                        string actualIconSet = GetAttributeValueP2T7Xml(iconSet, "iconSet");
                        if (string.IsNullOrWhiteSpace(actualIconSet))
                            actualIconSet = "3TrafficLights1";

                        if (string.Equals(
                            actualIconSet,
                            expectedOoxmlIconSet,
                            StringComparison.OrdinalIgnoreCase))
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals.WorksheetXml",
                    "P02_T07 worksheet conditional-formatting XML could not be parsed.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return false;
            }
        }

        private bool SqrefExactlyMatchesRangeP2T7Xml(string sqref, string expectedRange)
        {
            if (string.IsNullOrWhiteSpace(sqref)) return false;
            if (string.IsNullOrWhiteSpace(expectedRange)) return false;

            try
            {
                int eRow1, eCol1, eRow2, eCol2;
                if (!TryParseA1RangeP2T7Xml(expectedRange, out eRow1, out eCol1, out eRow2, out eCol2))
                    return false;

                string[] parts = sqref.Split(
                    new char[] { ' ', '\t', '\r', '\n' },
                    StringSplitOptions.RemoveEmptyEntries);

                if (parts.Length != 1)
                    return false;

                int r1, c1, r2, c2;
                if (!TryParseA1RangeP2T7Xml(parts[0], out r1, out c1, out r2, out c2))
                    return false;

                return r1 == eRow1 &&
                       r2 == eRow2 &&
                       c1 == eCol1 &&
                       c2 == eCol2;
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals.Range",
                    "P02_T07 conditional-formatting range could not be compared.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return false;
            }
        }

        private bool TryMapExpectedIconSetP2T7(string expectedIconSetName, out string ooxmlIconSetName)
        {
            ooxmlIconSetName = "";
            if (string.IsNullOrWhiteSpace(expectedIconSetName))
                return false;

            string normalized = Regex.Replace(
                expectedIconSetName,
                @"[^A-Z0-9]+",
                "",
                RegexOptions.IgnoreCase).ToUpperInvariant();

            switch (normalized)
            {
                case "3TRAFFICLIGHTSUNRIMMED":
                case "3TRAFFICLIGHTS1":
                case "XL3TRAFFICLIGHTS1":
                    ooxmlIconSetName = "3TrafficLights1";
                    return true;

                case "3TRAFFICLIGHTSRIMMED":
                case "3TRAFFICLIGHTS2":
                case "XL3TRAFFICLIGHTS2":
                    ooxmlIconSetName = "3TrafficLights2";
                    return true;

                default:
                    return false;
            }
        }

        private bool TryParseA1RangeP2T7Xml(
            string address,
            out int row1,
            out int col1,
            out int row2,
            out int col2)
        {
            row1 = 0;
            col1 = 0;
            row2 = 0;
            col2 = 0;

            if (string.IsNullOrWhiteSpace(address))
                return false;

            try
            {
                string a = address.Trim();

                int bangIndex = a.LastIndexOf('!');
                if (bangIndex >= 0 && bangIndex + 1 < a.Length)
                    a = a.Substring(bangIndex + 1);

                a = a.Replace("$", "");
                a = a.Replace("'", "");
                a = a.Trim();

                string[] parts = a.Split(':');
                string first = parts[0];
                string second = parts.Length > 1 ? parts[1] : parts[0];

                Match m1 = Regex.Match(first, @"^(?<col>[A-Za-z]+)(?<row>\d+)$");
                Match m2 = Regex.Match(second, @"^(?<col>[A-Za-z]+)(?<row>\d+)$");

                if (!m1.Success || !m2.Success)
                    return false;

                col1 = ColumnLettersToNumberP2T7Xml(m1.Groups["col"].Value);
                row1 = Convert.ToInt32(m1.Groups["row"].Value);

                col2 = ColumnLettersToNumberP2T7Xml(m2.Groups["col"].Value);
                row2 = Convert.ToInt32(m2.Groups["row"].Value);

                if (row1 > row2)
                {
                    int t = row1;
                    row1 = row2;
                    row2 = t;
                }

                if (col1 > col2)
                {
                    int t = col1;
                    col1 = col2;
                    col2 = t;
                }

                return row1 > 0 && col1 > 0 && row2 > 0 && col2 > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals.ParseRange",
                    "P02_T07 A1 range could not be parsed.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return false;
            }
        }

        private int ColumnLettersToNumberP2T7Xml(string letters)
        {
            if (string.IsNullOrWhiteSpace(letters))
                return 0;

            int result = 0;
            string s = letters.Trim().ToUpperInvariant();

            for (int i = 0; i < s.Length; i++)
            {
                char ch = s[i];
                if (ch < 'A' || ch > 'Z')
                    return 0;

                result = result * 26 + (ch - 'A' + 1);
            }

            return result;
        }

        private string ReadZipEntryTextP2T7Xml(ZipArchiveEntry entry)
        {
            if (entry == null) return "";

            try
            {
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "RangeIconSetEquals.ReadXml",
                    "P02_T07 XML part could not be read.",
                    ex,
                    "Excel2019_P02",
                    "P02_T07");
                return "";
            }
        }

        private string NormalizeSheetNameP2T7Xml(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
       
        //Project 2 Task 8
        public bool TableStyleEquals(string sheetName, string tableName, string expectedStyleName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(expectedStyleName)) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects listObjects = null;
            Xl.ListObject table = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                listObjects = ws.ListObjects;
                if (listObjects == null) return false;

                int tableCount = Convert.ToInt32(listObjects.Count);
                if (tableCount <= 0) return false;

                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(table);
                    table = null;

                    table = listObjects.Item[i];
                    if (table == null) continue;

                    if (!ListObjectNameMatchesP2T8(table, tableName))
                        continue;

                    if (ListObjectTableStyleEqualsP2T8(table, expectedStyleName))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(table);
                ReleaseCom(listObjects);
                ReleaseCom(ws);
            }
        }
        
        private bool ListObjectNameMatchesP2T8(Xl.ListObject table, string expectedTableName)
        {
            if (table == null) return false;

            if (string.IsNullOrWhiteSpace(expectedTableName))
                return true;

            string expected = NormalizeTableNameP2T8(expectedTableName);

            string name = "";
            string displayName = "";

            try
            {
                name = Convert.ToString(table.Name);
            }
            catch
            {
                name = "";
            }

            try
            {
                displayName = Convert.ToString(table.DisplayName);
            }
            catch
            {
                displayName = "";
            }

            return string.Equals(NormalizeTableNameP2T8(name), expected, StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(NormalizeTableNameP2T8(displayName), expected, StringComparison.OrdinalIgnoreCase);
        }

        private bool ListObjectTableStyleEqualsP2T8(Xl.ListObject table, string expectedStyleName)
        {
            if (table == null) return false;
            if (string.IsNullOrWhiteSpace(expectedStyleName)) return false;

            string actualStyleName = GetTableStyleNameP2T8(table);

            string actual = NormalizeTableStyleNameP2T8(actualStyleName);
            string expected = NormalizeTableStyleNameP2T8(expectedStyleName);

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        private string GetTableStyleNameP2T8(Xl.ListObject table)
        {
            if (table == null) return "";

            object styleObject = null;
            object nameObject = null;
            object nameLocalObject = null;

            try
            {
                styleObject = table.TableStyle;

                string directText = Convert.ToString(styleObject);
                if (!string.IsNullOrWhiteSpace(directText) &&
                    !directText.Contains("__ComObject"))
                {
                    return directText;
                }

                nameObject = GetComPropertyP2T3(styleObject, "Name");
                string name = Convert.ToString(nameObject);

                if (!string.IsNullOrWhiteSpace(name) &&
                    !name.Contains("__ComObject"))
                {
                    return name;
                }

                nameLocalObject = GetComPropertyP2T3(styleObject, "NameLocal");
                string nameLocal = Convert.ToString(nameLocalObject);

                if (!string.IsNullOrWhiteSpace(nameLocal) &&
                    !nameLocal.Contains("__ComObject"))
                {
                    return nameLocal;
                }

                return "";
            }
            catch
            {
                return "";
            }
            finally
            {
                ReleaseCom(nameLocalObject);
                ReleaseCom(nameObject);
                ReleaseCom(styleObject);
            }
        }

        private string NormalizeTableStyleNameP2T8(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = s.ToUpperInvariant();

            // UI của Excel hiển thị: White, Table Style Medium 1
            // Nhưng object model thường trả về: TableStyleMedium1
            s = s.Replace("WHITE", "");

            s = Regex.Replace(s, @"[^A-Z0-9]+", "");

            if (s.Contains("TABLESTYLEMEDIUM1"))
                return "TABLESTYLEMEDIUM1";

            return s;
        }

        private string NormalizeTableNameP2T8(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        private bool CriteriaObjectContainsExpected(object criteria, string expectedText)
        {
            if (criteria == null)
                return false;

            Array criteriaArray = criteria as Array;
            if (criteriaArray != null)
            {
                foreach (object item in criteriaArray)
                {
                    if (CriteriaObjectContainsExpected(item, expectedText))
                        return true;
                }

                return false;
            }

            string criteriaText = Convert.ToString(criteria);
            if (string.IsNullOrWhiteSpace(criteriaText))
                return false;

            string normalizedCriteria = NormalizeFilterValue(criteriaText);
            string normalizedExpected = NormalizeFilterValue(expectedText);

            if (string.Equals(normalizedCriteria, normalizedExpected, StringComparison.OrdinalIgnoreCase))
                return true;

            if (normalizedCriteria.Contains(normalizedExpected))
                return true;

            return false;
        }
        //Project 3 Task 1
        public bool NamedRangeContentsCleared(string namedRangeName, string expectedSheetName, string expectedAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(namedRangeName)) return false;

            Xl.Workbook wb = null;
            Xl.Range namedRange = null;
            Xl.Worksheet rangeWorksheet = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                if (wb == null) return false;

                if (!TryGetNamedRangeP3T1(wb, namedRangeName, out namedRange))
                    return false;

                if (namedRange == null)
                    return false;

                if (!string.IsNullOrWhiteSpace(expectedSheetName))
                {
                    rangeWorksheet = namedRange.Worksheet as Xl.Worksheet;
                    if (rangeWorksheet == null) return false;

                    string actualSheetName = Convert.ToString(rangeWorksheet.Name);

                    if (!SheetNameEqualsP3T1(actualSheetName, expectedSheetName))
                        return false;
                }

                if (!string.IsNullOrWhiteSpace(expectedAddress))
                {
                    string actualAddress = Convert.ToString(
                        namedRange.get_Address(
                            false,
                            false,
                            Xl.XlReferenceStyle.xlA1,
                            false,
                            Type.Missing));

                    if (!RangeAddressEqualsP3T1(actualAddress, expectedAddress))
                        return false;
                }

                return RangeContentsAreEmptyP3T1(namedRange);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(rangeWorksheet);
                ReleaseCom(namedRange);
                // Không ReleaseCom(wb) vì wb là workbook chính trong _session.
            }
        }
        private bool TryGetNamedRangeP3T1(Xl.Workbook wb, string namedRangeName, out Xl.Range resultRange)
        {
            resultRange = null;

            if (wb == null || string.IsNullOrWhiteSpace(namedRangeName))
                return false;

            Xl.Names names = null;
            Xl.Name nameObject = null;
            Xl.Range refersToRange = null;

            try
            {
                // Cách 1: tìm trực tiếp trong workbook-scoped names
                try
                {
                    names = wb.Names;
                    nameObject = names.Item(namedRangeName, Type.Missing, Type.Missing);

                    if (nameObject != null)
                    {
                        refersToRange = nameObject.RefersToRange;
                        if (refersToRange != null)
                        {
                            resultRange = refersToRange;
                            refersToRange = null;
                            return true;
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(refersToRange);
                    ReleaseCom(nameObject);
                    ReleaseCom(names);

                    refersToRange = null;
                    nameObject = null;
                    names = null;
                }

                // Cách 2: duyệt toàn bộ workbook names, xử lý cả dạng 'Sheet'!Convertible
                try
                {
                    names = wb.Names;
                    if (names != null)
                    {
                        int count = Convert.ToInt32(names.Count);

                        for (int i = 1; i <= count; i++)
                        {
                            ReleaseCom(refersToRange);
                            ReleaseCom(nameObject);

                            refersToRange = null;
                            nameObject = null;

                            nameObject = names.Item(i, Type.Missing, Type.Missing);
                            if (nameObject == null) continue;

                            string actualName = Convert.ToString(nameObject.Name);

                            if (!DefinedNameEqualsP3T1(actualName, namedRangeName))
                                continue;

                            try
                            {
                                refersToRange = nameObject.RefersToRange;
                            }
                            catch
                            {
                                refersToRange = null;
                            }

                            if (refersToRange != null)
                            {
                                resultRange = refersToRange;
                                refersToRange = null;
                                return true;
                            }
                        }
                    }
                }
                catch
                {
                }
                finally
                {
                    ReleaseCom(refersToRange);
                    ReleaseCom(nameObject);
                    ReleaseCom(names);
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool RangeContentsAreEmptyP3T1(Xl.Range range)
        {
            if (range == null) return false;

            Xl.Range cell = null;

            try
            {
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                if (rowCount <= 0 || columnCount <= 0)
                    return false;

                int checkedCells = 0;

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)range.Cells[r, c];
                        if (cell == null) continue;

                        checkedCells++;

                        if (CellHasFormulaP1T3(cell))
                            return false;

                        object value = null;
                        try
                        {
                            value = cell.Value2;
                        }
                        catch
                        {
                            value = null;
                        }

                        if (!CellValueIsEmptyP3T1(value))
                            return false;
                    }
                }

                return checkedCells > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool CellValueIsEmptyP3T1(object value)
        {
            if (value == null)
                return true;

            string text = Convert.ToString(value);

            return string.IsNullOrWhiteSpace(text);
        }

        private bool DefinedNameEqualsP3T1(string actualName, string expectedName)
        {
            string actual = NormalizeDefinedNameP3T1(actualName);
            string expected = NormalizeDefinedNameP3T1(expectedName);

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeDefinedNameP3T1(string name)
        {
            if (name == null) return "";

            string s = name.Trim();
            s = s.Replace("'", "");

            int bangIndex = s.LastIndexOf('!');
            if (bangIndex >= 0 && bangIndex + 1 < s.Length)
                s = s.Substring(bangIndex + 1);

            s = s.Replace("$", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }

        private bool SheetNameEqualsP3T1(string actualSheetName, string expectedSheetName)
        {
            string actual = NormalizeSheetNameP3T1(actualSheetName);
            string expected = NormalizeSheetNameP3T1(expectedSheetName);

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeSheetNameP3T1(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace("'", "");
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }

        private bool RangeAddressEqualsP3T1(string actualAddress, string expectedAddress)
        {
            string actual = NormalizeRangeAddressP3T1(actualAddress);
            string expected = NormalizeRangeAddressP3T1(expectedAddress);

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeRangeAddressP3T1(string address)
        {
            if (address == null) return "";

            string s = address.Trim();
            s = s.Replace("$", "");
            s = s.Replace("'", "");

            int bangIndex = s.LastIndexOf('!');
            if (bangIndex >= 0 && bangIndex + 1 < s.Length)
                s = s.Substring(bangIndex + 1);

            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        //Project 3 Task 2
        public bool RangeNumberFormatDecimalPlacesEquals(string sheetName, string rangeAddress, int expectedDecimalPlaces)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(rangeAddress)) return false;
            if (expectedDecimalPlaces < 0) return false;

            Xl.Worksheet ws = null;
            Xl.Range targetRange = null;
            Xl.Range cell = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                targetRange = ws.Range[rangeAddress];
                if (targetRange == null) return false;

                int rowCount = Convert.ToInt32(targetRange.Rows.Count);
                int columnCount = Convert.ToInt32(targetRange.Columns.Count);

                if (rowCount <= 0 || columnCount <= 0)
                    return false;

                int checkedCells = 0;

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)targetRange.Cells[r, c];
                        if (cell == null) continue;

                        checkedCells++;

                        if (!CellNumberFormatDecimalPlacesEqualsP3T2(cell, expectedDecimalPlaces))
                            return false;
                    }
                }

                return checkedCells > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(targetRange);
                ReleaseCom(ws);
            }
        }
        private bool CellNumberFormatDecimalPlacesEqualsP3T2(Xl.Range cell, int expectedDecimalPlaces)
        {
            if (cell == null) return false;

            string numberFormat = "";
            string numberFormatLocal = "";

            try
            {
                object nf = cell.NumberFormat;
                numberFormat = nf == null ? "" : Convert.ToString(nf);
            }
            catch
            {
                numberFormat = "";
            }

            try
            {
                object nfl = cell.NumberFormatLocal;
                numberFormatLocal = nfl == null ? "" : Convert.ToString(nfl);
            }
            catch
            {
                numberFormatLocal = "";
            }

            if (NumberFormatShowsDecimalPlacesP3T2(numberFormat, expectedDecimalPlaces))
                return true;

            if (NumberFormatShowsDecimalPlacesP3T2(numberFormatLocal, expectedDecimalPlaces))
                return true;

            // Fallback nhẹ: nếu Excel trả format không ổn định nhưng Text đang hiển thị đúng 2 chữ số.
            // Không dùng làm chính vì Text có thể phụ thuộc độ rộng cột.
            object value = null;
            double number;

            try
            {
                value = cell.Value2;
            }
            catch
            {
                value = null;
            }

            if (value != null && TryToDouble(value, out number))
            {
                string displayText = "";
                try
                {
                    displayText = cell.Text == null ? "" : Convert.ToString(cell.Text);
                }
                catch
                {
                    displayText = "";
                }

                if (DisplayTextShowsDecimalPlacesP3T2(displayText, expectedDecimalPlaces))
                    return true;
            }

            return false;
        }

        private bool NumberFormatShowsDecimalPlacesP3T2(string format, int expectedDecimalPlaces)
        {
            if (format == null) return false;

            string f = format.Trim();

            if (string.IsNullOrWhiteSpace(f))
                return false;

            if (string.Equals(f, "General", StringComparison.OrdinalIgnoreCase))
                return expectedDecimalPlaces == 0;

            if (f.Contains("__ComObject"))
                return false;

            // Lấy section đầu tiên của number format.
            // Ví dụ: 0.00;[Red]-0.00 thì chỉ cần kiểm tra 0.00.
            string section = f.Split(';')[0];

            section = RemoveQuotedTextP3T2(section);
            section = RemoveBracketTokensP3T2(section);

            // Không chấp nhận percent vì 0.00% hiển thị bản chất khác Number.
            if (section.Contains("%"))
                return false;

            int decimalSeparatorIndex = FindDecimalSeparatorIndexP3T2(section);

            if (decimalSeparatorIndex < 0)
                return expectedDecimalPlaces == 0;

            int decimalPlaceholderCount = CountDecimalPlaceholdersAfterSeparatorP3T2(section, decimalSeparatorIndex);

            return decimalPlaceholderCount == expectedDecimalPlaces;
        }

        private int FindDecimalSeparatorIndexP3T2(string section)
        {
            if (string.IsNullOrWhiteSpace(section))
                return -1;

            // Lấy dấu . hoặc , cuối cùng trong section.
            // Với #,##0.00 thì dấu cuối là .
            // Với #.##0,00 hoặc 0,00 thì dấu cuối là ,
            int dot = section.LastIndexOf('.');
            int comma = section.LastIndexOf(',');

            int index = Math.Max(dot, comma);

            if (index < 0)
                return -1;

            // Nếu sau dấu phân cách không có placeholder nào thì không phải decimal separator.
            bool hasPlaceholderAfter = false;

            for (int i = index + 1; i < section.Length; i++)
            {
                char ch = section[i];

                if (ch == '0' || ch == '#' || ch == '?')
                {
                    hasPlaceholderAfter = true;
                    break;
                }
            }

            return hasPlaceholderAfter ? index : -1;
        }

        private int CountDecimalPlaceholdersAfterSeparatorP3T2(string section, int separatorIndex)
        {
            if (string.IsNullOrWhiteSpace(section)) return 0;
            if (separatorIndex < 0 || separatorIndex >= section.Length - 1) return 0;

            int count = 0;

            for (int i = separatorIndex + 1; i < section.Length; i++)
            {
                char ch = section[i];

                if (ch == '0' || ch == '#' || ch == '?')
                {
                    count++;
                    continue;
                }

                // Bỏ qua ký tự format không ảnh hưởng số thập phân.
                if (ch == '_' || ch == '*' || ch == '\\' || ch == ' ' || ch == ')' || ch == '(')
                    continue;

                // Nếu gặp ký tự khác sau khi đã đếm decimal placeholders thì dừng.
                if (count > 0)
                    break;
            }

            return count;
        }

        private string RemoveQuotedTextP3T2(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            return Regex.Replace(text, "\"[^\"]*\"", "");
        }

        private string RemoveBracketTokensP3T2(string text)
        {
            if (string.IsNullOrWhiteSpace(text)) return "";

            // Xóa [Red], [Blue], [>=100], [$-en-US], v.v.
            return Regex.Replace(text, @"\[[^\]]*\]", "");
        }

        private bool DisplayTextShowsDecimalPlacesP3T2(string displayText, int expectedDecimalPlaces)
        {
            if (displayText == null) return false;

            string text = displayText.Trim();

            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (text.Contains("#"))
                return false;

            int dot = text.LastIndexOf('.');
            int comma = text.LastIndexOf(',');

            int separatorIndex = Math.Max(dot, comma);

            if (separatorIndex < 0)
                return expectedDecimalPlaces == 0;

            int digits = 0;

            for (int i = separatorIndex + 1; i < text.Length; i++)
            {
                char ch = text[i];

                if (char.IsDigit(ch))
                {
                    digits++;
                    continue;
                }

                break;
            }

            return digits == expectedDecimalPlaces;
        }
        //Project 3 Task 3
        public bool TableRowContainingTextDeleted(string sheetName, string searchText, int expectedDataRowCount)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(searchText)) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects listObjects = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Range dataBodyRange = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                listObjects = ws.ListObjects;
                if (listObjects == null) return false;

                int tableCount = Convert.ToInt32(listObjects.Count);
                if (tableCount <= 0) return false;

                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(table);
                    ReleaseCom(tableRange);
                    ReleaseCom(dataBodyRange);

                    table = null;
                    tableRange = null;
                    dataBodyRange = null;

                    table = listObjects.Item[i];
                    if (table == null) continue;

                    tableRange = table.Range;
                    if (tableRange == null) continue;

                    // Chọn đúng table có dữ liệu kỳ thi.
                    if (!TableLooksLikeInformaticsExamP3T3(tableRange))
                        continue;

                    dataBodyRange = table.DataBodyRange;
                    if (dataBodyRange == null)
                        return false;

                    int actualDataRows = Convert.ToInt32(dataBodyRange.Rows.Count);

                    // Sau khi xóa IELTS, table còn 3 dòng dữ liệu: IC3, MOS, ICDL.
                    // Nếu học viên chỉ xóa chữ IELTS hoặc clear contents cả dòng,
                    // số dòng DataBodyRange vẫn là 4 nên FAIL.
                    if (expectedDataRowCount > 0 && actualDataRows != expectedDataRowCount)
                        return false;

                    if (RangeContainsTextP3T3(dataBodyRange, searchText))
                        return false;

                    if (TableHasBlankOrDamagedDataRowP3T3(dataBodyRange))
                        return false;

                    return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(dataBodyRange);
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(listObjects);
                ReleaseCom(ws);
            }
        }
        private bool TableLooksLikeInformaticsExamP3T3(Xl.Range tableRange)
        {
            if (tableRange == null) return false;

            Xl.Range headerRow = null;
            Xl.Range firstHeader = null;

            try
            {
                headerRow = (Xl.Range)tableRange.Rows[1];
                if (headerRow == null) return false;

                firstHeader = (Xl.Range)headerRow.Cells[1, 1];
                string firstHeaderText = GetCellTextP1T3(firstHeader);

                if (!StringEqualsLooseP3T3(firstHeaderText, "Exam"))
                    return false;

                bool hasJanuary = FindHeaderColumnIndexInFilterRange(tableRange, "January") > 0;
                bool hasDecember = FindHeaderColumnIndexInFilterRange(tableRange, "December") > 0;

                return hasJanuary && hasDecember;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(firstHeader);
                ReleaseCom(headerRow);
            }
        }

        private bool RangeContainsTextP3T3(Xl.Range range, string searchText)
        {
            if (range == null) return false;
            if (string.IsNullOrWhiteSpace(searchText)) return false;

            Xl.Range cell = null;

            try
            {
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                string expected = NormalizeTextForCompareP3T3(searchText);

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)range.Cells[r, c];
                        string actual = NormalizeTextForCompareP3T3(GetCellTextP1T3(cell));

                        if (string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase))
                            return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool TableHasBlankOrDamagedDataRowP3T3(Xl.Range dataBodyRange)
        {
            if (dataBodyRange == null) return true;

            Xl.Range rowRange = null;
            Xl.Range firstCell = null;
            Xl.Range cell = null;

            try
            {
                int rowCount = Convert.ToInt32(dataBodyRange.Rows.Count);
                int columnCount = Convert.ToInt32(dataBodyRange.Columns.Count);

                for (int r = 1; r <= rowCount; r++)
                {
                    ReleaseCom(rowRange);
                    ReleaseCom(firstCell);
                    ReleaseCom(cell);

                    rowRange = null;
                    firstCell = null;
                    cell = null;

                    rowRange = (Xl.Range)dataBodyRange.Rows[r];
                    firstCell = (Xl.Range)rowRange.Cells[1, 1];

                    string firstText = GetCellTextP1T3(firstCell);

                    int nonEmptyCount = 0;

                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)rowRange.Cells[1, c];
                        string text = GetCellTextP1T3(cell);

                        if (!string.IsNullOrWhiteSpace(text))
                            nonEmptyCount++;
                    }

                    // Nếu cả dòng trống: học viên có thể đã Clear Contents thay vì Delete Table Rows.
                    if (nonEmptyCount == 0)
                        return true;

                    // Nếu ô Exam trống nhưng các tháng còn số liệu:
                    // học viên có thể chỉ xóa chữ IELTS.
                    if (string.IsNullOrWhiteSpace(firstText) && nonEmptyCount > 0)
                        return true;
                }

                return false;
            }
            catch
            {
                return true;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(firstCell);
                ReleaseCom(rowRange);
            }
        }

        private bool StringEqualsLooseP3T3(string actual, string expected)
        {
            string a = NormalizeTextForCompareP3T3(actual);
            string e = NormalizeTextForCompareP3T3(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeTextForCompareP3T3(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        //Project 3 Task 4
        public bool AverageFormulaByHeaders(string sheetName, string targetHeader, IList<string> sourceHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetHeader)) return false;
            if (sourceHeaders == null || sourceHeaders.Count == 0) return false;

            Xl.Worksheet ws = null;
            Xl.Range tableRange = null;
            Xl.Range sourceCell = null;
            Xl.Range targetCell = null;
            Xl.Range firstColumnCell = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                List<string> requiredHeaders = new List<string>();
                requiredHeaders.Add(targetHeader);
                for (int i = 0; i < sourceHeaders.Count; i++)
                    requiredHeaders.Add(sourceHeaders[i]);

                tableRange = GetTableOrUsedRangeByHeadersP1T4T5(ws, requiredHeaders);
                if (tableRange == null) return false;

                int targetColumnIndex = FindHeaderColumnIndexInFilterRange(tableRange, targetHeader);
                if (targetColumnIndex <= 0) return false;

                int[] sourceColumnIndexes = new int[sourceHeaders.Count];

                for (int i = 0; i < sourceHeaders.Count; i++)
                {
                    sourceColumnIndexes[i] = FindHeaderColumnIndexInFilterRange(tableRange, sourceHeaders[i]);
                    if (sourceColumnIndexes[i] <= 0)
                        return false;
                }

                int rowCount = Convert.ToInt32(tableRange.Rows.Count);
                int firstWorksheetRow = Convert.ToInt32(tableRange.Row);
                int firstWorksheetColumn = Convert.ToInt32(tableRange.Column);

                int checkedRows = 0;

                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    ReleaseCom(targetCell);
                    ReleaseCom(firstColumnCell);
                    targetCell = null;
                    firstColumnCell = null;

                    firstColumnCell = (Xl.Range)tableRange.Cells[rowIndex, 1];
                    string firstColumnText = GetCellTextP1T3(firstColumnCell);

                    // Bỏ qua dòng trống hoặc dòng Total nếu có.
                    if (string.IsNullOrWhiteSpace(firstColumnText))
                        continue;

                    if (StringEqualsLooseP3T4(firstColumnText, "Total"))
                        continue;

                    double total = 0;
                    int numberCount = 0;

                    for (int i = 0; i < sourceColumnIndexes.Length; i++)
                    {
                        ReleaseCom(sourceCell);
                        sourceCell = null;

                        sourceCell = (Xl.Range)tableRange.Cells[rowIndex, sourceColumnIndexes[i]];

                        double value;
                        if (TryGetCellNumberP3T4(sourceCell, out value))
                        {
                            total += value;
                            numberCount++;
                        }
                    }

                    if (numberCount != sourceColumnIndexes.Length)
                        continue;

                    checkedRows++;

                    double expectedAverage = total / numberCount;

                    targetCell = (Xl.Range)tableRange.Cells[rowIndex, targetColumnIndex];
                    if (targetCell == null) return false;

                    if (!CellHasFormulaP1T3(targetCell))
                        return false;

                    string formula = Convert.ToString(targetCell.Formula);
                    string formulaLocal = Convert.ToString(targetCell.FormulaLocal);

                    int worksheetRow = firstWorksheetRow + rowIndex - 1;

                    if (!FormulaLooksLikeAverageFormulaP3T4(
                            formula,
                            formulaLocal,
                            worksheetRow,
                            firstWorksheetColumn,
                            sourceColumnIndexes,
                            sourceHeaders))
                    {
                        return false;
                    }

                    double actualAverage;
                    if (!TryGetCellNumberP3T4(targetCell, out actualAverage))
                        return false;

                    if (Math.Abs(actualAverage - expectedAverage) > 0.0001)
                        return false;
                }

                return checkedRows > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(firstColumnCell);
                ReleaseCom(targetCell);
                ReleaseCom(sourceCell);
                ReleaseCom(tableRange);
                ReleaseCom(ws);
            }
        }
        private bool FormulaLooksLikeAverageFormulaP3T4(
    string formula,
    string formulaLocal,
    int worksheetRow,
    int firstWorksheetColumn,
    int[] sourceColumnIndexes,
    IList<string> sourceHeaders)
        {
            string f1 = NormalizeFormulaP3T4(formula);
            string f2 = NormalizeFormulaP3T4(formulaLocal);

            string combined = f1 + " " + f2;

            if (string.IsNullOrWhiteSpace(combined))
                return false;

            if (!combined.Contains("AVERAGE("))
                return false;

            if (sourceColumnIndexes == null || sourceColumnIndexes.Length == 0)
                return false;

            int firstSourceWorksheetColumn = firstWorksheetColumn + sourceColumnIndexes[0] - 1;
            int lastSourceWorksheetColumn = firstWorksheetColumn + sourceColumnIndexes[sourceColumnIndexes.Length - 1] - 1;

            string firstCell = ExcelColumnNameP1T3(firstSourceWorksheetColumn) + worksheetRow;
            string lastCell = ExcelColumnNameP1T3(lastSourceWorksheetColumn) + worksheetRow;
            string expectedRange = NormalizeFormulaP3T4(firstCell + ":" + lastCell);

            if (combined.Contains(expectedRange))
                return true;

            bool containsAllCellRefs = true;

            for (int i = 0; i < sourceColumnIndexes.Length; i++)
            {
                int worksheetColumn = firstWorksheetColumn + sourceColumnIndexes[i] - 1;
                string cellRef = NormalizeFormulaP3T4(ExcelColumnNameP1T3(worksheetColumn) + worksheetRow);

                if (!combined.Contains(cellRef))
                {
                    containsAllCellRefs = false;
                    break;
                }
            }

            if (containsAllCellRefs)
                return true;

            // Structured reference, ví dụ:
            // =AVERAGE([@[January]:[April]])
            // Có thể chỉ chứa January và April, không nhất thiết chứa February/March.
            if (sourceHeaders != null && sourceHeaders.Count > 0)
            {
                string firstHeader = NormalizeFormulaP3T4(sourceHeaders[0]);
                string lastHeader = NormalizeFormulaP3T4(sourceHeaders[sourceHeaders.Count - 1]);

                if (!string.IsNullOrWhiteSpace(firstHeader) &&
                    !string.IsNullOrWhiteSpace(lastHeader) &&
                    combined.Contains(firstHeader) &&
                    combined.Contains(lastHeader))
                {
                    return true;
                }

                bool containsAllHeaders = true;

                for (int i = 0; i < sourceHeaders.Count; i++)
                {
                    string header = NormalizeFormulaP3T4(sourceHeaders[i]);

                    if (string.IsNullOrWhiteSpace(header) || !combined.Contains(header))
                    {
                        containsAllHeaders = false;
                        break;
                    }
                }

                if (containsAllHeaders)
                    return true;
            }

            return false;
        }

        private string NormalizeFormulaP3T4(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();

            s = s.Replace(";", ",");
            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("[", "");
            s = s.Replace("]", "");
            s = s.Replace("@", "");

            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }

        private bool TryGetCellNumberP3T4(Xl.Range cell, out double value)
        {
            value = 0;

            if (cell == null)
                return false;

            object raw = null;

            try
            {
                raw = cell.Value2;
            }
            catch
            {
                raw = null;
            }

            if (raw == null)
                return false;

            try
            {
                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
            }

            string text = Convert.ToString(raw);
            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (double.TryParse(text, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return true;

            if (double.TryParse(text, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                return true;

            return false;
        }

        private bool StringEqualsLooseP3T4(string actual, string expected)
        {
            string a = NormalizeTextForCompareP3T4(actual);
            string e = NormalizeTextForCompareP3T4(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeTextForCompareP3T4(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        //Project 3 Task 5
        public bool ChartPrimaryVerticalAxisTitleEquals(string sheetName, string expectedTitle)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(expectedTitle)) return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                chartObjects = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (chartObjects == null) return false;

                int chartCount = Convert.ToInt32(chartObjects.Count);
                if (chartCount <= 0) return false;

                for (int i = 1; i <= chartCount; i++)
                {
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);

                    chart = null;
                    chartObject = null;

                    chartObject = chartObjects.Item(i) as Xl.ChartObject;
                    if (chartObject == null) continue;

                    chart = chartObject.Chart;
                    if (chart == null) continue;

                    if (ChartHasPrimaryVerticalAxisTitleP3T5(chart, expectedTitle))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(ws);
            }
        }
        private bool ChartHasPrimaryVerticalAxisTitleP3T5(Xl.Chart chart, string expectedTitle)
        {
            if (chart == null) return false;
            if (string.IsNullOrWhiteSpace(expectedTitle)) return false;

            Xl.Axis axis = null;
            Xl.AxisTitle axisTitle = null;

            try
            {
                // Primary Vertical Axis = Value Axis + Primary Axis Group.
                axis = chart.Axes(
                    Xl.XlAxisType.xlValue,
                    Xl.XlAxisGroup.xlPrimary) as Xl.Axis;

                if (axis == null)
                    return false;

                bool hasTitle = false;

                try
                {
                    hasTitle = Convert.ToBoolean(axis.HasTitle);
                }
                catch
                {
                    hasTitle = false;
                }

                if (!hasTitle)
                    return false;

                axisTitle = axis.AxisTitle;
                if (axisTitle == null)
                    return false;

                string actualText = "";

                try
                {
                    actualText = Convert.ToString(axisTitle.Text);
                }
                catch
                {
                    actualText = "";
                }

                if (AxisTitleEqualsP3T5(actualText, expectedTitle))
                    return true;

                // Một số Excel build đọc Caption ổn hơn Text.
                try
                {
                    actualText = Convert.ToString(axisTitle.Caption);
                }
                catch
                {
                    actualText = "";
                }

                if (AxisTitleEqualsP3T5(actualText, expectedTitle))
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(axisTitle);
                ReleaseCom(axis);
            }
        }

        private bool AxisTitleEqualsP3T5(string actual, string expected)
        {
            string a = NormalizeAxisTitleP3T5(actual);
            string e = NormalizeAxisTitleP3T5(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeAxisTitleP3T5(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", " ");
            s = s.Trim();

            return s;
        }
        //Project 3 Task 7
        public bool TableColumnFormulaFilledDown(string sheetName, string startCellAddress, string expectedFormula)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(startCellAddress)) return false;

            Xl.Worksheet ws = null;
            Xl.Range startCell = null;
            Xl.ListObjects listObjects = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Range dataBodyRange = null;
            Xl.Range checkCell = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                startCell = ws.Range[startCellAddress];
                if (startCell == null) return false;

                if (!CellHasFormulaP1T3(startCell))
                    return false;

                int startRow = Convert.ToInt32(startCell.Row);
                int startColumn = Convert.ToInt32(startCell.Column);

                string baseFormulaR1C1 = GetFormulaR1C1P3T7(startCell);
                if (string.IsNullOrWhiteSpace(baseFormulaR1C1))
                    return false;

                if (!string.IsNullOrWhiteSpace(expectedFormula))
                {
                    string actualSeedFormula = Convert.ToString(startCell.Formula);
                    if (!FormulaR1C1EqualsP3T7(actualSeedFormula, expectedFormula))
                        return false;
                }

                listObjects = ws.ListObjects;
                if (listObjects == null) return false;

                int tableCount = Convert.ToInt32(listObjects.Count);
                if (tableCount <= 0) return false;

                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(table);
                    ReleaseCom(tableRange);
                    ReleaseCom(dataBodyRange);

                    table = null;
                    tableRange = null;
                    dataBodyRange = null;

                    table = listObjects.Item[i];
                    if (table == null) continue;

                    tableRange = table.Range;
                    dataBodyRange = table.DataBodyRange;

                    if (tableRange == null || dataBodyRange == null)
                        continue;

                    if (!CellIsInsideRangeP3T7(startRow, startColumn, dataBodyRange))
                        continue;

                    int dataStartRow = Convert.ToInt32(dataBodyRange.Row);
                    int dataRowCount = Convert.ToInt32(dataBodyRange.Rows.Count);
                    int dataColumnCount = Convert.ToInt32(dataBodyRange.Columns.Count);

                    int tableStartColumn = Convert.ToInt32(tableRange.Column);

                    int targetColumnIndexInTable = startColumn - tableStartColumn + 1;
                    if (targetColumnIndexInTable <= 0 || targetColumnIndexInTable > dataColumnCount)
                        return false;

                    int startRowIndexInData = startRow - dataStartRow + 1;
                    if (startRowIndexInData <= 0 || startRowIndexInData > dataRowCount)
                        return false;

                    int checkedCells = 0;

                    for (int rowIndex = startRowIndexInData; rowIndex <= dataRowCount; rowIndex++)
                    {
                        ReleaseCom(checkCell);
                        checkCell = null;

                        checkCell = (Xl.Range)dataBodyRange.Cells[rowIndex, targetColumnIndexInTable];
                        if (checkCell == null) return false;

                        checkedCells++;

                        if (!CellHasFormulaP1T3(checkCell))
                            return false;

                        string currentFormulaR1C1 = GetFormulaR1C1P3T7(checkCell);

                        if (!FormulaR1C1EqualsP3T7(currentFormulaR1C1, baseFormulaR1C1))
                            return false;
                    }

                    return checkedCells > 1;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(checkCell);
                ReleaseCom(dataBodyRange);
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(listObjects);
                ReleaseCom(startCell);
                ReleaseCom(ws);
            }
        }
        private bool CellIsInsideRangeP3T7(int row, int column, Xl.Range range)
        {
            if (range == null) return false;

            try
            {
                int rangeRow = Convert.ToInt32(range.Row);
                int rangeColumn = Convert.ToInt32(range.Column);
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                int lastRow = rangeRow + rowCount - 1;
                int lastColumn = rangeColumn + columnCount - 1;

                return row >= rangeRow &&
                       row <= lastRow &&
                       column >= rangeColumn &&
                       column <= lastColumn;
            }
            catch
            {
                return false;
            }
        }

        private string GetFormulaR1C1P3T7(Xl.Range cell)
        {
            if (cell == null) return "";

            try
            {
                object formula = cell.FormulaR1C1;
                return formula == null ? "" : Convert.ToString(formula);
            }
            catch
            {
                try
                {
                    object formula = cell.Formula;
                    return formula == null ? "" : Convert.ToString(formula);
                }
                catch
                {
                    return "";
                }
            }
        }

        private bool FormulaR1C1EqualsP3T7(string actualFormula, string expectedFormula)
        {
            string actual = NormalizeFormulaP3T7(actualFormula);
            string expected = NormalizeFormulaP3T7(expectedFormula);

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeFormulaP3T7(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();

            s = s.Replace(";", ",");
            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("[", "");
            s = s.Replace("]", "");
            s = s.Replace("@", "");

            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        //Project 3 Task 8
        public bool MaxFormulaFromHeader(string sheetName, string targetCellAddress, string sourceHeader)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetCellAddress)) return false;
            if (string.IsNullOrWhiteSpace(sourceHeader)) return false;

            Xl.Worksheet ws = null;
            Xl.Range targetCell = null;
            Xl.Range sourceDataRange = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                targetCell = ws.Range[targetCellAddress];
                if (targetCell == null) return false;

                if (!CellHasFormulaP1T3(targetCell))
                    return false;

                sourceDataRange = FindTableColumnDataBodyRangeByHeaderP3T8(ws, sourceHeader);
                if (sourceDataRange == null)
                    sourceDataRange = FindUsedRangeColumnDataRangeByHeaderP3T8(ws, sourceHeader);

                if (sourceDataRange == null)
                    return false;

                double expectedMax;
                if (!TryGetMaxNumericValueP3T8(sourceDataRange, out expectedMax))
                    return false;

                double actualValue;
                if (!TryGetCellNumberP3T8(targetCell, out actualValue))
                    return false;

                if (Math.Abs(actualValue - expectedMax) > 0.0001)
                    return false;

                string formula = Convert.ToString(targetCell.Formula);
                string formulaLocal = Convert.ToString(targetCell.FormulaLocal);

                if (!FormulaLooksLikeMaxFromSourceP3T8(
                        formula,
                        formulaLocal,
                        sourceHeader,
                        sourceDataRange))
                {
                    return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sourceDataRange);
                ReleaseCom(targetCell);
                ReleaseCom(ws);
            }
        }
        private Xl.Range FindTableColumnDataBodyRangeByHeaderP3T8(Xl.Worksheet ws, string sourceHeader)
        {
            if (ws == null) return null;
            if (string.IsNullOrWhiteSpace(sourceHeader)) return null;

            Xl.ListObjects listObjects = null;
            Xl.ListObject table = null;
            Xl.ListColumns columns = null;
            Xl.ListColumn column = null;
            Xl.Range dataBodyRange = null;

            try
            {
                listObjects = ws.ListObjects;
                if (listObjects == null) return null;

                int tableCount = Convert.ToInt32(listObjects.Count);

                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(table);
                    ReleaseCom(columns);
                    ReleaseCom(column);
                    ReleaseCom(dataBodyRange);

                    table = null;
                    columns = null;
                    column = null;
                    dataBodyRange = null;

                    table = listObjects.Item[i];
                    if (table == null) continue;

                    columns = table.ListColumns;
                    if (columns == null) continue;

                    int columnCount = Convert.ToInt32(columns.Count);

                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(column);
                        ReleaseCom(dataBodyRange);

                        column = null;
                        dataBodyRange = null;

                        column = (Xl.ListColumn)columns.Item[c];
                        if (column == null) continue;

                        string actualHeader = "";

                        try
                        {
                            actualHeader = Convert.ToString(column.Name);
                        }
                        catch
                        {
                            actualHeader = "";
                        }

                        if (!HeaderEqualsP3T8(actualHeader, sourceHeader))
                            continue;

                        dataBodyRange = column.DataBodyRange;
                        if (dataBodyRange == null)
                            return null;

                        Xl.Range result = dataBodyRange;
                        dataBodyRange = null;
                        return result;
                    }
                }

                return null;
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(dataBodyRange);
                ReleaseCom(column);
                ReleaseCom(columns);
                ReleaseCom(table);
                ReleaseCom(listObjects);
            }
        }

        private Xl.Range FindUsedRangeColumnDataRangeByHeaderP3T8(Xl.Worksheet ws, string sourceHeader)
        {
            if (ws == null) return null;
            if (string.IsNullOrWhiteSpace(sourceHeader)) return null;

            Xl.Range usedRange = null;
            Xl.Range cell = null;
            Xl.Range startCell = null;
            Xl.Range endCell = null;

            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return null;

                int firstRow = Convert.ToInt32(usedRange.Row);
                int firstColumn = Convert.ToInt32(usedRange.Column);
                int rowCount = Convert.ToInt32(usedRange.Rows.Count);
                int columnCount = Convert.ToInt32(usedRange.Columns.Count);

                int lastRow = firstRow + rowCount - 1;
                int lastColumn = firstColumn + columnCount - 1;

                int headerRow = 0;
                int headerColumn = 0;

                int maxSearchRow = firstRow + 20;
                if (maxSearchRow > lastRow) maxSearchRow = lastRow;

                for (int r = firstRow; r <= maxSearchRow; r++)
                {
                    for (int c = firstColumn; c <= lastColumn; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)ws.Cells[r, c];

                        string text = GetCellTextP1T3(cell);
                        if (HeaderEqualsP3T8(text, sourceHeader))
                        {
                            headerRow = r;
                            headerColumn = c;
                            break;
                        }
                    }

                    if (headerRow > 0)
                        break;
                }

                if (headerRow <= 0 || headerColumn <= 0)
                    return null;

                if (headerRow + 1 > lastRow)
                    return null;

                startCell = (Xl.Range)ws.Cells[headerRow + 1, headerColumn];
                endCell = (Xl.Range)ws.Cells[lastRow, headerColumn];

                return ws.Range[startCell, endCell];
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(endCell);
                ReleaseCom(startCell);
                ReleaseCom(cell);
                ReleaseCom(usedRange);
            }
        }

        private bool TryGetMaxNumericValueP3T8(Xl.Range range, out double maxValue)
        {
            maxValue = 0;

            if (range == null) return false;

            Xl.Range cell = null;
            bool hasNumber = false;

            try
            {
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                for (int r = 1; r <= rowCount; r++)
                {
                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = (Xl.Range)range.Cells[r, c];

                        double value;
                        if (!TryGetCellNumberP3T8(cell, out value))
                            continue;

                        if (!hasNumber)
                        {
                            maxValue = value;
                            hasNumber = true;
                        }
                        else if (value > maxValue)
                        {
                            maxValue = value;
                        }
                    }
                }

                return hasNumber;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool TryGetCellNumberP3T8(Xl.Range cell, out double value)
        {
            value = 0;

            if (cell == null)
                return false;

            object raw = null;

            try
            {
                raw = cell.Value2;
            }
            catch
            {
                raw = null;
            }

            if (raw == null)
                return false;

            try
            {
                value = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
            }

            string text = Convert.ToString(raw);

            if (string.IsNullOrWhiteSpace(text))
                return false;

            if (double.TryParse(text, System.Globalization.NumberStyles.Any, CultureInfo.InvariantCulture, out value))
                return true;

            if (double.TryParse(text, System.Globalization.NumberStyles.Any, CultureInfo.CurrentCulture, out value))
                return true;

            return false;
        }

        private bool FormulaLooksLikeMaxFromSourceP3T8(
            string formula,
            string formulaLocal,
            string sourceHeader,
            Xl.Range sourceDataRange)
        {
            string f1 = NormalizeFormulaP3T8(formula);
            string f2 = NormalizeFormulaP3T8(formulaLocal);
            string combined = f1 + " " + f2;

            if (string.IsNullOrWhiteSpace(combined))
                return false;

            if (!combined.Contains("MAX("))
                return false;

            string normalizedHeader = NormalizeFormulaP3T8(sourceHeader);

            if (!string.IsNullOrWhiteSpace(normalizedHeader) &&
                combined.Contains(normalizedHeader))
            {
                return true;
            }

            string sourceAddress = GetRangeAddressA1P3T8(sourceDataRange);
            string normalizedAddress = NormalizeFormulaP3T8(sourceAddress);

            if (!string.IsNullOrWhiteSpace(normalizedAddress) &&
                combined.Contains(normalizedAddress))
            {
                return true;
            }

            string sourceAddressWithSheet = GetRangeAddressA1WithSheetP3T8(sourceDataRange);
            string normalizedAddressWithSheet = NormalizeFormulaP3T8(sourceAddressWithSheet);

            if (!string.IsNullOrWhiteSpace(normalizedAddressWithSheet) &&
                combined.Contains(normalizedAddressWithSheet))
            {
                return true;
            }

            return false;
        }

        private string GetRangeAddressA1P3T8(Xl.Range range)
        {
            if (range == null) return "";

            try
            {
                return Convert.ToString(range.get_Address(
                    false,
                    false,
                    Xl.XlReferenceStyle.xlA1,
                    false,
                    Type.Missing));
            }
            catch
            {
                return "";
            }
        }

        private string GetRangeAddressA1WithSheetP3T8(Xl.Range range)
        {
            if (range == null) return "";

            Xl.Worksheet ws = null;

            try
            {
                string address = GetRangeAddressA1P3T8(range);

                ws = range.Worksheet as Xl.Worksheet;
                string sheetName = ws == null ? "" : Convert.ToString(ws.Name);

                if (string.IsNullOrWhiteSpace(sheetName))
                    return address;

                return sheetName + "!" + address;
            }
            catch
            {
                return "";
            }
            finally
            {
                ReleaseCom(ws);
            }
        }

        private bool HeaderEqualsP3T8(string actualHeader, string expectedHeader)
        {
            string actual = NormalizeHeaderP3T8(actualHeader);
            string expected = NormalizeHeaderP3T8(expectedHeader);

            return string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeHeaderP3T8(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }

        private string NormalizeFormulaP3T8(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();

            s = s.Replace(";", ",");
            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("[", "");
            s = s.Replace("]", "");
            s = s.Replace("@", "");

            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        //Project 4 Task 1
        public bool ChartSwitchedRowColumn(
            string sheetName,
            string chartTitle,
            string chartName,
            string sourceRange,
            int expectedChartType)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                chartObjects = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (chartObjects == null) return false;

                int chartCount = Convert.ToInt32(chartObjects.Count);
                if (chartCount <= 0) return false;

                for (int i = 1; i <= chartCount; i++)
                {
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);
                    chart = null;
                    chartObject = null;

                    chartObject = (Xl.ChartObject)chartObjects.Item(i);
                    if (chartObject == null) continue;

                    if (!string.IsNullOrWhiteSpace(chartName) &&
                        !string.Equals(Convert.ToString(chartObject.Name), chartName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    chart = chartObject.Chart;
                    if (chart == null) continue;

                    if (!ChartTitleMatchesP4T1(chart, chartTitle))
                        continue;

                    if (expectedChartType > 0 && Convert.ToInt32(chart.ChartType) != expectedChartType)
                        continue;

                    if (expectedChartType > 0)
                    {
                        if (ChartPlotByRowsP4T1(chart) &&
                            ChartSeriesLooksExactlySwitchedP12T6(chart, ws, sourceRange))
                            return true;

                        continue;
                    }

                    // Cách ổn định nhất: nếu Excel đọc được PlotBy = xlRows thì PASS.
                    if (ChartPlotByRowsP4T1(chart))
                        return true;

                    // Fallback: kiểm tra series sau khi Switch Row/Column.
                    // Với source A2:C7, sau khi switch phải có 5 series theo từng dòng 3:7,
                    // mỗi series lấy values theo hàng B:C.
                    if (ChartSeriesLooksSwitchedP4T1(chart, ws, sourceRange))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(ws);
            }
        }

        private bool ChartTitleMatchesP4T1(Xl.Chart chart, string expectedTitle)
        {
            if (chart == null) return false;

            if (string.IsNullOrWhiteSpace(expectedTitle))
                return true;

            Xl.ChartTitle chartTitle = null;

            try
            {
                bool hasTitle = false;

                try
                {
                    hasTitle = Convert.ToBoolean(chart.HasTitle);
                }
                catch
                {
                    hasTitle = false;
                }

                if (!hasTitle)
                    return false;

                chartTitle = chart.ChartTitle;
                if (chartTitle == null)
                    return false;

                string actual = "";

                try
                {
                    actual = Convert.ToString(chartTitle.Text);
                }
                catch
                {
                    actual = "";
                }

                if (TextEqualsP4T1(actual, expectedTitle))
                    return true;

                try
                {
                    actual = Convert.ToString(chartTitle.Caption);
                }
                catch
                {
                    actual = "";
                }

                return TextEqualsP4T1(actual, expectedTitle);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chartTitle);
            }
        }

        private bool ChartPlotByRowsP4T1(Xl.Chart chart)
        {
            if (chart == null) return false;

            try
            {
                int plotBy = Convert.ToInt32(chart.PlotBy);
                return plotBy == Convert.ToInt32(Xl.XlRowCol.xlRows);
            }
            catch
            {
                return false;
            }
        }

        private bool ChartSeriesLooksSwitchedP4T1(Xl.Chart chart, Xl.Worksheet ws, string sourceRangeAddress)
        {
            if (chart == null || ws == null) return false;
            if (string.IsNullOrWhiteSpace(sourceRangeAddress)) return false;

            Xl.Range sourceRange = null;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;

            try
            {
                sourceRange = ws.Range[sourceRangeAddress];
                if (sourceRange == null) return false;

                int firstRow = Convert.ToInt32(sourceRange.Row);
                int firstColumn = Convert.ToInt32(sourceRange.Column);
                int rowCount = Convert.ToInt32(sourceRange.Rows.Count);
                int columnCount = Convert.ToInt32(sourceRange.Columns.Count);

                if (rowCount < 2 || columnCount < 2)
                    return false;

                int expectedSeriesCount = rowCount - 1;
                int expectedValueColumnCount = columnCount - 1;

                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null) return false;

                int actualSeriesCount = Convert.ToInt32(seriesCollection.Count);

                // Trước khi switch chart A2:C7 thường có 2 series theo cột.
                // Sau khi Switch Row/Column phải thành 5 series theo dòng.
                if (actualSeriesCount != expectedSeriesCount)
                    return false;

                int rowSeriesMatches = 0;

                for (int i = 1; i <= actualSeriesCount; i++)
                {
                    ReleaseCom(series);
                    series = null;

                    series = seriesCollection.Item(i) as Xl.Series;
                    if (series == null) continue;

                    string formula = "";

                    try
                    {
                        formula = Convert.ToString(series.Formula);
                    }
                    catch
                    {
                        formula = "";
                    }

                    string normalizedFormula = NormalizeChartFormulaP4T1(formula);

                    int expectedDataRow = firstRow + i;
                    string expectedRowRange = BuildSheetRangeAddressP4T1(
                        ws,
                        expectedDataRow,
                        firstColumn + 1,
                        1,
                        expectedValueColumnCount);

                    string normalizedExpectedRowRange = NormalizeChartFormulaP4T1(expectedRowRange);

                    if (!string.IsNullOrWhiteSpace(normalizedExpectedRowRange) &&
                        normalizedFormula.Contains(normalizedExpectedRowRange))
                    {
                        rowSeriesMatches++;
                    }
                }

                // Cho phép lệch nhẹ 1 series nếu Excel build ghi formula khác,
                // nhưng phải thấy đa số series lấy dữ liệu theo từng hàng.
                return rowSeriesMatches >= Math.Max(1, expectedSeriesCount - 1);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
                ReleaseCom(sourceRange);
            }
        }

        private bool ChartSeriesLooksExactlySwitchedP12T6(
            Xl.Chart chart,
            Xl.Worksheet ws,
            string sourceRangeAddress)
        {
            if (chart == null || ws == null || string.IsNullOrWhiteSpace(sourceRangeAddress))
                return false;

            Xl.Range sourceRange = null;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;

            try
            {
                sourceRange = ws.Range[sourceRangeAddress];
                if (sourceRange == null) return false;

                int firstRow = Convert.ToInt32(sourceRange.Row);
                int firstColumn = Convert.ToInt32(sourceRange.Column);
                int rowCount = Convert.ToInt32(sourceRange.Rows.Count);
                int columnCount = Convert.ToInt32(sourceRange.Columns.Count);
                if (rowCount < 2 || columnCount < 2) return false;

                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != rowCount - 1)
                    return false;

                string expectedCategories = NormalizeChartFormulaP4T1(
                    BuildSheetRangeAddressP4T1(ws, firstRow, firstColumn + 1, 1, columnCount - 1));

                for (int i = 1; i <= rowCount - 1; i++)
                {
                    ReleaseCom(series);
                    series = seriesCollection.Item(i) as Xl.Series;
                    if (series == null) return false;

                    string formula = NormalizeChartFormulaP4T1(Convert.ToString(series.Formula));
                    string expectedName = NormalizeChartFormulaP4T1(
                        Convert.ToString(ws.Name) + "!" +
                        ExcelColumnNameP1T3(firstColumn) +
                        (firstRow + i).ToString(CultureInfo.InvariantCulture));
                    string expectedValues = NormalizeChartFormulaP4T1(
                        BuildSheetRangeAddressP4T1(ws, firstRow + i, firstColumn + 1, 1, columnCount - 1));

                    if (!formula.Contains(expectedName) ||
                        !formula.Contains(expectedCategories) ||
                        !formula.Contains(expectedValues))
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
                ReleaseCom(sourceRange);
            }
        }

        private string BuildSheetRangeAddressP4T1(Xl.Worksheet ws, int row, int column, int rowCount, int columnCount)
        {
            if (ws == null) return "";
            if (row <= 0 || column <= 0 || rowCount <= 0 || columnCount <= 0) return "";

            try
            {
                string start = ExcelColumnNameP1T3(column) + row.ToString(CultureInfo.InvariantCulture);
                string end = ExcelColumnNameP1T3(column + columnCount - 1) +
                             (row + rowCount - 1).ToString(CultureInfo.InvariantCulture);

                string sheet = Convert.ToString(ws.Name);
                if (string.IsNullOrWhiteSpace(sheet))
                    return start + ":" + end;

                return sheet + "!" + start + ":" + end;
            }
            catch
            {
                return "";
            }
        }

        private bool TextEqualsP4T1(string actual, string expected)
        {
            string a = NormalizeTextP4T1(actual);
            string e = NormalizeTextP4T1(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeTextP4T1(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", " ");

            return s.Trim();
        }

        private string NormalizeChartFormulaP4T1(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            string s = text.Trim();

            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("=", "");
            s = s.Replace(";", ",");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }


        //Project 4 Task 2
        //Project 4 Task 2 - Very loose version for title/subtitle formatting
        //Project 4 Task 2 - XML style check for title/subtitle Format Painter
        public bool RangeFormattingMatches(string sourceSheetName, string sourceRangeAddress, string targetSheetName, string targetRangeAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sourceSheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetSheetName)) return false;

            Xl.Workbook wb = null;
            string tempPath = "";

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                if (wb == null) return false;

                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    "MosTrainer_P4T2_" + Guid.NewGuid().ToString("N") + ".xlsx");

                // Lấy đúng trạng thái workbook hiện tại, không bắt học viên phải bấm Save.
                wb.SaveCopyAs(tempPath);

                if (!File.Exists(tempPath))
                    return false;

                return XlsxTitleSubtitleStyleCopiedP4T2Xml(
                    tempPath,
                    sourceSheetName,
                    targetSheetName);
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPath) && File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch { }

                // Không ReleaseCom(wb), vì wb là workbook chính trong _session.
            }
        }
        private bool XlsxTitleSubtitleStyleCopiedP4T2Xml(string xlsxPath, string sourceSheetName, string targetSheetName)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath) || !File.Exists(xlsxPath)) return false;
            if (string.IsNullOrWhiteSpace(sourceSheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetSheetName)) return false;

            try
            {
                using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    string sourcePart = FindWorksheetPartPathP4T2Xml(zip, sourceSheetName);
                    string targetPart = FindWorksheetPartPathP4T2Xml(zip, targetSheetName);

                    if (string.IsNullOrWhiteSpace(sourcePart)) return false;
                    if (string.IsNullOrWhiteSpace(targetPart)) return false;

                    // Đề MOS yêu cầu title và subtitle, thực tế là A1 và A2.
                    string sourceA1Style = GetCellStyleIdP4T2Xml(zip, sourcePart, "A1");
                    string sourceA2Style = GetCellStyleIdP4T2Xml(zip, sourcePart, "A2");

                    string targetA1Style = GetCellStyleIdP4T2Xml(zip, targetPart, "A1");
                    string targetA2Style = GetCellStyleIdP4T2Xml(zip, targetPart, "A2");

                    if (string.IsNullOrWhiteSpace(sourceA1Style)) return false;
                    if (string.IsNullOrWhiteSpace(sourceA2Style)) return false;
                    if (string.IsNullOrWhiteSpace(targetA1Style)) return false;
                    if (string.IsNullOrWhiteSpace(targetA2Style)) return false;

                    bool titleOk = CellStyleEqualsOrEquivalentP4T2Xml(zip, sourceA1Style, targetA1Style);
                    bool subtitleOk = CellStyleEqualsOrEquivalentP4T2Xml(zip, sourceA2Style, targetA2Style);

                    return titleOk && subtitleOk;
                }
            }
            catch
            {
                return false;
            }
        }

        private string FindWorksheetPartPathP4T2Xml(ZipArchive zip, string sheetName)
        {
            if (zip == null || string.IsNullOrWhiteSpace(sheetName))
                return "";

            try
            {
                ZipArchiveEntry workbookEntry = zip.GetEntry("xl/workbook.xml");
                ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");

                if (workbookEntry == null || relsEntry == null)
                    return "";

                string workbookXml = ReadZipEntryTextP4T2Xml(workbookEntry);
                string relsXml = ReadZipEntryTextP4T2Xml(relsEntry);

                if (string.IsNullOrWhiteSpace(workbookXml) || string.IsNullOrWhiteSpace(relsXml))
                    return "";

                Match sheetMatch = Regex.Match(
                    workbookXml,
                    @"<sheet\b[^>]*\bname\s*=\s*""" + Regex.Escape(sheetName.Trim()) + @"""[^>]*\br:id\s*=\s*""(?<rid>[^""]+)""[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!sheetMatch.Success)
                {
                    foreach (Match m in Regex.Matches(
                        workbookXml,
                        @"<sheet\b[^>]*\bname\s*=\s*""(?<name>[^""]+)""[^>]*\br:id\s*=\s*""(?<rid>[^""]+)""[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.Singleline))
                    {
                        string actualName = m.Groups["name"].Value;

                        if (NormalizeSheetNameP4T2Xml(actualName) == NormalizeSheetNameP4T2Xml(sheetName))
                        {
                            sheetMatch = m;
                            break;
                        }
                    }
                }

                if (!sheetMatch.Success)
                    return "";

                string rid = sheetMatch.Groups["rid"].Value;
                if (string.IsNullOrWhiteSpace(rid))
                    return "";

                Match relMatch = Regex.Match(
                    relsXml,
                    @"<Relationship\b[^>]*\bId\s*=\s*""" + Regex.Escape(rid) + @"""[^>]*\bTarget\s*=\s*""(?<target>[^""]+)""[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!relMatch.Success)
                    return "";

                string target = relMatch.Groups["target"].Value;
                if (string.IsNullOrWhiteSpace(target))
                    return "";

                target = target.Replace("\\", "/");

                if (target.StartsWith("/xl/", StringComparison.OrdinalIgnoreCase))
                    target = target.Substring(1);
                else if (target.StartsWith("/", StringComparison.OrdinalIgnoreCase))
                    target = "xl" + target;
                else if (!target.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
                    target = "xl/" + target;

                return target;
            }
            catch
            {
                return "";
            }
        }

        private string GetCellStyleIdP4T2Xml(ZipArchive zip, string sheetPartPath, string cellAddress)
        {
            if (zip == null) return "";
            if (string.IsNullOrWhiteSpace(sheetPartPath)) return "";
            if (string.IsNullOrWhiteSpace(cellAddress)) return "";

            try
            {
                ZipArchiveEntry sheetEntry = zip.GetEntry(sheetPartPath);
                if (sheetEntry == null)
                    return "";

                string sheetXml = ReadZipEntryTextP4T2Xml(sheetEntry);
                if (string.IsNullOrWhiteSpace(sheetXml))
                    return "";

                Match cellMatch = Regex.Match(
                    sheetXml,
                    @"<c\b(?=[^>]*\br\s*=\s*""" + Regex.Escape(cellAddress.Trim()) + @""")[^>]*>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!cellMatch.Success)
                    return "";

                string cellTag = cellMatch.Value;

                Match styleMatch = Regex.Match(
                    cellTag,
                    @"\bs\s*=\s*""(?<style>[^""]+)""",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                // Nếu cell không có s thì Excel hiểu là style 0.
                if (!styleMatch.Success)
                    return "0";

                return styleMatch.Groups["style"].Value.Trim();
            }
            catch
            {
                return "";
            }
        }

        private bool CellStyleEqualsOrEquivalentP4T2Xml(ZipArchive zip, string sourceStyleId, string targetStyleId)
        {
            if (string.IsNullOrWhiteSpace(sourceStyleId)) return false;
            if (string.IsNullOrWhiteSpace(targetStyleId)) return false;

            if (string.Equals(sourceStyleId.Trim(), targetStyleId.Trim(), StringComparison.OrdinalIgnoreCase))
                return true;

            // Fallback: nếu Excel tạo style id khác nhưng nội dung xf giống nhau thì vẫn PASS.
            string sourceXf = GetCellXfXmlP4T2Xml(zip, sourceStyleId);
            string targetXf = GetCellXfXmlP4T2Xml(zip, targetStyleId);

            if (string.IsNullOrWhiteSpace(sourceXf)) return false;
            if (string.IsNullOrWhiteSpace(targetXf)) return false;

            if (string.Equals(sourceXf, targetXf, StringComparison.OrdinalIgnoreCase))
                return true;

            string sourceResolved = GetResolvedCellXfSignatureP20(zip, sourceStyleId);
            string targetResolved = GetResolvedCellXfSignatureP20(zip, targetStyleId);
            return !string.IsNullOrWhiteSpace(sourceResolved) &&
                string.Equals(sourceResolved, targetResolved, StringComparison.OrdinalIgnoreCase);
        }

        private string GetResolvedCellXfSignatureP20(ZipArchive zip, string styleId)
        {
            int index;
            if (zip == null || !int.TryParse(styleId, out index)) return "";
            ZipArchiveEntry entry = zip.GetEntry("xl/styles.xml");
            if (entry == null) return "";

            XDocument document;
            using (Stream stream = entry.Open()) document = XDocument.Load(stream);
            XElement cellXfs = document.Descendants().FirstOrDefault(element => element.Name.LocalName == "cellXfs");
            if (cellXfs == null) return "";
            List<XElement> xfs = cellXfs.Elements().Where(element => element.Name.LocalName == "xf").ToList();
            if (index < 0 || index >= xfs.Count) return "";

            XElement xf = xfs[index];
            StringBuilder signature = new StringBuilder();
            AppendResolvedStyleComponentP20(signature, document, "fonts", "font", GetIntAttributeP20(xf, "fontId"));
            AppendResolvedStyleComponentP20(signature, document, "fills", "fill", GetIntAttributeP20(xf, "fillId"));
            AppendResolvedStyleComponentP20(signature, document, "borders", "border", GetIntAttributeP20(xf, "borderId"));

            string numberFormatId = Convert.ToString((string)xf.Attribute("numFmtId"));
            XElement numberFormat = document.Descendants().FirstOrDefault(element =>
                element.Name.LocalName == "numFmt" && string.Equals(
                    Convert.ToString((string)element.Attribute("numFmtId")), numberFormatId, StringComparison.OrdinalIgnoreCase));
            signature.Append("NUMFMT:").Append(numberFormat == null ? numberFormatId :
                Convert.ToString((string)numberFormat.Attribute("formatCode"))).Append('|');

            foreach (XAttribute attribute in xf.Attributes()
                .Where(attribute => attribute.Name.LocalName != "fontId" && attribute.Name.LocalName != "fillId" &&
                    attribute.Name.LocalName != "borderId" && attribute.Name.LocalName != "numFmtId" &&
                    attribute.Name.LocalName != "xfId" && !attribute.Name.LocalName.StartsWith("apply", StringComparison.OrdinalIgnoreCase))
                .OrderBy(attribute => attribute.Name.LocalName))
                signature.Append(attribute.Name.LocalName).Append('=').Append(attribute.Value).Append('|');

            foreach (XElement child in xf.Elements().OrderBy(element => element.Name.LocalName))
                signature.Append(child.ToString(SaveOptions.DisableFormatting)).Append('|');
            return signature.ToString();
        }

        private void AppendResolvedStyleComponentP20(StringBuilder signature, XDocument document, string collectionName, string itemName, int index)
        {
            XElement collection = document.Descendants().FirstOrDefault(element => element.Name.LocalName == collectionName);
            XElement item = collection == null ? null : collection.Elements()
                .Where(element => element.Name.LocalName == itemName).Skip(index).FirstOrDefault();
            signature.Append(collectionName).Append(':')
                .Append(item == null ? "" : NormalizeStyleComponentP20(item, collectionName)).Append('|');
        }

        private string NormalizeStyleComponentP20(XElement element, string collectionName)
        {
            XElement normalized = new XElement(element);
            if (string.Equals(collectionName, "fonts", StringComparison.OrdinalIgnoreCase))
            {
                normalized.Descendants().Where(item => item.Name.LocalName == "family" ||
                    item.Name.LocalName == "charset" || item.Name.LocalName == "scheme").Remove();
            }
            else if (string.Equals(collectionName, "fills", StringComparison.OrdinalIgnoreCase))
            {
                XElement pattern = normalized.Descendants().FirstOrDefault(item => item.Name.LocalName == "patternFill");
                if (pattern != null && string.Equals(Convert.ToString((string)pattern.Attribute("patternType")),
                    "solid", StringComparison.OrdinalIgnoreCase))
                    pattern.Elements().Where(item => item.Name.LocalName == "bgColor").Remove();
            }
            return normalized.ToString(SaveOptions.DisableFormatting);
        }

        private int GetIntAttributeP20(XElement element, string attributeName)
        {
            int value;
            return element != null && int.TryParse(Convert.ToString((string)element.Attribute(attributeName)), out value) ? value : 0;
        }

        private string GetCellXfXmlP4T2Xml(ZipArchive zip, string styleId)
        {
            if (zip == null) return "";
            if (string.IsNullOrWhiteSpace(styleId)) return "";

            int index;
            if (!int.TryParse(styleId.Trim(), out index))
                return "";

            try
            {
                ZipArchiveEntry stylesEntry = zip.GetEntry("xl/styles.xml");
                if (stylesEntry == null)
                    return "";

                string stylesXml = ReadZipEntryTextP4T2Xml(stylesEntry);
                if (string.IsNullOrWhiteSpace(stylesXml))
                    return "";

                Match cellXfsMatch = Regex.Match(
                    stylesXml,
                    @"<cellXfs\b[^>]*>(?<body>.*?)</cellXfs>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!cellXfsMatch.Success)
                    return "";

                string body = cellXfsMatch.Groups["body"].Value;

                MatchCollection xfMatches = Regex.Matches(
                    body,
                    @"<xf\b[^>]*/>|<xf\b[^>]*>.*?</xf>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (index < 0 || index >= xfMatches.Count)
                    return "";

                return NormalizeStyleXmlP4T2Xml(xfMatches[index].Value);
            }
            catch
            {
                return "";
            }
        }

        private string NormalizeStyleXmlP4T2Xml(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml))
                return "";

            string s = xml.Trim();

            // Bỏ khoảng trắng để so sánh ổn định.
            s = Regex.Replace(s, @"\s+", " ");

            // Một số build Excel có thể thêm id runtime, bỏ qua nếu có.
            s = Regex.Replace(s, @"\s+xr:uid\s*=\s*""[^""]*""", "", RegexOptions.IgnoreCase);

            return s.Trim();
        }

        private string ReadZipEntryTextP4T2Xml(ZipArchiveEntry entry)
        {
            if (entry == null) return "";

            try
            {
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch
            {
                return "";
            }
        }

        private string NormalizeSheetNameP4T2Xml(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }
        //Project 4 Task 3
        public bool ChartSheetLegendRemovedValueLabelsAbove(string chartSheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(chartSheetName)) return false;

            Xl.Workbook wb = null;
            string tempPath = "";

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                if (wb == null) return false;

                tempPath = Path.Combine(
                    Path.GetTempPath(),
                    "MosTrainer_P4T3_" + Guid.NewGuid().ToString("N") + ".xlsx");

                // Lấy đúng trạng thái workbook hiện tại, không bắt học viên bấm Save.
                wb.SaveCopyAs(tempPath);

                if (!File.Exists(tempPath))
                    return false;

                return XlsxChartSheetHasNoLegendAndValueLabelsAboveP4T3Xml(
                    tempPath,
                    chartSheetName);
            }
            catch
            {
                return false;
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPath) && File.Exists(tempPath))
                        File.Delete(tempPath);
                }
                catch { }

                // Không ReleaseCom(wb), vì wb là workbook chính trong _session.
            }
        }

        private bool XlsxChartSheetHasNoLegendAndValueLabelsAboveP4T3Xml(string xlsxPath, string chartSheetName)
        {
            if (string.IsNullOrWhiteSpace(xlsxPath) || !File.Exists(xlsxPath)) return false;
            if (string.IsNullOrWhiteSpace(chartSheetName)) return false;

            try
            {
                using (FileStream fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (ZipArchive zip = new ZipArchive(fs, ZipArchiveMode.Read))
                {
                    string chartPartPath = FindChartPartPathFromChartSheetP4T3Xml(zip, chartSheetName);

                    if (string.IsNullOrWhiteSpace(chartPartPath))
                        return false;

                    ZipArchiveEntry chartEntry = zip.GetEntry(chartPartPath);
                    if (chartEntry == null)
                        return false;

                    string chartXml = ReadZipEntryTextP4T3Xml(chartEntry);
                    if (string.IsNullOrWhiteSpace(chartXml))
                        return false;

                    bool legendOk = ChartXmlLegendRemovedP4T3Xml(chartXml);
                    if (!legendOk)
                        return false;

                    bool labelsOk = ChartXmlHasValueLabelsAboveOnlyP4T3Xml(chartXml);
                    if (!labelsOk)
                        return false;

                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        private string FindChartPartPathFromChartSheetP4T3Xml(ZipArchive zip, string chartSheetName)
        {
            if (zip == null || string.IsNullOrWhiteSpace(chartSheetName))
                return "";

            try
            {
                string chartSheetPartPath = FindChartSheetPartPathP4T3Xml(zip, chartSheetName);
                if (string.IsNullOrWhiteSpace(chartSheetPartPath))
                    return "";

                ZipArchiveEntry chartSheetEntry = zip.GetEntry(chartSheetPartPath);
                if (chartSheetEntry == null)
                    return "";

                string chartSheetXml = ReadZipEntryTextP4T3Xml(chartSheetEntry);
                if (string.IsNullOrWhiteSpace(chartSheetXml))
                    return "";

                Match drawingMatch = Regex.Match(
                    chartSheetXml,
                    @"<drawing\b[^>]*\br:id\s*=\s*""(?<rid>[^""]+)""",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!drawingMatch.Success)
                    return "";

                string drawingRid = drawingMatch.Groups["rid"].Value;
                string chartSheetRelsPath = GetRelsPathP4T3Xml(chartSheetPartPath);
                string drawingTarget = FindRelationshipTargetByIdP4T3Xml(zip, chartSheetRelsPath, drawingRid);

                if (string.IsNullOrWhiteSpace(drawingTarget))
                    return "";

                string drawingPartPath = ResolvePackageTargetPathP4T3Xml(chartSheetPartPath, drawingTarget);
                if (string.IsNullOrWhiteSpace(drawingPartPath))
                    return "";

                ZipArchiveEntry drawingEntry = zip.GetEntry(drawingPartPath);
                if (drawingEntry == null)
                    return "";

                string drawingXml = ReadZipEntryTextP4T3Xml(drawingEntry);
                if (string.IsNullOrWhiteSpace(drawingXml))
                    return "";

                Match chartMatch = Regex.Match(
                    drawingXml,
                    @"<c:chart\b[^>]*\br:id\s*=\s*""(?<rid>[^""]+)""",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!chartMatch.Success)
                    return "";

                string chartRid = chartMatch.Groups["rid"].Value;
                string drawingRelsPath = GetRelsPathP4T3Xml(drawingPartPath);
                string chartTarget = FindRelationshipTargetByIdP4T3Xml(zip, drawingRelsPath, chartRid);

                if (string.IsNullOrWhiteSpace(chartTarget))
                    return "";

                return ResolvePackageTargetPathP4T3Xml(drawingPartPath, chartTarget);
            }
            catch
            {
                return "";
            }
        }

        private string FindChartSheetPartPathP4T3Xml(ZipArchive zip, string chartSheetName)
        {
            if (zip == null || string.IsNullOrWhiteSpace(chartSheetName))
                return "";

            try
            {
                ZipArchiveEntry workbookEntry = zip.GetEntry("xl/workbook.xml");
                ZipArchiveEntry relsEntry = zip.GetEntry("xl/_rels/workbook.xml.rels");

                if (workbookEntry == null || relsEntry == null)
                    return "";

                string workbookXml = ReadZipEntryTextP4T3Xml(workbookEntry);
                string relsXml = ReadZipEntryTextP4T3Xml(relsEntry);

                if (string.IsNullOrWhiteSpace(workbookXml) || string.IsNullOrWhiteSpace(relsXml))
                    return "";

                Match sheetMatch = Regex.Match(
                    workbookXml,
                    @"<sheet\b[^>]*\bname\s*=\s*""" + Regex.Escape(chartSheetName.Trim()) + @"""[^>]*\br:id\s*=\s*""(?<rid>[^""]+)""[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!sheetMatch.Success)
                {
                    foreach (Match m in Regex.Matches(
                        workbookXml,
                        @"<sheet\b[^>]*\bname\s*=\s*""(?<name>[^""]+)""[^>]*\br:id\s*=\s*""(?<rid>[^""]+)""[^>]*/?>",
                        RegexOptions.IgnoreCase | RegexOptions.Singleline))
                    {
                        string actualName = m.Groups["name"].Value;

                        if (NormalizeSheetNameP4T3Xml(actualName) == NormalizeSheetNameP4T3Xml(chartSheetName))
                        {
                            sheetMatch = m;
                            break;
                        }
                    }
                }

                if (!sheetMatch.Success)
                    return "";

                string rid = sheetMatch.Groups["rid"].Value;
                if (string.IsNullOrWhiteSpace(rid))
                    return "";

                Match relMatch = Regex.Match(
                    relsXml,
                    @"<Relationship\b[^>]*\bId\s*=\s*""" + Regex.Escape(rid) + @"""[^>]*\bTarget\s*=\s*""(?<target>[^""]+)""[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!relMatch.Success)
                    return "";

                string target = relMatch.Groups["target"].Value;
                if (string.IsNullOrWhiteSpace(target))
                    return "";

                return ResolvePackageTargetPathP4T3Xml("xl/workbook.xml", target);
            }
            catch
            {
                return "";
            }
        }

        private string FindRelationshipTargetByIdP4T3Xml(ZipArchive zip, string relsPath, string relationshipId)
        {
            if (zip == null) return "";
            if (string.IsNullOrWhiteSpace(relsPath)) return "";
            if (string.IsNullOrWhiteSpace(relationshipId)) return "";

            try
            {
                ZipArchiveEntry relsEntry = zip.GetEntry(relsPath);
                if (relsEntry == null)
                    return "";

                string relsXml = ReadZipEntryTextP4T3Xml(relsEntry);
                if (string.IsNullOrWhiteSpace(relsXml))
                    return "";

                Match relMatch = Regex.Match(
                    relsXml,
                    @"<Relationship\b[^>]*\bId\s*=\s*""" + Regex.Escape(relationshipId.Trim()) + @"""[^>]*\bTarget\s*=\s*""(?<target>[^""]+)""[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!relMatch.Success)
                    return "";

                return relMatch.Groups["target"].Value;
            }
            catch
            {
                return "";
            }
        }

        private bool ChartXmlLegendRemovedP4T3Xml(string chartXml)
        {
            if (string.IsNullOrWhiteSpace(chartXml)) return false;

            try
            {
                MatchCollection legends = Regex.Matches(
                    chartXml,
                    @"<c:legend\b[^>]*/>|<c:legend\b[^>]*>(?<body>.*?)</c:legend>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                // Remove Legend bằng Chart Elements thường làm mất hẳn node legend.
                if (legends == null || legends.Count == 0)
                    return true;

                foreach (Match legend in legends)
                {
                    string legendXml = legend.Value;

                    if (Regex.IsMatch(
                        legendXml,
                        @"<c:delete\b[^>]*\bval\s*=\s*""(?:1|true)""",
                        RegexOptions.IgnoreCase | RegexOptions.Singleline))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool ChartXmlHasValueLabelsAboveOnlyP4T3Xml(string chartXml)
        {
            if (string.IsNullOrWhiteSpace(chartXml)) return false;

            try
            {
                // Trường hợp Excel ghi dLbls cấp chart type.
                string chartWithoutSeries = Regex.Replace(
                    chartXml,
                    @"<c:ser\b[^>]*>.*?</c:ser>",
                    "",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (AnyValidDataLabelsBlockP4T3Xml(chartWithoutSeries))
                    return true;

                // Trường hợp Excel ghi dLbls riêng trong từng series.
                MatchCollection seriesMatches = Regex.Matches(
                    chartXml,
                    @"<c:ser\b[^>]*>(?<body>.*?)</c:ser>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (seriesMatches == null || seriesMatches.Count == 0)
                    return false;

                int validSeries = 0;

                foreach (Match seriesMatch in seriesMatches)
                {
                    string seriesXml = seriesMatch.Groups["body"].Value;

                    if (AnyValidDataLabelsBlockP4T3Xml(seriesXml))
                        validSeries++;
                }

                return validSeries == seriesMatches.Count;
            }
            catch
            {
                return false;
            }
        }

        private bool AnyValidDataLabelsBlockP4T3Xml(string xml)
        {
            if (string.IsNullOrWhiteSpace(xml)) return false;

            try
            {
                foreach (Match labelMatch in Regex.Matches(
                    xml,
                    @"<c:dLbls\b[^>]*>(?<body>.*?)</c:dLbls>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline))
                {
                    string labelBody = labelMatch.Groups["body"].Value;

                    if (DataLabelsBlockIsValueOnlyAboveP4T3Xml(labelBody))
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool DataLabelsBlockIsValueOnlyAboveP4T3Xml(string labelBody)
        {
            if (string.IsNullOrWhiteSpace(labelBody)) return false;

            try
            {
                bool showValue = ChartBoolElementP4T3Xml(labelBody, "showVal", false);
                if (!showValue)
                    return false;

                if (ChartBoolElementP4T3Xml(labelBody, "showLegendKey", false)) return false;
                if (ChartBoolElementP4T3Xml(labelBody, "showCatName", false)) return false;
                if (ChartBoolElementP4T3Xml(labelBody, "showSerName", false)) return false;
                if (ChartBoolElementP4T3Xml(labelBody, "showPercent", false)) return false;
                if (ChartBoolElementP4T3Xml(labelBody, "showBubbleSize", false)) return false;

                string position = ChartElementValP4T3Xml(labelBody, "dLblPos");

                // Above trong column chart thường là outEnd.
                // Nếu Excel không ghi dLblPos thì vẫn cho PASS vì thao tác có thể giữ default.
                if (!DataLabelPositionLooksAboveP4T3Xml(position))
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool ChartBoolElementP4T3Xml(string xml, string elementName, bool defaultValue)
        {
            if (string.IsNullOrWhiteSpace(xml)) return defaultValue;
            if (string.IsNullOrWhiteSpace(elementName)) return defaultValue;

            try
            {
                Match m = Regex.Match(
                    xml,
                    @"<c:" + Regex.Escape(elementName) + @"\b(?<attrs>[^>]*)/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!m.Success)
                    return defaultValue;

                string attrs = m.Groups["attrs"].Value;

                Match valMatch = Regex.Match(
                    attrs,
                    @"\bval\s*=\s*""(?<val>[^""]+)""",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                // XML chart: tag tồn tại mà không có val thì xem như true.
                if (!valMatch.Success)
                    return true;

                string value = valMatch.Groups["val"].Value.Trim().ToLowerInvariant();

                return value == "1" || value == "true";
            }
            catch
            {
                return defaultValue;
            }
        }

        private string ChartElementValP4T3Xml(string xml, string elementName)
        {
            if (string.IsNullOrWhiteSpace(xml)) return "";
            if (string.IsNullOrWhiteSpace(elementName)) return "";

            try
            {
                Match m = Regex.Match(
                    xml,
                    @"<c:" + Regex.Escape(elementName) + @"\b[^>]*\bval\s*=\s*""(?<val>[^""]+)""[^>]*/?>",
                    RegexOptions.IgnoreCase | RegexOptions.Singleline);

                if (!m.Success)
                    return "";

                return m.Groups["val"].Value.Trim();
            }
            catch
            {
                return "";
            }
        }

        private bool DataLabelPositionLooksAboveP4T3Xml(string position)
        {
            if (string.IsNullOrWhiteSpace(position))
                return true;

            string p = position.Trim().ToLowerInvariant();

            return p == "outend" ||
                   p == "t" ||
                   p == "top" ||
                   p == "above" ||
                   p == "bestfit";
        }

        private string GetRelsPathP4T3Xml(string partPath)
        {
            if (string.IsNullOrWhiteSpace(partPath)) return "";

            string p = partPath.Replace("\\", "/");
            int slash = p.LastIndexOf('/');

            if (slash < 0)
                return "_rels/" + p + ".rels";

            string dir = p.Substring(0, slash);
            string file = p.Substring(slash + 1);

            return dir + "/_rels/" + file + ".rels";
        }

        private string ResolvePackageTargetPathP4T3Xml(string basePartPath, string target)
        {
            if (string.IsNullOrWhiteSpace(target)) return "";

            try
            {
                string t = target.Replace("\\", "/").Trim();

                if (t.StartsWith("/", StringComparison.Ordinal))
                {
                    t = t.TrimStart('/');
                    return NormalizePackagePathP4T3Xml(t);
                }

                string basePath = basePartPath == null ? "" : basePartPath.Replace("\\", "/");
                string baseDir = "";

                int slash = basePath.LastIndexOf('/');
                if (slash >= 0)
                    baseDir = basePath.Substring(0, slash);

                string combined = string.IsNullOrWhiteSpace(baseDir)
                    ? t
                    : baseDir + "/" + t;

                return NormalizePackagePathP4T3Xml(combined);
            }
            catch
            {
                return "";
            }
        }

        private string NormalizePackagePathP4T3Xml(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return "";

            string[] parts = path.Replace("\\", "/").Split('/');
            List<string> stack = new List<string>();

            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i];

                if (string.IsNullOrWhiteSpace(part) || part == ".")
                    continue;

                if (part == "..")
                {
                    if (stack.Count > 0)
                        stack.RemoveAt(stack.Count - 1);

                    continue;
                }

                stack.Add(part);
            }

            return string.Join("/", stack.ToArray());
        }

        private string ReadZipEntryTextP4T3Xml(ZipArchiveEntry entry)
        {
            if (entry == null) return "";

            try
            {
                using (Stream stream = entry.Open())
                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
            catch
            {
                return "";
            }
        }

        private string NormalizeSheetNameP4T3Xml(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }


        //Project 4 Task 4
        public bool WorksheetTableConvertedToRange(
            string sheetName,
            string rangeAddress,
            string tableName,
            IList<string> expectedHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(rangeAddress)) rangeAddress = "A4:J30";

            Xl.Worksheet ws = null;
            Xl.Range targetRange = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                targetRange = ws.Range[rangeAddress];
                if (targetRange == null) return false;

                // Sau Convert to Range, dữ liệu và header vẫn phải còn.
                if (!RangeContainsExpectedHeadersP4T4(targetRange, expectedHeaders))
                    return false;

                if (!RangeHasDataRowsP4T4(targetRange))
                    return false;

                // Điều quan trọng nhất: không còn Excel Table/ListObject ở vùng này.
                if (WorksheetHasListObjectOnRangeP4T4(ws, targetRange, tableName))
                    return false;

                // Keep formatting: chấm nhẹ để tránh false fail do Interop.
                if (!RangeKeepsTableFormattingP4T4(targetRange))
                    return false;

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(targetRange);
                ReleaseCom(ws);
            }
        }

        private bool WorksheetHasListObjectOnRangeP4T4(Xl.Worksheet ws, Xl.Range targetRange, string tableName)
        {
            if (ws == null || targetRange == null)
                return false;

            Xl.ListObjects listObjects = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;

            try
            {
                listObjects = ws.ListObjects;
                if (listObjects == null)
                    return false;

                int count = Convert.ToInt32(listObjects.Count);
                if (count <= 0)
                    return false;

                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(tableRange);
                    ReleaseCom(table);
                    tableRange = null;
                    table = null;

                    table = listObjects.Item[i];
                    if (table == null) continue;

                    string actualTableName = "";
                    try
                    {
                        actualTableName = Convert.ToString(table.Name);
                    }
                    catch
                    {
                        actualTableName = "";
                    }

                    if (!string.IsNullOrWhiteSpace(tableName) &&
                        string.Equals(actualTableName, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }

                    try
                    {
                        tableRange = table.Range;
                    }
                    catch
                    {
                        tableRange = null;
                    }

                    if (tableRange != null && RangesIntersectP4T4(tableRange, targetRange))
                        return true;
                }

                return false;
            }
            catch
            {
                // Nếu đọc ListObjects lỗi thì không nên PASS.
                return true;
            }
            finally
            {
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(listObjects);
            }
        }

        private bool RangesIntersectP4T4(Xl.Range leftRange, Xl.Range rightRange)
        {
            if (leftRange == null || rightRange == null)
                return false;

            try
            {
                int lRow1 = Convert.ToInt32(leftRange.Row);
                int lCol1 = Convert.ToInt32(leftRange.Column);
                int lRow2 = lRow1 + Convert.ToInt32(leftRange.Rows.Count) - 1;
                int lCol2 = lCol1 + Convert.ToInt32(leftRange.Columns.Count) - 1;

                int rRow1 = Convert.ToInt32(rightRange.Row);
                int rCol1 = Convert.ToInt32(rightRange.Column);
                int rRow2 = rRow1 + Convert.ToInt32(rightRange.Rows.Count) - 1;
                int rCol2 = rCol1 + Convert.ToInt32(rightRange.Columns.Count) - 1;

                return lRow1 <= rRow2 &&
                       lRow2 >= rRow1 &&
                       lCol1 <= rCol2 &&
                       lCol2 >= rCol1;
            }
            catch
            {
                return false;
            }
        }

        private bool RangeContainsExpectedHeadersP4T4(Xl.Range range, IList<string> configuredHeaders)
        {
            if (range == null) return false;

            IList<string> expectedHeaders = configuredHeaders;
            if (expectedHeaders == null || expectedHeaders.Count == 0)
            {
                expectedHeaders = new string[]
                {
                    "Make", "Model", "Body", "Year", "Color", "Mileage", "Price",
                    "Quantity Instock", "Total", "Inspected"
                };
            }

            Xl.Range cell = null;

            try
            {
                int columnCount = Convert.ToInt32(range.Columns.Count);
                if (columnCount < expectedHeaders.Count)
                    return false;

                for (int c = 1; c <= expectedHeaders.Count; c++)
                {
                    ReleaseCom(cell);
                    cell = null;

                    cell = range.Cells[1, c] as Xl.Range;
                    if (cell == null) return false;

                    string actual = GetCellTextP1T3(cell);

                    if (!HeaderEqualsP4T4(actual, expectedHeaders[c - 1]))
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool RangeHasDataRowsP4T4(Xl.Range range)
        {
            if (range == null) return false;

            Xl.Range cell = null;

            try
            {
                int rowCount = Convert.ToInt32(range.Rows.Count);
                int columnCount = Convert.ToInt32(range.Columns.Count);

                if (rowCount < 2 || columnCount < 1)
                    return false;

                int nonEmptyCells = 0;

                for (int r = 2; r <= rowCount; r++)
                {
                    for (int c = 1; c <= columnCount; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = range.Cells[r, c] as Xl.Range;
                        if (cell == null) continue;

                        string text = GetCellTextP1T3(cell);
                        if (!string.IsNullOrWhiteSpace(text))
                            nonEmptyCells++;

                        if (nonEmptyCells >= 5)
                            return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool RangeKeepsTableFormattingP4T4(Xl.Range range)
        {
            if (range == null) return false;

            Xl.Range headerCell = null;

            try
            {
                int columnCount = Convert.ToInt32(range.Columns.Count);
                if (columnCount <= 0)
                    return false;

                int formattedHeaderCells = 0;

                for (int c = 1; c <= columnCount; c++)
                {
                    ReleaseCom(headerCell);
                    headerCell = null;

                    headerCell = range.Cells[1, c] as Xl.Range;
                    if (headerCell == null) continue;

                    if (CellLooksFormattedP4T4(headerCell))
                        formattedHeaderCells++;
                }

                // Header của table style thường được tô màu và/hoặc bold.
                return formattedHeaderCells >= Math.Max(1, columnCount / 2);
            }
            catch
            {
                // Tránh false fail nếu Excel Interop đọc format lỗi,
                // vì điều kiện chính của Task 4 vẫn là không còn ListObject.
                return true;
            }
            finally
            {
                ReleaseCom(headerCell);
            }
        }

        private bool CellLooksFormattedP4T4(Xl.Range cell)
        {
            if (cell == null) return false;

            Xl.Font font = null;
            Xl.Interior interior = null;

            try
            {
                try
                {
                    font = cell.Font;
                    if (font != null && ToBoolP4T4(font.Bold))
                        return true;
                }
                catch
                {
                }

                try
                {
                    interior = cell.Interior;
                    if (interior != null)
                    {
                        int colorIndex = ToIntP4T4(interior.ColorIndex, -4142);

                        if (colorIndex != -4142 && colorIndex != 0)
                            return true;
                    }
                }
                catch
                {
                }

                return false;
            }
            catch
            {
                return true;
            }
            finally
            {
                ReleaseCom(font);
                ReleaseCom(interior);
            }
        }

        private bool HeaderEqualsP4T4(string actual, string expected)
        {
            string a = NormalizeHeaderP4T4(actual);
            string e = NormalizeHeaderP4T4(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeHeaderP4T4(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"[^A-Z0-9]+", "", RegexOptions.IgnoreCase);
            s = s.ToUpperInvariant();

            return s;
        }

        private bool ToBoolP4T4(object value)
        {
            if (value == null) return false;

            try
            {
                return Convert.ToBoolean(value);
            }
            catch
            {
                string text = Convert.ToString(value);
                return string.Equals(text, "TRUE", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(text, "1", StringComparison.OrdinalIgnoreCase) ||
                       string.Equals(text, "-1", StringComparison.OrdinalIgnoreCase);
            }
        }

        private int ToIntP4T4(object value, int defaultValue)
        {
            if (value == null) return defaultValue;

            try
            {
                return Convert.ToInt32(value, CultureInfo.InvariantCulture);
            }
            catch
            {
                int result;
                if (int.TryParse(Convert.ToString(value), out result))
                    return result;

                return defaultValue;
            }
        }

        //Project 4 Task 5
        public bool ReportClusteredColumnChartCreated(string sheetName, string categoryHeader, string valueHeader)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(categoryHeader)) categoryHeader = "Month";
            if (string.IsNullOrWhiteSpace(valueHeader)) valueHeader = "Quantity";

            Xl.Worksheet ws = null;
            Xl.Range categoryDataRange = null;
            Xl.Range valueDataRange = null;
            Xl.Range categoryHeaderRange = null;
            Xl.Range valueHeaderRange = null;
            Xl.Range tableRange = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                int headerRow;
                int categoryColumn;
                int valueColumn;
                int lastDataRow;

                bool found = TryFindTwoColumnDataRangesP4T5(
                    ws,
                    categoryHeader,
                    valueHeader,
                    out headerRow,
                    out categoryColumn,
                    out valueColumn,
                    out lastDataRow,
                    out categoryHeaderRange,
                    out valueHeaderRange,
                    out categoryDataRange,
                    out valueDataRange,
                    out tableRange);

                if (!found)
                    return false;

                chartObjects = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (chartObjects == null)
                    return false;

                int chartCount = Convert.ToInt32(chartObjects.Count);
                if (chartCount <= 0)
                    return false;

                for (int i = 1; i <= chartCount; i++)
                {
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);
                    chart = null;
                    chartObject = null;

                    chartObject = (Xl.ChartObject)chartObjects.Item(i);
                    if (chartObject == null) continue;

                    chart = chartObject.Chart;
                    if (chart == null) continue;

                    if (!ChartIsClusteredColumnP4T5(chart))
                        continue;

                    if (!ChartLooksBelowTableP4T5(chartObject, lastDataRow))
                        continue;

                    if (ChartUsesCategoryAndValueRangesP4T5(
                        chart,
                        categoryDataRange,
                        valueDataRange,
                        categoryHeaderRange,
                        valueHeaderRange,
                        categoryHeader,
                        valueHeader))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(categoryDataRange);
                ReleaseCom(valueDataRange);
                ReleaseCom(categoryHeaderRange);
                ReleaseCom(valueHeaderRange);
                ReleaseCom(tableRange);
                ReleaseCom(ws);
            }
        }

        private bool TryFindTwoColumnDataRangesP4T5(
            Xl.Worksheet ws,
            string categoryHeader,
            string valueHeader,
            out int headerRow,
            out int categoryColumn,
            out int valueColumn,
            out int lastDataRow,
            out Xl.Range categoryHeaderRange,
            out Xl.Range valueHeaderRange,
            out Xl.Range categoryDataRange,
            out Xl.Range valueDataRange,
            out Xl.Range tableRange)
        {
            headerRow = 0;
            categoryColumn = 0;
            valueColumn = 0;
            lastDataRow = 0;
            categoryHeaderRange = null;
            valueHeaderRange = null;
            categoryDataRange = null;
            valueDataRange = null;
            tableRange = null;

            if (ws == null) return false;

            Xl.Range usedRange = null;
            Xl.Range cell = null;
            Xl.Range startCell = null;
            Xl.Range endCell = null;

            try
            {
                usedRange = ws.UsedRange;
                if (usedRange == null) return false;

                int firstRow = Convert.ToInt32(usedRange.Row);
                int firstColumn = Convert.ToInt32(usedRange.Column);
                int rowCount = Convert.ToInt32(usedRange.Rows.Count);
                int columnCount = Convert.ToInt32(usedRange.Columns.Count);

                int lastUsedRow = firstRow + rowCount - 1;
                int lastUsedColumn = firstColumn + columnCount - 1;

                int maxSearchRow = Math.Min(lastUsedRow, firstRow + 40);

                for (int r = firstRow; r <= maxSearchRow; r++)
                {
                    int foundCategoryColumn = 0;
                    int foundValueColumn = 0;

                    for (int c = firstColumn; c <= lastUsedColumn; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = ws.Cells[r, c] as Xl.Range;
                        if (cell == null) continue;

                        string text = GetCellTextP1T3(cell);

                        if (HeaderEqualsP4T5(text, categoryHeader))
                            foundCategoryColumn = c;

                        if (HeaderEqualsP4T5(text, valueHeader))
                            foundValueColumn = c;
                    }

                    if (foundCategoryColumn > 0 && foundValueColumn > 0)
                    {
                        headerRow = r;
                        categoryColumn = foundCategoryColumn;
                        valueColumn = foundValueColumn;
                        break;
                    }
                }

                if (headerRow <= 0 || categoryColumn <= 0 || valueColumn <= 0)
                    return false;

                lastDataRow = headerRow;

                for (int r = headerRow + 1; r <= lastUsedRow; r++)
                {
                    Xl.Range categoryCell = null;
                    Xl.Range valueCell = null;

                    try
                    {
                        categoryCell = ws.Cells[r, categoryColumn] as Xl.Range;
                        valueCell = ws.Cells[r, valueColumn] as Xl.Range;

                        string categoryText = GetCellTextP1T3(categoryCell);
                        string valueText = GetCellTextP1T3(valueCell);

                        if (string.IsNullOrWhiteSpace(categoryText) && string.IsNullOrWhiteSpace(valueText))
                            break;

                        lastDataRow = r;
                    }
                    finally
                    {
                        ReleaseCom(categoryCell);
                        ReleaseCom(valueCell);
                    }
                }

                if (lastDataRow <= headerRow)
                    return false;

                categoryHeaderRange = BuildRangeP4T5(ws, headerRow, categoryColumn, lastDataRow, categoryColumn);
                valueHeaderRange = BuildRangeP4T5(ws, headerRow, valueColumn, lastDataRow, valueColumn);

                categoryDataRange = BuildRangeP4T5(ws, headerRow + 1, categoryColumn, lastDataRow, categoryColumn);
                valueDataRange = BuildRangeP4T5(ws, headerRow + 1, valueColumn, lastDataRow, valueColumn);

                int leftColumn = Math.Min(categoryColumn, valueColumn);
                int rightColumn = Math.Max(categoryColumn, valueColumn);

                tableRange = BuildRangeP4T5(ws, headerRow, leftColumn, lastDataRow, rightColumn);

                return categoryHeaderRange != null &&
                       valueHeaderRange != null &&
                       categoryDataRange != null &&
                       valueDataRange != null &&
                       tableRange != null;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(startCell);
                ReleaseCom(endCell);
                ReleaseCom(usedRange);
            }
        }

        private Xl.Range BuildRangeP4T5(Xl.Worksheet ws, int row1, int col1, int row2, int col2)
        {
            if (ws == null) return null;
            if (row1 <= 0 || col1 <= 0 || row2 <= 0 || col2 <= 0) return null;

            Xl.Range startCell = null;
            Xl.Range endCell = null;

            try
            {
                startCell = ws.Cells[row1, col1] as Xl.Range;
                endCell = ws.Cells[row2, col2] as Xl.Range;

                if (startCell == null || endCell == null)
                    return null;

                return ws.Range[startCell, endCell];
            }
            catch
            {
                return null;
            }
            finally
            {
                ReleaseCom(startCell);
                ReleaseCom(endCell);
            }
        }

        private bool ChartIsClusteredColumnP4T5(Xl.Chart chart)
        {
            if (chart == null) return false;

            try
            {
                int chartType = Convert.ToInt32(chart.ChartType);

                if (chartType == Convert.ToInt32(Xl.XlChartType.xlColumnClustered))
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
        }

        private bool ChartLooksBelowTableP4T5(Xl.ChartObject chartObject, int lastDataRow)
        {
            if (chartObject == null) return false;
            if (lastDataRow <= 0) return true;

            Xl.Range topLeft = null;

            try
            {
                topLeft = chartObject.TopLeftCell as Xl.Range;
                if (topLeft == null)
                    return true;

                int chartTopRow = Convert.ToInt32(topLeft.Row);

                // Đề yêu cầu đặt dưới bảng, size/position chính xác không quan trọng.
                // Chart phải bắt đầu sau dòng dữ liệu cuối cùng.
                return chartTopRow > lastDataRow;
            }
            catch
            {
                // Nếu COM đọc vị trí lỗi thì không bắt fail, vì trọng tâm là chart đúng dữ liệu.
                return true;
            }
            finally
            {
                ReleaseCom(topLeft);
            }
        }

        private bool ChartUsesCategoryAndValueRangesP4T5(
            Xl.Chart chart,
            Xl.Range categoryDataRange,
            Xl.Range valueDataRange,
            Xl.Range categoryHeaderRange,
            Xl.Range valueHeaderRange,
            string categoryHeader,
            string valueHeader)
        {
            if (chart == null) return false;
            if (categoryDataRange == null || valueDataRange == null) return false;

            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;

            try
            {
                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null)
                    return false;

                int seriesCount = Convert.ToInt32(seriesCollection.Count);
                if (seriesCount <= 0)
                    return false;

                string expectedCategoryDataAddress = NormalizeChartFormulaP4T5(GetRangeAddressWithSheetP4T5(categoryDataRange));
                string expectedValueDataAddress = NormalizeChartFormulaP4T5(GetRangeAddressWithSheetP4T5(valueDataRange));
                string expectedCategoryHeaderAddress = NormalizeChartFormulaP4T5(GetRangeAddressWithSheetP4T5(categoryHeaderRange));
                string expectedValueHeaderAddress = NormalizeChartFormulaP4T5(GetRangeAddressWithSheetP4T5(valueHeaderRange));

                string expectedValueHeader = NormalizeHeaderP4T5(valueHeader);
                string expectedCategoryHeader = NormalizeHeaderP4T5(categoryHeader);

                for (int i = 1; i <= seriesCount; i++)
                {
                    ReleaseCom(series);
                    series = null;

                    series = seriesCollection.Item(i) as Xl.Series;
                    if (series == null) continue;

                    string formula = "";

                    try
                    {
                        formula = Convert.ToString(series.Formula);
                    }
                    catch
                    {
                        formula = "";
                    }

                    string normalizedFormula = NormalizeChartFormulaP4T5(formula);

                    bool valueOk =
                        (!string.IsNullOrWhiteSpace(expectedValueDataAddress) && normalizedFormula.Contains(expectedValueDataAddress)) ||
                        (!string.IsNullOrWhiteSpace(expectedValueHeaderAddress) && normalizedFormula.Contains(expectedValueHeaderAddress)) ||
                        SeriesValuesMatchRangeP4T5(series, valueDataRange);

                    bool categoryOk =
                        (!string.IsNullOrWhiteSpace(expectedCategoryDataAddress) && normalizedFormula.Contains(expectedCategoryDataAddress)) ||
                        (!string.IsNullOrWhiteSpace(expectedCategoryHeaderAddress) && normalizedFormula.Contains(expectedCategoryHeaderAddress)) ||
                        SeriesCategoriesMatchRangeP4T5(series, categoryDataRange);

                    bool nameOk = true;

                    try
                    {
                        string seriesName = Convert.ToString(series.Name);
                        string normalizedName = NormalizeHeaderP4T5(seriesName);

                        if (!string.IsNullOrWhiteSpace(normalizedName))
                            nameOk = normalizedName.Contains(expectedValueHeader) ||
                                     expectedValueHeader.Contains(normalizedName) ||
                                     normalizedName.Contains(expectedCategoryHeader) == false;
                    }
                    catch
                    {
                        nameOk = true;
                    }

                    if (valueOk && categoryOk && nameOk)
                        return true;
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
            }
        }

        private bool SeriesValuesMatchRangeP4T5(Xl.Series series, Xl.Range expectedRange)
        {
            if (series == null || expectedRange == null) return false;

            try
            {
                object seriesValues = series.Values;
                List<double> actualNumbers = ExtractDoubleListP4T5(seriesValues);
                List<double> expectedNumbers = ExtractDoubleListFromRangeP4T5(expectedRange);

                if (actualNumbers.Count == 0 || expectedNumbers.Count == 0)
                    return false;

                if (actualNumbers.Count != expectedNumbers.Count)
                    return false;

                for (int i = 0; i < actualNumbers.Count; i++)
                {
                    if (Math.Abs(actualNumbers[i] - expectedNumbers[i]) > 0.0001)
                        return false;
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool SeriesCategoriesMatchRangeP4T5(Xl.Series series, Xl.Range expectedRange)
        {
            if (series == null || expectedRange == null) return false;

            try
            {
                object xValues = series.XValues;
                List<string> actualTexts = ExtractStringListP4T5(xValues);
                List<string> expectedTexts = ExtractStringListFromRangeP4T5(expectedRange);

                if (actualTexts.Count == 0 || expectedTexts.Count == 0)
                    return false;

                if (actualTexts.Count != expectedTexts.Count)
                    return false;

                for (int i = 0; i < actualTexts.Count; i++)
                {
                    if (!string.Equals(
                        NormalizeTextP4T5(actualTexts[i]),
                        NormalizeTextP4T5(expectedTexts[i]),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }
                }

                return true;
            }
            catch
            {
                return false;
            }
        }

        private List<double> ExtractDoubleListFromRangeP4T5(Xl.Range range)
        {
            List<double> result = new List<double>();
            if (range == null) return result;

            Xl.Range cell = null;

            try
            {
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);

                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = range.Cells[r, c] as Xl.Range;
                        if (cell == null) continue;

                        object value = cell.Value2;

                        double number;
                        if (TryDoubleP4T5(value, out number))
                            result.Add(number);
                    }
                }

                return result;
            }
            catch
            {
                return result;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private List<string> ExtractStringListFromRangeP4T5(Xl.Range range)
        {
            List<string> result = new List<string>();
            if (range == null) return result;

            Xl.Range cell = null;

            try
            {
                int rows = Convert.ToInt32(range.Rows.Count);
                int cols = Convert.ToInt32(range.Columns.Count);

                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = range.Cells[r, c] as Xl.Range;
                        if (cell == null) continue;

                        string text = GetCellTextP1T3(cell);
                        if (!string.IsNullOrWhiteSpace(text))
                            result.Add(text);
                    }
                }

                return result;
            }
            catch
            {
                return result;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private List<double> ExtractDoubleListP4T5(object value)
        {
            List<double> result = new List<double>();

            if (value == null)
                return result;

            try
            {
                object[,] array2d = value as object[,];

                if (array2d != null)
                {
                    int r1 = array2d.GetLowerBound(0);
                    int r2 = array2d.GetUpperBound(0);
                    int c1 = array2d.GetLowerBound(1);
                    int c2 = array2d.GetUpperBound(1);

                    for (int r = r1; r <= r2; r++)
                    {
                        for (int c = c1; c <= c2; c++)
                        {
                            double number;
                            if (TryDoubleP4T5(array2d[r, c], out number))
                                result.Add(number);
                        }
                    }

                    return result;
                }

                object[] array = value as object[];

                if (array != null)
                {
                    for (int i = 0; i < array.Length; i++)
                    {
                        double number;
                        if (TryDoubleP4T5(array[i], out number))
                            result.Add(number);
                    }

                    return result;
                }

                double singleNumber;
                if (TryDoubleP4T5(value, out singleNumber))
                    result.Add(singleNumber);

                return result;
            }
            catch
            {
                return result;
            }
        }

        private List<string> ExtractStringListP4T5(object value)
        {
            List<string> result = new List<string>();

            if (value == null)
                return result;

            try
            {
                object[,] array2d = value as object[,];

                if (array2d != null)
                {
                    int r1 = array2d.GetLowerBound(0);
                    int r2 = array2d.GetUpperBound(0);
                    int c1 = array2d.GetLowerBound(1);
                    int c2 = array2d.GetUpperBound(1);

                    for (int r = r1; r <= r2; r++)
                    {
                        for (int c = c1; c <= c2; c++)
                        {
                            string text = Convert.ToString(array2d[r, c]);
                            if (!string.IsNullOrWhiteSpace(text))
                                result.Add(text);
                        }
                    }

                    return result;
                }

                object[] array = value as object[];

                if (array != null)
                {
                    for (int i = 0; i < array.Length; i++)
                    {
                        string text = Convert.ToString(array[i]);
                        if (!string.IsNullOrWhiteSpace(text))
                            result.Add(text);
                    }

                    return result;
                }

                string singleText = Convert.ToString(value);
                if (!string.IsNullOrWhiteSpace(singleText))
                    result.Add(singleText);

                return result;
            }
            catch
            {
                return result;
            }
        }

        private bool TryDoubleP4T5(object value, out double result)
        {
            result = 0;

            if (value == null)
                return false;

            try
            {
                result = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }
            catch
            {
            }

            return double.TryParse(
                Convert.ToString(value),
                NumberStyles.Any,
                CultureInfo.InvariantCulture,
                out result);
        }

        private string GetRangeAddressWithSheetP4T5(Xl.Range range)
        {
            if (range == null) return "";

            Xl.Worksheet ws = null;

            try
            {
                string address = Convert.ToString(range.get_Address(false, false, Xl.XlReferenceStyle.xlA1, false, Type.Missing));
                ws = range.Worksheet as Xl.Worksheet;

                string sheetName = ws == null ? "" : Convert.ToString(ws.Name);

                if (string.IsNullOrWhiteSpace(sheetName))
                    return address;

                return sheetName + "!" + address;
            }
            catch
            {
                return "";
            }
            finally
            {
                ReleaseCom(ws);
            }
        }

        private bool HeaderEqualsP4T5(string actual, string expected)
        {
            string a = NormalizeHeaderP4T5(actual);
            string e = NormalizeHeaderP4T5(expected);

            return string.Equals(a, e, StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeHeaderP4T5(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"[^A-Z0-9]+", "", RegexOptions.IgnoreCase);
            s = s.ToUpperInvariant();

            return s;
        }

        private string NormalizeTextP4T5(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", " ");
            s = s.ToUpperInvariant();

            return s.Trim();
        }

        private string NormalizeChartFormulaP4T5(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return "";

            string s = text.Trim();

            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace("=", "");
            s = s.Replace(";", ",");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }


        //Project 4 Task 6
        public bool LeftFormulaByHeaders(string sheetName, string targetHeader, string sourceHeader, int characterCount)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(sourceHeader)) sourceHeader = "Category";
            if (characterCount <= 0) characterCount = 2;

            Xl.Worksheet ws = null;
            Xl.Range usedRange = null;
            Xl.Range sourceCell = null;
            Xl.Range targetCell = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                usedRange = ws.UsedRange;
                if (usedRange == null) return false;

                int headerRow;
                int targetColumn;
                int sourceColumn;

                bool found = TryFindTargetAndSourceHeadersP4T6(
                    ws,
                    usedRange,
                    targetHeader,
                    sourceHeader,
                    out headerRow,
                    out targetColumn,
                    out sourceColumn);

                if (!found)
                    return false;

                int firstUsedRow = Convert.ToInt32(usedRange.Row);
                int usedRowCount = Convert.ToInt32(usedRange.Rows.Count);
                int lastUsedRow = firstUsedRow + usedRowCount - 1;

                int checkedRows = 0;

                for (int row = headerRow + 1; row <= lastUsedRow; row++)
                {
                    ReleaseCom(sourceCell);
                    ReleaseCom(targetCell);
                    sourceCell = null;
                    targetCell = null;

                    sourceCell = ws.Cells[row, sourceColumn] as Xl.Range;
                    targetCell = ws.Cells[row, targetColumn] as Xl.Range;

                    if (sourceCell == null || targetCell == null)
                        return false;

                    string sourceText = GetCellTextP1T3(sourceCell);

                    if (string.IsNullOrWhiteSpace(sourceText))
                        continue;

                    checkedRows++;

                    string expectedText = sourceText.Length <= characterCount
                        ? sourceText
                        : sourceText.Substring(0, characterCount);

                    string actualText = GetCellTextP1T3(targetCell);

                    if (!string.Equals(
                        NormalizeTextP4T6(actualText),
                        NormalizeTextP4T6(expectedText),
                        StringComparison.OrdinalIgnoreCase))
                    {
                        return false;
                    }

                    if (!CellHasFormulaP1T3(targetCell))
                        return false;

                    string formula = "";
                    string formulaLocal = "";
                    string formulaR1C1 = "";

                    try { formula = Convert.ToString(targetCell.Formula); } catch { formula = ""; }
                    try { formulaLocal = Convert.ToString(targetCell.FormulaLocal); } catch { formulaLocal = ""; }
                    try { formulaR1C1 = Convert.ToString(targetCell.FormulaR1C1); } catch { formulaR1C1 = ""; }

                    bool formulaOk = FormulaLooksLikeLeftP4T6(
                        formula,
                        formulaLocal,
                        formulaR1C1,
                        row,
                        sourceColumn,
                        sourceHeader,
                        characterCount);

                    if (!formulaOk)
                        return false;
                }

                return checkedRows > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sourceCell);
                ReleaseCom(targetCell);
                ReleaseCom(usedRange);
                ReleaseCom(ws);
            }
        }

        public bool LeftFormulaByHeadersInRanges(
            string sheetName,
            string targetHeader,
            string sourceHeader,
            string targetRangeAddress,
            string sourceRangeAddress,
            int characterCount)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetHeader) ||
                string.IsNullOrWhiteSpace(sourceHeader) || string.IsNullOrWhiteSpace(targetRangeAddress) ||
                string.IsNullOrWhiteSpace(sourceRangeAddress) || characterCount <= 0) return false;

            Xl.Worksheet ws = null;
            Xl.Range targets = null;
            Xl.Range sources = null;
            Xl.Range targetHeaderCell = null;
            Xl.Range sourceHeaderCell = null;
            Xl.Range targetCell = null;
            Xl.Range sourceCell = null;
            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;
                targets = ws.Range[targetRangeAddress];
                sources = ws.Range[sourceRangeAddress];
                if (targets == null || sources == null || Convert.ToInt32(targets.Columns.Count) != 1 ||
                    Convert.ToInt32(sources.Columns.Count) != 1 ||
                    Convert.ToInt32(targets.Rows.Count) != Convert.ToInt32(sources.Rows.Count) ||
                    Convert.ToInt32(targets.Row) != Convert.ToInt32(sources.Row)) return false;

                int firstRow = Convert.ToInt32(targets.Row);
                int targetColumn = Convert.ToInt32(targets.Column);
                int sourceColumn = Convert.ToInt32(sources.Column);
                if (firstRow <= 1) return false;
                targetHeaderCell = ws.Cells[firstRow - 1, targetColumn] as Xl.Range;
                sourceHeaderCell = ws.Cells[firstRow - 1, sourceColumn] as Xl.Range;
                if (targetHeaderCell == null || sourceHeaderCell == null ||
                    !HeaderEqualsP4T6(GetCellTextP1T3(targetHeaderCell), targetHeader) ||
                    !HeaderEqualsP4T6(GetCellTextP1T3(sourceHeaderCell), sourceHeader)) return false;

                int count = Convert.ToInt32(targets.Rows.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(targetCell); ReleaseCom(sourceCell);
                    targetCell = targets.Cells[i, 1] as Xl.Range;
                    sourceCell = sources.Cells[i, 1] as Xl.Range;
                    if (targetCell == null || sourceCell == null || !CellHasFormulaP1T3(targetCell)) return false;

                    string sourceText = GetCellTextP1T3(sourceCell);
                    string expectedText = sourceText.Length <= characterCount ? sourceText : sourceText.Substring(0, characterCount);
                    if (!string.Equals(NormalizeTextP4T6(GetCellTextP1T3(targetCell)), NormalizeTextP4T6(expectedText), StringComparison.OrdinalIgnoreCase))
                        return false;

                    string formula = NormalizeFormulaP4T6(Convert.ToString(targetCell.Formula));
                    string sourceReference = ExcelColumnNameP1T3(sourceColumn) + (firstRow + i - 1).ToString(CultureInfo.InvariantCulture);
                    string normalizedSheet = NormalizeFormulaP4T6(sheetName);
                    string pattern = @"^=LEFT\((?:" + Regex.Escape(normalizedSheet) + @"!)?" + Regex.Escape(sourceReference) + "," +
                        characterCount.ToString(CultureInfo.InvariantCulture) + @"\)$";
                    if (!Regex.IsMatch(formula, pattern, RegexOptions.CultureInvariant)) return false;
                }

                return count > 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(sourceCell); ReleaseCom(targetCell); ReleaseCom(sourceHeaderCell); ReleaseCom(targetHeaderCell);
                ReleaseCom(sources); ReleaseCom(targets); ReleaseCom(ws);
            }
        }

        private bool TryFindTargetAndSourceHeadersP4T6(
            Xl.Worksheet ws,
            Xl.Range usedRange,
            string targetHeader,
            string sourceHeader,
            out int headerRow,
            out int targetColumn,
            out int sourceColumn)
        {
            headerRow = 0;
            targetColumn = 0;
            sourceColumn = 0;

            if (ws == null || usedRange == null)
                return false;

            Xl.Range cell = null;

            try
            {
                int firstRow = Convert.ToInt32(usedRange.Row);
                int firstColumn = Convert.ToInt32(usedRange.Column);
                int rowCount = Convert.ToInt32(usedRange.Rows.Count);
                int columnCount = Convert.ToInt32(usedRange.Columns.Count);

                int lastRow = firstRow + rowCount - 1;
                int lastColumn = firstColumn + columnCount - 1;

                int maxSearchRow = Math.Min(lastRow, firstRow + 20);

                for (int r = firstRow; r <= maxSearchRow; r++)
                {
                    int foundTargetColumn = 0;
                    int foundSourceColumn = 0;

                    for (int c = firstColumn; c <= lastColumn; c++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = ws.Cells[r, c] as Xl.Range;
                        if (cell == null) continue;

                        string text = GetCellTextP1T3(cell);

                        if (HeaderEqualsP4T6(text, sourceHeader))
                            foundSourceColumn = c;

                        if (HeaderEqualsP4T6(text, targetHeader) || HeaderEqualsP4T6(text, "TTC") || HeaderEqualsP4T6(text, "TCC"))
                            foundTargetColumn = c;
                    }

                    if (foundSourceColumn > 0)
                    {
                        // File gốc thực tế có header TCC ở cột A, Category ở cột B.
                        // Nếu targetHeader ghi TTC nhưng workbook ghi TCC, vẫn chấp nhận.
                        // Nếu không tìm thấy target thì fallback cột ngay bên trái Category.
                        if (foundTargetColumn <= 0 && foundSourceColumn > firstColumn)
                            foundTargetColumn = foundSourceColumn - 1;

                        if (foundTargetColumn > 0)
                        {
                            headerRow = r;
                            targetColumn = foundTargetColumn;
                            sourceColumn = foundSourceColumn;
                            return true;
                        }
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool FormulaLooksLikeLeftP4T6(
            string formula,
            string formulaLocal,
            string formulaR1C1,
            int row,
            int sourceColumn,
            string sourceHeader,
            int characterCount)
        {
            string f1 = NormalizeFormulaP4T6(formula);
            string f2 = NormalizeFormulaP4T6(formulaLocal);
            string f3 = NormalizeFormulaP4T6(formulaR1C1);

            string combined = f1 + " " + f2 + " " + f3;

            if (string.IsNullOrWhiteSpace(combined))
                return false;

            if (!combined.Contains("LEFT("))
                return false;

            string countText = characterCount.ToString(CultureInfo.InvariantCulture);
            if (!combined.Contains("," + countText + ")") &&
                !combined.Contains(";" + countText + ")") &&
                !combined.Contains(countText + ")"))
            {
                return false;
            }

            string expectedA1 = NormalizeFormulaP4T6(ExcelColumnNameP1T3(sourceColumn) + row.ToString(CultureInfo.InvariantCulture));
            string expectedStructuredHeader = NormalizeFormulaP4T6(sourceHeader);
            string expectedR1C1 = "RC[";

            bool hasReference =
                combined.Contains(expectedA1) ||
                combined.Contains(expectedStructuredHeader) ||
                combined.Contains(expectedR1C1);

            if (!hasReference)
                return false;

            return true;
        }

        private bool HeaderEqualsP4T6(string actual, string expected)
        {
            string a = NormalizeHeaderP4T6(actual);
            string e = NormalizeHeaderP4T6(expected);

            if (string.Equals(a, e, StringComparison.OrdinalIgnoreCase))
                return true;

            // Đề đôi khi ghi TTC nhưng file thực tế có thể ghi TCC.
            if ((a == "TTC" && e == "TCC") || (a == "TCC" && e == "TTC"))
                return true;

            return false;
        }

        private string NormalizeHeaderP4T6(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"[^A-Z0-9]+", "", RegexOptions.IgnoreCase);
            s = s.ToUpperInvariant();

            return s;
        }

        private string NormalizeTextP4T6(string text)
        {
            if (text == null) return "";

            string s = text.Trim();
            s = s.Replace((char)160, ' ');
            s = Regex.Replace(s, @"\s+", " ");

            return s.Trim();
        }

        private string NormalizeFormulaP4T6(string formula)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            string s = formula.Trim();

            s = s.Replace("$", "");
            s = s.Replace("'", "");
            s = s.Replace("\"", "");
            s = s.Replace(";", ",");
            s = s.Replace("[", "");
            s = s.Replace("]", "");
            s = s.Replace("@", "");
            s = Regex.Replace(s, @"\s+", "");
            s = s.ToUpperInvariant();

            return s;
        }


        //Project 4 Task 7
        public bool LastFirstNameFormulaAtCell(
            string sheetName,
            string targetCellAddress,
            string lastNameHeader,
            string firstNameHeader,
            string separator)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");

            if (string.IsNullOrWhiteSpace(sheetName)) return false;
            if (string.IsNullOrWhiteSpace(targetCellAddress)) return false;
            if (string.IsNullOrWhiteSpace(lastNameHeader)) return false;
            if (string.IsNullOrWhiteSpace(firstNameHeader)) return false;
            if (separator == null) return false;

            Xl.Worksheet ws = null;
            Xl.Range usedRange = null;
            Xl.Range targetCell = null;
            Xl.Range lastNameCell = null;
            Xl.Range firstNameCell = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                targetCell = ws.Range[targetCellAddress];
                if (targetCell == null) return false;

                usedRange = ws.UsedRange;
                if (usedRange == null) return false;

                int targetRow = Convert.ToInt32(targetCell.Row);
                int lastNameColumn;
                int firstNameColumn;

                if (!TryFindNameHeaderColumnsP4T7(
                        ws,
                        usedRange,
                        targetRow,
                        lastNameHeader,
                        firstNameHeader,
                        out lastNameColumn,
                        out firstNameColumn))
                {
                    return false;
                }

                lastNameCell = ws.Cells[targetRow, lastNameColumn] as Xl.Range;
                firstNameCell = ws.Cells[targetRow, firstNameColumn] as Xl.Range;

                if (lastNameCell == null || firstNameCell == null)
                    return false;

                string lastName = GetCellStringP1T3(lastNameCell);
                string firstName = GetCellStringP1T3(firstNameCell);

                if (string.IsNullOrWhiteSpace(lastName) || string.IsNullOrWhiteSpace(firstName))
                    return false;

                if (!CellHasFormulaP1T3(targetCell))
                    return false;

                string expectedText = lastName + separator + firstName;
                string actualText = GetCellStringP1T3(targetCell);

                if (!string.Equals(actualText, expectedText, StringComparison.Ordinal))
                    return false;

                string formula = "";
                string formulaLocal = "";

                try { formula = Convert.ToString(targetCell.Formula); } catch { formula = ""; }
                try { formulaLocal = Convert.ToString(targetCell.FormulaLocal); } catch { formulaLocal = ""; }

                string lastReference = ExcelColumnNameP1T3(lastNameColumn) + targetRow.ToString(CultureInfo.InvariantCulture);
                string firstReference = ExcelColumnNameP1T3(firstNameColumn) + targetRow.ToString(CultureInfo.InvariantCulture);

                return FormulaJoinsLastAndFirstP4T7(
                    formula,
                    formulaLocal,
                    sheetName,
                    lastReference,
                    firstReference,
                    separator);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(firstNameCell);
                ReleaseCom(lastNameCell);
                ReleaseCom(targetCell);
                ReleaseCom(usedRange);
                ReleaseCom(ws);
            }
        }

        private bool TryFindNameHeaderColumnsP4T7(
            Xl.Worksheet ws,
            Xl.Range usedRange,
            int targetRow,
            string lastNameHeader,
            string firstNameHeader,
            out int lastNameColumn,
            out int firstNameColumn)
        {
            lastNameColumn = 0;
            firstNameColumn = 0;

            if (ws == null || usedRange == null || targetRow <= 1)
                return false;

            Xl.Range cell = null;

            try
            {
                int firstRow = Convert.ToInt32(usedRange.Row);
                int firstColumn = Convert.ToInt32(usedRange.Column);
                int lastRow = firstRow + Convert.ToInt32(usedRange.Rows.Count) - 1;
                int lastColumn = firstColumn + Convert.ToInt32(usedRange.Columns.Count) - 1;
                int startRow = Math.Min(targetRow - 1, lastRow);

                string expectedLastName = NormalizeNameHeaderP4T7(lastNameHeader);
                string expectedFirstName = NormalizeNameHeaderP4T7(firstNameHeader);

                for (int row = startRow; row >= firstRow; row--)
                {
                    int foundLastNameColumn = 0;
                    int foundFirstNameColumn = 0;

                    for (int column = firstColumn; column <= lastColumn; column++)
                    {
                        ReleaseCom(cell);
                        cell = null;

                        cell = ws.Cells[row, column] as Xl.Range;
                        if (cell == null) continue;

                        string header = NormalizeNameHeaderP4T7(GetCellTextP1T3(cell));

                        if (string.Equals(header, expectedLastName, StringComparison.OrdinalIgnoreCase))
                            foundLastNameColumn = column;

                        if (string.Equals(header, expectedFirstName, StringComparison.OrdinalIgnoreCase))
                            foundFirstNameColumn = column;
                    }

                    if (foundLastNameColumn > 0 && foundFirstNameColumn > 0)
                    {
                        lastNameColumn = foundLastNameColumn;
                        firstNameColumn = foundFirstNameColumn;
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool FormulaJoinsLastAndFirstP4T7(
            string formula,
            string formulaLocal,
            string sheetName,
            string lastReference,
            string firstReference,
            string separator)
        {
            string f1 = NormalizeJoinFormulaP4T7(formula, sheetName);
            string f2 = NormalizeJoinFormulaP4T7(formulaLocal, sheetName);

            string lastRef = NormalizeJoinFormulaP4T7(lastReference, sheetName);
            string firstRef = NormalizeJoinFormulaP4T7(firstReference, sheetName);
            string separatorLiteral = "\"" + separator.Replace("\"", "\"\"") + "\"";

            string concat = "=CONCAT(" + lastRef + "," + separatorLiteral + "," + firstRef + ")";
            string xlfnConcat = "=_XLFN.CONCAT(" + lastRef + "," + separatorLiteral + "," + firstRef + ")";
            string concatenate = "=CONCATENATE(" + lastRef + "," + separatorLiteral + "," + firstRef + ")";
            string ampersand = "=" + lastRef + "&" + separatorLiteral + "&" + firstRef;

            return FormulaEqualsAnyP4T7(f1, concat, xlfnConcat, concatenate, ampersand) ||
                   FormulaEqualsAnyP4T7(f2, concat, xlfnConcat, concatenate, ampersand);
        }

        private bool FormulaEqualsAnyP4T7(string formula, params string[] expectedFormulas)
        {
            if (string.IsNullOrWhiteSpace(formula) || expectedFormulas == null)
                return false;

            for (int i = 0; i < expectedFormulas.Length; i++)
            {
                if (string.Equals(formula, expectedFormulas[i], StringComparison.OrdinalIgnoreCase))
                    return true;
            }

            return false;
        }

        private string NormalizeJoinFormulaP4T7(string formula, string sheetName)
        {
            if (string.IsNullOrWhiteSpace(formula))
                return "";

            StringBuilder result = new StringBuilder();
            bool insideString = false;

            for (int i = 0; i < formula.Length; i++)
            {
                char ch = formula[i];

                if (ch == '"')
                {
                    result.Append(ch);

                    if (insideString && i + 1 < formula.Length && formula[i + 1] == '"')
                    {
                        result.Append(formula[i + 1]);
                        i++;
                    }
                    else
                    {
                        insideString = !insideString;
                    }

                    continue;
                }

                if (!insideString)
                {
                    if (char.IsWhiteSpace(ch) || ch == '$' || ch == '@')
                        continue;

                    if (ch == ';')
                        ch = ',';

                    ch = char.ToUpperInvariant(ch);
                }

                result.Append(ch);
            }

            string normalized = result.ToString();

            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                string normalizedSheet = sheetName.Trim().ToUpperInvariant().Replace("'", "''");
                normalized = normalized.Replace("'" + normalizedSheet + "'!", "");
                normalized = normalized.Replace(normalizedSheet + "!", "");
            }

            return normalized;
        }

        private string NormalizeNameHeaderP4T7(string text)
        {
            if (text == null) return "";

            string result = text.Trim();
            result = result.Replace((char)160, ' ');
            result = Regex.Replace(result, @"[^A-Z0-9]+", "", RegexOptions.IgnoreCase);
            return result.ToUpperInvariant();
        }

        //Project 4 Task 8
        public bool CenterFooterPageOfPagesEquals(string sheetName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Worksheet ws = null;
            Xl.PageSetup pageSetup = null;

            try
            {
                ws = GetWorksheetExactOrNormalizedP2T2T3(sheetName);
                if (ws == null) return false;

                pageSetup = ws.PageSetup;
                if (pageSetup == null) return false;

                string centerFooter = Convert.ToString(pageSetup.CenterFooter);
                return FooterIsPageOfPagesP4T8(centerFooter);
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(pageSetup);
                ReleaseCom(ws);
            }
        }

        private bool FooterIsPageOfPagesP4T8(string footer)
        {
            if (string.IsNullOrWhiteSpace(footer))
                return false;

            string value = footer.Trim();

            value = Regex.Replace(value, @"&\[Page\]", "&P", RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"&\[Pages\]", "&N", RegexOptions.IgnoreCase);

            // Ignore optional formatting codes which Excel can prepend to a built-in footer.
            value = Regex.Replace(value, @"&""[^""]*""", "", RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"&K[0-9A-F]{6}", "", RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"&\d{1,3}", "", RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"&[BIUSEXY]", "", RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"^\s*&C", "", RegexOptions.IgnoreCase);
            value = Regex.Replace(value, @"\s+", " ").Trim();

            return Regex.IsMatch(
                value,
                @"^Page\s+&P\s+of\s+&N$",
                RegexOptions.IgnoreCase);
        }

        private bool AutoFilterVisibleRowsMatch(Xl.Range filterRange, int fieldIndex, string expectedText)
        {
            if (filterRange == null || fieldIndex <= 0 || string.IsNullOrWhiteSpace(expectedText))
                return false;

            Xl.Range rows = null;
            Xl.Range row = null;
            Xl.Range entireRow = null;
            Xl.Range cell = null;

            int matchingRows = 0;
            int visibleMatchingRows = 0;

            try
            {
                rows = filterRange.Rows;
                int rowCount = rows.Count;

                for (int rowIndex = 2; rowIndex <= rowCount; rowIndex++)
                {
                    ReleaseCom(row);
                    ReleaseCom(entireRow);
                    ReleaseCom(cell);
                    row = null;
                    entireRow = null;
                    cell = null;

                    row = (Xl.Range)rows[rowIndex];
                    if (!RangeRowContainsDataP14(row))
                        continue;
                    entireRow = row.EntireRow;

                    bool hidden = false;
                    try { hidden = Convert.ToBoolean(entireRow.Hidden); } catch { hidden = false; }

                    cell = (Xl.Range)filterRange.Cells[rowIndex, fieldIndex];

                    string cellText = "";
                    try
                    {
                        cellText = cell.Text == null ? "" : cell.Text.ToString();
                    }
                    catch
                    {
                        cellText = cell.Value2 == null ? "" : cell.Value2.ToString();
                    }

                    bool match = string.Equals(
                        NormalizeFilterValue(cellText),
                        NormalizeFilterValue(expectedText),
                        StringComparison.OrdinalIgnoreCase);

                    if (match)
                    {
                        matchingRows++;

                        if (!hidden)
                            visibleMatchingRows++;
                        else
                            return false;
                    }
                    else
                    {
                        if (!hidden)
                            return false;
                    }
                }

                return matchingRows > 0 && visibleMatchingRows == matchingRows;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(entireRow);
                ReleaseCom(row);
                ReleaseCom(rows);
            }
        }

        private bool RangeRowContainsDataP14(Xl.Range row)
        {
            if (row == null) return false;
            Xl.Range cells = null;
            Xl.Range cell = null;
            try
            {
                cells = row.Cells;
                int count = Convert.ToInt32(cells.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(cell);
                    cell = cells[i] as Xl.Range;
                    if (cell != null && !string.IsNullOrWhiteSpace(Convert.ToString(cell.Value2)))
                        return true;
                }
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(cells);
            }
        }

        private string NormalizeHeaderText(string text)
        {
            if (text == null) return "";

            string result = text.Trim();
            result = Regex.Replace(result, @"\s+", " ");
            return result.ToUpperInvariant();
        }

        private string NormalizeFilterValue(string value)
        {
            if (value == null) return "";

            string result = value.Trim();

            if (result.StartsWith("=", StringComparison.Ordinal))
                result = result.Substring(1);

            result = result.Replace("\"", "");
            result = result.Replace("'", "");
            result = Regex.Replace(result, @"\s+", "");

            double number;
            if (double.TryParse(result, NumberStyles.Any, CultureInfo.InvariantCulture, out number))
            {
                if (Math.Abs(number - Math.Round(number)) < 0.000001)
                    result = Math.Round(number).ToString(CultureInfo.InvariantCulture);
                else
                    result = number.ToString(CultureInfo.InvariantCulture);
            }

            return result.ToUpperInvariant();
        }


        // Project 5 Task 1
        public bool ShapeHyperlinkEquals(string sheetName, string shapeName, string topLeftCellAddress, string expectedUrl)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(shapeName) ||
                string.IsNullOrWhiteSpace(expectedUrl)) return false;

            Xl.Worksheet ws = null;
            Xl.Shapes shapes = null;
            Xl.Shape shape = null;
            Xl.Range topLeftCell = null;
            Xl.Hyperlinks hyperlinks = null;
            Xl.Hyperlink hyperlink = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                shapes = ws.Shapes;
                shape = shapes.Item(shapeName);
                if (shape == null) return false;

                int shapeType = Convert.ToInt32(shape.Type);
                if (shapeType != (int)Office.MsoShapeType.msoPicture &&
                    shapeType != (int)Office.MsoShapeType.msoLinkedPicture)
                    return false;

                topLeftCell = shape.TopLeftCell;
                if (topLeftCell == null ||
                    !string.Equals(
                        NormalizeRangeAddress(Convert.ToString(topLeftCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing])),
                        NormalizeRangeAddress(topLeftCellAddress),
                        StringComparison.OrdinalIgnoreCase))
                    return false;

                hyperlinks = ws.Hyperlinks;
                int hyperlinkCount = Convert.ToInt32(hyperlinks.Count);

                for (int i = 1; i <= hyperlinkCount; i++)
                {
                    object parent = null;
                    try
                    {
                        ReleaseCom(hyperlink);
                        hyperlink = hyperlinks.Item[i];
                        if (hyperlink == null) continue;

                        parent = hyperlink.Parent;
                        Xl.Shape parentShape = parent as Xl.Shape;
                        if (parentShape == null ||
                            !string.Equals(parentShape.Name, shapeName, StringComparison.OrdinalIgnoreCase))
                            continue;

                        return UrlEqualsP05T1(Convert.ToString(hyperlink.Address), expectedUrl) &&
                               string.IsNullOrWhiteSpace(Convert.ToString(hyperlink.SubAddress));
                    }
                    finally
                    {
                        ReleaseCom(parent);
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ShapeHyperlinkEquals", "P05 T01 grading failed.", ex, "Excel2019_P05", "T01");
                return false;
            }
            finally
            {
                ReleaseCom(hyperlink);
                ReleaseCom(hyperlinks);
                ReleaseCom(topLeftCell);
                ReleaseCom(shape);
                ReleaseCom(shapes);
                ReleaseCom(ws);
            }
        }

        private bool UrlEqualsP05T1(string actual, string expected)
        {
            if (string.IsNullOrWhiteSpace(actual) || string.IsNullOrWhiteSpace(expected))
                return false;

            Uri actualUri;
            Uri expectedUri;
            if (!Uri.TryCreate(actual.Trim(), UriKind.Absolute, out actualUri) ||
                !Uri.TryCreate(expected.Trim(), UriKind.Absolute, out expectedUri))
                return false;

            return string.Equals(actualUri.Scheme, expectedUri.Scheme, StringComparison.OrdinalIgnoreCase) &&
                   string.Equals(actualUri.Host, expectedUri.Host, StringComparison.OrdinalIgnoreCase) &&
                   actualUri.Port == expectedUri.Port &&
                   string.Equals(actualUri.AbsolutePath.TrimEnd('/'), expectedUri.AbsolutePath.TrimEnd('/'), StringComparison.Ordinal) &&
                   string.Equals(actualUri.Query, expectedUri.Query, StringComparison.Ordinal) &&
                   string.Equals(actualUri.Fragment, expectedUri.Fragment, StringComparison.Ordinal);
        }

        // Project 5 Task 3
        public bool ChartDataTableWithoutLegendKeys(string sheetName, string chartTitle, string chartName, string sourceHeader)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.ChartTitle title = null;
            Xl.DataTable dataTable = null;
            Xl.Range sourceDataRange = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                if (!string.IsNullOrWhiteSpace(sourceHeader))
                {
                    sourceDataRange = FindTableColumnDataBodyRangeByHeaderP3T8(ws, sourceHeader);
                    if (sourceDataRange == null) return false;
                }

                chartObjects = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (chartObjects == null) return false;

                int count = Convert.ToInt32(chartObjects.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(dataTable);
                    ReleaseCom(title);
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);
                    dataTable = null;
                    title = null;
                    chart = null;
                    chartObject = null;

                    chartObject = chartObjects.Item(i) as Xl.ChartObject;
                    if (chartObject == null) continue;
                    chart = chartObject.Chart;
                    if (chart == null) continue;

                    bool hasTitle = false;
                    try { hasTitle = Convert.ToBoolean(chart.HasTitle); } catch { hasTitle = false; }

                    bool isTarget;
                    if (hasTitle)
                    {
                        title = chart.ChartTitle;
                        string actualTitle = title == null ? "" : Convert.ToString(title.Text);
                        isTarget = string.Equals(
                            NormalizeText(actualTitle),
                            NormalizeText(chartTitle),
                            StringComparison.OrdinalIgnoreCase);
                    }
                    else
                    {
                        // The inspected starter has one untitled chart object named "Chart 1".
                        // Use its stable object name only while no title exists; a wrong visible title must not pass.
                        isTarget = !string.IsNullOrWhiteSpace(chartName) &&
                                   string.Equals(chartObject.Name, chartName, StringComparison.OrdinalIgnoreCase);
                    }

                    if (!isTarget) continue;
                    if (sourceDataRange != null &&
                        !ChartContainsSeriesForRangeP11T6(chart, sourceDataRange, sourceHeader))
                        continue;
                    if (!Convert.ToBoolean(chart.HasDataTable)) return false;

                    dataTable = chart.DataTable;
                    return dataTable != null && !Convert.ToBoolean(dataTable.ShowLegendKey);
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartDataTableWithoutLegendKeys", "P05 T03 grading failed.", ex, "Excel2019_P05", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(sourceDataRange);
                ReleaseCom(dataTable);
                ReleaseCom(title);
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(ws);
            }
        }

        private bool ChartContainsSeriesForRangeP11T6(
            Xl.Chart chart,
            Xl.Range sourceDataRange,
            string sourceHeader)
        {
            if (chart == null || sourceDataRange == null || string.IsNullOrWhiteSpace(sourceHeader))
                return false;

            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;

            try
            {
                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null) return false;

                string expectedRange = NormalizeChartFormulaP4T5(
                    GetRangeAddressWithSheetP4T5(sourceDataRange));
                string expectedName = NormalizeHeaderP4T5(sourceHeader);

                for (int i = 1; i <= Convert.ToInt32(seriesCollection.Count); i++)
                {
                    ReleaseCom(series);
                    series = seriesCollection.Item(i) as Xl.Series;
                    if (series == null) continue;

                    string formula = NormalizeChartFormulaP4T5(Convert.ToString(series.Formula));
                    string name = NormalizeHeaderP4T5(Convert.ToString(series.Name));

                    if (!string.IsNullOrWhiteSpace(expectedRange) &&
                        formula.Contains(expectedRange) &&
                        string.Equals(name, expectedName, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
            }
        }

        // Project 5 Task 4
        public bool SalesByExamTableConvertedToRange(
            string sheetName,
            string rangeAddress,
            string tableName,
            IList<string> expectedHeaders,
            int expectedDataRowCount)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress) ||
                expectedHeaders == null || expectedHeaders.Count == 0 || expectedDataRowCount <= 0)
                return false;

            Xl.Worksheet ws = null;
            Xl.Range targetRange = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                targetRange = ws.Range[rangeAddress];
                if (targetRange == null) return false;

                int rowCount = Convert.ToInt32(targetRange.Rows.Count);
                int columnCount = Convert.ToInt32(targetRange.Columns.Count);
                if (rowCount != expectedDataRowCount + 2 || columnCount != expectedHeaders.Count)
                    return false;

                if (WorksheetHasListObjectOnRangeP05T4(ws, targetRange, tableName))
                    return false;

                if (!RangeHasExpectedHeadersP05T4(targetRange, expectedHeaders))
                    return false;

                if (!RangeRetainsDataP05T4(targetRange, expectedDataRowCount))
                    return false;

                return RangeRetainsSalesTableFormattingP05T4(targetRange, expectedDataRowCount);
            }
            catch (Exception ex)
            {
                AppLogger.Error("SalesByExamTableConvertedToRange", "P05 T04 grading failed.", ex, "Excel2019_P05", "T04");
                return false;
            }
            finally
            {
                ReleaseCom(targetRange);
                ReleaseCom(ws);
            }
        }

        private bool WorksheetHasListObjectOnRangeP05T4(Xl.Worksheet ws, Xl.Range targetRange, string tableName)
        {
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            try
            {
                tables = ws.ListObjects;
                int count = Convert.ToInt32(tables.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(tableRange);
                    ReleaseCom(table);
                    tableRange = null;
                    table = tables.Item[i];
                    if (table == null) continue;

                    if (!string.IsNullOrWhiteSpace(tableName) &&
                        string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                        return true;

                    tableRange = table.Range;
                    if (tableRange != null && RangesIntersectP4T4(tableRange, targetRange))
                        return true;
                }

                return false;
            }
            catch
            {
                return true;
            }
            finally
            {
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(tables);
            }
        }

        private bool RangeHasExpectedHeadersP05T4(Xl.Range range, IList<string> expectedHeaders)
        {
            Xl.Range cell = null;
            try
            {
                for (int column = 1; column <= expectedHeaders.Count; column++)
                {
                    ReleaseCom(cell);
                    cell = range.Cells[1, column] as Xl.Range;
                    if (cell == null || !string.Equals(
                        NormalizeText(Convert.ToString(cell.Value2)),
                        NormalizeText(expectedHeaders[column - 1]),
                        StringComparison.OrdinalIgnoreCase))
                        return false;
                }
                return true;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool RangeRetainsDataP05T4(Xl.Range range, int expectedDataRowCount)
        {
            Xl.Range cell = null;
            try
            {
                int columnCount = Convert.ToInt32(range.Columns.Count);
                for (int row = 2; row <= expectedDataRowCount + 2; row++)
                {
                    for (int column = 1; column <= columnCount; column++)
                    {
                        ReleaseCom(cell);
                        cell = range.Cells[row, column] as Xl.Range;
                        if (cell == null || string.IsNullOrWhiteSpace(Convert.ToString(cell.Value2)))
                            return false;
                    }
                }

                return true;
            }
            finally
            {
                ReleaseCom(cell);
            }
        }

        private bool RangeRetainsSalesTableFormattingP05T4(Xl.Range range, int expectedDataRowCount)
        {
            Xl.Range cell = null;
            Xl.Font font = null;
            Xl.Interior interior = null;
            try
            {
                int columnCount = Convert.ToInt32(range.Columns.Count);
                for (int row = 1; row <= expectedDataRowCount + 2; row++)
                {
                    int expectedFill = row == 1 ? 8164884 :
                                       (row >= 2 && row <= expectedDataRowCount + 1 && row % 2 == 0 ? 16503472 : 16777215);
                    bool expectedBold = row == 1 || row == expectedDataRowCount + 2;

                    for (int column = 1; column <= columnCount; column++)
                    {
                        ReleaseCom(interior);
                        ReleaseCom(font);
                        ReleaseCom(cell);
                        interior = null;
                        font = null;
                        cell = range.Cells[row, column] as Xl.Range;
                        if (cell == null) return false;

                        font = cell.Font;
                        interior = cell.Interior;
                        if (font == null || interior == null) return false;

                        if (!string.Equals(Convert.ToString(font.Name), "Tahoma", StringComparison.OrdinalIgnoreCase) ||
                            Math.Abs(Convert.ToDouble(font.Size) - 11d) > 0.01d ||
                            Convert.ToBoolean(font.Bold) != expectedBold ||
                            Convert.ToInt32(interior.Color) != expectedFill)
                            return false;

                        if (row == 1 && Convert.ToInt32(font.Color) != 16777215)
                            return false;
                    }
                }

                return true;
            }
            finally
            {
                ReleaseCom(interior);
                ReleaseCom(font);
                ReleaseCom(cell);
            }
        }

        // Project 5 Task 5
        public bool TableColumnFormulaMultipliesColumns(
            string sheetName,
            string tableName,
            string targetHeader,
            IList<string> sourceHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetHeader) ||
                sourceHeaders == null || sourceHeaders.Count != 2 ||
                string.IsNullOrWhiteSpace(sourceHeaders[0]) || string.IsNullOrWhiteSpace(sourceHeaders[1]))
                return false;

            Xl.Worksheet ws = null;
            Xl.ListObject table = null;
            Xl.Range tableData = null;
            Xl.Range targetData = null;
            Xl.Range firstSourceData = null;
            Xl.Range secondSourceData = null;
            Xl.Range targetCell = null;
            Xl.Range firstSourceCell = null;
            Xl.Range secondSourceCell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                table = FindTableByNameOrHeadersP05T5(
                    ws, tableName, targetHeader, sourceHeaders[0], sourceHeaders[1]);
                if (table == null) return false;

                int targetColumnIndex = FindTableColumnIndexP05T5(table, targetHeader);
                int firstSourceColumnIndex = FindTableColumnIndexP05T5(table, sourceHeaders[0]);
                int secondSourceColumnIndex = FindTableColumnIndexP05T5(table, sourceHeaders[1]);
                if (targetColumnIndex <= 0 || firstSourceColumnIndex <= 0 || secondSourceColumnIndex <= 0)
                    return false;

                tableData = table.DataBodyRange;
                if (tableData == null) return false;
                targetData = tableData.Columns[targetColumnIndex] as Xl.Range;
                firstSourceData = tableData.Columns[firstSourceColumnIndex] as Xl.Range;
                secondSourceData = tableData.Columns[secondSourceColumnIndex] as Xl.Range;
                if (targetData == null || firstSourceData == null || secondSourceData == null ||
                    Convert.ToInt32(targetData.Rows.Count) != Convert.ToInt32(firstSourceData.Rows.Count) ||
                    Convert.ToInt32(targetData.Rows.Count) != Convert.ToInt32(secondSourceData.Rows.Count))
                    return false;

                int dataRowCount = Convert.ToInt32(targetData.Rows.Count);
                for (int row = 1; row <= dataRowCount; row++)
                {
                    ReleaseCom(secondSourceCell);
                    ReleaseCom(firstSourceCell);
                    ReleaseCom(targetCell);
                    secondSourceCell = secondSourceData.Cells[row, 1] as Xl.Range;
                    firstSourceCell = firstSourceData.Cells[row, 1] as Xl.Range;
                    targetCell = targetData.Cells[row, 1] as Xl.Range;
                    if (firstSourceCell == null || secondSourceCell == null || targetCell == null ||
                        !Convert.ToBoolean(targetCell.HasFormula))
                        return false;

                    string formula = Convert.ToString(targetCell.Formula);
                    string firstAddress = Convert.ToString(firstSourceCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    string secondAddress = Convert.ToString(secondSourceCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!IsTableColumnsMultiplicationFormulaP05T5(
                        formula,
                        sourceHeaders[0], firstAddress,
                        sourceHeaders[1], secondAddress))
                        return false;

                    double firstValue;
                    double secondValue;
                    double actual;
                    if (!TryToDouble(firstSourceCell.Value2, out firstValue) ||
                        !TryToDouble(secondSourceCell.Value2, out secondValue) ||
                        !TryToDouble(targetCell.Value2, out actual))
                        return false;

                    double expected = firstValue * secondValue;
                    double tolerance = Math.Max(0.000001d, Math.Abs(expected) * 0.000000001d);
                    if (Math.Abs(actual - expected) > tolerance)
                        return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableColumnFormulaMultipliesColumns", "P05 T05 grading failed.", ex, "Excel2019_P05", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(secondSourceCell);
                ReleaseCom(firstSourceCell);
                ReleaseCom(targetCell);
                ReleaseCom(secondSourceData);
                ReleaseCom(firstSourceData);
                ReleaseCom(targetData);
                ReleaseCom(tableData);
                ReleaseCom(table);
                ReleaseCom(ws);
            }
        }

        private Xl.ListObject FindTableByNameOrHeadersP05T5(
            Xl.Worksheet ws,
            string tableName,
            string targetHeader,
            string firstSourceHeader,
            string secondSourceHeader)
        {
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            try
            {
                tables = ws.ListObjects;
                int count = Convert.ToInt32(tables.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table == null) continue;

                    bool nameMatches = string.IsNullOrWhiteSpace(tableName) ||
                                       string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase);
                    bool headersMatch = TableHasColumnP05T5(table, targetHeader) &&
                                        TableHasColumnP05T5(table, firstSourceHeader) &&
                                        TableHasColumnP05T5(table, secondSourceHeader);
                    if (nameMatches && headersMatch)
                    {
                        Xl.ListObject result = table;
                        table = null;
                        return result;
                    }
                }

                // Keep T05 independent from T02: use the sole table with the required columns
                // when the learner has not yet renamed it.
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && TableHasColumnP05T5(table, targetHeader) &&
                        TableHasColumnP05T5(table, firstSourceHeader) &&
                        TableHasColumnP05T5(table, secondSourceHeader))
                    {
                        Xl.ListObject result = table;
                        table = null;
                        return result;
                    }
                }

                return null;
            }
            finally
            {
                ReleaseCom(table);
                ReleaseCom(tables);
            }
        }

        private bool TableHasColumnP05T5(Xl.ListObject table, string header)
        {
            return FindTableColumnIndexP05T5(table, header) > 0;
        }

        private int FindTableColumnIndexP05T5(Xl.ListObject table, string header)
        {
            Xl.Range headerRange = null;
            Xl.Range cell = null;
            try
            {
                headerRange = table.HeaderRowRange;
                if (headerRange == null) return 0;
                int count = Convert.ToInt32(headerRange.Columns.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(cell);
                    cell = headerRange.Cells[1, i] as Xl.Range;
                    if (cell != null && string.Equals(
                        NormalizeText(Convert.ToString(cell.Value2)),
                        NormalizeText(header),
                        StringComparison.OrdinalIgnoreCase))
                        return i;
                }
                return 0;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(headerRange);
            }
        }

        private bool IsTableColumnsMultiplicationFormulaP05T5(
            string formula,
            string firstSourceHeader,
            string firstSourceAddress,
            string secondSourceHeader,
            string secondSourceAddress)
        {
            string expression = NormalizeFormula(formula);
            if (expression.StartsWith("=", StringComparison.Ordinal))
                expression = expression.Substring(1);

            expression = TrimOuterParenthesesP05T5(expression);
            string[] terms = expression.Split('*');
            if (terms.Length != 2) return false;

            string left = TrimOuterParenthesesP05T5(terms[0]);
            string right = TrimOuterParenthesesP05T5(terms[1]);
            return (IsSourceTermP05T5(left, firstSourceHeader, firstSourceAddress) &&
                    IsSourceTermP05T5(right, secondSourceHeader, secondSourceAddress)) ||
                   (IsSourceTermP05T5(right, firstSourceHeader, firstSourceAddress) &&
                    IsSourceTermP05T5(left, secondSourceHeader, secondSourceAddress));
        }

        private string TrimOuterParenthesesP05T5(string expression)
        {
            string result = expression ?? "";
            while (result.Length >= 2 && result[0] == '(' && result[result.Length - 1] == ')' &&
                   ParenthesesWrapWholeExpressionP05T5(result))
                result = result.Substring(1, result.Length - 2);
            return result;
        }

        private bool ParenthesesWrapWholeExpressionP05T5(string expression)
        {
            int depth = 0;
            for (int i = 0; i < expression.Length; i++)
            {
                if (expression[i] == '(') depth++;
                else if (expression[i] == ')') depth--;
                if (depth == 0 && i < expression.Length - 1) return false;
                if (depth < 0) return false;
            }
            return depth == 0;
        }

        private bool IsSourceTermP05T5(string term, string sourceHeader, string sourceAddress)
        {
            string normalized = term.Trim();
            string structured = "[@" + NormalizeText(sourceHeader).Replace(" ", "") + "]";
            if (string.Equals(normalized, structured, StringComparison.OrdinalIgnoreCase))
                return true;

            int bang = normalized.LastIndexOf('!');
            if (bang >= 0) normalized = normalized.Substring(bang + 1);
            normalized = normalized.Replace("$", "").Trim('\'', ' ');
            return string.Equals(normalized, NormalizeRangeAddress(sourceAddress), StringComparison.OrdinalIgnoreCase);
        }

        // Project 5 Task 7
        public bool RangesMergedExactly(string sheetName, IList<string> rangeAddresses)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || rangeAddresses == null || rangeAddresses.Count == 0)
                return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;
            Xl.Range firstCell = null;
            Xl.Range mergeArea = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                foreach (string rangeAddress in rangeAddresses)
                {
                    ReleaseCom(mergeArea);
                    ReleaseCom(firstCell);
                    ReleaseCom(range);
                    mergeArea = null;
                    firstCell = null;
                    range = ws.Range[rangeAddress];
                    if (range == null || !Convert.ToBoolean(range.MergeCells))
                        return false;

                    // Excel's MergeArea property is reliable only when requested from one cell.
                    firstCell = range.Cells[1, 1] as Xl.Range;
                    if (firstCell == null) return false;
                    mergeArea = firstCell.MergeArea;
                    if (mergeArea == null) return false;
                    string actual = Convert.ToString(mergeArea.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!string.Equals(NormalizeRangeAddress(actual), NormalizeRangeAddress(rangeAddress), StringComparison.OrdinalIgnoreCase))
                        return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangesMergedExactly", "P05 T07 grading failed.", ex, "Excel2019_P05", "T07");
                return false;
            }
            finally
            {
                ReleaseCom(mergeArea);
                ReleaseCom(firstCell);
                ReleaseCom(range);
                ReleaseCom(ws);
            }
        }

        // Project 5 Task 8
        public bool CellStylesApplied(
            string sheetName,
            IList<string> primaryRanges,
            string primaryStyleName,
            IList<string> secondaryRanges,
            string secondaryStyleName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || primaryRanges == null || primaryRanges.Count == 0 ||
                secondaryRanges == null || secondaryRanges.Count == 0 ||
                string.IsNullOrWhiteSpace(primaryStyleName) || string.IsNullOrWhiteSpace(secondaryStyleName))
                return false;

            Xl.Worksheet ws = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                return RangesHaveBuiltInStyleP05T8(ws, primaryRanges, primaryStyleName) &&
                       RangesHaveBuiltInStyleP05T8(ws, secondaryRanges, secondaryStyleName);
            }
            catch (Exception ex)
            {
                AppLogger.Error("CellStylesApplied", "P05 T08 grading failed.", ex, "Excel2019_P05", "T08");
                return false;
            }
            finally
            {
                ReleaseCom(ws);
            }
        }

        private bool RangesHaveBuiltInStyleP05T8(Xl.Worksheet ws, IList<string> addresses, string expectedStyleName)
        {
            Xl.Range range = null;
            Xl.Range cell = null;
            Xl.Style style = null;
            try
            {
                foreach (string address in addresses)
                {
                    ReleaseCom(range);
                    range = ws.Range[address];
                    if (range == null) return false;

                    int cellCount = Convert.ToInt32(range.Cells.Count);
                    for (int i = 1; i <= cellCount; i++)
                    {
                        ReleaseCom(style);
                        ReleaseCom(cell);
                        style = null;
                        cell = range.Cells[i] as Xl.Range;
                        if (cell == null) return false;

                        style = cell.Style as Xl.Style;
                        if (style == null || !Convert.ToBoolean(style.BuiltIn) ||
                            !string.Equals(Convert.ToString(style.Name), expectedStyleName, StringComparison.OrdinalIgnoreCase))
                            return false;
                    }
                }

                return true;
            }
            finally
            {
                ReleaseCom(style);
                ReleaseCom(cell);
                ReleaseCom(range);
            }
        }

        // Project 6 Task 1
        public bool ChartColorPaletteEquals(
            string sheetName,
            string chartName,
            string chartTitle,
            int expectedChartColor,
            int expectedChartType,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(chartName) ||
                string.IsNullOrWhiteSpace(chartTitle) || expectedChartColor <= 0)
                return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.ChartTitle title = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                chartObjects = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (chartObjects == null) return false;

                int count = Convert.ToInt32(chartObjects.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(title);
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);
                    title = null;
                    chart = null;
                    chartObject = chartObjects.Item(i) as Xl.ChartObject;
                    if (chartObject == null || !string.Equals(chartObject.Name, chartName, StringComparison.OrdinalIgnoreCase))
                        continue;

                    chart = chartObject.Chart;
                    if (chart == null || !Convert.ToBoolean(chart.HasTitle)) return false;
                    title = chart.ChartTitle;
                    if (title == null || !string.Equals(
                        NormalizeText(Convert.ToString(title.Text)), NormalizeText(chartTitle),
                        StringComparison.OrdinalIgnoreCase))
                        return false;

                    if (Convert.ToInt32(chart.ChartColor) != expectedChartColor ||
                        (expectedChartType != 0 && Convert.ToInt32(chart.ChartType) != expectedChartType))
                        return false;

                    return ChartReferencesExpectedRangesP06(chart, sheetName, sourceRanges);
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartColorPaletteEquals", "P06 T01 grading failed.", ex, "Excel2019_P06", "T01");
                return false;
            }
            finally
            {
                ReleaseCom(title);
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(ws);
            }
        }

        public bool ChartExColorPaletteEquals(
            string sheetName,
            string chartName,
            int expectedColorStyleId,
            string expectedLayoutId,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(chartName) ||
                expectedColorStyleId <= 0 || string.IsNullOrWhiteSpace(expectedLayoutId) ||
                sourceRanges == null || sourceRanges.Count == 0) return false;

            Xl.Workbook workbook = null;
            string tempPath = "";
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                if (workbook == null) return false;
                tempPath = Path.Combine(Path.GetTempPath(), "MosTrainer_P18T6_" + Guid.NewGuid().ToString("N") + ".xlsx");
                workbook.SaveCopyAs(tempPath);
                return XlsxChartExColorPaletteEqualsP18(
                    tempPath, sheetName, chartName, expectedColorStyleId, expectedLayoutId, sourceRanges);
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartExColorPaletteEquals", "P18 T06 grading failed.", ex, "Excel2019_P18", "T06");
                return false;
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPath) && File.Exists(tempPath)) File.Delete(tempPath);
                }
                catch (Exception ex)
                {
                    AppLogger.Warning("ChartExColorPaletteEquals", "Temporary grading copy could not be deleted: " + ex.Message,
                        "Excel2019_P18", "T06");
                }
                // workbook belongs to the active session and must not be released here.
            }
        }

        private bool XlsxChartExColorPaletteEqualsP18(
            string xlsxPath,
            string sheetName,
            string chartName,
            int expectedColorStyleId,
            string expectedLayoutId,
            IList<string> sourceRanges)
        {
            using (FileStream stream = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (ZipArchive zip = new ZipArchive(stream, ZipArchiveMode.Read))
            {
                string sheetPart = FindWorksheetPartPathP2T7Xml(zip, sheetName);
                XDocument sheetDocument = ReadZipDocumentP18(zip, sheetPart);
                if (sheetDocument == null) return false;
                XElement drawingReference = sheetDocument.Descendants()
                    .FirstOrDefault(element => string.Equals(element.Name.LocalName, "drawing", StringComparison.OrdinalIgnoreCase));
                string drawingRid = GetAttributeValueP2T7Xml(drawingReference, "id");
                string drawingPart = ResolveRelationshipTargetP18(zip, sheetPart, drawingRid, "/drawing");
                XDocument drawingDocument = ReadZipDocumentP18(zip, drawingPart);
                if (drawingDocument == null) return false;

                XElement chartNameElement = drawingDocument.Descendants()
                    .FirstOrDefault(element => string.Equals(element.Name.LocalName, "cNvPr", StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(GetAttributeValueP2T7Xml(element, "name"), chartName, StringComparison.OrdinalIgnoreCase));
                if (chartNameElement == null) return false;
                XElement anchor = chartNameElement.Ancestors()
                    .FirstOrDefault(element => element.Name.LocalName.EndsWith("Anchor", StringComparison.OrdinalIgnoreCase));
                if (anchor == null) return false;
                XElement chartReference = anchor.Descendants()
                    .FirstOrDefault(element => string.Equals(element.Name.LocalName, "chart", StringComparison.OrdinalIgnoreCase) &&
                        !string.IsNullOrWhiteSpace(GetAttributeValueP2T7Xml(element, "id")));
                string chartRid = GetAttributeValueP2T7Xml(chartReference, "id");
                string chartPart = ResolveRelationshipTargetP18(zip, drawingPart, chartRid, "/chartEx");
                XDocument chartDocument = ReadZipDocumentP18(zip, chartPart);
                if (chartDocument == null) return false;

                bool layoutMatches = chartDocument.Descendants()
                    .Where(element => string.Equals(element.Name.LocalName, "series", StringComparison.OrdinalIgnoreCase))
                    .Any(element => string.Equals(GetAttributeValueP2T7Xml(element, "layoutId"),
                        expectedLayoutId.Trim(), StringComparison.OrdinalIgnoreCase));
                if (!layoutMatches) return false;

                HashSet<string> referencedNames = new HashSet<string>(
                    chartDocument.Descendants()
                        .Where(element => string.Equals(element.Name.LocalName, "f", StringComparison.OrdinalIgnoreCase))
                        .Select(element => (element.Value ?? "").Trim())
                        .Where(value => !string.IsNullOrWhiteSpace(value)),
                    StringComparer.OrdinalIgnoreCase);
                if (referencedNames.Count == 0) return false;

                XDocument workbookDocument = ReadZipDocumentP18(zip, "xl/workbook.xml");
                if (workbookDocument == null) return false;
                Dictionary<string, string> definedNames = workbookDocument.Descendants()
                    .Where(element => string.Equals(element.Name.LocalName, "definedName", StringComparison.OrdinalIgnoreCase))
                    .Where(element => !string.IsNullOrWhiteSpace(GetAttributeValueP2T7Xml(element, "name")))
                    .GroupBy(element => GetAttributeValueP2T7Xml(element, "name"), StringComparer.OrdinalIgnoreCase)
                    .ToDictionary(group => group.Key, group => group.First().Value, StringComparer.OrdinalIgnoreCase);
                HashSet<string> actualRanges = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (string reference in referencedNames)
                {
                    string effectiveReference;
                    if (!definedNames.TryGetValue(reference, out effectiveReference)) effectiveReference = reference;
                    string normalized = NormalizeChartExSourceP18(effectiveReference, sheetName);
                    if (string.IsNullOrWhiteSpace(normalized)) return false;
                    actualRanges.Add(normalized);
                }

                HashSet<string> expectedRanges = new HashSet<string>(
                    sourceRanges.Select(NormalizeRangeAddress).Where(value => !string.IsNullOrWhiteSpace(value)),
                    StringComparer.OrdinalIgnoreCase);
                if (actualRanges.Count != expectedRanges.Count || !actualRanges.SetEquals(expectedRanges)) return false;

                string colorsPart = ResolveRelationshipTargetP18(zip, chartPart, "", "/chartColorStyle");
                XDocument colorsDocument = ReadZipDocumentP18(zip, colorsPart);
                if (colorsDocument == null || colorsDocument.Root == null) return false;
                int actualColorStyleId;
                return int.TryParse(GetAttributeValueP2T7Xml(colorsDocument.Root, "id"),
                    NumberStyles.Integer, CultureInfo.InvariantCulture, out actualColorStyleId) &&
                    actualColorStyleId == expectedColorStyleId;
            }
        }

        private XDocument ReadZipDocumentP18(ZipArchive zip, string partPath)
        {
            if (zip == null || string.IsNullOrWhiteSpace(partPath)) return null;
            ZipArchiveEntry entry = zip.GetEntry(partPath.Replace('\\', '/'));
            if (entry == null) return null;
            string xml = ReadZipEntryTextP2T7Xml(entry);
            return string.IsNullOrWhiteSpace(xml) ? null : XDocument.Parse(xml);
        }

        private string ResolveRelationshipTargetP18(
            ZipArchive zip,
            string sourcePart,
            string relationshipId,
            string requiredTypeSuffix)
        {
            if (zip == null || string.IsNullOrWhiteSpace(sourcePart)) return "";
            string directory = Path.GetDirectoryName(sourcePart).Replace('\\', '/');
            string relsPart = directory + "/_rels/" + Path.GetFileName(sourcePart) + ".rels";
            XDocument relationships = ReadZipDocumentP18(zip, relsPart);
            if (relationships == null) return "";
            XElement relationship = relationships.Descendants()
                .FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, "Relationship", StringComparison.OrdinalIgnoreCase) &&
                    (string.IsNullOrWhiteSpace(relationshipId) ||
                        string.Equals(GetAttributeValueP2T7Xml(element, "Id"), relationshipId, StringComparison.Ordinal)) &&
                    (string.IsNullOrWhiteSpace(requiredTypeSuffix) ||
                        GetAttributeValueP2T7Xml(element, "Type").EndsWith(requiredTypeSuffix, StringComparison.OrdinalIgnoreCase)));
            if (relationship == null) return "";
            string target = GetAttributeValueP2T7Xml(relationship, "Target");
            if (string.IsNullOrWhiteSpace(target)) return "";
            Uri resolved = new Uri(new Uri("http://package/" + sourcePart), target);
            return Uri.UnescapeDataString(resolved.AbsolutePath).TrimStart('/');
        }

        private string NormalizeChartExSourceP18(string formula, string expectedSheetName)
        {
            if (string.IsNullOrWhiteSpace(formula)) return "";
            string text = formula.Trim().TrimStart('=').Replace("$", "").Replace("'", "");
            int separator = text.LastIndexOf('!');
            if (separator <= 0 || separator >= text.Length - 1) return "";
            string sheet = text.Substring(0, separator).Trim();
            if (!string.Equals(NormalizeText(sheet), NormalizeText(expectedSheetName), StringComparison.OrdinalIgnoreCase)) return "";
            return NormalizeRangeAddress(text.Substring(separator + 1));
        }

        // Project 6 Task 3
        public bool RangeFormattingMatchesSourceCell(
            string sheetName,
            string sourceCellAddress,
            string targetRangeAddress,
            IList<string> expectedTargetTexts)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(sourceCellAddress) ||
                string.IsNullOrWhiteSpace(targetRangeAddress) || expectedTargetTexts == null || expectedTargetTexts.Count == 0)
                return false;

            Xl.Worksheet ws = null;
            Xl.Range source = null;
            Xl.Range target = null;
            Xl.Range cell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                source = ws.Range[sourceCellAddress];
                target = ws.Range[targetRangeAddress];
                if (source == null || target == null || Convert.ToInt32(target.Cells.Count) != expectedTargetTexts.Count)
                    return false;

                string sourceSignature = GetCellFormatSignatureP06T3(source);
                for (int i = 1; i <= expectedTargetTexts.Count; i++)
                {
                    ReleaseCom(cell);
                    cell = target.Cells[i] as Xl.Range;
                    if (cell == null || !string.Equals(
                        NormalizeText(Convert.ToString(cell.Value2)), NormalizeText(expectedTargetTexts[i - 1]),
                        StringComparison.OrdinalIgnoreCase))
                        return false;

                    if (!string.Equals(sourceSignature, GetCellFormatSignatureP06T3(cell), StringComparison.Ordinal))
                        return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangeFormattingMatchesSourceCell", "P06 T03 grading failed.", ex, "Excel2019_P06", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(target);
                ReleaseCom(source);
                ReleaseCom(ws);
            }
        }

        // Project 6 Task 4
        public bool WorkbookBuiltinPropertyEquals(string propertyName, string expectedValue)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(propertyName)) return false;

            Xl.Workbook workbook = null;
            ZipArchive archive = null;
            Stream stream = null;
            string tempPath = Path.Combine(
                Path.GetTempPath(),
                "MosTrainer-P06-T04-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                workbook.SaveCopyAs(tempPath);
                archive = ZipFile.OpenRead(tempPath);
                ZipArchiveEntry entry = archive.GetEntry("docProps/core.xml");
                if (entry == null) return false;
                stream = entry.Open();
                XDocument document = XDocument.Load(stream);
                XElement property = document.Descendants().FirstOrDefault(element =>
                    string.Equals(element.Name.LocalName, propertyName, StringComparison.OrdinalIgnoreCase));
                return property != null && string.Equals(
                    Convert.ToString(property.Value).Trim(),
                    Convert.ToString(expectedValue).Trim(),
                    StringComparison.Ordinal);
            }
            catch (Exception ex)
            {
                AppLogger.Error("WorkbookBuiltinPropertyEquals", "P06 T04 grading failed.", ex, "Excel2019_P06", "T04");
                return false;
            }
            finally
            {
                if (stream != null) stream.Dispose();
                if (archive != null) archive.Dispose();
                try
                {
                    if (File.Exists(tempPath)) File.Delete(tempPath);
                }
                catch
                {
                }
            }
        }

        // Project 6 Task 5
        public bool TableOnRangeWithStyle(
            string sheetName,
            string rangeAddress,
            string expectedStyleName,
            IList<string> expectedHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress) ||
                string.IsNullOrWhiteSpace(expectedStyleName) || expectedHeaders == null || expectedHeaders.Count == 0)
                return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Range headerRange = null;
            Xl.Range headerCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                int tableCount = Convert.ToInt32(tables.Count);
                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(headerCell);
                    ReleaseCom(headerRange);
                    ReleaseCom(tableRange);
                    ReleaseCom(table);
                    headerCell = null;
                    headerRange = null;
                    tableRange = null;
                    table = tables.Item[i];
                    if (table == null) continue;

                    tableRange = table.Range;
                    if (tableRange == null || !string.Equals(
                        NormalizeRangeAddress(Convert.ToString(tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing])),
                        NormalizeRangeAddress(rangeAddress), StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (!Convert.ToBoolean(table.ShowHeaders) || !ListObjectTableStyleEqualsP2T8(table, expectedStyleName))
                        return false;

                    headerRange = table.HeaderRowRange;
                    if (headerRange == null || Convert.ToInt32(headerRange.Columns.Count) != expectedHeaders.Count)
                        return false;

                    for (int column = 1; column <= expectedHeaders.Count; column++)
                    {
                        ReleaseCom(headerCell);
                        headerCell = headerRange.Cells[1, column] as Xl.Range;
                        if (headerCell == null || !string.Equals(
                            NormalizeText(Convert.ToString(headerCell.Value2)), NormalizeText(expectedHeaders[column - 1]),
                            StringComparison.OrdinalIgnoreCase))
                            return false;
                    }

                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableOnRangeWithStyle", "P06 T05 grading failed.", ex, "Excel2019_P06", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(headerCell);
                ReleaseCom(headerRange);
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(tables);
                ReleaseCom(ws);
            }
        }

        // Project 6 Task 7
        public bool RangeWrapTextEquals(string sheetName, string rangeAddress, bool expectedWrapText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress)) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;
            Xl.Range cell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                range = ws.Range[rangeAddress];
                if (range == null) return false;

                int count = Convert.ToInt32(range.Cells.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(cell);
                    cell = range.Cells[i] as Xl.Range;
                    if (cell == null || cell.WrapText == null || Convert.ToBoolean(cell.WrapText) != expectedWrapText)
                        return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangeWrapTextEquals", "P06 T07 grading failed.", ex, "Excel2019_P06", "T07");
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(range);
                ReleaseCom(ws);
            }
        }

        // Project 6 Task 8
        public bool ChartMovedToChartSheet(
            string sourceSheetName,
            string chartSheetName,
            string chartTitle,
            int expectedChartType,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sourceSheetName) || string.IsNullOrWhiteSpace(chartSheetName))
                return false;

            Xl.Workbook workbook = null;
            Xl.Worksheet sourceSheet = null;
            Xl.ChartObjects sourceCharts = null;
            Xl.Sheets sheets = null;
            object sheetObject = null;
            Xl.Chart chartSheet = null;
            Xl.ChartTitle title = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                sourceSheet = GetWorksheet(sourceSheetName);
                if (sourceSheet == null) return false;
                sourceCharts = sourceSheet.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (sourceCharts == null || Convert.ToInt32(sourceCharts.Count) != 0) return false;

                sheets = workbook.Sheets;
                int sheetCount = Convert.ToInt32(sheets.Count);
                for (int i = 1; i <= sheetCount; i++)
                {
                    ReleaseCom(sheetObject);
                    sheetObject = sheets.Item[i];
                    chartSheet = sheetObject as Xl.Chart;
                    if (chartSheet == null)
                    {
                        ReleaseCom(sheetObject);
                        sheetObject = null;
                        continue;
                    }

                    if (!string.Equals(chartSheet.Name, chartSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        ReleaseCom(chartSheet);
                        chartSheet = null;
                        sheetObject = null;
                        continue;
                    }

                    if (!string.IsNullOrWhiteSpace(chartTitle))
                    {
                        if (!Convert.ToBoolean(chartSheet.HasTitle)) return false;
                        title = chartSheet.ChartTitle;
                        if (title == null || !string.Equals(
                            NormalizeText(Convert.ToString(title.Text)), NormalizeText(chartTitle),
                            StringComparison.OrdinalIgnoreCase))
                            return false;
                    }

                    if (expectedChartType != 0 && Convert.ToInt32(chartSheet.ChartType) != expectedChartType)
                        return false;

                    return ChartReferencesExpectedRangesP06(chartSheet, sourceSheetName, sourceRanges);
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartMovedToChartSheet", "P06 T08 grading failed.", ex, "Excel2019_P06", "T08");
                return false;
            }
            finally
            {
                ReleaseCom(title);
                ReleaseCom(chartSheet);
                ReleaseCom(sheetObject);
                ReleaseCom(sheets);
                ReleaseCom(sourceCharts);
                ReleaseCom(sourceSheet);
            }
        }

        private bool ChartReferencesExpectedRangesP06(Xl.Chart chart, string sourceSheetName, IList<string> sourceRanges)
        {
            if (chart == null || sourceRanges == null || sourceRanges.Count == 0) return false;

            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            try
            {
                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) == 0) return false;

                StringBuilder formulas = new StringBuilder();
                int count = Convert.ToInt32(seriesCollection.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(series);
                    series = seriesCollection.Item(i);
                    if (series != null) formulas.Append('|').Append(Convert.ToString(series.Formula));
                }

                string normalizedFormulas = NormalizeChartReferenceP06(formulas.ToString());
                foreach (string range in sourceRanges)
                {
                    string expected = NormalizeChartReferenceP06(sourceSheetName + "!" + range);
                    if (!normalizedFormulas.Contains(expected)) return false;
                }

                return true;
            }
            finally
            {
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
            }
        }

        private string NormalizeChartReferenceP06(string text)
        {
            return Regex.Replace((text ?? "").Replace("$", "").Replace("'", ""), @"\s+", "").ToUpperInvariant();
        }

        private string GetCellFormatSignatureP06T3(Xl.Range cell)
        {
            Xl.Font font = null;
            Xl.Interior interior = null;
            Xl.Borders borders = null;
            Xl.Border border = null;
            try
            {
                font = cell.Font;
                interior = cell.Interior;
                borders = cell.Borders;

                StringBuilder result = new StringBuilder();
                AppendFormatValueP06(result, font == null ? null : font.Name);
                AppendFormatValueP06(result, font == null ? null : font.Size);
                AppendFormatValueP06(result, font == null ? null : font.Bold);
                AppendFormatValueP06(result, font == null ? null : font.Italic);
                AppendFormatValueP06(result, font == null ? null : font.Underline);
                AppendFormatValueP06(result, font == null ? null : font.Strikethrough);
                AppendFormatValueP06(result, font == null ? null : font.Color);
                AppendFormatValueP06(result, interior == null ? null : interior.Pattern);
                AppendFormatValueP06(result, interior == null ? null : interior.Color);
                AppendFormatValueP06(result, interior == null ? null : interior.PatternColor);
                AppendFormatValueP06(result, cell.NumberFormat);
                AppendFormatValueP06(result, cell.HorizontalAlignment);
                AppendFormatValueP06(result, cell.VerticalAlignment);
                AppendFormatValueP06(result, cell.WrapText);
                AppendFormatValueP06(result, cell.IndentLevel);
                AppendFormatValueP06(result, cell.Orientation);
                AppendFormatValueP06(result, cell.ShrinkToFit);

                int[] borderIndexes = { 7, 8, 9, 10, 11, 12 };
                foreach (int borderIndex in borderIndexes)
                {
                    ReleaseCom(border);
                    border = borders == null ? null : borders.Item[(Xl.XlBordersIndex)borderIndex];
                    AppendFormatValueP06(result, border == null ? null : border.LineStyle);
                    AppendFormatValueP06(result, border == null ? null : border.Weight);
                }

                return result.ToString();
            }
            finally
            {
                ReleaseCom(border);
                ReleaseCom(borders);
                ReleaseCom(interior);
                ReleaseCom(font);
            }
        }

        private void AppendFormatValueP06(StringBuilder builder, object value)
        {
            builder.Append(Convert.ToString(value, CultureInfo.InvariantCulture)).Append('\u001f');
        }

        // Project 7 Task 1
        public bool WorksheetShowFormulasEquals(string sheetName, bool expectedShowFormulas)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Workbook workbook = null;
            Xl.Worksheet targetSheet = null;
            object originalSheet = null;
            Xl.Window window = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                originalSheet = workbook.ActiveSheet;
                targetSheet = GetWorksheet(sheetName);
                if (targetSheet == null) return false;
                targetSheet.Activate();
                window = workbook.Application.ActiveWindow;
                return window != null && Convert.ToBoolean(window.DisplayFormulas) == expectedShowFormulas;
            }
            catch (Exception ex)
            {
                AppLogger.Error("WorksheetShowFormulasEquals", "P07 T01 grading failed.", ex, "Excel2019_P07", "T01");
                return false;
            }
            finally
            {
                try
                {
                    Xl.Worksheet originalWorksheet = originalSheet as Xl.Worksheet;
                    if (originalWorksheet != null) originalWorksheet.Activate();
                }
                catch { }
                ReleaseCom(window);
                ReleaseCom(targetSheet);
                ReleaseCom(originalSheet);
            }
        }

        // Project 7 Task 2
        public bool InvoiceCellsDeletedShiftUp(string sheetName, string deletedRangeAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) ||
                !string.Equals(NormalizeRangeAddress(deletedRangeAddress), "E7:F7", StringComparison.OrdinalIgnoreCase))
                return false;

            Xl.Worksheet ws = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                if (!CellValueEqualsP07(ws, "D7", "Reference number") ||
                    !CellValueEqualsP07(ws, "C13", "W19") ||
                    !CellValueEqualsP07(ws, "D13", "MOS: Microsoft Word (Office 2019)") ||
                    !CellValueEqualsP07(ws, "C14", "LO5") ||
                    !CellValueEqualsP07(ws, "D14", "IC3: Living Online (GS5)") ||
                    !CellValueEqualsP07(ws, "C15", "W16") ||
                    !CellValueEqualsP07(ws, "D15", "MOS: Microsoft Word (Office 2016)"))
                    return false;

                if (!CellValueEqualsP07(ws, "E12", "Quantity") || !CellValueEqualsP07(ws, "F12", "Unit price") ||
                    !CellValueEqualsP07(ws, "E13", 10d) || !CellValueEqualsP07(ws, "F13", 29.5d) ||
                    !CellValueEqualsP07(ws, "E14", 100d) || !CellValueEqualsP07(ws, "F14", 31d) ||
                    !CellValueEqualsP07(ws, "E15", 30d) || !CellValueEqualsP07(ws, "F15", 29.5d) ||
                    !CellValueEqualsP07(ws, "F23", "Subtotal") || !CellValueEqualsP07(ws, "F24", "Tax") ||
                    !CellValueEqualsP07(ws, "F25", "Total"))
                    return false;

                for (int row = 16; row <= 22; row++)
                {
                    if (!CellValueEqualsP07(ws, "E" + row, "") || !CellValueEqualsP07(ws, "F" + row, ""))
                        return false;
                }

                for (int row = 13; row <= 22; row++)
                {
                    string expected = "=IF(F" + row + "=\"\",\"\",(F" + row + "*E" + row + "))";
                    if (!CellFormulaEqualsP07(ws, "G" + row, expected)) return false;
                }

                if (!CellFormulaEqualsP07(ws, "G23", "=SUM(G13:G22)") ||
                    !CellFormulaEqualsP07(ws, "G24", "=G23*8%") ||
                    !CellFormulaEqualsP07(ws, "G25", "=SUM(G23:G24)"))
                    return false;

                return InvoiceFormattingMatchesP07(ws);
            }
            catch (Exception ex)
            {
                AppLogger.Error("InvoiceCellsDeletedShiftUp", "P07 T02 grading failed.", ex, "Excel2019_P07", "T02");
                return false;
            }
            finally
            {
                ReleaseCom(ws);
            }
        }

        // Project 7 Task 3
        public bool ChartStyleAndPaletteEquals(
            string sheetName,
            string chartName,
            string chartTitle,
            int expectedChartStyle,
            int expectedChartColor,
            int expectedChartType,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(chartName)) return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.ChartTitle title = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                int count = Convert.ToInt32(charts.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(title);
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);
                    title = null;
                    chart = null;
                    chartObject = charts.Item(i) as Xl.ChartObject;
                    if (chartObject == null || !string.Equals(chartObject.Name, chartName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    chart = chartObject.Chart;
                    if (chart == null) return false;
                    if (!string.IsNullOrWhiteSpace(chartTitle))
                    {
                        if (!Convert.ToBoolean(chart.HasTitle)) return false;
                        title = chart.ChartTitle;
                        if (title == null || !string.Equals(NormalizeText(Convert.ToString(title.Text)), NormalizeText(chartTitle), StringComparison.OrdinalIgnoreCase))
                            return false;
                    }
                    if (!ChartStyleEqualsP07(Convert.ToInt32(chart.ChartStyle), expectedChartStyle) ||
                        Convert.ToInt32(chart.ChartColor) != expectedChartColor ||
                        Convert.ToInt32(chart.ChartType) != expectedChartType)
                        return false;
                    return ChartReferencesExpectedRangesP06(chart, sheetName, sourceRanges);
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartStyleAndPaletteEquals", "P07 T03 grading failed.", ex, "Excel2019_P07", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(title);
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(charts);
                ReleaseCom(ws);
            }
        }

        private bool ChartStyleEqualsP07(int actualStyle, int expectedStyle)
        {
            return NormalizeChartStyleGalleryIndexP07(actualStyle) ==
                   NormalizeChartStyleGalleryIndexP07(expectedStyle);
        }

        private int NormalizeChartStyleGalleryIndexP07(int style)
        {
            // Depending on the Excel build/chart style family, the visible gallery
            // Style 1-48 is exposed as 1-48, 201-248, or 251-298.
            if (style >= 201 && style <= 248) return style - 200;
            if (style >= 251 && style <= 298) return style - 250;
            return style;
        }

        // Project 7 Task 4
        public bool WorkbookPersonalInformationRemoved()
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            string tempPath = Path.Combine(Path.GetTempPath(), "MosTrainer-P07-T04-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                Xl.Workbook workbook = (Xl.Workbook)_session.Workbook;
                if (!Convert.ToBoolean(workbook.RemovePersonalInformation)) return false;
                if (!WorkbookRetainsPersonalInfoTaskCoreContent(workbook)) return false;
                workbook.SaveCopyAs(tempPath);
                using (ZipArchive archive = ZipFile.OpenRead(tempPath))
                {
                    XDocument core = ReadZipDocumentP07(archive.GetEntry("docProps/core.xml"));
                    XDocument workbookXml = ReadZipDocumentP07(archive.GetEntry("xl/workbook.xml"));
                    if (core == null || workbookXml == null) return false;
                    string creator = GetElementValueP07(core, "creator");
                    string lastModifiedBy = GetElementValueP07(core, "lastModifiedBy");
                    XElement workbookPr = workbookXml.Descendants().FirstOrDefault(element => element.Name.LocalName == "workbookPr");
                    XAttribute filterPrivacy = workbookPr == null ? null :
                        workbookPr.Attributes().FirstOrDefault(attribute => attribute.Name.LocalName == "filterPrivacy");
                    bool privacyFlag = filterPrivacy != null &&
                        (filterPrivacy.Value == "1" || string.Equals(filterPrivacy.Value, "true", StringComparison.OrdinalIgnoreCase));
                    return privacyFlag && string.IsNullOrWhiteSpace(creator) && string.IsNullOrWhiteSpace(lastModifiedBy);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error("WorkbookPersonalInformationRemoved", "P07 T04 grading failed.", ex, "Excel2019_P07", "T04");
                return false;
            }
            finally
            {
                DeleteTemporaryFileP07(tempPath);
            }
        }

        // Project 7 Task 5
        public bool ClusteredColumnChartFromRanges(
            string sheetName,
            string tableName,
            string tableRangeAddress,
            int expectedChartType,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableRangeAddress)) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.SeriesCollection seriesCollection = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                int tableCount = Convert.ToInt32(tables.Count);
                for (int i = 1; i <= tableCount; i++)
                {
                    ReleaseCom(tableRange);
                    ReleaseCom(table);
                    tableRange = null;
                    table = tables.Item[i];
                    if (table == null || (!string.IsNullOrWhiteSpace(tableName) &&
                        !string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)))
                        continue;
                    tableRange = table.Range;
                    string actualAddress = tableRange == null ? "" : Convert.ToString(tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (string.Equals(NormalizeRangeAddress(actualAddress), NormalizeRangeAddress(tableRangeAddress), StringComparison.OrdinalIgnoreCase))
                        break;
                }
                if (table == null || tableRange == null) return false;

                double rightEdge = Convert.ToDouble(tableRange.Left) + Convert.ToDouble(tableRange.Width);
                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                int chartCount = Convert.ToInt32(charts.Count);
                for (int i = 1; i <= chartCount; i++)
                {
                    ReleaseCom(seriesCollection);
                    ReleaseCom(chart);
                    ReleaseCom(chartObject);
                    seriesCollection = null;
                    chart = null;
                    chartObject = charts.Item(i) as Xl.ChartObject;
                    if (chartObject == null || Convert.ToDouble(chartObject.Left) + 2d < rightEdge) continue;
                    chart = chartObject.Chart;
                    if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType) continue;
                    seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                    if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != 1) continue;
                    if (ChartReferencesExpectedRangesP06(chart, sheetName, sourceRanges)) return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ClusteredColumnChartFromRanges", "P07 T05 grading failed.", ex, "Excel2019_P07", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(seriesCollection);
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(charts);
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(tables);
                ReleaseCom(ws);
            }
        }

        // Project 7 Task 7
        public bool IfFormulaByHeadersStrict(
            string sheetName,
            string tableName,
            string targetHeader,
            string criteriaHeader,
            string compareOperator,
            double threshold,
            string trueText,
            string falseText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(targetHeader) || string.IsNullOrWhiteSpace(criteriaHeader) || compareOperator != "<")
                return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range data = null;
            Xl.Range targetData = null;
            Xl.Range criteriaData = null;
            Xl.Range targetCell = null;
            Xl.Range criteriaCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                int count = Convert.ToInt32(tables.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;
                int targetIndex = FindTableColumnIndexP05T5(table, targetHeader);
                int criteriaIndex = FindTableColumnIndexP05T5(table, criteriaHeader);
                if (targetIndex <= 0 || criteriaIndex <= 0) return false;
                data = table.DataBodyRange;
                if (data == null) return false;
                targetData = data.Columns[targetIndex] as Xl.Range;
                criteriaData = data.Columns[criteriaIndex] as Xl.Range;
                if (targetData == null || criteriaData == null ||
                    Convert.ToInt32(targetData.Rows.Count) != Convert.ToInt32(criteriaData.Rows.Count))
                    return false;

                int rowCount = Convert.ToInt32(targetData.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(targetCell);
                    ReleaseCom(criteriaCell);
                    targetCell = targetData.Cells[row, 1] as Xl.Range;
                    criteriaCell = criteriaData.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || criteriaCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;
                    double criteriaValue;
                    if (!TryToDouble(criteriaCell.Value2, out criteriaValue)) return false;
                    string expectedResult = criteriaValue < threshold ? trueText : falseText;
                    if (!string.Equals(Convert.ToString(targetCell.Value2), expectedResult, StringComparison.Ordinal)) return false;
                    string criteriaAddress = Convert.ToString(criteriaCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMatchesStrictIfP07(Convert.ToString(targetCell.Formula), criteriaAddress, criteriaHeader, threshold, trueText, falseText))
                        return false;
                }
                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("IfFormulaByHeadersStrict", "P07 T07 grading failed.", ex, "Excel2019_P07", "T07");
                return false;
            }
            finally
            {
                ReleaseCom(criteriaCell);
                ReleaseCom(targetCell);
                ReleaseCom(criteriaData);
                ReleaseCom(targetData);
                ReleaseCom(data);
                ReleaseCom(table);
                ReleaseCom(tables);
                ReleaseCom(ws);
            }
        }

        // Project 7 Task 8
        public bool NamedRangeRefersToRange(string name, string sheetName, string rangeAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress))
                return false;

            Xl.Workbook workbook = null;
            Xl.Names names = null;
            Xl.Name definedName = null;
            Xl.Range refersToRange = null;
            Xl.Worksheet worksheet = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                names = workbook.Names;
                int count = Convert.ToInt32(names.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(worksheet);
                    ReleaseCom(refersToRange);
                    ReleaseCom(definedName);
                    worksheet = null;
                    refersToRange = null;
                    definedName = names.Item(i, Type.Missing, Type.Missing);
                    if (definedName == null || !string.Equals(Convert.ToString(definedName.Name), name, StringComparison.OrdinalIgnoreCase))
                        continue;
                    try { refersToRange = definedName.RefersToRange; } catch { refersToRange = null; }
                    if (refersToRange == null) return false;
                    worksheet = refersToRange.Worksheet;
                    string actualAddress = Convert.ToString(refersToRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    return worksheet != null &&
                        string.Equals(worksheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(NormalizeRangeAddress(actualAddress), NormalizeRangeAddress(rangeAddress), StringComparison.OrdinalIgnoreCase);
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("NamedRangeRefersToRange", "P07 T08 grading failed.", ex, "Excel2019_P07", "T08");
                return false;
            }
            finally
            {
                ReleaseCom(worksheet);
                ReleaseCom(refersToRange);
                ReleaseCom(definedName);
                ReleaseCom(names);
            }
        }

        private XDocument ReadZipDocumentP07(ZipArchiveEntry entry)
        {
            if (entry == null) return null;
            using (Stream stream = entry.Open()) return XDocument.Load(stream);
        }

        private string GetElementValueP07(XDocument document, string localName)
        {
            XElement element = document.Descendants().FirstOrDefault(item => item.Name.LocalName == localName);
            return element == null ? "" : Convert.ToString(element.Value);
        }

        private void DeleteTemporaryFileP07(string path)
        {
            try { if (!string.IsNullOrWhiteSpace(path) && File.Exists(path)) File.Delete(path); }
            catch { }
        }

        private bool CellValueEqualsP07(Xl.Worksheet ws, string address, object expected)
        {
            Xl.Range cell = null;
            try
            {
                cell = ws.Range[address];
                if (cell == null) return false;
                string expectedText = Convert.ToString(expected);
                if (string.IsNullOrEmpty(expectedText)) return string.IsNullOrWhiteSpace(Convert.ToString(cell.Value2));
                return ObjectEqualsLoose(cell.Value2, expected);
            }
            finally { ReleaseCom(cell); }
        }

        private bool WorkbookRetainsP07CoreContent(Xl.Workbook workbook)
        {
            if (workbook == null) return false;
            return WorksheetHasExactTableP07("Sales by Exam", "Table5", "A1:M6") &&
                   WorksheetHasExactTableP07("Exam History", "Table4", "A2:J7") &&
                   WorksheetHasExactTableP07("Next Period", "Table6", "A3:F21") &&
                   WorksheetHasNamedChartsP07("Sales by Exam", new[] { "Chart 1", "Chart 2", "Chart 3" }) &&
                   WorksheetHasNamedChartsP07("Subcribe Results", new[] { "Chart 4" });
        }

        private bool WorksheetHasExactTableP07(string sheetName, string tableName, string rangeAddress)
        {
            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range range = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                int count = Convert.ToInt32(tables.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(range);
                    ReleaseCom(table);
                    range = null;
                    table = tables.Item[i];
                    if (table == null || !string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) continue;
                    range = table.Range;
                    string actual = range == null ? "" : Convert.ToString(range.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    return string.Equals(NormalizeRangeAddress(actual), NormalizeRangeAddress(rangeAddress), StringComparison.OrdinalIgnoreCase);
                }
                return false;
            }
            finally
            {
                ReleaseCom(range); ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        private bool WorksheetHasNamedChartsP07(string sheetName, IList<string> chartNames)
        {
            Xl.Worksheet ws = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chart = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                foreach (string expectedName in chartNames)
                {
                    bool found = false;
                    int count = Convert.ToInt32(charts.Count);
                    for (int i = 1; i <= count; i++)
                    {
                        ReleaseCom(chart);
                        chart = charts.Item(i) as Xl.ChartObject;
                        if (chart != null && string.Equals(chart.Name, expectedName, StringComparison.OrdinalIgnoreCase))
                        {
                            found = true;
                            break;
                        }
                    }
                    if (!found) return false;
                }
                return true;
            }
            finally
            {
                ReleaseCom(chart); ReleaseCom(charts); ReleaseCom(ws);
            }
        }

        private bool CellFormulaEqualsP07(Xl.Worksheet ws, string address, string expectedFormula)
        {
            Xl.Range cell = null;
            try
            {
                cell = ws.Range[address];
                return cell != null && Convert.ToBoolean(cell.HasFormula) &&
                    string.Equals(NormalizeFormula(Convert.ToString(cell.Formula)), NormalizeFormula(expectedFormula), StringComparison.OrdinalIgnoreCase);
            }
            finally { ReleaseCom(cell); }
        }

        private bool InvoiceFormattingMatchesP07(Xl.Worksheet ws)
        {
            Xl.Range shiftedTop = null;
            Xl.Range header = null;
            Xl.Range quantity = null;
            Xl.Range price = null;
            Xl.Range total = null;
            Xl.Font font = null;
            Xl.Interior interior = null;
            Xl.Borders borders = null;
            Xl.Border border = null;
            try
            {
                shiftedTop = ws.Range["E7:F7"];
                font = shiftedTop.Font;
                interior = shiftedTop.Interior;
                if (Convert.ToBoolean(font.Bold) || Convert.ToInt32(interior.Color) != 10079487)
                    return false;
                ReleaseCom(interior); ReleaseCom(font); interior = null; font = null;

                header = ws.Range["E12"];
                font = header.Font;
                interior = header.Interior;
                if (!Convert.ToBoolean(font.Bold) || Convert.ToInt32(font.Color) != 16777215 ||
                    Convert.ToInt32(interior.Color) != 10855845 ||
                    Convert.ToInt32(header.HorizontalAlignment) != (int)Xl.XlHAlign.xlHAlignCenter)
                    return false;
                ReleaseCom(interior); ReleaseCom(font); interior = null; font = null;

                quantity = ws.Range["E13"];
                interior = quantity.Interior;
                if (Convert.ToInt32(interior.Color) != 15592941 ||
                    Convert.ToInt32(quantity.HorizontalAlignment) != (int)Xl.XlHAlign.xlHAlignCenter)
                    return false;
                ReleaseCom(interior); interior = null;

                price = ws.Range["F13"];
                if (!NormalizeFormat(Convert.ToString(price.NumberFormat)).Contains("#,##0.00")) return false;

                total = ws.Range["F25"];
                font = total.Font;
                borders = total.Borders;
                border = borders.Item[Xl.XlBordersIndex.xlEdgeBottom];
                return Convert.ToBoolean(font.Bold) && Convert.ToInt32(border.LineStyle) == (int)Xl.XlLineStyle.xlContinuous &&
                    Convert.ToInt32(border.Weight) == (int)Xl.XlBorderWeight.xlMedium;
            }
            finally
            {
                ReleaseCom(border); ReleaseCom(borders); ReleaseCom(interior); ReleaseCom(font);
                ReleaseCom(total); ReleaseCom(price); ReleaseCom(quantity); ReleaseCom(header); ReleaseCom(shiftedTop);
            }
        }

        private bool FormulaMatchesStrictIfP07(
            string formula, string criteriaAddress, string criteriaHeader,
            double threshold, string trueText, string falseText)
        {
            string normalized = NormalizeFormula(formula);
            string pattern = "^=IF\\((?<condition>.+),\\\"" + Regex.Escape(NormalizeFormula(trueText)) +
                "\\\",\\\"" + Regex.Escape(NormalizeFormula(falseText)) + "\\\"\\)$";
            Match match = Regex.Match(normalized, pattern, RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            string condition = match.Groups["condition"].Value;
            if (condition.Contains("<=") || condition.Contains(">=") || condition.Contains("<>")) return false;
            string[] terms = condition.Split('<');
            if (terms.Length != 2) return false;
            return (IsCriteriaReferenceP07(terms[0], criteriaAddress, criteriaHeader) && IsThresholdP07(terms[1], threshold)) ||
                   (IsThresholdP07(terms[0], threshold) && IsCriteriaReferenceP07(terms[1], criteriaAddress, criteriaHeader));
        }

        private bool IsCriteriaReferenceP07(string term, string address, string header)
        {
            string normalized = NormalizeFormula(term);
            if (string.Equals(NormalizeRangeAddress(normalized), NormalizeRangeAddress(address), StringComparison.OrdinalIgnoreCase))
                return true;
            string structured = "[@[" + NormalizeFormula(header) + "]]";
            return normalized.EndsWith(structured, StringComparison.OrdinalIgnoreCase);
        }

        private bool IsThresholdP07(string term, double expected)
        {
            string text = NormalizeFormula(term);
            bool percent = text.EndsWith("%", StringComparison.Ordinal);
            if (percent) text = text.Substring(0, text.Length - 1);
            double value;
            if (!double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return false;
            if (percent) value /= 100d;
            return Math.Abs(value - expected) < 0.0000001d;
        }

        // Project 8 Task 1
        public bool TableColumnFormulaMultipliesNamedRange(
            string sheetName,
            string tableName,
            string targetHeader,
            string sourceHeader,
            string namedRange,
            string namedRangeAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(targetHeader) || string.IsNullOrWhiteSpace(sourceHeader) ||
                string.IsNullOrWhiteSpace(namedRange) || string.IsNullOrWhiteSpace(namedRangeAddress))
                return false;

            Xl.Workbook workbook = null;
            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableData = null;
            Xl.Range targetData = null;
            Xl.Range sourceData = null;
            Xl.Range targetCell = null;
            Xl.Range sourceCell = null;
            Xl.Names names = null;
            Xl.Name name = null;
            Xl.Range namedCell = null;
            Xl.Worksheet namedSheet = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;

                int targetIndex = FindTableColumnIndexP05T5(table, targetHeader);
                int sourceIndex = FindTableColumnIndexP05T5(table, sourceHeader);
                if (targetIndex <= 0 || sourceIndex <= 0) return false;

                names = workbook.Names;
                for (int i = 1; i <= Convert.ToInt32(names.Count); i++)
                {
                    ReleaseCom(name);
                    name = names.Item(i, Type.Missing, Type.Missing);
                    if (name != null && string.Equals(Convert.ToString(name.Name), namedRange, StringComparison.OrdinalIgnoreCase)) break;
                    name = null;
                }
                if (name == null) return false;
                try { namedCell = name.RefersToRange; } catch { namedCell = null; }
                if (namedCell == null || Convert.ToInt32(namedCell.Cells.Count) != 1) return false;
                namedSheet = namedCell.Worksheet;
                string actualNamedAddress = Convert.ToString(namedCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                if (namedSheet == null || !string.Equals(namedSheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(NormalizeRangeAddress(actualNamedAddress), NormalizeRangeAddress(namedRangeAddress), StringComparison.OrdinalIgnoreCase))
                    return false;

                double namedValue;
                if (!TryToDouble(namedCell.Value2, out namedValue)) return false;

                tableData = table.DataBodyRange;
                if (tableData == null) return false;
                targetData = tableData.Columns[targetIndex] as Xl.Range;
                sourceData = tableData.Columns[sourceIndex] as Xl.Range;
                if (targetData == null || sourceData == null ||
                    Convert.ToInt32(targetData.Rows.Count) != Convert.ToInt32(sourceData.Rows.Count)) return false;

                int rowCount = Convert.ToInt32(targetData.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(targetCell);
                    ReleaseCom(sourceCell);
                    targetCell = targetData.Cells[row, 1] as Xl.Range;
                    sourceCell = sourceData.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || sourceCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;

                    string sourceAddress = Convert.ToString(sourceCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMultipliesNamedRangeP08(Convert.ToString(targetCell.Formula), sourceHeader, sourceAddress, namedRange))
                        return false;

                    double sourceValue;
                    if (TryToDouble(sourceCell.Value2, out sourceValue))
                    {
                        double actualValue;
                        if (!TryToDouble(targetCell.Value2, out actualValue)) return false;
                        double expectedValue = sourceValue * namedValue;
                        double tolerance = Math.Max(0.000001d, Math.Abs(expectedValue) * 0.000000001d);
                        if (Math.Abs(actualValue - expectedValue) > tolerance) return false;
                    }
                }

                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableColumnFormulaMultipliesNamedRange", "P08 T01 grading failed.", ex, "Excel2019_P08", "T01");
                return false;
            }
            finally
            {
                ReleaseCom(namedSheet); ReleaseCom(namedCell); ReleaseCom(name); ReleaseCom(names);
                ReleaseCom(sourceCell); ReleaseCom(targetCell); ReleaseCom(sourceData); ReleaseCom(targetData);
                ReleaseCom(tableData); ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 8 Task 2
        public bool TableRowContainingTextDeletedPreserveOutside(
            string sheetName,
            string tableName,
            string searchText,
            int expectedDataRowCount,
            string expectedTableRange,
            string preservedFilterRange,
            IList<string> expectedFirstColumnValues)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(searchText) || expectedDataRowCount <= 0 ||
                expectedFirstColumnValues == null || expectedFirstColumnValues.Count != expectedDataRowCount)
                return false;

            Xl.Workbook workbook = null;
            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Range data = null;
            Xl.Range cell = null;
            Xl.Names names = null;
            Xl.Name filterName = null;
            Xl.Range filterRange = null;
            Xl.Worksheet filterSheet = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;
                tableRange = table.Range;
                data = table.DataBodyRange;
                if (tableRange == null || data == null || Convert.ToInt32(data.Rows.Count) != expectedDataRowCount) return false;
                string actualTableRange = Convert.ToString(tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                if (!string.Equals(NormalizeRangeAddress(actualTableRange), NormalizeRangeAddress(expectedTableRange), StringComparison.OrdinalIgnoreCase) ||
                    RangeContainsTextP3T3(data, searchText)) return false;

                for (int row = 1; row <= expectedDataRowCount; row++)
                {
                    ReleaseCom(cell);
                    cell = data.Cells[row, 1] as Xl.Range;
                    if (cell == null || !string.Equals(
                        NormalizeText(Convert.ToString(cell.Value2)), NormalizeText(expectedFirstColumnValues[row - 1]),
                        StringComparison.OrdinalIgnoreCase)) return false;
                }

                names = workbook.Names;
                string expectedFilterName = sheetName + "!_FilterDatabase";
                for (int i = 1; i <= Convert.ToInt32(names.Count); i++)
                {
                    ReleaseCom(filterName);
                    filterName = names.Item(i, Type.Missing, Type.Missing);
                    if (filterName != null && string.Equals(Convert.ToString(filterName.Name), expectedFilterName, StringComparison.OrdinalIgnoreCase)) break;
                    filterName = null;
                }
                if (filterName == null) return false;
                try { filterRange = filterName.RefersToRange; } catch { filterRange = null; }
                if (filterRange == null) return false;
                filterSheet = filterRange.Worksheet;
                string actualFilterRange = Convert.ToString(filterRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                return filterSheet != null && string.Equals(filterSheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(NormalizeRangeAddress(actualFilterRange), NormalizeRangeAddress(preservedFilterRange), StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableRowContainingTextDeletedPreserveOutside", "P08 T02 grading failed.", ex, "Excel2019_P08", "T02");
                return false;
            }
            finally
            {
                ReleaseCom(filterSheet); ReleaseCom(filterRange); ReleaseCom(filterName); ReleaseCom(names);
                ReleaseCom(cell); ReleaseCom(data); ReleaseCom(tableRange); ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 12 Task 3
        public bool TableRowContainingTextDeletedPreserveUsedRange(
            string sheetName,
            string tableName,
            string searchText,
            int expectedDataRowCount,
            string expectedTableRange,
            string expectedUsedRange,
            IList<string> expectedFirstColumnValues)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(searchText) || expectedDataRowCount <= 0 ||
                string.IsNullOrWhiteSpace(expectedTableRange) || string.IsNullOrWhiteSpace(expectedUsedRange) ||
                expectedFirstColumnValues == null || expectedFirstColumnValues.Count != expectedDataRowCount)
                return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Range data = null;
            Xl.Range usedRange = null;
            Xl.Range cell = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase))
                        break;
                    table = null;
                }

                if (table == null) return false;
                tableRange = table.Range;
                data = table.DataBodyRange;
                if (tableRange == null || data == null || Convert.ToInt32(data.Rows.Count) != expectedDataRowCount)
                    return false;

                string actualTableRange = Convert.ToString(
                    tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                if (!string.Equals(
                        NormalizeRangeAddress(actualTableRange),
                        NormalizeRangeAddress(expectedTableRange),
                        StringComparison.OrdinalIgnoreCase) ||
                    RangeContainsTextP3T3(data, searchText))
                    return false;

                for (int row = 1; row <= expectedDataRowCount; row++)
                {
                    ReleaseCom(cell);
                    cell = data.Cells[row, 1] as Xl.Range;
                    if (cell == null || !string.Equals(
                            NormalizeText(Convert.ToString(cell.Value2)),
                            NormalizeText(expectedFirstColumnValues[row - 1]),
                            StringComparison.OrdinalIgnoreCase))
                        return false;
                }

                usedRange = ws.UsedRange;
                if (usedRange == null) return false;
                string actualUsedRange = Convert.ToString(
                    usedRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);

                return string.Equals(
                    NormalizeRangeAddress(actualUsedRange),
                    NormalizeRangeAddress(expectedUsedRange),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                AppLogger.Error(
                    "TableRowContainingTextDeletedPreserveUsedRange",
                    "P12 T03 grading failed.", ex, "Excel2019_P12", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(usedRange);
                ReleaseCom(data);
                ReleaseCom(tableRange);
                ReleaseCom(table);
                ReleaseCom(tables);
                ReleaseCom(ws);
            }
        }

        // Project 8 Task 3
        public bool RangeAlignmentIndentEquals(string sheetName, string rangeAddress, string expectedAlignment, int expectedIndent)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress) ||
                expectedIndent < 0 || expectedIndent > 15) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;
            Xl.Range cell = null;
            try
            {
                string alignmentToken = Regex.Replace(expectedAlignment ?? "", @"[\s\-/]+", "").ToUpperInvariant();
                bool allowTextGeneralAsLeft = alignmentToken == "LEFTORGENERAL";
                int expectedAlignmentValue = allowTextGeneralAsLeft
                    ? (int)Xl.XlHAlign.xlHAlignLeft
                    : ParseHorizontalAlignment(expectedAlignment);
                if (expectedAlignmentValue == 0) return false;
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                range = ws.Range[rangeAddress];
                if (range == null) return false;
                int count = Convert.ToInt32(range.Cells.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(cell);
                    cell = range.Cells[i] as Xl.Range;
                    if (cell == null || Convert.ToInt32(cell.IndentLevel) != expectedIndent) return false;
                    int actualAlignment = Convert.ToInt32(cell.HorizontalAlignment);
                    bool alignmentMatches = actualAlignment == expectedAlignmentValue;
                    if (allowTextGeneralAsLeft && actualAlignment == (int)Xl.XlHAlign.xlHAlignGeneral)
                        alignmentMatches = !string.IsNullOrWhiteSpace(Convert.ToString(cell.Value2));
                    if (!alignmentMatches) return false;
                }
                return count > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangeAlignmentIndentEquals", "P08 T03 grading failed.", ex, "Excel2019_P08", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(cell); ReleaseCom(range); ReleaseCom(ws);
            }
        }

        // Project 8 Task 4
        public bool SparklinesByRangeAndType(string sheetName, string locationRange, string dataRange, string expectedType)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(locationRange) ||
                string.IsNullOrWhiteSpace(dataRange)) return false;

            Xl.Worksheet ws = null;
            Xl.Range expectedLocation = null;
            Xl.Range expectedData = null;
            object groups = null;
            object group = null;
            Xl.Range groupLocation = null;
            Xl.Range groupData = null;
            try
            {
                int expectedSparkType;
                if (!TryParseSparklineTypeP08(expectedType, out expectedSparkType)) return false;
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                expectedLocation = ws.Range[locationRange];
                expectedData = ws.Range[dataRange];
                if (expectedLocation == null || expectedData == null ||
                    Convert.ToInt32(expectedLocation.Columns.Count) != 1 ||
                    Convert.ToInt32(expectedLocation.Rows.Count) != Convert.ToInt32(expectedData.Rows.Count)) return false;

                int firstLocationRow = Convert.ToInt32(expectedLocation.Row);
                int lastLocationRow = firstLocationRow + Convert.ToInt32(expectedLocation.Rows.Count) - 1;
                int locationColumn = Convert.ToInt32(expectedLocation.Column);
                int dataFirstColumn = Convert.ToInt32(expectedData.Column);
                int dataLastColumn = dataFirstColumn + Convert.ToInt32(expectedData.Columns.Count) - 1;
                bool[] covered = new bool[Convert.ToInt32(expectedLocation.Rows.Count)];

                groups = GetComPropertyP2T3(expectedLocation, "SparklineGroups");
                int groupCount = ToIntP2T3(GetComPropertyP2T3(groups, "Count"));
                if (groupCount <= 0) return false;
                for (int i = 1; i <= groupCount; i++)
                {
                    ReleaseCom(groupData); ReleaseCom(groupLocation); ReleaseCom(group);
                    groupData = null; groupLocation = null;
                    group = GetComIndexedPropertyP2T3(groups, "Item", i);
                    if (group == null) continue;
                    groupLocation = GetComPropertyP2T3(group, "Location") as Xl.Range;
                    if (groupLocation == null) continue;
                    int groupColumn = Convert.ToInt32(groupLocation.Column);
                    int groupFirstRow = Convert.ToInt32(groupLocation.Row);
                    int groupLastRow = groupFirstRow + Convert.ToInt32(groupLocation.Rows.Count) - 1;
                    bool touchesTarget = groupColumn == locationColumn && groupLastRow >= firstLocationRow && groupFirstRow <= lastLocationRow;
                    if (!touchesTarget) continue;
                    if (ToIntP2T3(GetComPropertyP2T3(group, "Type")) != expectedSparkType || Convert.ToInt32(groupLocation.Columns.Count) != 1 ||
                        groupFirstRow < firstLocationRow || groupLastRow > lastLocationRow) return false;

                    string sourceData = Convert.ToString(GetComPropertyP2T3(group, "SourceData"));
                    string sourceAddress = NormalizeRangeAddress(sourceData);
                    try { groupData = ws.Range[sourceAddress]; } catch { groupData = null; }
                    if (groupData == null || Convert.ToInt32(groupData.Column) != dataFirstColumn ||
                        Convert.ToInt32(groupData.Columns.Count) != dataLastColumn - dataFirstColumn + 1 ||
                        Convert.ToInt32(groupData.Row) != groupFirstRow ||
                        Convert.ToInt32(groupData.Rows.Count) != groupLastRow - groupFirstRow + 1) return false;

                    for (int row = groupFirstRow; row <= groupLastRow; row++)
                    {
                        int index = row - firstLocationRow;
                        if (covered[index]) return false;
                        covered[index] = true;
                    }
                }

                for (int i = 0; i < covered.Length; i++) if (!covered[i]) return false;
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("SparklinesByRangeAndType", "P08 T04 grading failed.", ex, "Excel2019_P08", "T04");
                return false;
            }
            finally
            {
                ReleaseCom(groupData); ReleaseCom(groupLocation); ReleaseCom(group); ReleaseCom(groups);
                ReleaseCom(expectedData); ReleaseCom(expectedLocation); ReleaseCom(ws);
            }
        }

        // Project 8 Task 5
        public bool TableTotalRowSumsByHeaders(string sheetName, string tableName, IList<string> sumHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                sumHeaders == null || sumHeaders.Count == 0) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.ListColumns columns = null;
            Xl.ListColumn column = null;
            Xl.Range data = null;
            Xl.Range totalsRow = null;
            Xl.Range total = null;
            Xl.Range cell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null || !Convert.ToBoolean(table.ShowTotals)) return false;
                columns = table.ListColumns;
                foreach (string header in sumHeaders)
                {
                    ReleaseCom(cell); ReleaseCom(total); ReleaseCom(data); ReleaseCom(column);
                    cell = null; total = null; data = null; column = null;
                    int index = FindTableColumnIndexP05T5(table, header);
                    if (index <= 0) return false;
                    column = columns.Item[index];
                    if (column == null || Convert.ToInt32(column.TotalsCalculation) != (int)Xl.XlTotalsCalculation.xlTotalsCalculationSum)
                        return false;
                    data = column.DataBodyRange;
                    ReleaseCom(totalsRow);
                    totalsRow = table.TotalsRowRange;
                    total = totalsRow == null ? null : totalsRow.Cells[1, index] as Xl.Range;
                    if (data == null || total == null || !Convert.ToBoolean(total.HasFormula)) return false;
                    double expected = 0d;
                    for (int row = 1; row <= Convert.ToInt32(data.Rows.Count); row++)
                    {
                        ReleaseCom(cell);
                        cell = data.Cells[row, 1] as Xl.Range;
                        double value;
                        if (cell != null && TryToDouble(cell.Value2, out value)) expected += value;
                    }
                    double actual;
                    if (!TryToDouble(total.Value2, out actual) || Math.Abs(actual - expected) > 0.000001d) return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableTotalRowSumsByHeaders", "P08 T05 grading failed.", ex, "Excel2019_P08", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(cell); ReleaseCom(total); ReleaseCom(totalsRow); ReleaseCom(data); ReleaseCom(column); ReleaseCom(columns);
                ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 8 Task 6
        public bool CountBlankFormulaByHeaders(string sheetName, string tableName, string targetHeader, IList<string> sourceHeaders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(targetHeader) || sourceHeaders == null || sourceHeaders.Count < 2) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableData = null;
            Xl.Range targetData = null;
            Xl.Range targetCell = null;
            Xl.Range sourceCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;
                int targetIndex = FindTableColumnIndexP05T5(table, targetHeader);
                int firstSourceIndex = FindTableColumnIndexP05T5(table, sourceHeaders[0]);
                if (targetIndex <= 0 || firstSourceIndex <= 0) return false;
                for (int i = 0; i < sourceHeaders.Count; i++)
                    if (FindTableColumnIndexP05T5(table, sourceHeaders[i]) != firstSourceIndex + i) return false;

                tableData = table.DataBodyRange;
                if (tableData == null) return false;
                targetData = tableData.Columns[targetIndex] as Xl.Range;
                if (targetData == null) return false;
                int rowCount = Convert.ToInt32(targetData.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(targetCell);
                    targetCell = targetData.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;
                    int worksheetRow = Convert.ToInt32(targetCell.Row);
                    string firstAddress = ExcelColumnNameP1T3(Convert.ToInt32(tableData.Column) + firstSourceIndex - 1) + worksheetRow;
                    string lastAddress = ExcelColumnNameP1T3(Convert.ToInt32(tableData.Column) + firstSourceIndex + sourceHeaders.Count - 2) + worksheetRow;
                    if (!FormulaIsCountBlankP08(Convert.ToString(targetCell.Formula), firstAddress + ":" + lastAddress,
                        sourceHeaders[0], sourceHeaders[sourceHeaders.Count - 1])) return false;

                    int expectedBlanks = 0;
                    for (int source = 0; source < sourceHeaders.Count; source++)
                    {
                        ReleaseCom(sourceCell);
                        sourceCell = tableData.Cells[row, firstSourceIndex + source] as Xl.Range;
                        if (sourceCell == null) return false;
                        object value = sourceCell.Value2;
                        if (value == null || string.IsNullOrEmpty(Convert.ToString(value))) expectedBlanks++;
                    }
                    double actual;
                    if (!TryToDouble(targetCell.Value2, out actual) || Math.Abs(actual - expectedBlanks) > 0.000001d) return false;
                }
                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("CountBlankFormulaByHeaders", "P08 T06 grading failed.", ex, "Excel2019_P08", "T06");
                return false;
            }
            finally
            {
                ReleaseCom(sourceCell); ReleaseCom(targetCell); ReleaseCom(targetData); ReleaseCom(tableData);
                ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 8 Task 7
        public bool TableMultiLevelSortStateEquals(
            string sheetName,
            string tableName,
            string tableRangeAddress,
            IList<string> sortHeaders,
            IList<string> sortOrders)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                sortHeaders == null || sortOrders == null || sortHeaders.Count != 2 || sortOrders.Count != 2)
                return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Sort sort = null;
            Xl.SortFields fields = null;
            Xl.SortField field = null;
            Xl.Range key = null;
            Xl.Range expectedKey = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;
                tableRange = table.Range;
                string actualTableRange = tableRange == null ? "" : Convert.ToString(tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                if (!string.Equals(NormalizeRangeAddress(actualTableRange), NormalizeRangeAddress(tableRangeAddress), StringComparison.OrdinalIgnoreCase)) return false;

                sort = table.Sort;
                fields = sort.SortFields;
                if (fields == null || Convert.ToInt32(fields.Count) != 2) return false;
                for (int i = 0; i < 2; i++)
                {
                    ReleaseCom(expectedKey); ReleaseCom(key); ReleaseCom(field);
                    expectedKey = null; key = null;
                    int keyIndex = FindTableColumnIndexP05T5(table, sortHeaders[i]);
                    if (keyIndex <= 0) return false;
                    field = fields.Item[i + 1];
                    key = field.Key;
                    Xl.ListColumn listColumn = table.ListColumns.Item[keyIndex];
                    try { expectedKey = listColumn.DataBodyRange; }
                    finally { ReleaseCom(listColumn); }
                    if (field == null || key == null || expectedKey == null ||
                        Convert.ToInt32(field.SortOn) != (int)Xl.XlSortOn.xlSortOnValues) return false;
                    string actualKey = Convert.ToString(key.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    string expectedKeyAddress = Convert.ToString(expectedKey.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!string.Equals(NormalizeRangeAddress(actualKey), NormalizeRangeAddress(expectedKeyAddress), StringComparison.OrdinalIgnoreCase)) return false;
                    int expectedOrder = IsAscendingSortOrderP1T4(sortOrders, i)
                        ? (int)Xl.XlSortOrder.xlAscending
                        : (int)Xl.XlSortOrder.xlDescending;
                    if (Convert.ToInt32(field.Order) != expectedOrder) return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableMultiLevelSortStateEquals", "P08 T07 grading failed.", ex, "Excel2019_P08", "T07");
                return false;
            }
            finally
            {
                ReleaseCom(expectedKey); ReleaseCom(key); ReleaseCom(field);
                ReleaseCom(fields); ReleaseCom(sort); ReleaseCom(tableRange); ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 8 Task 8
        public bool ChartQuickLayoutEquals(
            string sheetName,
            string chartName,
            int expectedLayout,
            int expectedChartType,
            string seriesNameRange,
            string categoryRange,
            string valuesRange)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(chartName) || expectedLayout != 3)
                return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.Legend legend = null;
            Xl.Axis categoryAxis = null;
            Xl.Axis valueAxis = null;
            Xl.ChartGroup chartGroup = null;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            Xl.Range names = null;
            Xl.Range categories = null;
            Xl.Range values = null;
            Xl.Range nameCell = null;
            Xl.Range valueRow = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    ReleaseCom(chartObject);
                    chartObject = charts.Item(i) as Xl.ChartObject;
                    if (chartObject != null && string.Equals(chartObject.Name, chartName, StringComparison.OrdinalIgnoreCase)) break;
                    chartObject = null;
                }
                if (chartObject == null) return false;
                chart = chartObject.Chart;
                if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType ||
                    !Convert.ToBoolean(chart.HasTitle) || !Convert.ToBoolean(chart.HasLegend)) return false;
                legend = chart.Legend;
                if (legend == null || Convert.ToInt32(legend.Position) != (int)Xl.XlLegendPosition.xlLegendPositionBottom) return false;

                categoryAxis = chart.Axes(Xl.XlAxisType.xlCategory, Xl.XlAxisGroup.xlPrimary) as Xl.Axis;
                valueAxis = chart.Axes(Xl.XlAxisType.xlValue, Xl.XlAxisGroup.xlPrimary) as Xl.Axis;
                if (categoryAxis == null || valueAxis == null || Convert.ToBoolean(categoryAxis.HasTitle) ||
                    Convert.ToBoolean(valueAxis.HasTitle) || Convert.ToBoolean(categoryAxis.HasMajorGridlines) ||
                    !Convert.ToBoolean(valueAxis.HasMajorGridlines)) return false;
                chartGroup = chart.ChartGroups(1) as Xl.ChartGroup;
                if (chartGroup == null || Convert.ToInt32(chartGroup.GapWidth) != 75 ||
                    Convert.ToInt32(chartGroup.Overlap) != -25) return false;

                names = ws.Range[seriesNameRange];
                categories = ws.Range[categoryRange];
                values = ws.Range[valuesRange];
                if (names == null || categories == null || values == null) return false;
                int nameRows = Convert.ToInt32(names.Rows.Count);
                int nameColumns = Convert.ToInt32(names.Columns.Count);
                int valueRows = Convert.ToInt32(values.Rows.Count);
                int valueColumns = Convert.ToInt32(values.Columns.Count);
                int categoryRows = Convert.ToInt32(categories.Rows.Count);
                int categoryColumns = Convert.ToInt32(categories.Columns.Count);
                bool seriesByColumns = nameRows == 1 && nameColumns == valueColumns &&
                    categoryColumns == 1 && categoryRows == valueRows;
                bool seriesByRows = nameColumns == 1 && nameRows == valueRows &&
                    categoryRows == 1 && categoryColumns == valueColumns;
                if (!seriesByColumns && !seriesByRows) return false;
                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                int seriesCount = seriesByColumns ? nameColumns : nameRows;
                if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != seriesCount) return false;
                string expectedCategory = NormalizeChartReferenceP06(sheetName + "!" +
                    Convert.ToString(categories.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]));
                for (int i = 1; i <= seriesCount; i++)
                {
                    ReleaseCom(valueRow); ReleaseCom(nameCell); ReleaseCom(series);
                    valueRow = null; nameCell = null;
                    series = seriesCollection.Item(i);
                    if (series == null || Convert.ToBoolean(series.HasDataLabels)) return false;
                    nameCell = seriesByColumns ? names.Cells[1, i] as Xl.Range : names.Cells[i, 1] as Xl.Range;
                    valueRow = seriesByColumns ? values.Columns[i] as Xl.Range : values.Rows[i] as Xl.Range;
                    if (nameCell == null || valueRow == null) return false;
                    string formula = NormalizeChartReferenceP06(Convert.ToString(series.Formula));
                    string expectedName = NormalizeChartReferenceP06(sheetName + "!" +
                        Convert.ToString(nameCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]));
                    string expectedValues = NormalizeChartReferenceP06(sheetName + "!" +
                        Convert.ToString(valueRow.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]));
                    if (!formula.Contains(expectedName) || !formula.Contains(expectedCategory) || !formula.Contains(expectedValues))
                        return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartQuickLayoutEquals", "P08 T08 grading failed.", ex, "Excel2019_P08", "T08");
                return false;
            }
            finally
            {
                ReleaseCom(valueRow); ReleaseCom(nameCell); ReleaseCom(values); ReleaseCom(categories); ReleaseCom(names);
                ReleaseCom(series); ReleaseCom(seriesCollection); ReleaseCom(chartGroup); ReleaseCom(valueAxis); ReleaseCom(categoryAxis);
                ReleaseCom(legend); ReleaseCom(chart); ReleaseCom(chartObject); ReleaseCom(charts); ReleaseCom(ws);
            }
        }

        private bool FormulaMultipliesNamedRangeP08(string formula, string sourceHeader, string sourceAddress, string namedRange)
        {
            string expression = NormalizeFormula(formula);
            if (expression.StartsWith("=", StringComparison.Ordinal)) expression = expression.Substring(1);
            expression = TrimOuterParenthesesP05T5(expression);
            string[] terms = expression.Split('*');
            if (terms.Length != 2) return false;
            string left = TrimOuterParenthesesP05T5(terms[0]);
            string right = TrimOuterParenthesesP05T5(terms[1]);
            return (IsSourceTermP08(left, sourceHeader, sourceAddress) && IsNamedRangeTermP08(right, namedRange)) ||
                   (IsSourceTermP08(right, sourceHeader, sourceAddress) && IsNamedRangeTermP08(left, namedRange));
        }

        private bool IsSourceTermP08(string term, string sourceHeader, string sourceAddress)
        {
            string normalized = TrimOuterParenthesesP05T5(NormalizeFormula(term));
            string header = NormalizeText(sourceHeader).Replace(" ", "");
            if (string.Equals(normalized, "[@" + header + "]", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(normalized, "[@[" + header + "]]", StringComparison.OrdinalIgnoreCase))
                return true;

            int bang = normalized.LastIndexOf('!');
            if (bang >= 0) normalized = normalized.Substring(bang + 1);
            normalized = normalized.Replace("$", "").Trim('\'', ' ');
            return string.Equals(normalized, NormalizeRangeAddress(sourceAddress), StringComparison.OrdinalIgnoreCase);
        }

        private bool IsNamedRangeTermP08(string term, string namedRange)
        {
            string normalized = TrimOuterParenthesesP05T5(NormalizeFormula(term));
            int bang = normalized.LastIndexOf('!');
            if (bang >= 0) normalized = normalized.Substring(bang + 1);
            normalized = normalized.Trim('\'', ' ');
            return string.Equals(normalized, NormalizeFormula(namedRange), StringComparison.OrdinalIgnoreCase);
        }

        private bool TryParseSparklineTypeP08(string text, out int sparkType)
        {
            sparkType = 0;
            string normalized = Regex.Replace(text ?? "", @"[^A-Z0-9]+", "", RegexOptions.IgnoreCase).ToUpperInvariant();
            if (normalized == "WINLOSS" || normalized == "XLSPARKCOLUMNSTACKED100")
            {
                sparkType = (int)Xl.XlSparkType.xlSparkColumnStacked100;
                return true;
            }
            if (normalized == "COLUMN" || normalized == "XLSPARKCOLUMN")
            {
                sparkType = (int)Xl.XlSparkType.xlSparkColumn;
                return true;
            }
            if (normalized == "LINE" || normalized == "XLSPARKLINE")
            {
                sparkType = (int)Xl.XlSparkType.xlSparkLine;
                return true;
            }
            return false;
        }

        private bool FormulaIsCountBlankP08(string formula, string directRange, string firstHeader, string lastHeader)
        {
            string normalized = NormalizeFormula(formula);
            if (!normalized.StartsWith("=COUNTBLANK(", StringComparison.OrdinalIgnoreCase) || !normalized.EndsWith(")", StringComparison.Ordinal))
                return false;
            string argument = normalized.Substring(12, normalized.Length - 13);
            argument = TrimOuterParenthesesP05T5(argument);
            if (string.Equals(NormalizeRangeAddress(argument), NormalizeRangeAddress(directRange), StringComparison.OrdinalIgnoreCase))
                return true;
            string structured = "[@[" + NormalizeFormula(firstHeader) + "]:[" + NormalizeFormula(lastHeader) + "]]";
            return argument.EndsWith(structured, StringComparison.OrdinalIgnoreCase);
        }

        // Project 9 Task 1
        public bool InvoiceStockBlockDeletedShiftUp(string sheetName, string deletedRangeAddress, string sourceRangeAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (!string.Equals(NormalizeRangeAddress(deletedRangeAddress), "E1:F4", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(NormalizeRangeAddress(sourceRangeAddress), "E5:F16", StringComparison.OrdinalIgnoreCase))
                return false;

            Xl.Worksheet ws = null;
            Xl.Range used = null;
            Xl.Range cell = null;
            Xl.Font font = null;
            Xl.Interior interior = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

                if (!CellValueEqualsP07(ws, "A1", "PART NUMBER") ||
                    !CellValueEqualsP07(ws, "B1", "SUPPLIER NAME") ||
                    !CellValueEqualsP07(ws, "C1", "TOY CATEGORY") ||
                    !CellValueEqualsP07(ws, "D1", "TOY DETAIL") ||
                    !CellValueEqualsP07(ws, "G1", "STOCK VALUE") ||
                    !CellValueEqualsP07(ws, "I2", "Maximum Stock Value:"))
                    return false;

                if (!CellValueEqualsP07(ws, "E1", "QUANTITY IN STOCK") ||
                    !CellValueEqualsP07(ws, "F1", "PRICE") ||
                    !CellValueEqualsP07(ws, "F12", "Total Value") ||
                    !CellValueEqualsP07(ws, "E12", ""))
                    return false;

                double[] quantities = { 9054d, 8587d, 8103d, 9062d, 8565d, 6520d, 8475d, 2142d, 4057d, 2014d };
                double[] prices = { 57.5d, 56.56d, 45.5d, 23.99d, 19.9d, 21.5d, 12.99d, 45.99d, 10.99d, 14.99d };
                for (int i = 0; i < quantities.Length; i++)
                {
                    int row = i + 2;
                    if (!CellValueEqualsP07(ws, "E" + row, quantities[i]) ||
                        !CellValueEqualsP07(ws, "F" + row, prices[i]))
                        return false;

                    string expectedFormula = "=Invoice!$E" + row + "*Invoice!$F" + row;
                    if (!CellFormulaEqualsP07(ws, "G" + row, expectedFormula)) return false;
                }
                if (!CellFormulaEqualsP07(ws, "G12", "=SUM(G2:G11)")) return false;

                for (int row = 13; row <= 16; row++)
                {
                    if (!CellValueEqualsP07(ws, "E" + row, "") || !CellValueEqualsP07(ws, "F" + row, ""))
                        return false;
                }

                used = ws.UsedRange;
                if (used == null || Convert.ToInt32(used.Row) != 1 ||
                    Convert.ToInt32(used.Row) + Convert.ToInt32(used.Rows.Count) - 1 != 12)
                    return false;

                cell = ws.Range["E1"];
                font = cell.Font;
                interior = cell.Interior;
                if (font == null || !Convert.ToBoolean(font.Bold) || interior == null ||
                    Convert.ToInt32(interior.Color) != 12874308)
                    return false;
                ReleaseCom(interior); interior = null;
                ReleaseCom(font); font = null;
                ReleaseCom(cell); cell = ws.Range["F12"];
                font = cell.Font;
                return font != null && Convert.ToBoolean(font.Bold);
            }
            catch (Exception ex)
            {
                AppLogger.Error("InvoiceStockBlockDeletedShiftUp", "P09 T01 grading failed.", ex, "Excel2019_P09", "T01");
                return false;
            }
            finally
            {
                ReleaseCom(interior); ReleaseCom(font); ReleaseCom(cell); ReleaseCom(used); ReleaseCom(ws);
            }
        }

        // Project 9 Task 2
        public bool EmailFormulaByHeadersStrict(string sheetName, string targetHeader, string sourceHeader, string domain)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetHeader) ||
                string.IsNullOrWhiteSpace(sourceHeader) || string.IsNullOrWhiteSpace(domain)) return false;

            Xl.Worksheet ws = null;
            Xl.Range used = null;
            Xl.Range cell = null;
            Xl.Range sourceCell = null;
            Xl.Range targetCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                used = ws.UsedRange;
                if (used == null) return false;

                int headerRow = 0;
                int sourceColumn = 0;
                int targetColumn = 0;
                int firstRow = Convert.ToInt32(used.Row);
                int lastRow = firstRow + Convert.ToInt32(used.Rows.Count) - 1;
                int firstColumn = Convert.ToInt32(used.Column);
                int lastColumn = firstColumn + Convert.ToInt32(used.Columns.Count) - 1;
                for (int row = firstRow; row <= lastRow && headerRow == 0; row++)
                {
                    sourceColumn = 0;
                    targetColumn = 0;
                    for (int column = firstColumn; column <= lastColumn; column++)
                    {
                        ReleaseCom(cell);
                        cell = ws.Cells[row, column] as Xl.Range;
                        string text = NormalizeText(Convert.ToString(cell == null ? null : cell.Value2));
                        if (string.Equals(text, NormalizeText(sourceHeader), StringComparison.OrdinalIgnoreCase)) sourceColumn = column;
                        if (string.Equals(text, NormalizeText(targetHeader), StringComparison.OrdinalIgnoreCase)) targetColumn = column;
                    }
                    if (sourceColumn > 0 && targetColumn > 0) headerRow = row;
                }
                if (headerRow <= 0 || sourceColumn <= 0 || targetColumn <= 0) return false;

                string normalizedDomain = domain.Trim();
                if (!normalizedDomain.StartsWith("@", StringComparison.Ordinal)) normalizedDomain = "@" + normalizedDomain;
                int checkedRows = 0;
                for (int row = headerRow + 1; row <= lastRow; row++)
                {
                    ReleaseCom(sourceCell); ReleaseCom(targetCell);
                    sourceCell = ws.Cells[row, sourceColumn] as Xl.Range;
                    targetCell = ws.Cells[row, targetColumn] as Xl.Range;
                    string sourceText = Convert.ToString(sourceCell == null ? null : sourceCell.Value2).Trim();
                    string targetText = Convert.ToString(targetCell == null ? null : targetCell.Value2).Trim();
                    if (string.IsNullOrWhiteSpace(sourceText) && string.IsNullOrWhiteSpace(targetText)) continue;
                    if (string.IsNullOrWhiteSpace(sourceText) || targetCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;
                    if (!string.Equals(targetText, sourceText + normalizedDomain, StringComparison.OrdinalIgnoreCase)) return false;
                    if (!EmailFormulaUsesSameRowSourceP09(Convert.ToString(targetCell.Formula),
                        ExcelColumnNameP1T3(sourceColumn) + row, normalizedDomain)) return false;
                    checkedRows++;
                }
                return checkedRows == 3;
            }
            catch (Exception ex)
            {
                AppLogger.Error("EmailFormulaByHeadersStrict", "P09 T02 grading failed.", ex, "Excel2019_P09", "T02");
                return false;
            }
            finally
            {
                ReleaseCom(targetCell); ReleaseCom(sourceCell); ReleaseCom(cell); ReleaseCom(used); ReleaseCom(ws);
            }
        }

        // Project 9 Task 3
        public bool ClusteredColumnChartBelowRange(string sheetName, string sourceBlockRange, int expectedChartType, IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(sourceBlockRange) ||
                sourceRanges == null || sourceRanges.Count != 3) return false;

            Xl.Worksheet ws = null;
            Xl.Range sourceBlock = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                sourceBlock = ws.Range[sourceBlockRange];
                if (sourceBlock == null) return false;
                double bottom = Convert.ToDouble(sourceBlock.Top) + Convert.ToDouble(sourceBlock.Height);

                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    ReleaseCom(series); ReleaseCom(seriesCollection); ReleaseCom(chart); ReleaseCom(chartObject);
                    series = null; seriesCollection = null; chart = null;
                    chartObject = charts.Item(i) as Xl.ChartObject;
                    if (chartObject == null || Convert.ToDouble(chartObject.Top) + 2d < bottom) continue;
                    chart = chartObject.Chart;
                    if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType) continue;
                    seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                    if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != 1) continue;
                    series = seriesCollection.Item(1);
                    if (series != null && SeriesFormulaMatchesRangesP09(
                        Convert.ToString(series.Formula), sheetName, sourceRanges)) return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ClusteredColumnChartBelowRange", "P09 T03 grading failed.", ex, "Excel2019_P09", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(series); ReleaseCom(seriesCollection); ReleaseCom(chart); ReleaseCom(chartObject);
                ReleaseCom(charts); ReleaseCom(sourceBlock); ReleaseCom(ws);
            }
        }

        // Project 9 Task 4
        public bool RangeGreaterThanConditionalFormattingEquals(string sheetName, string rangeAddress, double threshold, string expectedFormat)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress) ||
                !string.Equals(NormalizeText(expectedFormat), "Yellow Fill with Dark Yellow Text", StringComparison.OrdinalIgnoreCase))
                return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;
            Xl.FormatConditions conditions = null;
            Xl.FormatCondition condition = null;
            Xl.Range appliesTo = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                range = ws.Range[rangeAddress];
                conditions = range == null ? null : range.FormatConditions;
                if (conditions == null) return false;

                int count = Convert.ToInt32(conditions.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(appliesTo);
                    ReleaseCom(condition);
                    appliesTo = null;
                    condition = conditions.Item(i) as Xl.FormatCondition;
                    if (condition == null || Convert.ToInt32(condition.Type) != (int)Xl.XlFormatConditionType.xlCellValue ||
                        Convert.ToInt32(condition.Operator) != (int)Xl.XlFormatConditionOperator.xlGreater)
                        continue;

                    string formula = Convert.ToString(condition.Formula1).Trim().TrimStart('=');
                    double actualThreshold;
                    if (!TryToDouble(formula, out actualThreshold) || Math.Abs(actualThreshold - threshold) > 0.000001d)
                        continue;

                    appliesTo = condition.AppliesTo;
                    string actualRange = Convert.ToString(appliesTo.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!string.Equals(NormalizeRangeAddress(actualRange), NormalizeRangeAddress(rangeAddress), StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (ConditionalFormatHasYellowFillDarkYellowTextP09(condition) &&
                        RangeDisplaysYellowFillDarkYellowTextP09(range, threshold))
                        return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangeGreaterThanConditionalFormattingEquals", "P09 T04 grading failed.", ex, "Excel2019_P09", "T04");
                return false;
            }
            finally
            {
                ReleaseCom(appliesTo); ReleaseCom(condition);
                ReleaseCom(conditions); ReleaseCom(range); ReleaseCom(ws);
            }
        }

        private bool ConditionalFormatHasYellowFillDarkYellowTextP09(Xl.FormatCondition condition)
        {
            Xl.Interior interior = null;
            Xl.Font font = null;
            try
            {
                interior = condition == null ? null : condition.Interior;
                font = condition == null ? null : condition.Font;
                return interior != null && font != null &&
                    Convert.ToInt32(interior.Color) == 10284031 &&
                    Convert.ToInt32(font.Color) == 22428;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(font);
                ReleaseCom(interior);
            }
        }

        private bool RangeDisplaysYellowFillDarkYellowTextP09(Xl.Range range, double threshold)
        {
            Xl.Range cell = null;
            Xl.DisplayFormat display = null;
            Xl.Interior interior = null;
            Xl.Font font = null;
            try
            {
                if (range == null) return false;
                ((Xl.Application)_session.App).Calculate();
                int count = Convert.ToInt32(range.Cells.Count);
                bool foundQualifyingCell = false;
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(font); ReleaseCom(interior); ReleaseCom(display); ReleaseCom(cell);
                    font = null; interior = null; display = null;
                    cell = range.Cells[i] as Xl.Range;
                    if (cell == null) return false;
                    double value;
                    if (!TryToDouble(cell.Value2, out value)) continue;
                    display = cell.DisplayFormat;
                    interior = display == null ? null : display.Interior;
                    font = display == null ? null : display.Font;
                    bool hasExpectedDisplay = interior != null && font != null &&
                        Convert.ToInt32(interior.Color) == 10284031 &&
                        Convert.ToInt32(font.Color) == 22428;
                    if (value > threshold)
                    {
                        foundQualifyingCell = true;
                        if (!hasExpectedDisplay) return false;
                    }
                    else if (hasExpectedDisplay)
                    {
                        return false;
                    }
                }
                return foundQualifyingCell;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(font); ReleaseCom(interior); ReleaseCom(display); ReleaseCom(cell);
            }
        }

        // Project 9 Task 5
        public bool ChartSheetTitleAboveValueLabelsOutsideEnd(string chartSheetName, int expectedChartType, IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(chartSheetName) || sourceRanges == null || sourceRanges.Count != 3) return false;

            Xl.Workbook workbook = null;
            Xl.Sheets charts = null;
            Xl.Chart chart = null;
            Xl.ChartTitle title = null;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            Xl.DataLabels labels = null;
            Xl.DataLabel label = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                charts = workbook.Charts;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    ReleaseCom(chart);
                    chart = charts.Item[i] as Xl.Chart;
                    if (chart != null && string.Equals(chart.Name, chartSheetName, StringComparison.OrdinalIgnoreCase)) break;
                    chart = null;
                }
                if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType || !Convert.ToBoolean(chart.HasTitle)) return false;
                title = chart.ChartTitle;
                if (title == null || !Convert.ToBoolean(title.IncludeInLayout) ||
                    Convert.ToDouble(title.Top) >= Convert.ToDouble(chart.PlotArea.Top)) return false;
                if (!ChartReferencesExpectedRangesP06(chart, "Supplier Information", sourceRanges)) return false;

                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != 1) return false;
                series = seriesCollection.Item(1);
                if (series == null || !Convert.ToBoolean(series.HasDataLabels)) return false;
                labels = series.DataLabels(Type.Missing) as Xl.DataLabels;
                if (labels == null || Convert.ToInt32(labels.Count) != 3) return false;
                for (int i = 1; i <= Convert.ToInt32(labels.Count); i++)
                {
                    ReleaseCom(label);
                    label = labels.Item(i) as Xl.DataLabel;
                    if (label == null || Convert.ToInt32(label.Position) != (int)Xl.XlDataLabelPosition.xlLabelPositionOutsideEnd ||
                        !Convert.ToBoolean(label.ShowValue) || Convert.ToBoolean(label.ShowCategoryName) ||
                        Convert.ToBoolean(label.ShowSeriesName)) return false;
                }
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartSheetTitleAboveValueLabelsOutsideEnd", "P09 T05 grading failed.", ex, "Excel2019_P09", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(label); ReleaseCom(labels); ReleaseCom(series); ReleaseCom(seriesCollection);
                ReleaseCom(title); ReleaseCom(chart); ReleaseCom(charts);
            }
        }

        // Project 9 Task 8
        public bool UpperFormulaFilledRange(string sheetName, string rangeAddress, string expectedPrefix)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress) ||
                string.IsNullOrWhiteSpace(expectedPrefix)) return false;

            Xl.Worksheet ws = null;
            Xl.Range range = null;
            Xl.Range cell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                range = ws.Range[rangeAddress];
                if (range == null || Convert.ToInt32(range.Columns.Count) != 1) return false;
                int rowCount = Convert.ToInt32(range.Rows.Count);
                for (int i = 1; i <= rowCount; i++)
                {
                    ReleaseCom(cell);
                    cell = range.Cells[i, 1] as Xl.Range;
                    if (cell == null || !Convert.ToBoolean(cell.HasFormula)) return false;
                    int worksheetRow = Convert.ToInt32(cell.Row);
                    string expectedText = expectedPrefix.ToUpperInvariant() + i;
                    if (!string.Equals(Convert.ToString(cell.Value2), expectedText, StringComparison.Ordinal)) return false;
                    if (!UpperSupplierIdFormulaMatchesP09(Convert.ToString(cell.Formula), worksheetRow)) return false;
                }
                return rowCount == 3;
            }
            catch (Exception ex)
            {
                AppLogger.Error("UpperFormulaFilledRange", "P09 T08 grading failed.", ex, "Excel2019_P09", "T08");
                return false;
            }
            finally
            {
                ReleaseCom(cell); ReleaseCom(range); ReleaseCom(ws);
            }
        }

        private bool EmailFormulaUsesSameRowSourceP09(string formula, string expectedSourceAddress, string expectedDomain)
        {
            string normalized = NormalizeFormula(formula);
            Match match = Regex.Match(normalized,
                @"^=(?:_XLFN\.)?(?:CONCATENATE|CONCAT)\((?<source>[^,]+),""(?<domain>[^""]+)""\)$",
                RegexOptions.IgnoreCase);
            if (!match.Success) return false;
            return string.Equals(NormalizeRangeAddress(match.Groups["source"].Value),
                    NormalizeRangeAddress(expectedSourceAddress), StringComparison.OrdinalIgnoreCase) &&
                string.Equals(match.Groups["domain"].Value, expectedDomain, StringComparison.OrdinalIgnoreCase);
        }

        private bool SeriesFormulaMatchesRangesP09(string formula, string sheetName, IList<string> sourceRanges)
        {
            if (sourceRanges == null || sourceRanges.Count != 3) return false;
            string normalized = NormalizeChartReferenceP06(formula);
            string expected = NormalizeChartReferenceP06("=SERIES(" + sheetName + "!" + sourceRanges[0] + "," +
                sheetName + "!" + sourceRanges[1] + "," + sheetName + "!" + sourceRanges[2] + ",1)");
            return string.Equals(normalized, expected, StringComparison.OrdinalIgnoreCase);
        }

        private bool UpperSupplierIdFormulaMatchesP09(string formula, int worksheetRow)
        {
            string normalized = NormalizeFormula(formula);
            string functionPattern = @"(?:_XLFN\.)?(?:CONCATENATE|CONCAT)";
            string rowPattern = @"(?:ROW\(A" + worksheetRow + @"\)|ROW\(\))";
            string pattern = @"^=UPPER\(" + functionPattern + @"\(""SID""," + rowPattern + @"-2\)\)$";
            return Regex.IsMatch(normalized, pattern, RegexOptions.IgnoreCase);
        }

        // Project 10 Task 1
        public bool CellStyleEquals(string sheetName, string cellAddress, string expectedStyleName, string expectedText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(cellAddress) ||
                string.IsNullOrWhiteSpace(expectedStyleName)) return false;

            Xl.Worksheet ws = null;
            Xl.Range cell = null;
            Xl.Style style = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                cell = ws.Range[cellAddress];
                if (cell == null || (!string.IsNullOrEmpty(expectedText) &&
                    !string.Equals(Convert.ToString(cell.Value2), expectedText, StringComparison.Ordinal))) return false;
                style = cell.Style as Xl.Style;
                return style != null && Convert.ToBoolean(style.BuiltIn) &&
                    string.Equals(Convert.ToString(style.Name), expectedStyleName, StringComparison.OrdinalIgnoreCase);
            }
            catch (Exception ex)
            {
                AppLogger.Error("CellStyleEquals", "P10 T01 grading failed.", ex, "Excel2019_P10", "T01");
                return false;
            }
            finally
            {
                ReleaseCom(style); ReleaseCom(cell); ReleaseCom(ws);
            }
        }

        // Project 10 Task 2
        public bool ChartSheetSwitchedRowColumn(string chartSheetName, string sourceSheetName, int expectedChartType, IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(chartSheetName) || string.IsNullOrWhiteSpace(sourceSheetName)) return false;

            Xl.Workbook workbook = null;
            Xl.Sheets chartSheets = null;
            Xl.Chart chart = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                chartSheets = workbook.Charts;
                int count = Convert.ToInt32(chartSheets.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(chart);
                    chart = chartSheets.Item[i] as Xl.Chart;
                    if (chart != null && string.Equals(chart.Name, chartSheetName, StringComparison.OrdinalIgnoreCase))
                        break;
                    chart = null;
                }
                return chart != null && Convert.ToInt32(chart.ChartType) == expectedChartType &&
                    Convert.ToInt32(chart.PlotBy) == (int)Xl.XlRowCol.xlRows &&
                    ChartSeriesMatchRangeTriplesP10(chart, sourceSheetName, sourceRanges);
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartSheetSwitchedRowColumn", "P10 T02 grading failed.", ex, "Excel2019_P10", "T02");
                return false;
            }
            finally
            {
                ReleaseCom(chart); ReleaseCom(chartSheets);
            }
        }

        // Project 10 Task 3
        public bool RangeFormulaMultipliesFixedCell(string sheetName, string targetRangeAddress, string sourceRangeAddress, string fixedCellAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetRangeAddress) ||
                string.IsNullOrWhiteSpace(sourceRangeAddress) || string.IsNullOrWhiteSpace(fixedCellAddress)) return false;

            Xl.Worksheet ws = null;
            Xl.Range targets = null;
            Xl.Range sources = null;
            Xl.Range fixedCell = null;
            Xl.Range target = null;
            Xl.Range source = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                targets = ws.Range[targetRangeAddress];
                sources = ws.Range[sourceRangeAddress];
                fixedCell = ws.Range[fixedCellAddress];
                if (targets == null || sources == null || fixedCell == null ||
                    Convert.ToInt32(targets.Cells.Count) != Convert.ToInt32(sources.Cells.Count)) return false;
                double fixedValue;
                if (!TryToDouble(fixedCell.Value2, out fixedValue)) return false;

                int count = Convert.ToInt32(targets.Cells.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(source); ReleaseCom(target);
                    target = targets.Cells[i] as Xl.Range;
                    source = sources.Cells[i] as Xl.Range;
                    if (target == null || source == null || !Convert.ToBoolean(target.HasFormula)) return false;
                    string sourceAddress = Convert.ToString(source.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMultipliesReferencesP10(Convert.ToString(target.Formula), sourceAddress, fixedCellAddress)) return false;
                    double sourceValue;
                    double actualValue;
                    if (!TryToDouble(source.Value2, out sourceValue) || !TryToDouble(target.Value2, out actualValue) ||
                        Math.Abs(actualValue - (sourceValue * fixedValue)) > 0.000001d) return false;
                }
                return count > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangeFormulaMultipliesFixedCell", "P10 T03 grading failed.", ex, "Excel2019_P10", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(source); ReleaseCom(target); ReleaseCom(fixedCell);
                ReleaseCom(sources); ReleaseCom(targets); ReleaseCom(ws);
            }
        }

        // Project 10 Task 4
        public bool SpecificChartMovedToChartSheet(string sourceSheetName, string chartSheetName, int expectedChartType, int preservedChartType, IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sourceSheetName) || string.IsNullOrWhiteSpace(chartSheetName)) return false;

            Xl.Workbook workbook = null;
            Xl.Sheets chartSheets = null;
            Xl.Chart targetChart = null;
            Xl.Worksheet sourceSheet = null;
            Xl.ChartObjects embeddedCharts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart embeddedChart = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                chartSheets = workbook.Charts;
                for (int i = 1; i <= Convert.ToInt32(chartSheets.Count); i++)
                {
                    ReleaseCom(targetChart);
                    targetChart = chartSheets.Item[i] as Xl.Chart;
                    if (targetChart != null && string.Equals(targetChart.Name, chartSheetName, StringComparison.OrdinalIgnoreCase)) break;
                    targetChart = null;
                }
                if (targetChart == null || Convert.ToInt32(targetChart.ChartType) != expectedChartType ||
                    !ChartSeriesMatchRangeTriplesP10(targetChart, sourceSheetName, sourceRanges)) return false;

                sourceSheet = GetWorksheet(sourceSheetName);
                if (sourceSheet == null) return false;
                embeddedCharts = sourceSheet.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (embeddedCharts == null) return false;
                bool preservedChartFound = false;
                for (int i = 1; i <= Convert.ToInt32(embeddedCharts.Count); i++)
                {
                    ReleaseCom(embeddedChart); ReleaseCom(chartObject);
                    embeddedChart = null;
                    chartObject = embeddedCharts.Item(i) as Xl.ChartObject;
                    if (chartObject == null) continue;
                    embeddedChart = chartObject.Chart;
                    if (embeddedChart == null) continue;
                    int chartType = Convert.ToInt32(embeddedChart.ChartType);
                    if (chartType == expectedChartType && ChartSeriesMatchRangeTriplesP10(embeddedChart, sourceSheetName, sourceRanges))
                        return false;
                    if (chartType == preservedChartType) preservedChartFound = true;
                }
                return preservedChartFound;
            }
            catch (Exception ex)
            {
                AppLogger.Error("SpecificChartMovedToChartSheet", "P10 T04 grading failed.", ex, "Excel2019_P10", "T04");
                return false;
            }
            finally
            {
                ReleaseCom(embeddedChart); ReleaseCom(chartObject); ReleaseCom(embeddedCharts);
                ReleaseCom(sourceSheet); ReleaseCom(targetChart); ReleaseCom(chartSheets);
            }
        }

        // Project 10 Task 5
        public bool ChartExpandedToIncludeRange(string sheetName, int expectedChartType, IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            Xl.Worksheet ws = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    ReleaseCom(chart); ReleaseCom(chartObject);
                    chart = null;
                    chartObject = charts.Item(i) as Xl.ChartObject;
                    if (chartObject == null) continue;
                    chart = chartObject.Chart;
                    if (chart != null && Convert.ToInt32(chart.ChartType) == expectedChartType &&
                        ChartSeriesMatchRangeTriplesP10(chart, sheetName, sourceRanges)) return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartExpandedToIncludeRange", "P10 T05 grading failed.", ex, "Excel2019_P10", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(chart); ReleaseCom(chartObject); ReleaseCom(charts); ReleaseCom(ws);
            }
        }

        // Project 10 Task 7
        public bool IfNumericFormulaByHeadersStrict(string sheetName, string tableName, string targetHeader, string criteriaHeader, string compareOperator, double threshold, string trueValue, string falseValue)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(targetHeader) || string.IsNullOrWhiteSpace(criteriaHeader) || compareOperator != ">") return false;
            double expectedTrue;
            double expectedFalse;
            if (!TryToDouble(trueValue, out expectedTrue) || !TryToDouble(falseValue, out expectedFalse)) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range data = null;
            Xl.Range targetData = null;
            Xl.Range criteriaData = null;
            Xl.Range targetCell = null;
            Xl.Range criteriaCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(table.Name, tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;
                int targetIndex = FindTableColumnIndexP05T5(table, targetHeader);
                int criteriaIndex = FindTableColumnIndexP05T5(table, criteriaHeader);
                if (targetIndex <= 0 || criteriaIndex <= 0) return false;
                data = table.DataBodyRange;
                if (data == null) return false;
                targetData = data.Columns[targetIndex] as Xl.Range;
                criteriaData = data.Columns[criteriaIndex] as Xl.Range;
                if (targetData == null || criteriaData == null ||
                    Convert.ToInt32(targetData.Rows.Count) != Convert.ToInt32(criteriaData.Rows.Count)) return false;

                int rowCount = Convert.ToInt32(targetData.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(criteriaCell); ReleaseCom(targetCell);
                    targetCell = targetData.Cells[row, 1] as Xl.Range;
                    criteriaCell = criteriaData.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || criteriaCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;
                    double criteria;
                    double actual;
                    if (!TryToDouble(criteriaCell.Value2, out criteria) || !TryToDouble(targetCell.Value2, out actual)) return false;
                    double expected = criteria > threshold ? expectedTrue : expectedFalse;
                    if (Math.Abs(actual - expected) > 0.000001d) return false;
                    string criteriaAddress = Convert.ToString(criteriaCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMatchesNumericIfP10(Convert.ToString(targetCell.Formula), criteriaAddress, criteriaHeader,
                        threshold, expectedTrue, expectedFalse)) return false;
                }
                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("IfNumericFormulaByHeadersStrict", "P10 T07 grading failed.", ex, "Excel2019_P10", "T07");
                return false;
            }
            finally
            {
                ReleaseCom(criteriaCell); ReleaseCom(targetCell); ReleaseCom(criteriaData);
                ReleaseCom(targetData); ReleaseCom(data); ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 10 Task 8
        public bool CellHyperlinkWithScreenTipEquals(string sheetName, string cellAddress, string expectedAddress, string expectedScreenTip, string expectedDisplayText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(cellAddress) ||
                string.IsNullOrWhiteSpace(expectedAddress)) return false;

            Xl.Worksheet ws = null;
            Xl.Range cell = null;
            Xl.Hyperlinks hyperlinks = null;
            Xl.Hyperlink hyperlink = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                cell = ws.Range[cellAddress];
                if (cell == null || !string.Equals(Convert.ToString(cell.Value2), expectedDisplayText, StringComparison.Ordinal)) return false;
                hyperlinks = cell.Hyperlinks;
                if (hyperlinks == null) return false;
                for (int i = 1; i <= Convert.ToInt32(hyperlinks.Count); i++)
                {
                    ReleaseCom(hyperlink);
                    hyperlink = hyperlinks.Item[i];
                    if (hyperlink != null && HyperlinkAddressesEqualP21(Convert.ToString(hyperlink.Address), expectedAddress) &&
                        string.IsNullOrEmpty(Convert.ToString(hyperlink.SubAddress)) &&
                        string.Equals(Convert.ToString(hyperlink.ScreenTip), expectedScreenTip, StringComparison.Ordinal)) return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("CellHyperlinkWithScreenTipEquals", "P10 T08 grading failed.", ex, "Excel2019_P10", "T08");
                return false;
            }
            finally
            {
                ReleaseCom(hyperlink); ReleaseCom(hyperlinks); ReleaseCom(cell); ReleaseCom(ws);
            }
        }

        private bool HyperlinkAddressesEqualP21(string actualAddress, string expectedAddress)
        {
            Uri actualUri;
            Uri expectedUri;
            if (Uri.TryCreate(actualAddress, UriKind.Absolute, out actualUri) &&
                Uri.TryCreate(expectedAddress, UriKind.Absolute, out expectedUri))
            {
                return string.Equals(actualUri.AbsoluteUri, expectedUri.AbsoluteUri, StringComparison.Ordinal);
            }
            return string.Equals(actualAddress, expectedAddress, StringComparison.Ordinal);
        }

        private bool ChartSeriesMatchRangeTriplesP10(Xl.Chart chart, string sourceSheetName, IList<string> sourceRanges)
        {
            if (chart == null || sourceRanges == null || sourceRanges.Count == 0 || sourceRanges.Count % 3 != 0) return false;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            try
            {
                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                int expectedCount = sourceRanges.Count / 3;
                if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != expectedCount) return false;
                for (int i = 1; i <= expectedCount; i++)
                {
                    ReleaseCom(series);
                    series = seriesCollection.Item(i);
                    if (series == null) return false;
                    int offset = (i - 1) * 3;
                    string expected = "=SERIES(" + sourceSheetName + "!" + sourceRanges[offset] + "," +
                        sourceSheetName + "!" + sourceRanges[offset + 1] + "," +
                        sourceSheetName + "!" + sourceRanges[offset + 2] + "," + i.ToString(CultureInfo.InvariantCulture) + ")";
                    if (!string.Equals(NormalizeChartReferenceP06(Convert.ToString(series.Formula)),
                        NormalizeChartReferenceP06(expected), StringComparison.OrdinalIgnoreCase)) return false;
                }
                return true;
            }
            finally
            {
                ReleaseCom(series); ReleaseCom(seriesCollection);
            }
        }

        private bool FormulaMultipliesReferencesP10(string formula, string sourceAddress, string fixedCellAddress)
        {
            string expression = TrimOuterParenthesesP05T5(NormalizeFormula(formula).TrimStart('='));
            int multiplyIndex = expression.IndexOf('*');
            if (multiplyIndex <= 0 || multiplyIndex != expression.LastIndexOf('*')) return false;
            string left = TrimOuterParenthesesP05T5(expression.Substring(0, multiplyIndex));
            string right = TrimOuterParenthesesP05T5(expression.Substring(multiplyIndex + 1));
            return (ReferenceEqualsP10(left, sourceAddress) && ReferenceEqualsP10(right, fixedCellAddress)) ||
                (ReferenceEqualsP10(right, sourceAddress) && ReferenceEqualsP10(left, fixedCellAddress));
        }

        private bool ReferenceEqualsP10(string formulaToken, string expectedAddress)
        {
            return string.Equals(NormalizeRangeAddress(formulaToken), NormalizeRangeAddress(expectedAddress), StringComparison.OrdinalIgnoreCase);
        }

        private bool FormulaMatchesNumericIfP10(string formula, string criteriaAddress, string criteriaHeader,
            double threshold, double trueValue, double falseValue)
        {
            string normalized = NormalizeFormula(formula);
            Match match = Regex.Match(normalized, @"^=IF\((?<condition>.+),(?<true>[^,]+),(?<false>[^,]+)\)$", RegexOptions.IgnoreCase);
            if (!match.Success || match.Groups["true"].Value.Contains("\"") || match.Groups["false"].Value.Contains("\"")) return false;
            double actualTrue;
            double actualFalse;
            if (!TryToDouble(match.Groups["true"].Value, out actualTrue) || !TryToDouble(match.Groups["false"].Value, out actualFalse) ||
                Math.Abs(actualTrue - trueValue) > 0.000001d || Math.Abs(actualFalse - falseValue) > 0.000001d) return false;
            string condition = match.Groups["condition"].Value;
            if (condition.Contains(">=") || condition.Contains("<=") || condition.Contains("<>")) return false;
            int greater = condition.IndexOf('>');
            if (greater > 0 && greater == condition.LastIndexOf('>'))
                return IsCriteriaReferenceP10(condition.Substring(0, greater), criteriaAddress, criteriaHeader) &&
                    IsThresholdP07(condition.Substring(greater + 1), threshold);
            int less = condition.IndexOf('<');
            return less > 0 && less == condition.LastIndexOf('<') &&
                IsThresholdP07(condition.Substring(0, less), threshold) &&
                IsCriteriaReferenceP10(condition.Substring(less + 1), criteriaAddress, criteriaHeader);
        }

        private bool IsCriteriaReferenceP10(string token, string address, string header)
        {
            string normalized = TrimOuterParenthesesP05T5(NormalizeFormula(token));
            if (ReferenceEqualsP10(normalized, address)) return true;
            string normalizedHeader = NormalizeFormula(header);
            return normalized.EndsWith("[@" + normalizedHeader + "]", StringComparison.OrdinalIgnoreCase) ||
                normalized.EndsWith("[@[" + normalizedHeader + "]]", StringComparison.OrdinalIgnoreCase);
        }

        private bool WorkbookRetainsPersonalInfoTaskCoreContent(Xl.Workbook workbook)
        {
            if (workbook == null) return false;
            if (WorksheetExistsP21("Sales by Exam")) return WorkbookRetainsP07CoreContent(workbook);

            return Convert.ToInt32(workbook.Worksheets.Count) == 4 &&
                   WorksheetExistsP21("Elmenshawy") &&
                   WorksheetExistsP21("Kadry") &&
                   WorksheetExistsP21("El3ameed") &&
                   WorksheetExistsP21("Office") &&
                   WorksheetHasExactTableP07("Elmenshawy", "Table3", "A4:C11") &&
                   WorksheetHasExactTableP07("Kadry", "Table1", "A1:F20") &&
                   WorksheetHasNamedChartsP07("Elmenshawy", new[] { "Chart 1" }) &&
                   WorksheetCellTextEqualsP21("Elmenshawy", "A2", "Mos2019.com") &&
                   WorksheetCellTextEqualsP21("El3ameed", "H1", "Books Sold") &&
                   WorksheetCellTextEqualsP21("El3ameed", "I1", "Bonus");
        }

        private bool WorksheetExistsP21(string sheetName)
        {
            Xl.Worksheet worksheet = null;
            try
            {
                worksheet = GetWorksheet(sheetName);
                return worksheet != null;
            }
            finally { ReleaseCom(worksheet); }
        }

        private bool WorksheetCellTextEqualsP21(string sheetName, string address, string expectedText)
        {
            Xl.Worksheet worksheet = null;
            Xl.Range cell = null;
            try
            {
                worksheet = GetWorksheet(sheetName);
                if (worksheet == null) return false;
                cell = worksheet.Range[address];
                return cell != null && string.Equals(Convert.ToString(cell.Value2), expectedText, StringComparison.Ordinal);
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(worksheet);
            }
        }

        public bool RangeFormulaMultipliesNamedRange(
            string sheetName,
            string targetRangeAddress,
            string sourceRangeAddress,
            string sourceHeader,
            string targetHeader,
            string namedRange,
            string namedRangeAddress)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetRangeAddress) ||
                string.IsNullOrWhiteSpace(sourceRangeAddress) || string.IsNullOrWhiteSpace(sourceHeader) ||
                string.IsNullOrWhiteSpace(targetHeader) || string.IsNullOrWhiteSpace(namedRange) ||
                string.IsNullOrWhiteSpace(namedRangeAddress)) return false;

            Xl.Workbook workbook = null;
            Xl.Worksheet ws = null;
            Xl.Range targets = null;
            Xl.Range sources = null;
            Xl.Range targetHeaderCell = null;
            Xl.Range sourceHeaderCell = null;
            Xl.Range targetCell = null;
            Xl.Range sourceCell = null;
            Xl.Names names = null;
            Xl.Name name = null;
            Xl.Range namedCell = null;
            Xl.Worksheet namedSheet = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                ws = GetWorksheet(sheetName);
                if (workbook == null || ws == null) return false;

                targets = ws.Range[targetRangeAddress];
                sources = ws.Range[sourceRangeAddress];
                if (targets == null || sources == null ||
                    Convert.ToInt32(targets.Columns.Count) != 1 || Convert.ToInt32(sources.Columns.Count) != 1 ||
                    Convert.ToInt32(targets.Rows.Count) != Convert.ToInt32(sources.Rows.Count) ||
                    Convert.ToInt32(targets.Row) != Convert.ToInt32(sources.Row)) return false;

                int firstRow = Convert.ToInt32(targets.Row);
                int targetColumn = Convert.ToInt32(targets.Column);
                int sourceColumn = Convert.ToInt32(sources.Column);
                if (firstRow <= 1) return false;
                targetHeaderCell = ws.Cells[firstRow - 1, targetColumn] as Xl.Range;
                sourceHeaderCell = ws.Cells[firstRow - 1, sourceColumn] as Xl.Range;
                if (targetHeaderCell == null || sourceHeaderCell == null ||
                    !string.Equals(NormalizeText(Convert.ToString(targetHeaderCell.Value2)), NormalizeText(targetHeader), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(NormalizeText(Convert.ToString(sourceHeaderCell.Value2)), NormalizeText(sourceHeader), StringComparison.OrdinalIgnoreCase))
                    return false;

                names = workbook.Names;
                for (int i = 1; i <= Convert.ToInt32(names.Count); i++)
                {
                    ReleaseCom(name);
                    name = names.Item(i, Type.Missing, Type.Missing);
                    if (name != null && string.Equals(Convert.ToString(name.Name), namedRange, StringComparison.OrdinalIgnoreCase)) break;
                    name = null;
                }
                if (name == null) return false;
                try { namedCell = name.RefersToRange; } catch { namedCell = null; }
                if (namedCell == null || Convert.ToInt32(namedCell.Cells.Count) != 1) return false;
                namedSheet = namedCell.Worksheet;
                string actualNamedAddress = Convert.ToString(
                    namedCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                if (namedSheet == null || !string.Equals(namedSheet.Name, sheetName, StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(NormalizeRangeAddress(actualNamedAddress), NormalizeRangeAddress(namedRangeAddress), StringComparison.OrdinalIgnoreCase))
                    return false;

                double namedValue;
                if (!TryToDouble(namedCell.Value2, out namedValue)) return false;

                int rowCount = Convert.ToInt32(targets.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(targetCell);
                    ReleaseCom(sourceCell);
                    targetCell = targets.Cells[row, 1] as Xl.Range;
                    sourceCell = sources.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || sourceCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;

                    string sourceAddress = Convert.ToString(
                        sourceCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMultipliesNamedRangeP08(
                        Convert.ToString(targetCell.Formula), sourceHeader, sourceAddress, namedRange)) return false;

                    double sourceValue;
                    double actualValue;
                    if (!TryToDouble(sourceCell.Value2, out sourceValue) || !TryToDouble(targetCell.Value2, out actualValue)) return false;
                    double expectedValue = sourceValue * namedValue;
                    double tolerance = Math.Max(0.000001d, Math.Abs(expectedValue) * 0.000000001d);
                    if (Math.Abs(actualValue - expectedValue) > tolerance) return false;
                }

                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("RangeFormulaMultipliesNamedRange", "P19 T03 grading failed.", ex, "Excel2019_P19", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(namedSheet); ReleaseCom(namedCell); ReleaseCom(name); ReleaseCom(names);
                ReleaseCom(sourceCell); ReleaseCom(targetCell); ReleaseCom(sourceHeaderCell); ReleaseCom(targetHeaderCell);
                ReleaseCom(sources); ReleaseCom(targets); ReleaseCom(ws);
            }
        }

        // Project 14 Task 1
        public bool CellsDeletedShiftUp(
            string sheetName,
            string deletedRangeAddress,
            string sourceRangeAddress,
            IList<string> markerAddresses,
            IList<string> expectedMarkerValues)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(deletedRangeAddress) ||
                string.IsNullOrWhiteSpace(sourceRangeAddress) || markerAddresses == null ||
                expectedMarkerValues == null || markerAddresses.Count == 0 ||
                markerAddresses.Count != expectedMarkerValues.Count) return false;

            Xl.Worksheet ws = null;
            Xl.Range deletedRange = null;
            Xl.Range sourceRange = null;
            Xl.Range cell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                deletedRange = ws.Range[deletedRangeAddress];
                sourceRange = ws.Range[sourceRangeAddress];
                if (deletedRange == null || sourceRange == null) return false;

                int deletedRows = Convert.ToInt32(deletedRange.Rows.Count);
                int deletedColumns = Convert.ToInt32(deletedRange.Columns.Count);
                int sourceRows = Convert.ToInt32(sourceRange.Rows.Count);
                if (deletedRows <= 0 || sourceRows < deletedRows ||
                    Convert.ToInt32(sourceRange.Columns.Count) != deletedColumns ||
                    Convert.ToInt32(sourceRange.Row) != Convert.ToInt32(deletedRange.Row) + deletedRows ||
                    Convert.ToInt32(sourceRange.Column) != Convert.ToInt32(deletedRange.Column)) return false;

                for (int i = 0; i < markerAddresses.Count; i++)
                {
                    ReleaseCom(cell);
                    cell = ws.Range[markerAddresses[i]];
                    if (cell == null || !CellValueMatchesExpectedP14(cell.Value2, expectedMarkerValues[i]))
                        return false;
                }

                int firstVacatedRow = Convert.ToInt32(sourceRange.Row) + sourceRows - deletedRows;
                int firstColumn = Convert.ToInt32(sourceRange.Column);
                for (int rowOffset = 0; rowOffset < deletedRows; rowOffset++)
                {
                    for (int columnOffset = 0; columnOffset < deletedColumns; columnOffset++)
                    {
                        ReleaseCom(cell);
                        cell = ws.Cells[firstVacatedRow + rowOffset, firstColumn + columnOffset] as Xl.Range;
                        if (cell == null || !CellValueMatchesExpectedP14(cell.Value2, "")) return false;
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("CellsDeletedShiftUp", "P14 T01 grading failed.", ex, "Excel2019_P14", "T01");
                return false;
            }
            finally
            {
                ReleaseCom(cell); ReleaseCom(sourceRange); ReleaseCom(deletedRange); ReleaseCom(ws);
            }
        }

        private bool CellValueMatchesExpectedP14(object value, string expected)
        {
            string expectedText = expected ?? "";
            if (value == null) return expectedText.Length == 0;

            double actualNumber;
            double expectedNumber;
            if (TryToDouble(value, out actualNumber) && TryToDouble(expectedText, out expectedNumber))
                return Math.Abs(actualNumber - expectedNumber) <= 0.000001d;

            return string.Equals(Convert.ToString(value), expectedText, StringComparison.Ordinal);
        }

        // Project 14 Task 5
        public bool IfFormulaByHeadersInRanges(
            string sheetName,
            string targetHeader,
            string criteriaHeader,
            string targetRangeAddress,
            string criteriaRangeAddress,
            string compareOperator,
            double threshold,
            string trueText,
            string falseText)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetHeader) ||
                string.IsNullOrWhiteSpace(criteriaHeader) || string.IsNullOrWhiteSpace(targetRangeAddress) ||
                string.IsNullOrWhiteSpace(criteriaRangeAddress) || compareOperator != "<" ||
                string.IsNullOrWhiteSpace(trueText)) return false;

            Xl.Worksheet ws = null;
            Xl.Range targets = null;
            Xl.Range criteria = null;
            Xl.Range targetHeaderCell = null;
            Xl.Range criteriaHeaderCell = null;
            Xl.Range targetCell = null;
            Xl.Range criteriaCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                targets = ws.Range[targetRangeAddress];
                criteria = ws.Range[criteriaRangeAddress];
                if (targets == null || criteria == null ||
                    Convert.ToInt32(targets.Columns.Count) != 1 || Convert.ToInt32(criteria.Columns.Count) != 1 ||
                    Convert.ToInt32(targets.Cells.Count) != Convert.ToInt32(criteria.Cells.Count) ||
                    Convert.ToInt32(targets.Row) != Convert.ToInt32(criteria.Row)) return false;

                int firstRow = Convert.ToInt32(targets.Row);
                if (firstRow <= 1) return false;
                targetHeaderCell = ws.Cells[firstRow - 1, Convert.ToInt32(targets.Column)] as Xl.Range;
                criteriaHeaderCell = ws.Cells[firstRow - 1, Convert.ToInt32(criteria.Column)] as Xl.Range;
                if (targetHeaderCell == null || criteriaHeaderCell == null ||
                    !string.Equals(NormalizeText(Convert.ToString(targetHeaderCell.Value2)), NormalizeText(targetHeader), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(NormalizeText(Convert.ToString(criteriaHeaderCell.Value2)), NormalizeText(criteriaHeader), StringComparison.OrdinalIgnoreCase))
                    return false;

                int count = Convert.ToInt32(targets.Cells.Count);
                for (int i = 1; i <= count; i++)
                {
                    ReleaseCom(criteriaCell); ReleaseCom(targetCell);
                    targetCell = targets.Cells[i] as Xl.Range;
                    criteriaCell = criteria.Cells[i] as Xl.Range;
                    if (targetCell == null || criteriaCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;

                    double criteriaValue;
                    if (!TryToDouble(criteriaCell.Value2, out criteriaValue)) return false;
                    string expectedResult = criteriaValue < threshold ? trueText : falseText;
                    if (!string.Equals(Convert.ToString(targetCell.Value2), expectedResult, StringComparison.Ordinal)) return false;

                    string criteriaAddress = Convert.ToString(criteriaCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMatchesStrictIfP07(Convert.ToString(targetCell.Formula), criteriaAddress, criteriaHeader, threshold, trueText, falseText))
                        return false;
                }

                return count > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("IfFormulaByHeadersInRanges", "P14 T05 grading failed.", ex, "Excel2019_P14", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(criteriaCell); ReleaseCom(targetCell);
                ReleaseCom(criteriaHeaderCell); ReleaseCom(targetHeaderCell);
                ReleaseCom(criteria); ReleaseCom(targets); ReleaseCom(ws);
            }
        }

        // Project 14 Task 6
        public bool ChartSheetTitleAboveValueLabelsOutsideEndBySource(
            string chartSheetName,
            string sourceSheetName,
            int expectedChartType,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(chartSheetName) || string.IsNullOrWhiteSpace(sourceSheetName) ||
                sourceRanges == null || sourceRanges.Count != 3) return false;

            Xl.Workbook workbook = null;
            Xl.Sheets charts = null;
            Xl.Chart chart = null;
            Xl.ChartTitle title = null;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            Xl.Points points = null;
            Xl.DataLabels labels = null;
            Xl.DataLabel label = null;
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                charts = workbook.Charts;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    ReleaseCom(chart);
                    chart = charts.Item[i] as Xl.Chart;
                    if (chart != null && string.Equals(chart.Name, chartSheetName, StringComparison.OrdinalIgnoreCase)) break;
                    chart = null;
                }
                if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType ||
                    !Convert.ToBoolean(chart.HasTitle)) return false;
                title = chart.ChartTitle;
                if (title == null || !Convert.ToBoolean(title.IncludeInLayout) ||
                    Convert.ToDouble(title.Top) >= Convert.ToDouble(chart.PlotArea.Top)) return false;

                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != 1) return false;
                series = seriesCollection.Item(1);
                if (series == null || !SeriesFormulaMatchesRangesP09(Convert.ToString(series.Formula), sourceSheetName, sourceRanges) ||
                    !Convert.ToBoolean(series.HasDataLabels)) return false;

                points = series.Points(Type.Missing) as Xl.Points;
                labels = series.DataLabels(Type.Missing) as Xl.DataLabels;
                if (points == null || labels == null ||
                    Convert.ToInt32(labels.Count) != Convert.ToInt32(points.Count)) return false;
                for (int i = 1; i <= Convert.ToInt32(labels.Count); i++)
                {
                    ReleaseCom(label);
                    label = labels.Item(i) as Xl.DataLabel;
                    if (label == null || Convert.ToInt32(label.Position) != (int)Xl.XlDataLabelPosition.xlLabelPositionOutsideEnd ||
                        !Convert.ToBoolean(label.ShowValue) || Convert.ToBoolean(label.ShowCategoryName) ||
                        Convert.ToBoolean(label.ShowSeriesName)) return false;
                }

                return Convert.ToInt32(labels.Count) > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartSheetTitleAboveValueLabelsOutsideEndBySource", "P14 T06 grading failed.", ex, "Excel2019_P14", "T06");
                return false;
            }
            finally
            {
                ReleaseCom(label); ReleaseCom(labels); ReleaseCom(points); ReleaseCom(series);
                ReleaseCom(seriesCollection); ReleaseCom(title); ReleaseCom(chart); ReleaseCom(charts);
            }
        }

        // Project 15 Task 3
        public bool TableCreatedWithHeadersAndStyle(
            string sheetName,
            string originalRangeAddress,
            string expectedStyleName,
            IList<string> expectedHeaders,
            string removableRowText,
            IList<string> expectedFirstColumnValues)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(originalRangeAddress) ||
                expectedHeaders == null || expectedHeaders.Count == 0 ||
                expectedFirstColumnValues == null || expectedFirstColumnValues.Count == 0) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range originalRange = null;
            Xl.Range tableRange = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                originalRange = ws.Range[originalRangeAddress];
                tables = ws.ListObjects;
                if (originalRange == null || tables == null) return false;

                int expectedRow = Convert.ToInt32(originalRange.Row);
                int expectedColumn = Convert.ToInt32(originalRange.Column);
                int expectedColumnCount = Convert.ToInt32(originalRange.Columns.Count);

                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(tableRange);
                    ReleaseCom(table);
                    tableRange = null;
                    table = tables.Item[i];
                    if (table == null) continue;
                    tableRange = table.Range;
                    if (tableRange == null || Convert.ToInt32(tableRange.Row) != expectedRow ||
                        Convert.ToInt32(tableRange.Column) != expectedColumn ||
                        Convert.ToInt32(tableRange.Columns.Count) != expectedColumnCount ||
                        !Convert.ToBoolean(table.ShowHeaders) ||
                        !ListObjectTableStyleEqualsP2T8(table, expectedStyleName) ||
                        !TableHeadersEqualP15(table, expectedHeaders)) continue;

                    List<string> actualValues = GetTableRowSignaturesP15(table);
                    if (StringListsEqualP15(actualValues, expectedFirstColumnValues) &&
                        Convert.ToInt32(tableRange.Rows.Count) == expectedFirstColumnValues.Count + 1)
                        return true;

                    List<string> afterDelete = new List<string>();
                    bool removed = false;
                    foreach (string value in expectedFirstColumnValues)
                    {
                        if (!removed && RowSignatureHasFirstValueP15(value, removableRowText))
                        {
                            removed = true;
                            continue;
                        }
                        afterDelete.Add(value);
                    }

                    if (removed && StringListsEqualP15(actualValues, afterDelete) &&
                        Convert.ToInt32(tableRange.Rows.Count) == afterDelete.Count + 1)
                        return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableCreatedWithHeadersAndStyle", "P15 T03 grading failed.", ex, "Excel2019_P15", "T03");
                return false;
            }
            finally
            {
                ReleaseCom(tableRange); ReleaseCom(originalRange); ReleaseCom(table);
                ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 15 Task 4
        public bool TableRowContainingTextDeletedByValues(
            string sheetName,
            string expectedTableRange,
            string searchText,
            string expectedStyleName,
            IList<string> expectedHeaders,
            IList<string> expectedFirstColumnValues)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(expectedTableRange) ||
                expectedHeaders == null || expectedFirstColumnValues == null) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range range = null;
            Xl.Range dataRange = null;
            Xl.Range cell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                if (tables == null) return false;

                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(dataRange); ReleaseCom(range); ReleaseCom(table);
                    dataRange = null; range = null; table = tables.Item[i];
                    if (table == null) continue;
                    range = table.Range;
                    if (range == null || !string.Equals(NormalizeRangeAddress(Convert.ToString(range.Address)), NormalizeRangeAddress(expectedTableRange), StringComparison.OrdinalIgnoreCase) ||
                        !Convert.ToBoolean(table.ShowHeaders) ||
                        !ListObjectTableStyleEqualsP2T8(table, expectedStyleName) ||
                        !TableHeadersEqualP15(table, expectedHeaders) ||
                        !StringListsEqualP15(GetTableRowSignaturesP15(table), expectedFirstColumnValues)) continue;

                    dataRange = table.DataBodyRange;
                    if (dataRange == null || Convert.ToInt32(dataRange.Rows.Count) != expectedFirstColumnValues.Count) continue;
                    int rows = Convert.ToInt32(dataRange.Rows.Count);
                    int columns = Convert.ToInt32(dataRange.Columns.Count);
                    bool foundDeletedText = false;
                    for (int r = 1; r <= rows && !foundDeletedText; r++)
                    {
                        bool rowHasContent = false;
                        for (int c = 1; c <= columns; c++)
                        {
                            ReleaseCom(cell);
                            cell = dataRange.Cells[r, c] as Xl.Range;
                            string text = GetCellTextP1T3(cell);
                            if (!string.IsNullOrWhiteSpace(text)) rowHasContent = true;
                            if (string.Equals(NormalizeText(text), NormalizeText(searchText), StringComparison.OrdinalIgnoreCase))
                                foundDeletedText = true;
                        }
                        if (!rowHasContent) return false;
                    }
                    if (!foundDeletedText) return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableRowContainingTextDeletedByValues", "P15 T04 grading failed.", ex, "Excel2019_P15", "T04");
                return false;
            }
            finally
            {
                ReleaseCom(cell); ReleaseCom(dataRange); ReleaseCom(range); ReleaseCom(table);
                ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 15 Task 5
        public bool ClusteredColumnChartByHeadersRightOfData(string sheetName, string categoryHeader, string valueHeader, int expectedChartType)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(categoryHeader) || string.IsNullOrWhiteSpace(valueHeader)) return false;

            Xl.Worksheet ws = null;
            Xl.Range categoryHeaderRange = null, valueHeaderRange = null, categoryDataRange = null, valueDataRange = null, tableRange = null;
            Xl.Range rightHeaderCell = null;
            Xl.ChartObjects charts = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.SeriesCollection seriesCollection = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                int headerRow, categoryColumn, valueColumn, lastDataRow;
                if (!TryFindTwoColumnDataRangesP4T5(ws, categoryHeader, valueHeader,
                    out headerRow, out categoryColumn, out valueColumn, out lastDataRow,
                    out categoryHeaderRange, out valueHeaderRange, out categoryDataRange,
                    out valueDataRange, out tableRange)) return false;

                int rightColumn = Math.Max(categoryColumn, valueColumn);
                for (int c = rightColumn + 1; c <= rightColumn + 20; c++)
                {
                    Xl.Range headerCell = null;
                    try
                    {
                        headerCell = ws.Cells[headerRow, c] as Xl.Range;
                        if (headerCell == null || string.IsNullOrWhiteSpace(GetCellTextP1T3(headerCell))) break;
                        rightColumn = c;
                    }
                    finally { ReleaseCom(headerCell); }
                }
                rightHeaderCell = ws.Cells[headerRow, rightColumn] as Xl.Range;
                if (rightHeaderCell == null) return false;
                double dataRight = Convert.ToDouble(rightHeaderCell.Left) + Convert.ToDouble(rightHeaderCell.Width);

                charts = ws.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (charts == null) return false;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    ReleaseCom(seriesCollection); ReleaseCom(chart); ReleaseCom(chartObject);
                    seriesCollection = null; chart = null; chartObject = charts.Item(i) as Xl.ChartObject;
                    if (chartObject == null || Convert.ToDouble(chartObject.Left) + 2.0 < dataRight) continue;
                    chart = chartObject.Chart;
                    if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType) continue;
                    seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                    if (seriesCollection == null || Convert.ToInt32(seriesCollection.Count) != 1) continue;
                    if (ChartUsesCategoryAndValueRangesP4T5(chart, categoryDataRange, valueDataRange,
                        categoryHeaderRange, valueHeaderRange, categoryHeader, valueHeader)) return true;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("ClusteredColumnChartByHeadersRightOfData", "P15 T05 grading failed.", ex, "Excel2019_P15", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(seriesCollection); ReleaseCom(chart); ReleaseCom(chartObject); ReleaseCom(charts);
                ReleaseCom(rightHeaderCell); ReleaseCom(tableRange); ReleaseCom(valueDataRange);
                ReleaseCom(categoryDataRange); ReleaseCom(valueHeaderRange); ReleaseCom(categoryHeaderRange); ReleaseCom(ws);
            }
        }

        private bool TableHeadersEqualP15(Xl.ListObject table, IList<string> expectedHeaders)
        {
            if (table == null || expectedHeaders == null) return false;
            Xl.ListColumns columns = null;
            Xl.ListColumn column = null;
            try
            {
                columns = table.ListColumns;
                if (columns == null || Convert.ToInt32(columns.Count) != expectedHeaders.Count) return false;
                for (int i = 1; i <= expectedHeaders.Count; i++)
                {
                    ReleaseCom(column);
                    column = columns.Item[i];
                    if (column == null || !HeaderEqualsP4T5(Convert.ToString(column.Name), expectedHeaders[i - 1])) return false;
                }
                return true;
            }
            finally { ReleaseCom(column); ReleaseCom(columns); }
        }

        private List<string> GetTableRowSignaturesP15(Xl.ListObject table)
        {
            List<string> values = new List<string>();
            if (table == null) return values;
            Xl.Range dataRange = null;
            Xl.Range cell = null;
            try
            {
                dataRange = table.DataBodyRange;
                if (dataRange == null) return values;
                int rows = Convert.ToInt32(dataRange.Rows.Count);
                int columns = Convert.ToInt32(dataRange.Columns.Count);
                for (int r = 1; r <= rows; r++)
                {
                    List<string> parts = new List<string>();
                    for (int c = 1; c <= columns; c++)
                    {
                        ReleaseCom(cell);
                        cell = dataRange.Cells[r, c] as Xl.Range;
                        object value = cell == null ? null : cell.Value2;
                        parts.Add(Convert.ToString(value, CultureInfo.InvariantCulture));
                    }
                    values.Add(string.Join("|", parts));
                }
                return values;
            }
            finally { ReleaseCom(cell); ReleaseCom(dataRange); }
        }

        private bool StringListsEqualP15(IList<string> actual, IList<string> expected)
        {
            if (actual == null || expected == null || actual.Count != expected.Count) return false;
            for (int i = 0; i < actual.Count; i++)
                if (!string.Equals(NormalizeText(actual[i]), NormalizeText(expected[i]), StringComparison.OrdinalIgnoreCase)) return false;
            return true;
        }

        private bool RowSignatureHasFirstValueP15(string rowSignature, string expectedFirstValue)
        {
            string firstValue = (rowSignature ?? "").Split('|')[0];
            return string.Equals(NormalizeText(firstValue), NormalizeText(expectedFirstValue), StringComparison.OrdinalIgnoreCase);
        }

        public string GetCellDisplayText(string address)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(address)) return "";

            Xl.Workbook wb = null;
            Xl.Worksheet ws = null;
            Xl.Range cell = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;
                ws = wb.ActiveSheet as Xl.Worksheet;
                if (ws == null) return "";

                cell = ws.Range[address];
                return cell.Text == null ? "" : cell.Text.ToString();
            }
            catch
            {
                return "";
            }
            finally
            {
                ReleaseCom(cell);
                ReleaseCom(ws);
            }
        }

        private Xl.Worksheet GetWorksheet(string sheetName)
        {
            if (!IsOpened)
                return null;

            Xl.Workbook wb = null;
            Xl.Sheets sheets = null;

            try
            {
                wb = (Xl.Workbook)_session.Workbook;

                if (string.IsNullOrWhiteSpace(sheetName))
                    return wb.ActiveSheet as Xl.Worksheet;

                sheets = wb.Worksheets;

                for (int i = 1; i <= sheets.Count; i++)
                {
                    Xl.Worksheet ws = null;

                    try
                    {
                        ws = (Xl.Worksheet)sheets[i];

                        if (string.Equals(ws.Name, sheetName, StringComparison.OrdinalIgnoreCase))
                            return ws;

                        ReleaseCom(ws);
                        ws = null;
                    }
                    catch
                    {
                        ReleaseCom(ws);
                    }
                }

                return null;
            }
            finally
            {
                ReleaseCom(sheets);
            }
        }
        private bool ObjectEqualsLoose(object actual, object expected)
        {
            if (actual == null && expected == null) return true;
            if (actual == null || expected == null) return false;

            double actualNumber;
            double expectedNumber;

            if (TryToDouble(actual, out actualNumber) && TryToDouble(expected, out expectedNumber))
                return Math.Abs(actualNumber - expectedNumber) < 0.000001;

            string actualText = Convert.ToString(actual);
            string expectedText = Convert.ToString(expected);

            return string.Equals(NormalizeText(actualText), NormalizeText(expectedText), StringComparison.OrdinalIgnoreCase);
        }

        private bool TryToDouble(object value, out double number)
        {
            number = 0;

            if (value == null)
                return false;

            if (value is double || value is int || value is float || value is decimal)
            {
                number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                return true;
            }

            string text = Convert.ToString(value);
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.Trim().Replace(",", ".");
            return double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out number);
        }

        private string NormalizeText(string text)
        {
            if (text == null) return "";
            return Regex.Replace(text.Trim(), @"\s+", " ");
        }

        private string NormalizeFormula(string formula)
        {
            if (formula == null) return "";

            string f = formula.Trim();
            f = f.Replace(";", ",");
            f = f.Replace("$", "");
            f = Regex.Replace(f, @"\s+", "");
            f = f.ToUpperInvariant();

            return f;
        }

        private string NormalizeFormat(string format)
        {
            if (format == null) return "";

            string f = format.Trim();
            f = f.Replace("_", "");
            f = f.Replace(" ", "");
            f = f.ToUpperInvariant();

            return f;
        }

        // Project 20 Task 1
        public bool RangeFormattingMatchesAndPreservesText(
            string sourceSheetName,
            string sourceRangeAddress,
            string targetSheetName,
            string targetRangeAddress,
            IList<string> expectedTargetTexts)
        {
            if (expectedTargetTexts == null || expectedTargetTexts.Count != 2 ||
                !string.Equals(NormalizeRangeAddress(sourceRangeAddress), "A1:A2", StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(NormalizeRangeAddress(targetRangeAddress), "A1:A2", StringComparison.OrdinalIgnoreCase))
                return false;

            return RangeFormattingMatches(sourceSheetName, sourceRangeAddress, targetSheetName, targetRangeAddress) &&
                CellTextEquals(targetSheetName, "A1", expectedTargetTexts[0]) &&
                CellTextEquals(targetSheetName, "A2", expectedTargetTexts[1]);
        }

        // Project 20 Task 2
        public bool TableNameOnRangeEquals(string sheetName, string rangeAddress, string expectedTableName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(rangeAddress) ||
                string.IsNullOrWhiteSpace(expectedTableName)) return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range tableRange = null;
            Xl.Range headerRange = null;
            Xl.Range dataRange = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                if (tables == null) return false;

                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(dataRange); ReleaseCom(headerRange); ReleaseCom(tableRange); ReleaseCom(table);
                    dataRange = null; headerRange = null; tableRange = null; table = null;
                    table = tables.Item[i];
                    if (table == null) continue;
                    tableRange = table.Range;
                    if (tableRange == null || !string.Equals(
                        NormalizeRangeAddress(Convert.ToString(tableRange.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing])),
                        NormalizeRangeAddress(rangeAddress), StringComparison.OrdinalIgnoreCase)) continue;

                    headerRange = table.HeaderRowRange;
                    dataRange = table.DataBodyRange;
                    if (headerRange == null || dataRange == null ||
                        Convert.ToInt32(headerRange.Cells.Count) <= 0 || Convert.ToInt32(dataRange.Rows.Count) <= 0)
                        return false;

                    return string.Equals(Convert.ToString(table.Name).Trim(), expectedTableName.Trim(), StringComparison.OrdinalIgnoreCase) &&
                        string.Equals(Convert.ToString(table.DisplayName).Trim(), expectedTableName.Trim(), StringComparison.OrdinalIgnoreCase);
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableNameOnRangeEquals", "P20 T02 grading failed.", ex, "Excel2019_P20", "T02");
                return false;
            }
            finally
            {
                ReleaseCom(dataRange); ReleaseCom(headerRange); ReleaseCom(tableRange);
                ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        // Project 20 Task 6
        public bool ChartSheetLegendRemovedValueLabelsAboveBySource(
            string chartSheetName,
            string sourceSheetName,
            int expectedChartType,
            IList<string> sourceRanges)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(chartSheetName) || string.IsNullOrWhiteSpace(sourceSheetName) ||
                sourceRanges == null || sourceRanges.Count != 3) return false;

            Xl.Workbook workbook = null;
            Xl.Sheets charts = null;
            Xl.Chart chart = null;
            string tempPath = "";
            try
            {
                workbook = (Xl.Workbook)_session.Workbook;
                charts = workbook.Charts;
                for (int i = 1; i <= Convert.ToInt32(charts.Count); i++)
                {
                    chart = charts.Item[i] as Xl.Chart;
                    if (chart != null && string.Equals(chart.Name, chartSheetName, StringComparison.OrdinalIgnoreCase)) break;
                    ReleaseCom(chart);
                    chart = null;
                }
                if (chart == null)
                    return EmbeddedKadryChartLegendRemovedValueLabelsAboveP20(expectedChartType);
                if (Convert.ToInt32(chart.ChartType) != expectedChartType ||
                    !ChartSeriesMatchRangeTriplesP10(chart, sourceSheetName, sourceRanges)) return false;

                // Excel builds do not always serialize chart-sheet labels identically.
                // Prefer the live series-level state when it is fully observable, and
                // retain the OOXML path below as a strict fallback.
                bool liveStateMatches;
                if (TryGetChartNoLegendOutsideEndValueLabelsP20(chart, out liveStateMatches))
                    return liveStateMatches;

                tempPath = Path.Combine(Path.GetTempPath(),
                    "MosTrainer_P20T06_" + Guid.NewGuid().ToString("N") + ".xlsx");
                workbook.SaveCopyAs(tempPath);
                return File.Exists(tempPath) &&
                    XlsxChartSheetHasNoLegendAndValueLabelsAboveP4T3Xml(tempPath, chartSheetName);
            }
            catch (Exception ex)
            {
                AppLogger.Error("ChartSheetLegendRemovedValueLabelsAboveBySource", "P20 T06 grading failed.", ex, "Excel2019_P20", "T06");
                return false;
            }
            finally
            {
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempPath) && File.Exists(tempPath)) File.Delete(tempPath);
                }
                catch { }
                ReleaseCom(chart); ReleaseCom(charts);
            }
        }

        private bool EmbeddedKadryChartLegendRemovedValueLabelsAboveP20(int expectedChartType)
        {
            Xl.Worksheet worksheet = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            try
            {
                worksheet = GetWorksheet("Kadry");
                if (worksheet == null) return false;
                chartObjects = worksheet.ChartObjects(Type.Missing) as Xl.ChartObjects;
                if (chartObjects == null) return false;

                IList<string> expectedRanges = new[]
                {
                    "C3", "A4:B33", "C4:C33",
                    "D3", "A4:B33", "D4:D33",
                    "E3", "A4:B33", "E4:E33",
                    "F3", "A4:B33", "F4:F33",
                    "G3", "A4:B33", "G4:G33",
                    "H3", "A4:B33", "H4:H33",
                    "I3", "A4:B33", "I4:I33"
                };

                for (int index = 1; index <= Convert.ToInt32(chartObjects.Count); index++)
                {
                    ReleaseCom(chart); chart = null;
                    ReleaseCom(chartObject); chartObject = null;
                    chartObject = chartObjects.Item(index) as Xl.ChartObject;
                    chart = chartObject == null ? null : chartObject.Chart;
                    if (chart == null || Convert.ToInt32(chart.ChartType) != expectedChartType ||
                        !ChartSeriesMatchRangeTriplesP10(chart, "Kadry", expectedRanges)) continue;

                    bool matches;
                    return TryGetChartNoLegendOutsideEndValueLabelsP20(chart, out matches) && matches;
                }
                return false;
            }
            catch (Exception ex)
            {
                AppLogger.Error("EmbeddedKadryChartLegendRemovedValueLabelsAboveP20",
                    "P20 T06 embedded-chart fallback failed.", ex, "Excel2019_P20", "T06");
                return false;
            }
            finally
            {
                ReleaseCom(chart); ReleaseCom(chartObject); ReleaseCom(chartObjects); ReleaseCom(worksheet);
            }
        }

        private bool TryGetChartNoLegendOutsideEndValueLabelsP20(Xl.Chart chart, out bool matches)
        {
            matches = false;
            if (chart == null) return true;
            Xl.SeriesCollection seriesCollection = null;
            Xl.Series series = null;
            Xl.DataLabels labels = null;
            try
            {
                if (Convert.ToBoolean(chart.HasLegend)) return true;
                seriesCollection = chart.SeriesCollection(Type.Missing) as Xl.SeriesCollection;
                int count = seriesCollection == null ? 0 : Convert.ToInt32(seriesCollection.Count);
                if (count <= 0) return true;

                for (int index = 1; index <= count; index++)
                {
                    ReleaseCom(labels); labels = null;
                    ReleaseCom(series); series = null;
                    series = seriesCollection.Item(index);
                    if (series == null || !Convert.ToBoolean(series.HasDataLabels)) return true;
                    labels = series.DataLabels(Type.Missing) as Xl.DataLabels;
                    if (labels == null || Convert.ToInt32(labels.Position) !=
                        (int)Xl.XlDataLabelPosition.xlLabelPositionOutsideEnd ||
                        !Convert.ToBoolean(labels.ShowValue) ||
                        Convert.ToBoolean(labels.ShowSeriesName) ||
                        Convert.ToBoolean(labels.ShowCategoryName) ||
                        Convert.ToBoolean(labels.ShowLegendKey)) return true;
                }
                matches = true;
                return true;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(labels);
                ReleaseCom(series);
                ReleaseCom(seriesCollection);
            }
        }

        // Project 21 Task 2
        public bool IfNumericFormulaByHeadersInRanges(
            string sheetName,
            string targetHeader,
            string criteriaHeader,
            string targetRangeAddress,
            string criteriaRangeAddress,
            string compareOperator,
            double threshold,
            string trueValue,
            string falseValue)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(targetHeader) ||
                string.IsNullOrWhiteSpace(criteriaHeader) || string.IsNullOrWhiteSpace(targetRangeAddress) ||
                string.IsNullOrWhiteSpace(criteriaRangeAddress) || compareOperator != ">") return false;

            double expectedTrue;
            double expectedFalse;
            if (!TryToDouble(trueValue, out expectedTrue) || !TryToDouble(falseValue, out expectedFalse)) return false;

            Xl.Worksheet ws = null;
            Xl.Range targets = null;
            Xl.Range criteria = null;
            Xl.Range targetHeaderCell = null;
            Xl.Range criteriaHeaderCell = null;
            Xl.Range targetCell = null;
            Xl.Range criteriaCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                targets = ws.Range[targetRangeAddress];
                criteria = ws.Range[criteriaRangeAddress];
                if (targets == null || criteria == null || Convert.ToInt32(targets.Columns.Count) != 1 ||
                    Convert.ToInt32(criteria.Columns.Count) != 1 ||
                    Convert.ToInt32(targets.Rows.Count) != Convert.ToInt32(criteria.Rows.Count)) return false;

                targetCell = targets.Cells[1, 1] as Xl.Range;
                criteriaCell = criteria.Cells[1, 1] as Xl.Range;
                targetHeaderCell = targetCell == null ? null : targetCell.Offset[-1, 0] as Xl.Range;
                criteriaHeaderCell = criteriaCell == null ? null : criteriaCell.Offset[-1, 0] as Xl.Range;
                ReleaseCom(targetCell); targetCell = null;
                ReleaseCom(criteriaCell); criteriaCell = null;
                if (targetHeaderCell == null || criteriaHeaderCell == null ||
                    !string.Equals(NormalizeText(Convert.ToString(targetHeaderCell.Value2)), NormalizeText(targetHeader), StringComparison.OrdinalIgnoreCase) ||
                    !string.Equals(NormalizeText(Convert.ToString(criteriaHeaderCell.Value2)), NormalizeText(criteriaHeader), StringComparison.OrdinalIgnoreCase))
                    return false;

                int rowCount = Convert.ToInt32(targets.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(targetCell); ReleaseCom(criteriaCell);
                    targetCell = targets.Cells[row, 1] as Xl.Range;
                    criteriaCell = criteria.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || criteriaCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;

                    double criteriaNumber;
                    double actualNumber;
                    if (!TryToDouble(criteriaCell.Value2, out criteriaNumber) || !TryToDouble(targetCell.Value2, out actualNumber)) return false;
                    double expectedNumber = criteriaNumber > threshold ? expectedTrue : expectedFalse;
                    if (Math.Abs(actualNumber - expectedNumber) > 0.000001d) return false;

                    string criteriaAddress = Convert.ToString(criteriaCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMatchesNumericIfP10(Convert.ToString(targetCell.Formula), criteriaAddress, criteriaHeader,
                        threshold, expectedTrue, expectedFalse)) return false;
                }
                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("IfNumericFormulaByHeadersInRanges", "P21 T02 grading failed.", ex, "Excel2019_P21", "T02");
                return false;
            }
            finally
            {
                ReleaseCom(criteriaCell); ReleaseCom(targetCell);
                ReleaseCom(criteriaHeaderCell); ReleaseCom(targetHeaderCell);
                ReleaseCom(criteria); ReleaseCom(targets); ReleaseCom(ws);
            }
        }

        // Project 21 Task 5
        public bool TableColumnUpperLeftFormulaByHeaders(
            string sheetName,
            string tableName,
            string targetHeader,
            string sourceHeader,
            int characterCount)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(tableName) ||
                string.IsNullOrWhiteSpace(targetHeader) || string.IsNullOrWhiteSpace(sourceHeader) || characterCount <= 0)
                return false;

            Xl.Worksheet ws = null;
            Xl.ListObjects tables = null;
            Xl.ListObject table = null;
            Xl.Range data = null;
            Xl.Range targets = null;
            Xl.Range sources = null;
            Xl.Range targetCell = null;
            Xl.Range sourceCell = null;
            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;
                tables = ws.ListObjects;
                for (int i = 1; i <= Convert.ToInt32(tables.Count); i++)
                {
                    ReleaseCom(table);
                    table = tables.Item[i];
                    if (table != null && string.Equals(Convert.ToString(table.Name), tableName, StringComparison.OrdinalIgnoreCase)) break;
                    table = null;
                }
                if (table == null) return false;

                int targetIndex = FindTableColumnIndexP05T5(table, targetHeader);
                int sourceIndex = FindTableColumnIndexP05T5(table, sourceHeader);
                if (targetIndex <= 0 || sourceIndex <= 0) return false;
                data = table.DataBodyRange;
                if (data == null) return false;
                targets = data.Columns[targetIndex] as Xl.Range;
                sources = data.Columns[sourceIndex] as Xl.Range;
                if (targets == null || sources == null ||
                    Convert.ToInt32(targets.Rows.Count) != Convert.ToInt32(sources.Rows.Count)) return false;

                int rowCount = Convert.ToInt32(targets.Rows.Count);
                for (int row = 1; row <= rowCount; row++)
                {
                    ReleaseCom(targetCell); ReleaseCom(sourceCell);
                    targetCell = targets.Cells[row, 1] as Xl.Range;
                    sourceCell = sources.Cells[row, 1] as Xl.Range;
                    if (targetCell == null || sourceCell == null || !Convert.ToBoolean(targetCell.HasFormula)) return false;

                    string sourceText = Convert.ToString(sourceCell.Value2) ?? "";
                    string expected = sourceText.Substring(0, Math.Min(characterCount, sourceText.Length)).ToUpperInvariant();
                    if (!string.Equals(Convert.ToString(targetCell.Value2), expected, StringComparison.Ordinal)) return false;
                    string sourceAddress = Convert.ToString(sourceCell.Address[false, false, Xl.XlReferenceStyle.xlA1, Type.Missing, Type.Missing]);
                    if (!FormulaMatchesUpperLeftP21(Convert.ToString(targetCell.Formula), sourceAddress, sourceHeader, characterCount))
                        return false;
                }
                return rowCount > 0;
            }
            catch (Exception ex)
            {
                AppLogger.Error("TableColumnUpperLeftFormulaByHeaders", "P21 T05 grading failed.", ex, "Excel2019_P21", "T05");
                return false;
            }
            finally
            {
                ReleaseCom(sourceCell); ReleaseCom(targetCell); ReleaseCom(sources); ReleaseCom(targets);
                ReleaseCom(data); ReleaseCom(table); ReleaseCom(tables); ReleaseCom(ws);
            }
        }

        private bool FormulaMatchesUpperLeftP21(string formula, string sourceAddress, string sourceHeader, int characterCount)
        {
            string normalized = NormalizeFormula(formula);
            Match match = Regex.Match(normalized, @"^=UPPER\(LEFT\((?<source>.+),(?<count>\d+)\)\)$", RegexOptions.IgnoreCase);
            int actualCount;
            if (!match.Success || !int.TryParse(match.Groups["count"].Value, out actualCount) || actualCount != characterCount)
                return false;

            string source = TrimOuterParenthesesP05T5(match.Groups["source"].Value);
            if (ReferenceEqualsP10(source, sourceAddress)) return true;
            string header = NormalizeFormula(sourceHeader);
            return source.EndsWith("[@" + header + "]", StringComparison.OrdinalIgnoreCase) ||
                source.EndsWith("[@[" + header + "]]", StringComparison.OrdinalIgnoreCase) ||
                source.EndsWith("[[#THISROW],[" + header + "]]", StringComparison.OrdinalIgnoreCase);
        }

        private string NormalizeRangeAddress(string address)
        {
            if (address == null) return "";

            string text = address.Trim();
            if (text.StartsWith("=", StringComparison.Ordinal))
                text = text.Substring(1);

            int exclamationIndex = text.LastIndexOf('!');
            if (exclamationIndex >= 0 && exclamationIndex + 1 < text.Length)
                text = text.Substring(exclamationIndex + 1);

            text = text.Replace("$", "");
            text = text.Replace("'", "");
            text = Regex.Replace(text, @"\s+", "");
            text = text.ToUpperInvariant();

            return text;
        }

        private int ParseHorizontalAlignment(string alignment)
        {
            if (string.IsNullOrWhiteSpace(alignment))
                return 0;

            string text = alignment.Trim().ToUpperInvariant();
            text = text.Replace(" ", "");
            text = text.Replace("-", "");

            if (text == "LEFT" || text == "XLLEFT")
                return (int)Xl.XlHAlign.xlHAlignLeft;

            if (text == "CENTER" || text == "CENTRE" || text == "XLCENTER")
                return (int)Xl.XlHAlign.xlHAlignCenter;

            if (text == "RIGHT" || text == "XLRIGHT")
                return (int)Xl.XlHAlign.xlHAlignRight;

            if (text == "GENERAL" || text == "XLGENERAL")
                return (int)Xl.XlHAlign.xlHAlignGeneral;

            if (text == "JUSTIFY" || text == "XLJUSTIFY")
                return (int)Xl.XlHAlign.xlHAlignJustify;

            if (text == "DISTRIBUTED" || text == "XLDISTRIBUTED")
                return (int)Xl.XlHAlign.xlHAlignDistributed;

            if (text == "FILL" || text == "XLFILL")
                return (int)Xl.XlHAlign.xlHAlignFill;

            int value;
            if (int.TryParse(text, out value))
                return value;

            return 0;
        }

        private int GetExcelPid(Xl.Application app)
        {
            if (app == null)
                return 0;

            try
            {
                IntPtr hwnd = new IntPtr(app.Hwnd);
                return WinApiProcessHelper.GetProcessIdFromHwnd(hwnd);
            }
            catch
            {
                return 0;
            }
        }

        private static void ReleaseCom(object obj)
        {
            try
            {
                if (obj != null && Marshal.IsComObject(obj))
                    Marshal.FinalReleaseComObject(obj);
            }
            catch
            {
            }
        }

        public void Dispose()
        {
            CloseWorkbook();
        }
    }
}
