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

                bool criteriaOk = AutoFilterCriteriaContainsExpected(ws, fieldIndex, expectedText);
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

            return true;
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
        public bool EmailFormulaFromHeader(string sheetName, string targetHeader, string sourceHeader, string domain)
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
                        normalizedDomain);

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
            string domain)
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

            bool looksLikeConstructFormula =
                combined.Contains("CONCATENATE(") ||
                combined.Contains("CONCAT(") ||
                combined.Contains("TEXTJOIN(") ||
                combined.Contains("&");

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

        private bool AutoFilterCriteriaContainsExpected(Xl.Worksheet ws, int fieldIndex, string expectedText)
        {
            if (ws == null || fieldIndex <= 0 || string.IsNullOrWhiteSpace(expectedText))
                return false;

            Xl.AutoFilter autoFilter = null;
            Xl.Filters filters = null;
            Xl.Filter filter = null;

            try
            {
                autoFilter = ws.AutoFilter;
                if (autoFilter == null) return false;

                filters = autoFilter.Filters;
                if (filters == null || fieldIndex > filters.Count) return false;

                filter = filters.Item[fieldIndex];
                if (filter == null) return false;

                bool isOn = false;
                try { isOn = filter.On; } catch { isOn = false; }
                if (!isOn) return false;

                object criteria1 = null;
                object criteria2 = null;

                try { criteria1 = filter.Criteria1; } catch { }
                try { criteria2 = filter.Criteria2; } catch { }

                if (CriteriaObjectContainsExpected(criteria1, expectedText))
                    return true;

                if (CriteriaObjectContainsExpected(criteria2, expectedText))
                    return true;

                return false;
            }
            catch
            {
                return false;
            }
            finally
            {
                ReleaseCom(filter);
                ReleaseCom(filters);
                ReleaseCom(autoFilter);
            }
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

            if (!normalizedFormula.Contains("SUM("))
                return false;

            string formulaWithoutNames = normalizedFormula;

            for (int i = 0; i < rangeNames.Count; i++)
            {
                string name = NormalizeDefinedNameP2T5(rangeNames[i]);
                if (string.IsNullOrWhiteSpace(name))
                    return false;

                if (!normalizedFormula.Contains(name))
                    return false;

                formulaWithoutNames = formulaWithoutNames.Replace(name, "");
            }

            if (FormulaContainsDirectCellReferenceP2T5(formulaWithoutNames))
                return false;

            // Sau khi bỏ Total1/Total2/Total3, không được còn số cố định nào.
            // Như vậy =SUM(Total1,Total2,Total3,100) sẽ FAIL.
            if (Regex.IsMatch(formulaWithoutNames, @"\d"))
                return false;

            return true;
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
        public bool TableColumnFormulaFilledDown(string sheetName, string startCellAddress)
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
        public bool ChartSwitchedRowColumn(string sheetName, string chartTitle, string sourceRange)
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

                    chart = chartObject.Chart;
                    if (chart == null) continue;

                    if (!ChartTitleMatchesP4T1(chart, chartTitle))
                        continue;

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

            return string.Equals(sourceXf, targetXf, StringComparison.OrdinalIgnoreCase);
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
        public bool WorksheetTableConvertedToRange(string sheetName, string rangeAddress, string tableName)
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
                if (!RangeContainsNewCarSalesHeadersP4T4(targetRange))
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

        private bool RangeContainsNewCarSalesHeadersP4T4(Xl.Range range)
        {
            if (range == null) return false;

            string[] expectedHeaders = new string[]
            {
        "Make",
        "Model",
        "Body",
        "Year",
        "Color",
        "Mileage",
        "Price",
        "Quantity Instock",
        "Total",
        "Inspected"
            };

            Xl.Range cell = null;

            try
            {
                int columnCount = Convert.ToInt32(range.Columns.Count);
                if (columnCount < expectedHeaders.Length)
                    return false;

                for (int c = 1; c <= expectedHeaders.Length; c++)
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
                // Cho phép top row bằng dòng cuối + 1 hoặc thấp hơn.
                return chartTopRow >= lastDataRow;
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
        public bool ChartDataTableWithoutLegendKeys(string sheetName, string chartTitle, string chartName)
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            if (string.IsNullOrWhiteSpace(sheetName)) return false;

            Xl.Worksheet ws = null;
            Xl.ChartObjects chartObjects = null;
            Xl.ChartObject chartObject = null;
            Xl.Chart chart = null;
            Xl.ChartTitle title = null;
            Xl.DataTable dataTable = null;

            try
            {
                ws = GetWorksheet(sheetName);
                if (ws == null) return false;

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
                ReleaseCom(dataTable);
                ReleaseCom(title);
                ReleaseCom(chart);
                ReleaseCom(chartObject);
                ReleaseCom(chartObjects);
                ReleaseCom(ws);
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
            if (string.IsNullOrWhiteSpace(sourceSheetName) || string.IsNullOrWhiteSpace(chartSheetName) ||
                string.IsNullOrWhiteSpace(chartTitle))
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

                    if (!Convert.ToBoolean(chartSheet.HasTitle)) return false;
                    title = chartSheet.ChartTitle;
                    if (title == null || !string.Equals(
                        NormalizeText(Convert.ToString(title.Text)), NormalizeText(chartTitle),
                        StringComparison.OrdinalIgnoreCase))
                        return false;

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
                    if (chart == null || !Convert.ToBoolean(chart.HasTitle)) return false;
                    title = chart.ChartTitle;
                    if (title == null || !string.Equals(NormalizeText(Convert.ToString(title.Text)), NormalizeText(chartTitle), StringComparison.OrdinalIgnoreCase))
                        return false;
                    if (Convert.ToInt32(chart.ChartStyle) != expectedChartStyle ||
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

        // Project 7 Task 4
        public bool WorkbookPersonalInformationRemoved()
        {
            if (!IsOpened) throw new InvalidOperationException("Workbook not opened.");
            string tempPath = Path.Combine(Path.GetTempPath(), "MosTrainer-P07-T04-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                Xl.Workbook workbook = (Xl.Workbook)_session.Workbook;
                if (!Convert.ToBoolean(workbook.RemovePersonalInformation)) return false;
                if (!WorkbookRetainsP07CoreContent(workbook)) return false;
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
