using System;
using System.Collections.Generic;

namespace MosTrainer.Core.Interfaces
{
    public interface IExcelController
    {
        bool IsOpened { get; }

        void StartExcel();
        void OpenWorkbook(string filePath);
        void CloseWorkbook();

        bool WorksheetExists(string sheetName);
        bool CellValueEquals(string sheetName, string address, object expectedValue);
        bool CellTextEquals(string sheetName, string address, string expectedText);
        bool CellFormulaEquals(string sheetName, string address, string expectedFormula);
        bool CellFormulaEqualsNormalized(string sheetName, string address, string expectedFormula);

        bool RangeNumberFormatEquals(string sheetName, string address, string expectedFormat);
        bool RangeFillColorEquals(string sheetName, string address, int expectedColor);
        bool RangeFontBoldEquals(string sheetName, string address, bool expectedBold);
        bool RangeHorizontalAlignmentEquals(string sheetName, string address, string expectedAlignment);

        bool TableExists(string sheetName, string tableName);
        bool ChartExists(string sheetName);
        bool NamedRangeExists(string name);

        // Excel 2019 - Project 1 Task 1
        bool PrintAreaEquals(string sheetName, string expectedRange);
        // Excel 2019 - Project 1 Task 2
        bool AutoFilterByHeaderEquals(string sheetName, string columnHeader, string expectedText);
        // Excel 2019 - Project 1 Task 3
        bool IfFormulaByHeaders(string sheetName,string targetHeader,string criteriaHeader,string compareOperator,double threshold,string trueText,string falseText);
        // Excel 2019 - Project 1 Task 4
        bool MultiLevelSortByHeaders(string sheetName, IList<string> sortHeaders, IList<string> sortOrders);

        // Excel 2019 - Project 1 Task 5
        bool EmailFormulaFromHeader(string sheetName, string targetHeader, string sourceHeader, string domain);
        // Excel 2019 - Project 1 Task 6
        bool TableBandedRows(string sheetName);

        // Excel 2019 - Project 1 Task 7
        bool ChartSheetExists(string chartSheetName, string sourceSheetName);
        // Excel 2019 - Project 1 Task 8
        bool NoConditionalFormatting(string sheetName);
        // Excel 2019 - Project 2 Task 1
        bool ImportedCsvAtCell(string projectId,string sheetName,string startCell,string sourceFileName,bool firstRowAsHeaders);
        // Excel 2019 - Project 2 Task 2
        bool ColumnWidthEquals(string sheetName, string columnRange, double expectedWidth);

        // Excel 2019 - Project 2 Task 3
        bool ColumnSparklinesByRange(string sheetName, string locationRange, string dataRange);
        // Excel 2019 - Project 2 Task 4
        bool FreezePanesEquals(string sheetName, int expectedFreezeRows, int expectedFreezeColumns);
        // Excel 2019 - Project 2 Task 5
        bool NamedRangesSumFormula(string sheetName, string cellAddress, IList<string> rangeNames);
        // Excel 2019 - Project 2 Task 6
        bool ChartAltTextEquals(string sheetName, string expectedText);
        // Excel 2019 - Project 2 Task 7
        bool TableColumnIconSetEquals(string sheetName, string columnHeader, string expectedIconSetName);
        // Excel 2019 - Project 2 Task 7 - fixed by range
        bool RangeIconSetEquals(string sheetName, string rangeAddress, string expectedIconSetName);
        // Excel 2019 - Project 2 Task 8
        bool TableStyleEquals(string sheetName, string tableName, string expectedStyleName);
        // Excel 2019 - Project 3 Task 1
        bool NamedRangeContentsCleared(string namedRangeName, string expectedSheetName, string expectedAddress);

        // Excel 2019 - Project 3 Task 2
        bool RangeNumberFormatDecimalPlacesEquals(string sheetName, string rangeAddress, int expectedDecimalPlaces);
        // Excel 2019 - Project 3 Task 3
        bool TableRowContainingTextDeleted(string sheetName, string searchText, int expectedDataRowCount);

        // Excel 2019 - Project 3 Task 4
        bool AverageFormulaByHeaders(string sheetName, string targetHeader, IList<string> sourceHeaders);
        // Excel 2019 - Project 3 Task 5
        bool ChartPrimaryVerticalAxisTitleEquals(string sheetName, string expectedTitle);
        // Excel 2019 - Project 3 Task 7
        bool TableColumnFormulaFilledDown(string sheetName, string startCellAddress);

        // Excel 2019 - Project 3 Task 8
        bool MaxFormulaFromHeader(string sheetName, string targetCellAddress, string sourceHeader);
        // Excel 2019 - Project 4 Task 1
        bool ChartSwitchedRowColumn(string sheetName, string chartTitle, string sourceRange);

        // Excel 2019 - Project 4 Task 2
        bool RangeFormattingMatches(string sourceSheetName, string sourceRange, string targetSheetName, string targetRange);
        // Excel 2019 - Project 4 Task 3
        bool ChartSheetLegendRemovedValueLabelsAbove(string chartSheetName);

        // Excel 2019 - Project 4 Task 4
        bool WorksheetTableConvertedToRange(string sheetName, string rangeAddress, string tableName);
        // Excel 2019 - Project 4 Task 5
        bool ReportClusteredColumnChartCreated(string sheetName, string categoryHeader, string valueHeader);

        // Excel 2019 - Project 4 Task 6
        bool LeftFormulaByHeaders(string sheetName, string targetHeader, string sourceHeader, int characterCount);
        // Excel 2019 - Project 4 Task 7
        bool LastFirstNameFormulaAtCell(string sheetName, string targetCellAddress, string lastNameHeader, string firstNameHeader, string separator);
        // Excel 2019 - Project 4 Task 8
        bool CenterFooterPageOfPagesEquals(string sheetName);
        // Excel 2019 - Project 5 Task 1
        bool ShapeHyperlinkEquals(string sheetName, string shapeName, string topLeftCellAddress, string expectedUrl);
        // Excel 2019 - Project 5 Task 3
        bool ChartDataTableWithoutLegendKeys(string sheetName, string chartTitle, string chartName);
        // Excel 2019 - Project 5 Task 4
        bool SalesByExamTableConvertedToRange(string sheetName, string rangeAddress, string tableName, IList<string> expectedHeaders, int expectedDataRowCount);
        // Excel 2019 - Project 5 Task 5
        bool TableColumnFormulaMultipliesColumns(string sheetName, string tableName, string targetHeader, IList<string> sourceHeaders);
        // Excel 2019 - Project 5 Task 7
        bool RangesMergedExactly(string sheetName, IList<string> rangeAddresses);
        // Excel 2019 - Project 5 Task 8
        bool CellStylesApplied(string sheetName, IList<string> primaryRanges, string primaryStyleName, IList<string> secondaryRanges, string secondaryStyleName);
        // Excel 2019 - Project 6 Task 1
        bool ChartColorPaletteEquals(string sheetName, string chartName, string chartTitle, int expectedChartColor, int expectedChartType, IList<string> sourceRanges);
        // Excel 2019 - Project 6 Task 3
        bool RangeFormattingMatchesSourceCell(string sheetName, string sourceCellAddress, string targetRangeAddress, IList<string> expectedTargetTexts);
        // Excel 2019 - Project 6 Task 4
        bool WorkbookBuiltinPropertyEquals(string propertyName, string expectedValue);
        // Excel 2019 - Project 6 Task 5
        bool TableOnRangeWithStyle(string sheetName, string rangeAddress, string expectedStyleName, IList<string> expectedHeaders);
        // Excel 2019 - Project 6 Task 7
        bool RangeWrapTextEquals(string sheetName, string rangeAddress, bool expectedWrapText);
        // Excel 2019 - Project 6 Task 8
        bool ChartMovedToChartSheet(string sourceSheetName, string chartSheetName, string chartTitle, int expectedChartType, IList<string> sourceRanges);
        // Excel 2019 - Project 7 Task 1
        bool WorksheetShowFormulasEquals(string sheetName, bool expectedShowFormulas);
        // Excel 2019 - Project 7 Task 2
        bool InvoiceCellsDeletedShiftUp(string sheetName, string deletedRangeAddress);
        // Excel 2019 - Project 7 Task 3
        bool ChartStyleAndPaletteEquals(string sheetName, string chartName, string chartTitle, int expectedChartStyle, int expectedChartColor, int expectedChartType, IList<string> sourceRanges);
        // Excel 2019 - Project 7 Task 4
        bool WorkbookPersonalInformationRemoved();
        // Excel 2019 - Project 7 Task 5
        bool ClusteredColumnChartFromRanges(string sheetName, string tableName, string tableRangeAddress, int expectedChartType, IList<string> sourceRanges);
        // Excel 2019 - Project 7 Task 7
        bool IfFormulaByHeadersStrict(string sheetName, string tableName, string targetHeader, string criteriaHeader, string compareOperator, double threshold, string trueText, string falseText);
        // Excel 2019 - Project 7 Task 8
        bool NamedRangeRefersToRange(string name, string sheetName, string rangeAddress);
        // Excel 2019 - Project 8 Task 1
        bool TableColumnFormulaMultipliesNamedRange(string sheetName, string tableName, string targetHeader, string sourceHeader, string namedRange, string namedRangeAddress);
        // Excel 2019 - Project 8 Task 2
        bool TableRowContainingTextDeletedPreserveOutside(string sheetName, string tableName, string searchText, int expectedDataRowCount, string expectedTableRange, string preservedFilterRange, IList<string> expectedFirstColumnValues);
        // Excel 2019 - Project 8 Task 3
        bool RangeAlignmentIndentEquals(string sheetName, string rangeAddress, string expectedAlignment, int expectedIndent);
        // Excel 2019 - Project 8 Task 4
        bool SparklinesByRangeAndType(string sheetName, string locationRange, string dataRange, string expectedType);
        // Excel 2019 - Project 8 Task 5
        bool TableTotalRowSumsByHeaders(string sheetName, string tableName, IList<string> sumHeaders);
        // Excel 2019 - Project 8 Task 6
        bool CountBlankFormulaByHeaders(string sheetName, string tableName, string targetHeader, IList<string> sourceHeaders);
        // Excel 2019 - Project 8 Task 7
        bool TableMultiLevelSortStateEquals(string sheetName, string tableName, string tableRange, IList<string> sortHeaders, IList<string> sortOrders);
        // Excel 2019 - Project 8 Task 8
        bool ChartQuickLayoutEquals(string sheetName, string chartName, int expectedLayout, int expectedChartType, string seriesNameRange, string categoryRange, string valuesRange);
        string GetCellDisplayText(string address);
    }
}
