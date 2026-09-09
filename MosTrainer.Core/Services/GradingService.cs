using MosTrainer.Core.Interfaces;
using MosTrainer.Core.Models;
using System;

namespace MosTrainer.Core.Services
{
    public class GradingService
    {
        private readonly IExcelController _excel;

        public GradingService(IExcelController excel)
        {
            _excel = excel;
        }

        public (bool pass, string message) CheckTask(TaskDefinition task)
        {
            if (task == null)
                return (false, "Task is null.");

            if (!_excel.IsOpened)
                return (false, "Excel is not opened.");

            string projectId = (task.ProjectId ?? "").Trim();

            if (projectId.StartsWith("Excel2019_", StringComparison.OrdinalIgnoreCase))
                return CheckExcel2019Task(task);

            return (false, "Unsupported project: " + projectId);
        }

        private (bool pass, string message) CheckExcel2019Task(TaskDefinition task)
        {
            switch ((task.ProjectId ?? "").Trim())
            {
                case "Excel2019_P01":
                    return CheckExcel2019P01(task);
                case "Excel2019_P02":
                    return CheckExcel2019P02(task);
                case "Excel2019_P03":
                    return CheckExcel2019P03(task);
                case "Excel2019_P04":
                    return CheckExcel2019P04(task);

                default:
                    return (false, "Unsupported Excel 2019 project: " + task.ProjectId);
            }
        }

        private (bool pass, string message) CheckExcel2019P01(TaskDefinition task)
        {
            // P01 dùng assertionType để route từng task.
            // T01 hiện dùng PrintAreaEquals.
            return CheckByAssertion(task);
        }
        private (bool pass, string message) CheckExcel2019P02(TaskDefinition task)
        {
            // P02 cũng route theo assertionType để mỗi task độc lập, dễ mở rộng.
            return CheckByAssertion(task);
        }
        private (bool pass, string message) CheckExcel2019P03(TaskDefinition task)
        {
            return CheckByAssertion(task);
        }
        private (bool pass, string message) CheckExcel2019P04(TaskDefinition task)
        {
            return CheckByAssertion(task);
        }
        private (bool pass, string message) CheckByAssertion(TaskDefinition task)
        {
            string type = (task.AssertionType ?? "").Trim();

            switch (type)
            {
                case "WorksheetExists":
                    return Result(_excel.WorksheetExists(task.SheetName));

                case "CellValueEquals":
                    return Result(_excel.CellValueEquals(task.SheetName, task.Cell, task.ExpectedValue));

                case "CellTextEquals":
                    return Result(_excel.CellTextEquals(task.SheetName, task.Cell, task.ExpectedText));

                case "CellFormulaEquals":
                    return Result(_excel.CellFormulaEquals(task.SheetName, task.Cell, task.ExpectedFormula));

                case "CellFormulaEqualsNormalized":
                    return Result(_excel.CellFormulaEqualsNormalized(task.SheetName, task.Cell, task.ExpectedFormula));

                case "NumberFormatEquals":
                    return Result(_excel.RangeNumberFormatEquals(task.SheetName, GetAddress(task), task.ExpectedFormat));

                case "FillColorEquals":
                    return Result(_excel.RangeFillColorEquals(task.SheetName, GetAddress(task), ToInt(task.ExpectedValue)));

                case "FontBoldEquals":
                    return Result(_excel.RangeFontBoldEquals(task.SheetName, GetAddress(task), ToBool(task.ExpectedValue)));

                case "HorizontalAlignmentEquals":
                    return Result(_excel.RangeHorizontalAlignmentEquals(task.SheetName, GetAddress(task), task.ExpectedText));

                case "TableExists":
                    return Result(_excel.TableExists(task.SheetName, task.TableName));

                case "ChartExists":
                    return Result(_excel.ChartExists(task.SheetName));

                case "NamedRangeExists":
                    return Result(_excel.NamedRangeExists(task.NamedRange));
                //Project 1 task 1
                case "PrintAreaEquals":
                    return Result(_excel.PrintAreaEquals(task.SheetName, task.Range));
                //Project 1 task 2
                case "AutoFilterByHeaderEquals":
                    return Result(_excel.AutoFilterByHeaderEquals(task.SheetName, task.ColumnHeader, task.ExpectedText));
                //Project 1 task 3
                case "IfFormulaByHeaders":
                    return Result(_excel.IfFormulaByHeaders(task.SheetName,task.TargetHeader,task.CriteriaHeader,task.Operator,task.Threshold,task.TrueText,task.FalseText));
                //Project 1 task 4
                case "MultiLevelSortByHeaders":return Result(_excel.MultiLevelSortByHeaders(task.SheetName,task.SortHeaders,task.SortOrders));
                //Project 1 task 5
                case "EmailFormulaFromHeader":
                    return Result(_excel.EmailFormulaFromHeader(task.SheetName,task.TargetHeader,task.SourceHeader,task.Domain));
                //Project 1 task 6
                case "TableBandedRows":
                    return Result(_excel.TableBandedRows(task.SheetName));
                //Project 1 task 7
                case "ChartSheetExists":
                    return Result(_excel.ChartSheetExists(task.SheetName, task.SourceSheetName));
                //Project 1 task 8
                case "NoConditionalFormatting":
                    return Result(_excel.NoConditionalFormatting(task.SheetName));
                //Project 2 task 1
                case "ImportedCsvAtCell":
                case "ImportCsvAtCell":
                case "ImportTextFileAtCell":
                    return Result(_excel.ImportedCsvAtCell(task.ProjectId,task.SheetName,task.Cell,task.SourceFileName, task.FirstRowAsHeaders));
                //Project 2 task 2
                case "ColumnWidthEquals":
                case "ColumnWidthExactly":
                    return Result(_excel.ColumnWidthEquals(
                        task.SheetName,
                        task.Range,
                        task.ExpectedWidth));

                //Project 2 task 3
                case "ColumnSparklinesByRange":
                case "ColumnSparklines":
                case "SparklineColumnByRange":
                    return Result(_excel.ColumnSparklinesByRange(
                        task.SheetName,
                        string.IsNullOrWhiteSpace(task.LocationRange) ? task.Range : task.LocationRange,
                        task.DataRange));
                //Project 2 task 4
                case "FreezePanesEquals":
                case "FreezeRowsAndColumns":
                case "FreezeTopRows":
                case "FreezeRows":
                    return Result(_excel.FreezePanesEquals(
                        task.SheetName,
                        task.FreezeRows,
                        task.FreezeColumns));
                //Project 2 task 5
                case "NamedRangesSumFormula":
                case "SumNamedRangesFormula":
                case "NamedRangeSumFormula":
                    return Result(_excel.NamedRangesSumFormula(
                        task.SheetName,
                        task.Cell,
                        task.RangeNames));

                //Project 2 task 6
                case "ChartAltTextEquals":
                case "ChartAltTextDescriptionEquals":
                    return Result(_excel.ChartAltTextEquals(
                        task.SheetName,
                        task.ExpectedText));
                //Project 2 task 7
                case "RangeIconSetEquals":
                case "RangeTrafficLightsUnrimmed":
                case "RangeIconSet":
                case "TableColumnIconSetEquals":
                case "TableColumnIconSet":
                case "PriceColumnTrafficLights":
                case "TrafficLightsUnrimmed":
                    return Result(_excel.RangeIconSetEquals(
                        task.SheetName,
                        string.IsNullOrWhiteSpace(task.Range) ? "G10:G40" : task.Range,
                        task.ExpectedText));

                //Project 2 task 8
                case "TableStyleEquals":
                case "TableStyle":
                case "TableStyleMedium1":
                    return Result(_excel.TableStyleEquals(
                        task.SheetName,
                        task.TableName,
                        task.ExpectedText));
                //Project 3 task 1
                case "NamedRangeContentsCleared":
                case "NamedRangeContentDeleted":
                case "NamedRangeCleared":
                case "NamedRangeDeletedContents":
                    return Result(_excel.NamedRangeContentsCleared(
                        task.NamedRange,
                        task.SheetName,
                        task.Range));

                //Project 3 task 2
                case "RangeNumberFormatDecimalPlacesEquals":
                case "RangeDecimalPlacesEquals":
                case "NumberDecimalPlacesEquals":
                case "DecimalPlacesEquals":
                    return Result(_excel.RangeNumberFormatDecimalPlacesEquals(
                        task.SheetName,
                        GetAddress(task),
                        task.DecimalPlaces));

                //Project 3 task 3
                case "TableRowContainingTextDeleted":
                case "DeleteTableRowContainingText":
                case "TableRowDeletedByText":
                case "RemoveTableRowContainingText":
                    return Result(_excel.TableRowContainingTextDeleted(
                        task.SheetName,
                        task.ExpectedText,
                        task.ExpectedRowCount));

                //Project 3 task 4
                case "AverageFormulaByHeaders":
                case "AverageFormulaFromHeaders":
                case "MonthlyAverageFormula":
                case "AverageMonthlyQuantity":
                    return Result(_excel.AverageFormulaByHeaders(
                        task.SheetName,
                        task.TargetHeader,
                        task.SourceHeaders));
                //Project 3 task 5
                case "ChartPrimaryVerticalAxisTitleEquals":
                case "PrimaryVerticalAxisTitleEquals":
                case "ChartVerticalAxisTitleEquals":
                case "ValueAxisTitleEquals":
                    return Result(_excel.ChartPrimaryVerticalAxisTitleEquals(
                        task.SheetName,
                        task.ExpectedText));
                //Project 3 task 6 dung lai cai cua task 5 de check title cua vertical axis, de test lai xem co dung ko truoc khi lam task 6 moi
                //Project 3 task 7
                case "TableColumnFormulaFilledDown":
                case "FormulaFilledDownToEndOfTableColumn":
                case "FillFormulaDownTableColumn":
                case "ExtendFormulaToEndOfTableColumn":
                    return Result(_excel.TableColumnFormulaFilledDown(
                        task.SheetName,
                        task.Cell));

                //Project 3 task 8
                case "MaxFormulaFromHeader":
                case "MaxFunctionFromHeader":
                case "HighestNumberFromColumn":
                case "MaxStockValue":
                    return Result(_excel.MaxFormulaFromHeader(
                        task.SheetName,
                        task.Cell,
                        task.SourceHeader));
                //Project 4 task 1
                case "ChartSwitchedRowColumn":
                case "ChartSwitchRowColumn":
                case "SwitchRowColumnChart":
                case "ChartRowsColumnsSwapped":
                case "SwapChartDataOverAxis":
                    return Result(_excel.ChartSwitchedRowColumn(
                        task.SheetName,
                        task.ChartTitle,
                        string.IsNullOrWhiteSpace(task.Range) ? task.DataRange : task.Range));
                //Project 4 task 2
                case "RangeFormattingMatches":
                case "FormattingCopiedFromRange":
                case "TitleSubtitleFormattingCopied":
                case "CopyTitleSubtitleFormatting":
                case "FormatPainterTitleSubtitle":
                    return Result(_excel.RangeFormattingMatches(
                        task.SourceSheetName,
                        string.IsNullOrWhiteSpace(task.SourceRange) ? task.Range : task.SourceRange,
                        task.SheetName,
                        GetAddress(task)));

                //Project 4 task 3
                case "ChartSheetLegendRemovedValueLabelsAbove":
                case "ChartSheetNoLegendValueLabelsAbove":
                case "ChartDataLabelsValuesAboveNoLegend":
                case "ChartLegendRemovedValueLabelsOnly":
                case "NoLegendValueLabelsAbove":
                    return Result(_excel.ChartSheetLegendRemovedValueLabelsAbove(
                        task.SheetName));

                //Project 4 task 4
                case "WorksheetTableConvertedToRange":
                case "TableConvertedToRange":
                case "ConvertTableToRangeKeepFormatting":
                case "TableConvertedToRangeKeepFormatting":
                case "ConvertToRangeKeepFormatting":
                    return Result(_excel.WorksheetTableConvertedToRange(
                        task.SheetName,
                        string.IsNullOrWhiteSpace(task.Range) ? "A4:J30" : task.Range,
                        task.TableName));
                //Project 4 task 5
                case "ReportClusteredColumnChartCreated":
                case "ClusteredColumnChartByHeaders":
                case "ClusteredColumnChartMonthQuantity":
                case "ReportQuantityMonthlyClusteredColumnChart":
                case "CreateMonthlyQuantityClusteredColumnChart":
                    {
                        string categoryHeader = "Month";
                        string valueHeader = "Quantity";

                        if (task.SourceHeaders != null && task.SourceHeaders.Count > 0)
                            categoryHeader = task.SourceHeaders[0];

                        if (task.SourceHeaders != null && task.SourceHeaders.Count > 1)
                            valueHeader = task.SourceHeaders[1];

                        return Result(_excel.ReportClusteredColumnChartCreated(
                            task.SheetName,
                            categoryHeader,
                            valueHeader));
                    }

                //Project 4 task 6
                case "LeftFormulaByHeaders":
                case "LeftFunctionByHeaders":
                case "FirstCharactersFromHeader":
                case "First2CharactersFromCategory":
                case "TtcFromCategoryLeft":
                    return Result(_excel.LeftFormulaByHeaders(
                        task.SheetName,
                        task.TargetHeader,
                        task.SourceHeader,
                        task.CharacterCount <= 0 ? 2 : task.CharacterCount));

                //Project 4 task 7
                case "LastFirstNameFormulaAtCell":
                    return Result(_excel.LastFirstNameFormulaAtCell(
                        task.SheetName,
                        task.Cell,
                        task.SourceHeaders != null && task.SourceHeaders.Count > 0 ? task.SourceHeaders[0] : "LastName",
                        task.SourceHeaders != null && task.SourceHeaders.Count > 1 ? task.SourceHeaders[1] : "Firstname",
                        string.IsNullOrEmpty(task.ExpectedText) ? ", " : task.ExpectedText));

                //Project 4 task 8
                case "CenterFooterPageOfPagesEquals":
                    return Result(_excel.CenterFooterPageOfPagesEquals(task.SheetName));


                default:
                    return (false, "Unsupported assertion type: " + task.AssertionType);
            }
        }

        private string GetAddress(TaskDefinition task)
        {
            if (!string.IsNullOrWhiteSpace(task.Range))
                return task.Range;

            return task.Cell;
        }

        private int ToInt(string text)
        {
            int value;
            int.TryParse(text, out value);
            return value;
        }

        private bool ToBool(string text)
        {
            bool value;
            if (bool.TryParse(text, out value))
                return value;

            return string.Equals(text, "1", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "yes", StringComparison.OrdinalIgnoreCase) ||
                   string.Equals(text, "true", StringComparison.OrdinalIgnoreCase);
        }

        private (bool pass, string message) Result(bool ok)
        {
            return ok ? (true, "PASS") : (false, "FAIL");
        }
    }
}
