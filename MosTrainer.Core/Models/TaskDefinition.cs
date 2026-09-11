using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Collections.Generic;

namespace MosTrainer.Core.Models
{
    public class TaskDefinition
    {
        public string ProjectId { get; set; } = "";
        public string TaskId { get; set; } = "";

        public string TitleKey { get; set; } = "";
        public string InstructionKey { get; set; } = "";

        public string AssertionType { get; set; } = "";

        public string SheetName { get; set; } = "";
        public string Cell { get; set; } = "";
        public string Range { get; set; } = "";

        public string ExpectedText { get; set; } = "";
        public string ExpectedFormula { get; set; } = "";
        public string ExpectedFormat { get; set; } = "";
        public string ExpectedValue { get; set; } = "";

        public string TableName { get; set; } = "";
        public string ChartTitle { get; set; } = "";
        public string ChartName { get; set; } = "";
        public int ExpectedChartType { get; set; }
        public string ShapeName { get; set; } = "";
        public string PropertyName { get; set; } = "";
        public string NamedRange { get; set; } = "";
        public string ColumnHeader { get; set; } = "";
        public string TargetHeader { get; set; } = "";
        public string CriteriaHeader { get; set; } = "";
        public string Operator { get; set; } = "";
        public double Threshold { get; set; }
        public string TrueText { get; set; } = "";
        public string FalseText { get; set; } = "";
        public List<string> SortHeaders { get; set; } = new List<string>();
        public List<string> SortOrders { get; set; } = new List<string>();
        public string SourceHeader { get; set; } = "";
        public string Domain { get; set; } = "";
        public string SourceSheetName { get; set; } = "";
        public string SourceFileName { get; set; } = "";
        public bool FirstRowAsHeaders { get; set; } = true;
        public double ExpectedWidth { get; set; }
        public string DataRange { get; set; } = "";
        public string LocationRange { get; set; } = "";
        public List<string> RangeNames { get; set; } = new List<string>();
        public List<string> TargetRanges { get; set; } = new List<string>();
        public List<string> SecondaryRanges { get; set; } = new List<string>();
        public List<string> ExpectedTexts { get; set; } = new List<string>();
        public string SecondaryExpectedFormat { get; set; } = "";
        public int FreezeRows { get; set; }
        public int FreezeColumns { get; set; }
        public int DecimalPlaces { get; set; }
        public int ExpectedRowCount { get; set; }
        public List<string> SourceHeaders { get; set; } = new List<string>();
        public string SourceRange { get; set; } = "";
        public int CharacterCount { get; set; }

        [JsonExtensionData]
        public IDictionary<string, JToken> Extra { get; set; }
    }
}
