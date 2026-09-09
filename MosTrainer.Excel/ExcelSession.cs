using System;

namespace MosTrainer.Excel
{
    internal class ExcelSession
    {
        public object App;        // Excel.Application (để tránh nullable + giữ C# 7.3 đơn giản)
        public object Workbook;   // Excel.Workbook
        public int Pid;
        public string WorkbookPath;
        public DateTime StartedAt;

        public ExcelSession()
        {
            App = null;
            Workbook = null;
            Pid = 0;
            WorkbookPath = "";
            StartedAt = DateTime.Now;
        }
    }
}
