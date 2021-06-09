using System.IO;
using System.Collections.Generic;
using ClosedXML.Excel;

namespace ReportFormatter
{
    public class ReportFormatter
    {
        private string _reportFile;
        private List<string> _orderNumbers;
        public bool FinalizeReport(string reportFile, List<string> orderNumbers)
        {
            if (string.IsNullOrWhiteSpace(reportFile) || !File.Exists(reportFile))
                return false;
            if (orderNumbers is null || orderNumbers.Count < 1)
                return false;

            _reportFile = reportFile;
            _orderNumbers = orderNumbers;
            DuplicateWorksheet();
            return ValidateNewReport();
        }

        private bool ValidateNewReport()
        {
            //check if final report matches expected output
            var wb = new XLWorkbook(_reportFile);
            var worksheets = wb.Worksheets.Count;
            return worksheets == _orderNumbers.Count + 2;
        }

        private void DuplicateWorksheet()
        {
            var workbook = new XLWorkbook(_reportFile);
            var baseWorksheet = workbook.Worksheet(2);
            foreach (var orderNumber in _orderNumbers)
            {
                baseWorksheet.CopyTo(orderNumber);
                var currentWorksheet = workbook.Worksheet(orderNumber);
                currentWorksheet.Cell(4, 6).Value = orderNumber.Replace('-', '/');
            }

            workbook.Save();
        }
    }
}
