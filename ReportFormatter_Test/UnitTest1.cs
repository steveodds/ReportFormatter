using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Collections.Generic;

namespace ReportFormatter_Test
{
    [TestClass]
    public class UnitTest1
    {
        private string reportFile = @"C:\Users\user\Downloads\COLGATE.xlsx";
        private List<string> orderNumbers = new List<string>()
        {
            "NAR-1-CPEAL-01361-MRO-TV-011555",
            "NAR-1-CPEAL-01333-MRO-TV-011111",
            "NAR-1-CPEAL-01364-MRO-TV-011999",
            "NAR-1-CPEAL-01366-MRO-TV-011000",
            "NAR-1-CPEAL-01361-MRO-TV-011267",
            "NAR-1-CPEAL-01361-MRO-TV-011954",
            "NAR-1-CPEAL-01367-MRO-TV-011467",
            "NAR-1-CPEAL-00000-MRO-TV-011224"
        };


        [TestMethod]
        public void MainTest()
        {
            var formatter = new ReportFormatter.ReportFormatter();
            var result = formatter.FinalizeReport(reportFile, orderNumbers);
            Assert.IsTrue(result);
        }
    }
}
