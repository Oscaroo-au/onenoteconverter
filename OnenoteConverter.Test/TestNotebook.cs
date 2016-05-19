using System;
using System.Diagnostics;
using System.Threading.Tasks;
using Microsoft.VisualStudio.TestTools.UnitTesting;

using OnenoteConverter.Lib;

namespace OnenoteConverter.Test
{
    [TestClass]
    public class TestNotebook
    {
        string m_TestRoot = @"C:\Data\Programming\SVN checkouts\ConversionTest";
        /// <summary>
        /// Tests whether I can open a notebook on onenote.
        /// </summary>
        [TestMethod]
        public void TestOpenNotebook()
        {
            var src_path = new Uri(m_TestRoot + @"\Notebooks\TEST");

            var app = new OneNote();
            var nb = new Notebook(OneNote.Instance, src_path);
            nb.Show();

            // Visual check to see whether we can see it.
        }

        [TestMethod]
        public async Task TestPublishAs2007()
        {
            var src_path = new Uri(m_TestRoot + @"\Notebooks\TEST");
            var dest_path = new Uri(m_TestRoot + @"\Notebooks\TEST_As2007");

            var app = new OneNote();
            var nb = new Notebook(OneNote.Instance, src_path);

            await nb.Publish2007Async(dest_path);

            // Visual check to see whether it has been converted.
            nb.Close();
        }

        [TestMethod]
        public async Task TestPublishAs2013()
        {
            var src_path = new Uri(m_TestRoot + @"\Notebooks\TEST");
            var dest_path = new Uri(m_TestRoot + @"\Notebooks\TEST_As2013");

            var app = new OneNote();
            var nb = new Notebook(OneNote.Instance, src_path);

            await nb.Publish2013Async(dest_path);

            // Visual check to see whether it has been converted.
            nb.Close();
        }


        [TestMethod]
        public async Task TestConvertSmallSample()
        {
            Uri dest = new Uri(m_TestRoot + @"\Output");
            Uri src = new Uri(@"G:\RVS Notebooks - Confidential");

            var conv = new Lib.OnenoteConverter();
            conv.DestinationPath = dest;
            conv.SourcePath = src;
            conv.FilterExpression = "^BRA";

            ProgressReporter report = new ProgressReporter();

            report.Progress += Report_Progress;

            await conv.ConvertAsync(report);
            
        }

        private void Report_Progress(object sender, ProgressEventArgs e)
        {
            Trace.WriteLine(e.Message, "Status:");
        }
    }
}
