using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Server
{
    class ReportGenerate
    {
        byte[] result;
        string reportPath;
        string format;
        string historyID;
        string deviceInfo;
        string encoding;
        string mimeType;
        string extension;
        public string fileName;
        string[] streamIDs;
        // Declarer the array of parameter values
        ReportServer.ParameterValue[] paramArray;

        public ReportGenerate()
        {
            // Default Constructor for the class ctor
            result = null;
            reportPath = "/DSI/CompanyList";
            format = "PDF";
            historyID = null;
            deviceInfo = null;
            streamIDs = null;
            fileName = "C:\\DSIAPPS\\ReportTest.pdf";
            initParameterValues();
        }
        public void GenerateReport()
        {
            ReportServer.Warning[] warnings = null;
            // Parameters variable for later use
            // ReportServer.ParameterValue[] reportHistoryParameters = null;
            // Setup the Report Execution Service rs
            ReportServer.ReportExecutionService rs = new ReportServer.ReportExecutionService();
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
            // Set up an Execution Context
            ReportServer.ExecutionHeader execHeader = new ReportServer.ExecutionHeader();
            rs.ExecutionHeaderValue = execHeader;
            // Load the report
            rs.LoadReport(reportPath, historyID);
            // Render the report
            result = rs.Render(format, deviceInfo, out extension, out mimeType, out encoding, out warnings, out streamIDs);
        }
        public void SaveReportToDisk()
        {
            FileStream stream = File.Create(fileName, result.Length);
            stream.Write(result, 0, result.Length);
            stream.Close();
        }
        private void initParameterValues()
        {
            // Nothing to do here! But we will receive the message in as a string and process it.
        }

    }
}
