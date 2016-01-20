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
        string filePath;
        string fileSuffix;
        string fileBaseName;
        public string fileName;
        string[] streamIDs;
        // Declare the list for processing the individual parameters
        List<ReportServer.ParameterValue> parameters;
        // Declare a parameter value instance for later use
        ReportServer.ParameterValue parameter;

        public ReportGenerate()
        {
            // Default Constructor for the class ctor
            result = null;
            reportPath = "/DSI/PickupList";
            format = "PDF";
            historyID = null;
            deviceInfo = null;
            streamIDs = null;
            fileName = "C:\\DSIAPPS\\ReportTest.pdf";
            filePath = @"C:\DSIAPPS";
            fileBaseName = "TestReport";
            fileSuffix = ".pdf";
            // For use in formatting the file name
            // DateTime.Now.ToString("yyyyMMddHHmmssfff")

        }
        public void GenerateReport()
        {
            ReportServer.Warning[] warnings = null;
            // Setup the Report Execution Service rs
            ReportServer.ReportExecutionService rs = new ReportServer.ReportExecutionService();
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
            // Set up an Execution Context
            ReportServer.ExecutionHeader execHeader = new ReportServer.ExecutionHeader();
            rs.ExecutionHeaderValue = execHeader;
            // Load the report
            rs.LoadReport(reportPath, historyID);
            // If the report has parameters then load the parameters
            rs.SetExecutionParameters(parameters.ToArray(), "en-us");
            // Render the report
            result = rs.Render(format, deviceInfo, out extension, out mimeType, out encoding, out warnings, out streamIDs);
        }
        public void SaveReportToDisk()
        {
            fileName = fileBaseName + DateTime.Now.ToString("yyyyMMddHHmmss") + fileSuffix;
            fileName = Path.Combine(filePath, fileName);
            FileStream stream = File.Create(fileName, result.Length);
            stream.Write(result, 0, result.Length);
            stream.Close();
        }
        public void initParameterValues(string message)
        {
            // Set up for processing the report
            // Configure for parameters
            parameters = new List<ReportServer.ParameterValue>();
            // Parse the incoming message into parameters for use on the report
            Dictionary<string, string> keyValuePairs = message.Split(',')
                .Select(value => value.Split('='))
                .ToDictionary(pair => pair[0], pair => pair[1]);
            foreach (KeyValuePair<string,string> param in keyValuePairs)
            {
                // Prep for finding report parameters
                parameter = new ReportServer.ParameterValue();
                // Process the special parameters used to override the report name and report path
                // values are name value pairs name=value
                // filePath = the folder on disk where the report should be stored
                // reportName = the full path to the report on the report server ie /DSI/CompanyList
                // remaining values passed through are parameter values
                switch (param.Key)
                {
                    case "filePath":
                        filePath = param.Value;
                        break;
                    case "reportName":
                        reportPath = param.Value;
                        break;
                    default:
                        parameter.Name = param.Key;
                        parameter.Value = param.Value;
                        parameters.Add(parameter);
                        break;
                }
                
            }

        }

    }
}
