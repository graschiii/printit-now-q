using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Drawing.Printing;
using System.Web.Services.Protocols;
using System.Runtime.InteropServices; // For Marshal.Copy
using System.Security.Principal;
using eDSITran.ReportServer;

namespace eDSITran
{
    /// <summary>
    /// Experimental (for the moment) class for testing the feasibility of sending 
    /// Reporting Services reports directly to a target printer.
    /// </summary>
    public class ReportPrinter
    {
        ReportExecutionService rs;
        private byte[][] m_renderedReport;
        private Graphics.EnumerateMetafileProc m_delegate = null;
        private MemoryStream m_currentPageStream;
        private Metafile m_metafile = null;
        int m_numberOfPages;
        private int m_currentPrintingPage;
        private int m_lastPrintingPage;
        /* The following, in combination with the setDebug method is defined to easily allow emission of 
         * application log messages via statements like:
         *
         *  if (debugOn) eventLog.WriteEntry("Message text", EventLogEntryType.Information, 7001);
        */
        protected static EventLog eventLog = new EventLog("Application", ".", "Aspnet_wp");
        protected bool debugOn;

        public ReportPrinter()
        {
            this.debugOn = System.Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["debug"].Trim());
            rs = new ReportExecutionService();
            rs.Url = System.Configuration.ConfigurationManager.AppSettings["eDSITran_ReportServer_ReportingService"].Trim();
            rs.Credentials = System.Net.CredentialCache.DefaultCredentials;
        }
        public ReportPrinter(System.Net.ICredentials rsCr)
        {
            this.debugOn = System.Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["debug"].Trim());
            rs = new ReportExecutionService();
            rs.Url = System.Configuration.ConfigurationManager.AppSettings["eDSITran_ReportServer_ReportingService"].Trim();
            rs.Credentials = rsCr;
        }
        public ReportPrinter(System.Net.ICredentials rsCr, string reportServerUrl)
        {
            this.debugOn = System.Convert.ToBoolean(System.Configuration.ConfigurationManager.AppSettings["debug"].Trim());
            rs = new ReportExecutionService();
            rs.Url = reportServerUrl;
            rs.Credentials = rsCr;
        }

        /// <summary>
        /// This first version of RenderReport is strictly for legacy support of NIFTransferShipment.  NIFTransferShipment should be re-worked to use the rendorReport
        /// version that accepts an array of ParameterValues.
        /// </summary>
        /// <param name="reportPath"></param>
        /// <param name="identifier"></param>
        /// <returns></returns>
        public byte[][] renderReport(string reportPath, ParameterValue[] reportParameters)
        {
            // Private variables for rendering
            string historyID = null;
            ExecutionHeader execHeader = new ExecutionHeader();

            try
            {
                rs.Timeout = 300000;
                rs.ExecutionHeaderValue = execHeader;

                ExecutionInfo execInfo = new ExecutionInfo();
                execInfo = rs.LoadReport(reportPath, historyID);

                if (reportParameters != null)
                {
                    rs.SetExecutionParameters(reportParameters, "en-us");
                }

                Byte[][] pages = new Byte[0][];
                string format = "IMAGE";
                int numberOfPages = 1;
                byte[] currentPageStream = new byte[1] { 0x00 }; // this single byte will prime the while loop
                string extension = null;
                string encoding = null;
                string mimeType = null;
                string[] streamIDs = null;
                Warning[] warnings = null;

                while (currentPageStream.Length > 0)
                {
                    string deviceInfo = String.Format(@"<DeviceInfo><OutputFormat>EMF</OutputFormat><PrintDpiX>96</PrintDpiX><PrintDpiY>96</PrintDpiY><StartPage>{0}</StartPage></DeviceInfo>", numberOfPages);

                    //Execute the report and get page count.
                    currentPageStream = rs.Render(
                       format,
                       deviceInfo,
                       out extension,
                       out encoding,
                       out mimeType,
                       out warnings,
                       out streamIDs);
                    
                    if (currentPageStream.Length == 0 && numberOfPages == 1)
                    {
                        break;  // nothing rendered
                    }

                    if (currentPageStream.Length > 0)
                    {
                        Array.Resize(ref pages, pages.Length + 1);
                        pages[pages.Length - 1] = currentPageStream;
                        numberOfPages++;
                    }

                }

                m_numberOfPages = numberOfPages - 1;

                return pages;
            }
            catch (SoapException ex)
            {
                eventLog.WriteEntry(ex.Detail.InnerXml, EventLogEntryType.Information, 7001);
                Console.WriteLine(ex.Detail.InnerXml);
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.Message, EventLogEntryType.Information, 7001);
                Console.WriteLine(ex.Message);
            }
            finally
            {
            }

            return null;
        }

        public bool printReport(string printer, string report, Int16 copies, bool landscape, List<string> parameterNames, List<string> parameterValues)
        {
            bool success = true;
            if (debugOn) eventLog.WriteEntry("Parameters: " + parameterNames.Count.ToString(), EventLogEntryType.Information, 7001);
            ParameterValue[] reportParameters = new ParameterValue[parameterNames.Count];

            for (int i = 0; i < parameterNames.Count; i++)
            {
                ParameterValue a = new ParameterValue();
                a.Name = parameterNames[i];
                a.Value = parameterValues[i];
                reportParameters.SetValue(a, i);
            }

            try
            {
                success = this.printReport(printer, report, copies, landscape, reportParameters);
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.ToString(), EventLogEntryType.Information, 7001);
                Console.WriteLine(ex);
            }

            return success;
        }
        public bool printReport(string printer, string report, Int16 copies, bool landscape, ParameterValue[] reportParameters)
        {
            bool success = true;
            this.RenderedReport = this.renderReport(report, reportParameters);

            // Stop impersonation
            WindowsImpersonationContext ctx = WindowsIdentity.Impersonate(IntPtr.Zero);
            
            try
            {
                // Wait for the report to completely render.
                if (m_numberOfPages < 1)
                    return false;
                PrinterSettings printerSettings = new PrinterSettings();
                printerSettings.MaximumPage = m_numberOfPages;
                printerSettings.MinimumPage = 1;
                printerSettings.PrintRange = PrintRange.SomePages;
                printerSettings.FromPage = 1;
                printerSettings.ToPage = m_numberOfPages;
                printerSettings.Copies = copies;

                if (printer != "*DEFAULT")
                {
                    printerSettings.PrinterName = printer;
                }
                else
                {
                    string printerList = "";
                    for (int i = 0; i < PrinterSettings.InstalledPrinters.Count; i++)
                    {
                        if (!printerSettings.IsDefaultPrinter) {
                            printerSettings.PrinterName = PrinterSettings.InstalledPrinters[i];
                        }

                        if (i > 0) { printerList = printerList + ", "; }
                        printerList = printerList + PrinterSettings.InstalledPrinters[i];
                    }

                    if (debugOn) eventLog.WriteEntry("Installed printers: " + printerList, EventLogEntryType.Information, 7001);
                }


                PrintDocument pd = new PrintDocument();
                m_currentPrintingPage = 1;
                m_lastPrintingPage = m_numberOfPages;
                pd.PrinterSettings = printerSettings;
                pd.DocumentName = report;
                pd.DefaultPageSettings.Landscape = landscape;

                // Print report
                if (debugOn) eventLog.WriteEntry("Printing " + m_numberOfPages.ToString() + " pages to " + printerSettings.PrinterName, EventLogEntryType.Information, 7001);

                pd.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
                pd.Print();
            }
            catch (Exception ex)
            {
                eventLog.WriteEntry(ex.Message, EventLogEntryType.Error, 7001);
            }
            finally
            {
                // Resume impersonation
                ctx.Undo();
            }

            return success;
        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            ev.HasMorePages = false;
            if (m_currentPrintingPage <= m_lastPrintingPage && MoveToPage(m_currentPrintingPage))
            {
                // Draw the page
                ReportDrawPage(ev.Graphics);
                // If the next page is less than or equal to the last page, 
                // print another page.
                if (++m_currentPrintingPage <= m_lastPrintingPage)
                    ev.HasMorePages = true;
            }
        }

        // Method to draw the current emf memory stream 
        private void ReportDrawPage(Graphics g)
        {
            if (null == m_currentPageStream || 0 == m_currentPageStream.Length || null == m_metafile)
                return;
            lock (this)
            {
                // Set the metafile delegate.
                int width = m_metafile.Width;
                int height = m_metafile.Height;
                m_delegate = new Graphics.EnumerateMetafileProc(MetafileCallback);
                // Draw in the rectangle
                Point[] points = new Point[3];
                Point destPoint = new Point(0, 0);
                Point destPoint1 = new Point(width, 0);
                Point destPoint2 = new Point(0, height);

                points[0] = destPoint;
                points[1] = destPoint1;
                points[2] = destPoint2;

                g.EnumerateMetafile(m_metafile, points, m_delegate);
                // Clean up
                m_delegate = null;
            }
        }
        private bool MoveToPage(Int32 page)
        {
            // Check to make sure that the current page exists in
            // the array list
            if (null == this.RenderedReport[m_currentPrintingPage - 1])
                return false;
            // Set current page stream equal to the rendered page
            m_currentPageStream = new MemoryStream(this.RenderedReport[m_currentPrintingPage - 1]);
            // Set its postion to start.
            m_currentPageStream.Position = 0;
            // Initialize the metafile
            if (null != m_metafile)
            {
                m_metafile.Dispose();
                m_metafile = null;
            }
            // Load the metafile image for this page
            m_metafile = new Metafile((Stream)m_currentPageStream);
            return true;
        }

        private bool MetafileCallback(
         EmfPlusRecordType recordType,
         int flags,
         int dataSize,
         IntPtr data,
         PlayRecordCallback callbackData)
        {
            byte[] dataArray = null;
            // Dance around unmanaged code.
            if (data != IntPtr.Zero)
            {
                // Copy the unmanaged record to a managed byte buffer 
                // that can be used by PlayRecord.
                dataArray = new byte[dataSize];
                Marshal.Copy(data, dataArray, 0, dataSize);
            }
            // play the record.      
            m_metafile.PlayRecord(recordType, flags, dataSize, dataArray);

            return true;
        }

        public byte[][] RenderedReport
        {
            get
            {
                return m_renderedReport;
            }
            set
            {
                m_renderedReport = value;
            }
        }
    }
}
