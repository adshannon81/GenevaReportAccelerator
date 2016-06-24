
//  If LoggingOn is defined, then Run.Log is created in the Output Folder and holds the contents of the status variable
//#define LoggingOn

//  If DebugOn is defined, then extra item lists are captured in status variable
//#define DebugOn


using System;
using System.Collections.Generic;
using Advent.Geneva.WFM.Framework.BaseImplementation;
using Advent.Geneva.WFM.Framework.Interfaces;
using Advent.Geneva.WFM.SQLDataAccess;
using Advent.Geneva.WFM.GenevaDataAccess;
using System.IO;
using GenevaAutoReports.RE2005;
using System.Text.RegularExpressions;

namespace GenevaAutoReports
{
    /// <summary>
    /// Abstract class ActivityBase implements IActivity. Deriving from ActivityBase class instead of IActivity, gives a good starting point to implement an activity.
    /// Some of the methods that could be common across activity implementations are given a default definition. For overrides, the implemented methods are marked virtual. 
    /// Deriving from this class does not force you to implement the overloaded 'Run' method which takes a step name, in case you are not using activity steps.
    /// </summary>
    public class RunClientReports : ActivityBase
    {
        private static ReportExecutionService rsExec;

        private DateTime ReportRunDate = DateTime.Now;
        private string status = "";
        private string outputFolder = "";

        private string portfolio;
        private static DateTime dtStartDate;
        private static DateTime dtEndDate;
        private static DateTime dtKnowledgeDate;
        private static DateTime dtPriorKnowledgeDate;

        private double reportCountCSV = 0;
        private double reportTotalCSV = 0;
        private double reportCountPDF = 0;
        private double reportTotalPDF = 0;

        string strStartDate;
        string strEndDate;
        string strKnowledgeDate;
        string strPriorKnowledgeDate;
        private List<ReportParameters> reportListCSV;
        private string reportName;

        private enum reportFormat
        {
            PDF, CSV
        }

        public override void Init()
        {}

        public override void Start()
        {}

        public override void Run(ActivityRun activityRun, IGenevaActions genevaInstance)
        {
            
            status = "Executed by " + genevaInstance.GetCurrentUserName() + Environment.NewLine
                    + "Files saved to " + GetSettingValue("OutputFolder") + Environment.NewLine;
            
            try
            {
                //Set Activity Start Time
                activityRun.StartDateTime = DateTime.Now;
                reportCountCSV = 0;   
                reportCountPDF = 0;  
                activityRun.CurrentStep = "Run";
                activityRun.StartDateTime = DateTime.Now;
                outputFolder = GetSettingValue("OutputFolder");
                                
                //#########  Set User Parameters #############
                UpdateProgress("Setting Parameters....", activityRun);

                //Read Activty Parameter
                portfolio = activityRun.GetParameterValue("Portfolio");
                strStartDate = activityRun.GetParameterValue("StartDate");
                strEndDate = activityRun.GetParameterValue("EndDate");
                strKnowledgeDate = activityRun.GetParameterValue("KnowledgeDate");
                strPriorKnowledgeDate = activityRun.GetParameterValue("PriorKnowledgeDate");
                
                dtStartDate = DateTime.ParseExact(strStartDate, "yyyy-MM-dd:HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                dtEndDate = DateTime.ParseExact(strEndDate, "yyyy-MM-dd:HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                dtKnowledgeDate = DateTime.ParseExact(strKnowledgeDate, "yyyy-MM-dd:HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                dtPriorKnowledgeDate = DateTime.ParseExact(strPriorKnowledgeDate, "yyyy-MM-dd:HH:mm:ss", System.Globalization.CultureInfo.InvariantCulture);
                
                //Creating Generic (base) ReportParameter List; add required parameter to this object
                ReportParameterList base_parameters = new ReportParameterList();
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("ConnectionString", Properties.Settings.Default.GenevaConnection));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("Portfolio", portfolio));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("PeriodStartDate", dtStartDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("PeriodEndDate", dtEndDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("KnowledgeDate", dtKnowledgeDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("PriorKnowledgeDate", dtPriorKnowledgeDate));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("AccountingRunType", "ClosedPeriod"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("AccountingCalendar", portfolio));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("RegionalSettings", "en-IE"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("DisableHyperlinks", "True"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("QuantityPrecision", "4"));
                base_parameters.Add(new Advent.Geneva.WFM.GenevaDataAccess.ReportParameter("AddendumPages", "None"));


                //#########  Get Report List CSV #############
                UpdateProgress("Getting Report list for CSV...", activityRun);
                
                //List<ReportParameters> reportListCSV = getReportNames("CSV", base_parameters);
                reportListCSV = getReportNames("CSV", base_parameters);
                        
                foreach (ReportParameters report in reportListCSV)
                {
                    //CSVReport(activityRun, genevaInstance, report);
                    CSVReportSSRS(activityRun, genevaInstance, report);
                    reportCountCSV++; 
                }

                activityRun.Save();


                //#########  Get Report List SSRSPDF #############
                UpdateProgress("Getting Report List for PDF...", activityRun);
               
                List<ReportParameters> reportListPDF = getReportNames("PDF", base_parameters);

                foreach (ReportParameters report in reportListPDF)
                {
                    PDFReport(activityRun, genevaInstance, report);
                    reportCountPDF++;                     
                }

                UpdateProgress("Completed!" + Environment.NewLine, activityRun);
                activityRun.UpdateSuccessfulActivityRun();
                activityRun.Save();



#if LoggingOn
                TextWriter RunWriter = new StreamWriter(outputFolder + "\\" + "Run.log");

                RunWriter.Write(status);
                RunWriter.Close();
#endif

            }
            catch (Exception exp)
            {
                //Set Failure Flag
                Exception e = new Exception(status + Environment.NewLine + "-----------Exception Message--------------" + Environment.NewLine + exp.Message);
                activityRun.SaveActivityStep(false, reportName);
                activityRun.UpdateFailedActivityRun(e);
                activityRun.Save();
            }
            finally
            {
                //Set Activity End Time and Save Activity
                activityRun.EndDateTime = DateTime.Now;
                activityRun.Save();
            }
        }


        public override void Stop()
        {}

        public override void ShutDown()
        {}


        public override List<ValidateResult> Validate()
        {
            throw new NotImplementedException();
        }

        public override List<ValidateResult> ValidateActivityRequestParams(RequestParameters parameters)
        {
            List<ValidateResult> list = new List<ValidateResult>();
            //Read Activty Parameter
            string portfolio = parameters.GetParameterValue("Portfolio");
            string strStartDate = parameters.GetParameterValue("StartDate");
            string strEndDate = parameters.GetParameterValue("EndDate");
            string strKnowledgeDate = parameters.GetParameterValue("KnowledgeDate");
            string strPriorKnowledgeDate = parameters.GetParameterValue("PriorKnowledgeDate");

            if (string.IsNullOrEmpty(portfolio) == true)
            {
                list.Add(new ValidateResult("Missing Required Parameter : Portfolio"));
            }

            if (string.IsNullOrEmpty(strStartDate) == true)
            {
                list.Add(new ValidateResult("Missing Required Parameter : Start Date"));
            }
            if (string.IsNullOrEmpty(strEndDate) == true)
            {
                list.Add(new ValidateResult("Missing Required Parameter : End Date"));
            }
            if (string.IsNullOrEmpty(strKnowledgeDate) == true)
            {
                list.Add(new ValidateResult("Missing Required Parameter : Knowledge Date"));
            }
            if (string.IsNullOrEmpty(strPriorKnowledgeDate) == true)
            {
                list.Add(new ValidateResult("Missing Required Parameter : Prior Knowledge Date"));
            }
            if (string.IsNullOrEmpty(strStartDate) == false && string.IsNullOrEmpty(strEndDate) == false
                 && string.IsNullOrEmpty(strKnowledgeDate) == false && string.IsNullOrEmpty(strPriorKnowledgeDate) == false)
            {
                try
                {
                    DateTime dtStartDate = DateTime.ParseExact(strStartDate, "yyyy-MM-dd:HH:mm:ss",
                        new System.Globalization.CultureInfo(System.Globalization.CultureInfo.InvariantCulture.Name),
                        System.Globalization.DateTimeStyles.None);

                    DateTime dtEndDate = DateTime.ParseExact(strEndDate, "yyyy-MM-dd:HH:mm:ss",
                        new System.Globalization.CultureInfo(System.Globalization.CultureInfo.InvariantCulture.Name),
                        System.Globalization.DateTimeStyles.None);

                    DateTime dtKnowledgeDate = DateTime.ParseExact(strKnowledgeDate, "yyyy-MM-dd:HH:mm:ss",
                        new System.Globalization.CultureInfo(System.Globalization.CultureInfo.InvariantCulture.Name),
                        System.Globalization.DateTimeStyles.None);

                    DateTime dtPriorKnowledgeDate = DateTime.ParseExact(strPriorKnowledgeDate, "yyyy-MM-dd:HH:mm:ss",
                        new System.Globalization.CultureInfo(System.Globalization.CultureInfo.InvariantCulture.Name),
                        System.Globalization.DateTimeStyles.None);

                    if (dtEndDate.CompareTo(dtStartDate) < 0)
                    {
                        list.Add(new ValidateResult("Error: Start Date should be less than End Date"));
                    }
                    if (strKnowledgeDate.CompareTo(strPriorKnowledgeDate) < 0)
                    {
                        list.Add(new ValidateResult("Error: Prior Knowledge Date should be less than Knowledge Date"));
                    }
                }
                catch (Exception exp)
                {
                    list.Add(new ValidateResult(exp.Message));
                }
            }
            return list;
        }

        private Exception RunSSRSReport(ReportParameters Report, ActivityRun activityRun, reportFormat format)
        {
            rsExec = new ReportExecutionService();
            rsExec.Url = Properties.Settings.Default.RE2005_ReportExecutionService;

            rsExec.UseDefaultCredentials = true;

            string historyID = null;
            string deviceInfo = null;
            Byte[] results;
            string encoding = String.Empty;
            string mimeType = String.Empty;
            string extension = String.Empty;
            Warning[] warnings = null;
            string[] streamIDs = null;

            var p = "";
            for (int i = 0; i < Report.Parameters.Count; i++)
            {
                p = "Param[" + i.ToString() + "]" + Report.Parameters[i].Name + "|" + Report.Parameters[i].Value.ToString() + Environment.NewLine + p;

            }

            // Path of the Report - XLS, PDF etc.
            Report.OutputFilePath = GetOutputFileName(Report);
            string FilePath = outputFolder + Report.OutputFilePath + "." + format.ToString().ToLower();

            // Name of the report - Please note this is not the RDL file.
            string _reportName = @"/GenevaReports/" + Report.Name;

            try
            {
                ExecutionInfo ei = rsExec.LoadReport(_reportName, historyID);
                ParameterValue[] parameters = new ParameterValue[Report.Parameters.Count];

                for (int i = 0; i < Report.Parameters.Count; i++)
                {

                    parameters[i] = new ParameterValue();
                    parameters[i].Name = Report.Parameters[i].Name;
                    parameters[i].Value = (string)Report.Parameters[i].Value;
                }

                rsExec.SetExecutionParameters(parameters, "en-IE");

                DataSourceCredentials dataSourceCredentials2 = new DataSourceCredentials();
                dataSourceCredentials2.DataSourceName = Properties.Settings.Default.DataSourceName;
                dataSourceCredentials2.UserName = Properties.Settings.Default.GenevaUser;
                dataSourceCredentials2.Password = Properties.Settings.Default.GenevaPass;


                DataSourceCredentials[] _credentials2 = new DataSourceCredentials[] { dataSourceCredentials2 };

                var c = "";
                for (int i = 0; i < _credentials2.Length; i++)
                {
                    c = "_credentials2[" + i.ToString() + "]:" + _credentials2[i].DataSourceName + "|" +
                                                                      _credentials2[i].UserName + "|" +
                                                                      _credentials2[i].Password + "|" + Environment.NewLine + c;
                }

                rsExec.SetExecutionCredentials(_credentials2);
                //rsExec.UseDefaultCredentials = true;

                results = rsExec.Render(format.ToString(), deviceInfo, out extension,
                                                            out encoding,
                                                            out mimeType,
                                                            out warnings,
                                                            out streamIDs);

                using (FileStream stream = File.OpenWrite(FilePath))
                {
                    stream.Write(results, 0, results.Length);
                }

            }
            catch (Exception ex)
            {
                status = "--- ERROR ---" + Environment.NewLine +
                        "Running PDF Report: " + Report.Name + Environment.NewLine +
                        ex.Message + Environment.NewLine +
                        "--------------" + Environment.NewLine +
                        status + Environment.NewLine;
                return new Exception(status);
            }

            return null;

        }

        private List<ReportParameters> getReportNames(string ReportType, ReportParameterList base_ParameterList)
        {
            List<ReportParameters> reports = new List<ReportParameters>();
            int i = 0;

            reports.Add(new ReportParameters("Custom Unsettled Income Report #0008",
                                                    "0008CustomUnsettledIncomeReport.rsl"));
            reports[i].AddParameterList(base_ParameterList);
            i++;

            reports.Add(new ReportParameters("Custom Portfolio Valuation Report #0014",
                                                "0014CustomPortfolioValuationReport.rsl"));
            reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("Group4", "Currency");
            //reports[i].AddParameters("IncludeNotionalValuesForBasketSwaps", "1");
            i++;

            reports.Add(new ReportParameters("Custom Unsettled Transactions Report #0012",
                                                "0012CustomUnsettledTransactionsReport.rsl"));
            reports[i].AddParameterList(base_ParameterList);
            i++;

            //reports.Add(new ReportParameters("Custom Profit and Loss Report #0015",
            //                                    "0015CustomProfitandLossReport.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("AccountingCalendar", portfolio);
            ////reports[i].AddParameters("requestType", "ssrs");
            //i++;

            //reports.Add(new ReportParameters("Custom Realised Gain Loss Ledger #0005",
            //                                    "0005CustomRealisedGainLossLedger.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            ////reports[i].AddParameters("Group1", "InvestmentType");
            ////reports[i].AddParameters("NettingRuleView", "1");
            ////reports[i].AddParameters("requestType", "ssrs");
            ////reports[i].AddParameters("WashSalesTaxableEndDate", "June 3, 2016 11:59:59 pm");
            ////reports[i].AddParameters("WashSalesCutoverDate", "January 1, 1901 12:00:00 am");
            //i++;

            //reports.Add(new ReportParameters("Custom Appreciation and Depreciation on Foreign Currency Contracts #0018",
            //                                    "0018CustomAppreciationandDepreciationonFCC.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            ////reports[i].AddParameters("Group4", "TransactionType");
            ////reports[i].AddParameters("GroupbyStrategy", "0");
            //i++;

            //reports.Add(new ReportParameters("Custom Performance Attribution Report #0056",
            //                                    "0056CustomPerformanceAttributionReport.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            ////reports[i].AddParameters("requestType", "ssrs");
            //i++;

            reports.Add(new ReportParameters("Trial Balance",
                                                "glmap_fundtrialbal.rsl"));
            reports[i].AddParameterList(base_ParameterList);
            reports[i].AddParameters("FundLegalEntity", portfolio);
            //reports[i].AddParameters("FiscalCalendar", "JanYearStart");
            //reports[i].AddParameters("DisableInvestmentFilters", "1");
            //reports[i].AddParameters("requestType", "ssrs");
            //reports[i].AddParameters("YearStartRetainedEarnings", "1");
            i++;

            //reports.Add(new ReportParameters("Cash Appraisal",
            //                                    "glmap_cashapp.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            ////reports[i].AddParameters("DisableLocAcctFilters", "1");
            ////reports[i].AddParameters("DisableStrategyFilters", "1");
            //i++;

            //reports.Add(new ReportParameters("Fund Structure NAV",
            //                                    "nav.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("AccountingRunType", "NAV");
            //reports[i].AddParameters("KnowledgeDate", ReportRunDate.ToString("yyyy-MM-dd HH:mm:ss"));
            ////reports[i].AddParameters("DisableInvestmentFilters", "1");
            ////reports[i].AddParameters("DisableLocAcctFilters", "1");
            ////reports[i].AddParameters("DisableStrategyFilters", "1");
            //i++;

            //reports.Add(new ReportParameters("Fund Capital Ledger",
            //                                    "fundcapldgr.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("AccountingRunType", "NAV");
            //reports[i].AddParameters("KnowledgeDate", ReportRunDate.ToString("yyyy-MM-dd HH:mm:ss"));
            ////reports[i].AddParameters("DisableInvestmentFilters", "1");
            ////reports[i].AddParameters("DisableLocAcctFilters", "1");
            ////reports[i].AddParameters("DisableStrategyFilters", "1");
            ////reports[i].AddParameters("requestType", "ssrs");
            //i++;

            //reports.Add(new ReportParameters("Fund Allocation Percentages",
            //                                    "fundalloc.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("AccountingRunType", "NAV");
            //reports[i].AddParameters("KnowledgeDate", ReportRunDate.ToString("yyyy-MM-dd HH:mm:ss"));
            ////reports[i].AddParameters("DisableInvestmentFilters", "1");
            ////reports[i].AddParameters("DisableLocAcctFilters", "1");
            ////reports[i].AddParameters("DisableStrategyFilters", "1");
            ////reports[i].AddParameters("requestType", "ssrs");
            //i++;

            //reports.Add(new ReportParameters("Fund Allocated Income Detail",
            //                                    "fundincdet.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("AccountingRunType", "NAV");
            //reports[i].AddParameters("KnowledgeDate", ReportRunDate.ToString("yyyy-MM-dd HH:mm:ss"));
            //i++;

            reports.Add(new ReportParameters("Statement of Net Assets",
                                                "glmap_netassets.rsl"));
            reports[i].AddParameterList(base_ParameterList);
            //reports[i].AddParameters("DisableInvestmentFilters", "1");
            //reports[i].AddParameters("requestType", "ssrs");
            i++;

            //reports.Add(new ReportParameters("Statement of Changes in Net Assets",
            //                                    "glmap_chginassets.rsl"));
            //reports[i].AddParameterList(base_ParameterList);
            ////reports[i].AddParameters("DisableInvestmentFilters", "1");
            ////reports[i].AddParameters("ShowYearToDate", "0");
            ////reports[i].AddParameters("ReportType", "Summary");
            ////reports[i].AddParameters("requestType", "ssrs");
            //i++;

            reports.Add(new ReportParameters("Local Position Appraisal",
                                                "locposapp.rsl"));
            reports[i].AddParameterList(base_ParameterList);
            reports[i].AddParameters("Consolidate", "None");
            //reports[i].AddParameters("requestType", "ssrs");
            i++;


            if (ReportType == "PDF")
            {
                //reports.Add(new ReportParameters("Custom Other Assets And Liabilities #0050",
                //                                    "0050CustomOtherAssetsAndLiabilities.rsl"));
                //reports[i].AddParameterList(base_ParameterList);
                ////reports[i].AddParameters("requestType", "ssrs");
                //i++;

                reportTotalPDF = i;
            }
            else if (ReportType == "CSV")
            {
                reportTotalCSV = i;
            }

            return reports;
            
        }

        private void UpdateProgress(string Message, ActivityRun activityRun)
        {
            double csvPercent = 0;
            if (reportCountCSV != 0)
                csvPercent = (reportCountCSV / reportTotalCSV);

            double pdfPercent = 0;
            if (reportCountPDF != 0)
                pdfPercent = (reportCountPDF / reportTotalPDF);

            activityRun.Note = status + Environment.NewLine +
                "CSV Reports Completed: " + csvPercent.ToString("P0") + Environment.NewLine +
                "PDF Reports Completed: " + pdfPercent.ToString("P0") + Environment.NewLine +
                Environment.NewLine +
                Message;

            activityRun.Save();
        }

        private string GetOutputFileName(ReportParameters Report)
        {
            string OutputFileName = portfolio + "_";

            Regex rgx = new Regex("[^a-zA-Z0-9]");
            string cleanReportName = rgx.Replace(Report.Name, "");

            OutputFileName += cleanReportName + "_" + dtEndDate.ToString("yyyyMMdd") + "_" + ReportRunDate.ToString("yyyyMMddHHmmss");

            return OutputFileName;
        }

        public void CSVReport(ActivityRun ar, IGenevaActions genevaInstance, ReportParameters report)
        {
            reportName = report.Name;
            ar.CurrentStep = "CSVReport";
            UpdateProgress("Running CSV - " + report.Name, ar);

            TextReader tReader;
            genevaInstance.ExecuteReport(report.FileName, report.ParameterList, ReportOutputFormat.CSV, out tReader);
            var data = tReader.ReadToEnd();
            tReader.Close();

            report.OutputFilePath = GetOutputFileName(report);
            string FilePath = outputFolder + report.OutputFilePath + ".csv";
            TextWriter tWriter = new StreamWriter(FilePath);

            tWriter.Write(data);
            tWriter.Close();

            ar.SaveActivityStep(true, report.Name, report.OutputFilePath, "");
        }

        public void CSVReportSSRS(ActivityRun ar, IGenevaActions genevaInstance, ReportParameters report)
        {
            reportName = report.Name;
            ar.CurrentStep = "CSVReport";

            UpdateProgress("Running CSV - " + report.Name, ar);

            Exception rdlException = RunSSRSReport(report, ar, reportFormat.CSV);
            if (rdlException != null)
            {
                throw rdlException;
            }

            ar.SaveActivityStep(true, report.Name, report.OutputFilePath, "");
        }
        
        public void PDFReport(ActivityRun ar, IGenevaActions genevaInstance, ReportParameters report)
        {
            reportName = report.Name;
            ar.CurrentStep = "PDFReport";

            UpdateProgress("Running PDF - " + report.Name, ar);

            Exception rdlException = RunSSRSReport(report, ar, reportFormat.PDF);
            if (rdlException != null)
            {
                throw rdlException;
            }

            ar.SaveActivityStep(true, report.Name, report.OutputFilePath, "");
        }

    }
}
