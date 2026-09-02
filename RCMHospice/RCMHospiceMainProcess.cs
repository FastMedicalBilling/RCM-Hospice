using DocumentFormat.OpenXml.Spreadsheet;
using OfficeOpenXml;
using System.Data;
using System.Globalization;
using static RCMHospice.RCMHospiceHelpers;

namespace RCMHospice
{
    public static class RCMHospiceMainProcess
    {
        public static void NOEProcess(string searchReportPath, string[] agencyFiles, string queryFilePath)
        {
            Console.WriteLine("NOE Process Started");

            Directory.CreateDirectory("ErrorLogs");
            string logFileName = $"logNOEProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = Path.Combine("ErrorLogs", logFileName);

            ArchiveFiles(searchReportPath, "SearchReport");

            string originalChangesFilename = ModifyExcelFile(searchReportPath, "Search");

            for (int indexOfAgencies = 0; indexOfAgencies < agencyFiles.Length; indexOfAgencies++)
            {
                ArchiveFiles(agencyFiles[indexOfAgencies], "Agencies");
            }

            DataTable csvData = ReadCsvFile(queryFilePath, "QueryName", "SqlQuery");

            for (int indexOfAgencies = 0; indexOfAgencies < agencyFiles.Length; indexOfAgencies++)
            {
                string agencyFilePath = agencyFiles[indexOfAgencies];

                string agencyName = agencyFilePath
                    .Replace("C:\\Automation\\Files\\AgenciesHospice\\", "")
                    .Replace(" HH", "")
                    .Replace(".xlsx", "")
                    .Trim();

                Console.WriteLine($"Starting NOE work on {agencyName}");
                RCMHospiceHelpers.ForceColumnToText(searchReportPath, "TOB");

                string noeQuery = GetSqlQueryByQueryName(csvData, "NOE Filter");

                DataTable resultFromSearchReport = ExecuteExcelQuery(searchReportPath, noeQuery);
                resultFromSearchReport = ConvertDatesToDateOnly(resultFromSearchReport);

                resultFromSearchReport = RCMHospiceHelpers.FilterNOESearchReportRows(
                    resultFromSearchReport,
                    agencyName
                );

                if (resultFromSearchReport.Rows.Count == 0)
                {
                    Console.WriteLine($"No NOE rows found for {agencyName}");
                    continue;
                }

                string newFilePath = "RecalculatedFile.xlsx";
                File.Copy(agencyFilePath, newFilePath, true);

                using (ExcelPackage package = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromSearchReport.Rows.Count; indexOfPatients++)
                    {
                        string patientName = resultFromSearchReport.Rows[indexOfPatients]["Patient Name"].ToString().Trim();
                        string hicMbi = resultFromSearchReport.Rows[indexOfPatients]["HIC/MBI"].ToString().Trim();
                        string sLoc = resultFromSearchReport.Rows[indexOfPatients]["S/Loc"].ToString().Trim();

                        if (string.IsNullOrWhiteSpace(hicMbi))
                            continue;

                        if (!DateTime.TryParse(resultFromSearchReport.Rows[indexOfPatients]["Start Date"].ToString(), out DateTime startDate))
                            continue;

                        bool startsWithS = sLoc.StartsWith("S", StringComparison.OrdinalIgnoreCase);
                        bool startsWithP = sLoc.StartsWith("P", StringComparison.OrdinalIgnoreCase);

                        if (!startsWithS && !startsWithP)
                            continue;

                        string statusToWrite = startsWithP ? "Approved" : "Accepted";

                        var (capYear, sheetName) = RCMHospiceHelpers.GetCapYearAndSheet(startDate);

                        ExcelWorksheet ws;

                        try
                        {
                            ws = RCMHospiceHelpers.EnsureClaimsYearSheetExists(package, sheetName);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            LogError(ex.Message, logFilePath);
                            continue;
                        }

                        int patientsNameCol = RCMHospiceHelpers.GetColumnByHeader(ws, "PATIENTS NAME");
                        int statusCol = RCMHospiceHelpers.GetColumnByHeader(ws, "Status");
                        int socCol = RCMHospiceHelpers.GetColumnByHeader(ws, "SOC");
                        int dcStatusCol = RCMHospiceHelpers.GetColumnByHeader(ws, "DC Status");
                        int hicMbiCol = RCMHospiceHelpers.GetColumnByHeader(ws, "HIC/MBI");

                        List<int> monthColumns = RCMHospiceHelpers.GetMonthColumns(ws);

                        if (patientsNameCol <= 0 || statusCol <= 0 || socCol <= 0 || dcStatusCol <= 0 || hicMbiCol <= 0 || monthColumns.Count == 0)
                        {
                            Console.WriteLine($"Required columns were not found on sheet {sheetName} for {agencyName}");
                            LogError($"Required columns were not found on sheet {sheetName} for {agencyName}", logFilePath);
                            continue;
                        }

                        int existingRow = RCMHospiceHelpers.FindRowByHicMbi(ws, hicMbi, hicMbiCol);
                        int targetRow = existingRow > 0 ? existingRow : RCMHospiceHelpers.GetNextPatientRow(ws, patientsNameCol);
                        Console.WriteLine($"Writing patient {patientName} to row {targetRow} on {sheetName}");

                        List<DateTime> monthDates = RCMHospiceHelpers.GenerateMonthDates(startDate, capYear);

                        ws.Cells[targetRow, patientsNameCol].Value = patientName;
                        ws.Cells[targetRow, statusCol].Value = statusToWrite;
                        ws.Cells[targetRow, socCol].Value = startDate;
                        ws.Cells[targetRow, socCol].Style.Numberformat.Format = "m/d/yyyy";
                        ws.Cells[targetRow, dcStatusCol].Value = "Active";
                        ws.Cells[targetRow, hicMbiCol].Value = hicMbi;

                        RCMHospiceHelpers.WriteMonthDatesToMonthColumns(ws, targetRow, monthDates, monthColumns);

                        if (existingRow > 0)
                        {
                            Console.WriteLine($"Updated patient {patientName} with HIC/MBI {hicMbi} in {agencyName}, sheet {sheetName}, status {statusToWrite}");
                            LogError($"Updated patient {patientName} with HIC/MBI {hicMbi} in {agencyName}, sheet {sheetName}, status {statusToWrite}", logFilePath);
                        }
                        else
                        {
                            Console.WriteLine($"Inserted patient {patientName} with HIC/MBI {hicMbi} in {agencyName}, sheet {sheetName}, status {statusToWrite}");
                            LogError($"Inserted patient {patientName} with HIC/MBI {hicMbi} in {agencyName}, sheet {sheetName}, status {statusToWrite}", logFilePath);
                        }
                    }

                    package.Save();
                }

                File.Copy(newFilePath, agencyFilePath, true);
                File.Delete(newFilePath);
            }

            UndoModifyExcelFile(originalChangesFilename);
            Console.WriteLine("NOE Process Completed");
        }

        public static void FutureSummaryProcess(string xlsFileSearchReport, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Future Summary Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFutureSummaryProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);
            string body = "Hey Boss,\nThe following Agencies had Not Equal today in the Future Summary Report :\n\n";

            DataTable csvData = new DataTable();


            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string changesQuery = GetSqlQueryByQueryName(csvData, "search table by TOB");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                int FutureSummaryTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "Future Summary");

                Console.WriteLine($"Starting work on {agencyName}");
                string statusCheck = CheckStatuses(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[FutureSummaryTabInd], agencyName);

                if (statusCheck != "")
                {
                    body += statusCheck;
                    body += "\n";
                }
            }

            body += "\nRegards,\nYour Friendly Neighbourhood Automation";

            SendEmail("Automation Reports", "max@hhabilling.com;jayberko@gmail.com", $"Not Equal report for {DateOnly.FromDateTime(DateTime.Today)}", body, logFilePath);
        }

        public static void SuspenseProcess(string xlsFileSuspenseReport, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Suspense Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logSuspenseProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);

            string originalChangesFilename = ModifyExcelFile(xlsFileSuspenseReport, "Search");
            ArchiveFiles(xlsFileSuspenseReport, "SuspenseReport");

            string submitDateValueFromSuspense = string.Empty;
            string reimbValueFromSuspense = string.Empty;
            string TOBValueFromSuspense = string.Empty;
            string HICValueFromSuspense = string.Empty;
            string startDateFromSuspense = string.Empty;
            string patientNameFromSuspense = string.Empty;
            DataTable csvData = new DataTable();
            DataTable resultFromSuspense = new DataTable();
            DataTable resultFromAgency = new DataTable();

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string changesQuery = GetSqlQueryByQueryName(csvData, "Suspense Table");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                resultFromSuspense = ExecuteExcelQuery(xlsFileSuspenseReport, changesQuery);
                resultFromSuspense = ConvertDatesToDateOnly(resultFromSuspense);
                ClearTableLeaveHeader(xlsfilePathAgency[indexOfAgencies]);

                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                int paymentsTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "PAYMENTS");

                Console.WriteLine($"Starting work on {agencyName}");
                ClearColumnsDandE(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[paymentsTabInd]);

                if (resultFromSuspense.Rows.Count > 0)
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromSuspense.Rows.Count; indexOfPatients++)
                    {
                        submitDateValueFromSuspense = resultFromSuspense.Rows[indexOfPatients]["Submit Date"].ToString();
                        TOBValueFromSuspense = resultFromSuspense.Rows[indexOfPatients]["TOB"].ToString();
                        reimbValueFromSuspense = resultFromSuspense.Rows[indexOfPatients]["Reimb"].ToString();
                        HICValueFromSuspense = resultFromSuspense.Rows[indexOfPatients]["HIC/MBI"].ToString();
                        startDateFromSuspense = resultFromSuspense.Rows[indexOfPatients]["Start Date"].ToString();
                        patientNameFromSuspense = resultFromSuspense.Rows[indexOfPatients]["Patient Name"].ToString();

                        string newDate = AddDaysToDate(submitDateValueFromSuspense, 14);
                        if (reimbValueFromSuspense != "")
                        {
                            if (TOBValueFromSuspense == "811" || TOBValueFromSuspense == "812" || TOBValueFromSuspense == "813" || TOBValueFromSuspense == "814")
                            {
                                UpdatePaymentsSheetAggregate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[paymentsTabInd], newDate, reimbValueFromSuspense, "suspense projection");
                                UpdateTotalsFromSuspense(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[paymentsTabInd]);
                                WriteToInProcessTab(xlsfilePathAgency[indexOfAgencies], patientNameFromSuspense, startDateFromSuspense, submitDateValueFromSuspense, reimbValueFromSuspense);
                            }
                        }
                        else if (TOBValueFromSuspense == "811" || TOBValueFromSuspense == "812" || TOBValueFromSuspense == "813" || TOBValueFromSuspense == "814")
                        {
                            WriteToInProcessTab(xlsfilePathAgency[indexOfAgencies], patientNameFromSuspense, startDateFromSuspense, submitDateValueFromSuspense, reimbValueFromSuspense);
                        }
                    }
                }

                string inProcessSheetName = sheetNamesOfAgencyFile.FirstOrDefault(x => x.Trim().Equals("In Process", StringComparison.OrdinalIgnoreCase));

                if (string.IsNullOrWhiteSpace(inProcessSheetName))
                    throw new Exception("Worksheet 'In Process' was not found in sheetNamesOfAgencyFile.");

                SortExcelRowsByPaidDate(xlsfilePathAgency[indexOfAgencies], inProcessSheetName, "Submit Date", "asc");
                ChangePatientsNameToProperCase(xlsfilePathAgency[indexOfAgencies], inProcessSheetName);
            }
            UndoModifyExcelFile(originalChangesFilename);
        }

        public static void HICPullProcess(string[] xlsfilePathAgency, string xlsxHICFile, string csvFilePath)
        {
            Console.WriteLine($"HIC Pull Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFuturePaymentProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);

            string originalHICFilename = ModifyExcelFile(xlsxHICFile, "Change");
            ArchiveFiles(xlsxHICFile, "GetHIC");

            DataTable csvData = new DataTable();
            DataTable resultFromHICFile = new DataTable();
            DataTable resultFromAgencyFileClaims = new DataTable();
            DataTable resultFromAgencyFileSummary = new DataTable();
            DataTable resultFromAgencyFileFuture = new DataTable();

            csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
            string changesQuery = GetSqlQueryByQueryName(csvData, "select all");
            List<string> sheetNamesOfHICFile = GetSheetNames(xlsxHICFile);
            changesQuery = changesQuery.Replace("worksheet", sheetNamesOfHICFile[0]);
            resultFromHICFile = ExecuteExcelQuery(xlsxHICFile, changesQuery);
            resultFromHICFile = ConvertDatesToDateOnly(resultFromHICFile);
            var distinctResultsHICFile = resultFromHICFile.AsEnumerable().GroupBy(row => row.Field<string>("Patient Name")).Select(group => group.First()).CopyToDataTable();

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                List<string> sheetNamesOfAgency = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                var resultsPerAgencyFromHIC = DeleteRowsByName(resultFromHICFile, "Agency", agencyName);
                Console.WriteLine($"Starting work on {agencyName}");

                changesQuery = GetSqlQueryByQueryName(csvData, "select all");
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfAgency[0]);
                resultFromAgencyFileClaims = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], changesQuery);
                resultFromAgencyFileClaims = ConvertDatesToDateOnly(resultFromAgencyFileClaims);
                changesQuery = GetSqlQueryByQueryName(csvData, "select all");
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfAgency[2]);
                resultFromAgencyFileSummary = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], changesQuery);
                resultFromAgencyFileSummary = ConvertDatesToDateOnly(resultFromAgencyFileSummary);
                changesQuery = GetSqlQueryByQueryName(csvData, "select all");
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfAgency[3]);
                resultFromAgencyFileFuture = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], changesQuery);
                resultFromAgencyFileFuture = ConvertDatesToDateOnly(resultFromAgencyFileFuture);

                //var distinctResultsAgency = resultFromAgencyFile.AsEnumerable().GroupBy(row => row.Field<string>("PATIENTS NAME")).Select(group => group.First()).CopyToDataTable();
                CompareAndInsert(distinctResultsHICFile, resultFromAgencyFileClaims);
                CompareAndInsert(distinctResultsHICFile, resultFromAgencyFileFuture);
                CompareAndInsert(distinctResultsHICFile, resultFromAgencyFileSummary);
                CopyDataTableToWorksheet(xlsfilePathAgency[indexOfAgencies], resultFromAgencyFileClaims, sheetNamesOfAgency[0]);
                CopyDataTableToWorksheet(xlsfilePathAgency[indexOfAgencies], resultFromAgencyFileSummary, sheetNamesOfAgency[2]);
                CopyDataTableToWorksheet(xlsfilePathAgency[indexOfAgencies], resultFromAgencyFileFuture, sheetNamesOfAgency[3]);
            }
            UndoModifyExcelFile(originalHICFilename);
        }

        public static void EmailAgencyList(string xlsFileAgencyList, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Email and Text Agency List Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFuturePaymentProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);

            //string originalChangesFilename = ModifyExcelFile(xlsFilePaymentSummary, "Summary");

            DataTable csvData = new DataTable();
            DataTable resultFromAgencyList = new DataTable();
            DataTable resultFromPayment = new DataTable();
            Tuple<int, int> indexOfCell = new Tuple<int, int>(0, 0);

            csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
            string changesQuery = GetSqlQueryByQueryName(csvData, "select all");
            List<string> sheetNamesOfAgencyList = GetSheetNames(xlsFileAgencyList);
            changesQuery = changesQuery.Replace("worksheet", sheetNamesOfAgencyList[0]);
            resultFromAgencyList = ExecuteExcelQuery(xlsFileAgencyList, changesQuery);

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                DataTable configurationInfo = ReadCsvFile("ConfigFile.csv", "SmtpInfo", "Details");
                string sendToEmail = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToEmail");
                string sendToName = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToName");

                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                changesQuery = GetSqlQueryByQueryName(csvData, "select all");
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                int paymentsTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "PAYMENTS");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                Console.WriteLine($"Starting work on {agencyName}");
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfAgencyFile[paymentsTabInd]);
                resultFromPayment = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], changesQuery);
                resultFromPayment = ConvertDatesToDateOnly(resultFromPayment);

                // Create a new DataTable to store filtered rows
                DataTable filteredresultFromPayment = resultFromPayment.Clone(); // Clones the structure of yourDataTable

                foreach (DataRow row in resultFromPayment.Rows)
                {
                    string payDateString = row["DATE"].ToString();
                    string depositDateString = row["DEPOSIT"].ToString();
                    string amountString = row["AMOUNT"].ToString();
                    string suspenseString = row["SUSPENSE"].ToString();
                    DateTime depositDate;

                    if (depositDateString.Equals("2/3/2026"))
                    {
                        int i = 1;
                    }

                    // Check if the "Deposit Date" column value matches the specified date format
                    if (DateTime.TryParseExact(depositDateString, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out depositDate)
                        && !string.IsNullOrWhiteSpace(amountString) || !string.IsNullOrWhiteSpace(suspenseString)) // Check if amount or suspense is not empty
                    {
                        // Check if the date is today or in the future
                        if (depositDate.Date >= DateTime.Today)
                        {
                            // If the conditions are met, import the row to the filteredDataTable
                            filteredresultFromPayment.Rows.Add(row.ItemArray); // Add the entire row to maintain data integrity
                        }
                    }
                }

                if (filteredresultFromPayment.Rows.Count > 0)
                {
                    for (int indexOfAgencyList = 0; indexOfAgencyList < resultFromAgencyList.Rows.Count; indexOfAgencyList++)
                    {
                        string AgencyNameFromList = resultFromAgencyList.Rows[indexOfAgencyList]["Agency Name"].ToString();
                        if (AgencyNameFromList == agencyName)
                        {
                            string emailFromList = resultFromAgencyList.Rows[indexOfAgencyList]["Email"].ToString();
                            string cellFromList = resultFromAgencyList.Rows[indexOfAgencyList]["Cell"].ToString();
                            string isStarted = resultFromAgencyList.Rows[indexOfAgencyList]["Started"].ToString();

                            // Constructing the initial part of the message
                            string body = $"Good Afternoon {agencyName},\n\nAttached is the updated spreadsheet for today\n\nHere are the upcoming payments:\n";

                            int rowCount = filteredresultFromPayment.Rows.Count;
                            int currentIndex = 0;

                            // Adding datatable content to the body
                            foreach (DataRow row in filteredresultFromPayment.Rows)
                            {
                                string formattedAmount = string.Format("${0:N2}", row["AMOUNT"]);
                                string formattedAmountSUSPENSE = string.Format("{0:N2}", row["SUSPENSE"]);
                                string formattedAmountTOTAL = string.Format("{0:N2}", row["TOTAL"]);
                                string formattedAmountPASTDUE = string.Empty;

                                int index = filteredresultFromPayment.Columns.IndexOf("PAST DUE");

                                if (index >= 0 && index < filteredresultFromPayment.Columns.Count - 1)
                                {
                                    // Get the heading name to the right
                                    formattedAmountPASTDUE = filteredresultFromPayment.Columns[index + 1].ColumnName.Replace('#', '.');
                                }

                                if (row["NOTES"].ToString() != "")
                                    body += $"\nPay Date: {row["DATE"]}, Deposit Date: {row["DEPOSIT"]}, Amount: {formattedAmount}, Amount In Process: {formattedAmountSUSPENSE}, Total: {formattedAmountTOTAL}, Notes: {row["NOTES"]}\n";
                                else
                                    body += $"\nPay Date: {row["DATE"]}, Deposit Date: {row["DEPOSIT"]}, Amount: {formattedAmount}\n";

                                currentIndex++;

                                if (currentIndex == rowCount)
                                {
                                    body += $"\nTotal Amount In Process Past Due: {formattedAmountPASTDUE}";
                                }
                            }

                            body += "\nRegards,\nFast Medical Billing";
                            string subject = $"Today's Payment Projections for {agencyName} - {DateTime.Today.Month + "/" + DateTime.Today.Day + "/" + DateTime.Today.Year}";
                            string textMessageBody = TruncateMessageBody(body);

                            if (cellFromList.Contains(';'))
                            {
                                bool isListMatch = true;
                                var listOfCells = cellFromList.Split(';');
                                var listOfYes = isStarted.Split(';');
                                if (listOfCells.Length != listOfYes.Length)
                                {
                                    isListMatch = false;
                                }

                                for (int i = 0; i < listOfCells.Length; i++)
                                {
                                    if (listOfYes[i].ToLower().Contains("yes") && listOfCells[i] != "" && isListMatch == true)
                                    {
                                        SendSms("18184312850", listOfCells[i], textMessageBody);
                                    }
                                    else
                                    {
                                        Console.WriteLine($"Mismatching list of cellphone to list of approvals for {agencyName}, text was not sent");
                                    }
                                }
                            }
                            else
                            {
                                if (isStarted.ToLower().Contains("yes") && cellFromList != "")
                                {
                                    SendSms("18184312850", cellFromList, textMessageBody);
                                }
                                else
                                {
                                    Console.WriteLine($"No cells provided for {agencyName}, text was not sent");
                                }
                            }

                            string htmlToSend = BuildHtmlEmailFromBody(body, subject, DateTime.Today, ".\\template.html");

                            SendEmail(agencyName, emailFromList, subject, htmlToSend, logFilePath, xlsfilePathAgency[indexOfAgencies], isHtml: true);

                        }
                    }
                }
                else
                {
                    Console.WriteLine($"No data today for {agencyName}, email and text was not sent");
                }
            }
        }

        public static void PaymentSummaryProcess(string xlsFilePaymentSummary, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Payment Summary Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFuturePaymentProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);

            string originalChangesFilename = ModifyExcelFile(xlsFilePaymentSummary, "Summary");
            ApplyHolidayAggregation_HighLevelPaymentSummary(xlsFilePaymentSummary, "Pay Date");
            ArchiveFiles(xlsFilePaymentSummary, "PaymentSummary");

            string agencyFromPaymentSummary = string.Empty;
            string payDateValueFromPaymentSummary = string.Empty;
            string checkNumberFromPaymentSummary = string.Empty;
            string projectedAmountFromPaymentSummary = string.Empty;
            string checkAmountFromPaymentSummary = string.Empty;
            string payDateValueFromAgency = string.Empty;
            DataTable csvData = new DataTable();
            DataTable resultFromPaymentSummary = new DataTable();
            Tuple<int, int> indexOfCell = new Tuple<int, int>(0, 0);
            string alphaOfNameCell = string.Empty;
            string alphaOfStatusCell = string.Empty;

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                DataTable configurationInfo = ReadCsvFile("ConfigFile.csv", "SmtpInfo", "Details");
                string sendToEmail = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToEmail");
                string sendToName = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToName");

                List<string> sheetNamesOfHighLevel = GetSheetNames((xlsFilePaymentSummary));

                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string changesQuery = GetSqlQueryByQueryName(csvData, "Payments table");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfHighLevel[0]);
                resultFromPaymentSummary = ExecuteExcelQuery(xlsFilePaymentSummary, changesQuery);
                resultFromPaymentSummary = ConvertDatesToDateOnly(resultFromPaymentSummary);
                DataView dv = resultFromPaymentSummary.DefaultView;
                dv.Sort = " [Pay Date] asc";
                resultFromPaymentSummary = dv.ToTable();
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                int paymentsTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "PAYMENTS");

                Console.WriteLine($"Starting work on {agencyName}");

                bool paymentForToday = HasAnyScheduledNumberForToday(resultFromPaymentSummary);

                if (paymentForToday)
                {
                    string ccnNumber = RCMHospiceHelpers.GetCcnFromCapTab(xlsfilePathAgency[indexOfAgencies]);

                    if (ccnNumber != "")
                    {
                        RCMHospiceHelpers.WriteCompanyCodeToMappingFile(ccnNumber);

                        bool success = RunEIDMReportsDownloader(out string err);

                        if (!success)
                        {
                            Console.WriteLine($"EIDM downloader failed: {err} skipping this company");
                            continue;
                        }
                        else
                        {
                            //if downloaded successfully wait 5 mins
                            Thread.Sleep(TimeSpan.FromMinutes(5));
                        }

                        string fileName = RCMHospiceHelpers.GetCompanyFileNameFromMappingFile();

                        string csvFile1 = RCMHospiceProcess.EIDMPath + fileName;
                        string csvFile2 = RCMHospiceProcess.EIDMPath + fileName.Replace("Count", "Total").Insert(fileName.LastIndexOf('.'), "_-_IP_OP");

                        PopulateCapTabFromCsvFiles(xlsfilePathAgency[indexOfAgencies], csvFile1, csvFile2);
                        DeleteCsvFiles(RCMHospiceProcess.EIDMPath + fileName, csvFile2);
                    }
                    else
                        Console.WriteLine($"No CCN for {xlsfilePathAgency[indexOfAgencies]}");
                }



                if (resultFromPaymentSummary.Rows.Count > 0)
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromPaymentSummary.Rows.Count; indexOfPatients++)
                    {
                        payDateValueFromPaymentSummary = resultFromPaymentSummary.Rows[indexOfPatients]["Pay Date"].ToString();
                        checkNumberFromPaymentSummary = resultFromPaymentSummary.Rows[indexOfPatients]["Check #"].ToString();
                        projectedAmountFromPaymentSummary = resultFromPaymentSummary.Rows[indexOfPatients]["Projected"].ToString();
                        checkAmountFromPaymentSummary = resultFromPaymentSummary.Rows[indexOfPatients]["Check Amount"].ToString();

                        if (checkNumberFromPaymentSummary != "")
                        {
                            UpdatePaymentsSheet(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[paymentsTabInd], payDateValueFromPaymentSummary, checkAmountFromPaymentSummary, "received");
                        }
                        else
                        {
                            // Convert the payDateValueFromPaymentSummary string to a DateTime object
                            DateTime payDate = DateTime.ParseExact(payDateValueFromPaymentSummary, "MM/dd/yyyy", null);

                            // Get the current date
                            DateTime currentDate = DateTime.Today;

                            // Compare the two dates
                            if (payDate.Date == currentDate)
                            {
                                SendEmail(sendToName, sendToEmail, $"No check number for Today's date for {agencyName}", $"Hello {sendToName},\n\nFor Today's date that was present in the High Level Payment Summary report there was no check number.\nPlease take care.\n\nYour Friendly Neighborhood Automation", logFilePath);
                            }
                            else
                            {
                                UpdatePaymentsSheet(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[paymentsTabInd], payDateValueFromPaymentSummary, projectedAmountFromPaymentSummary, "projected");
                            }
                        }

                    }
                }
            }
            UndoModifyExcelFile(originalChangesFilename);
        }

        public static void MoveFuturePaymentToSummaryProcess(string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Move Future Payment Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFuturePaymentProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                Console.WriteLine($"Starting work on {agencyName}");
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                int futurepaymentTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "Future payment");
                int PaymentSummaryTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "Payment Summary");
                CopyRowToSummary(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[futurepaymentTabInd], sheetNamesOfAgencyFile[PaymentSummaryTabInd]);
            }
        }

        public static void FuturePaymentProcess(string xlsFileSearchReport, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Future Payment Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFuturePaymentProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);


            string originalChangesFilename = ModifyExcelFile(xlsFileSearchReport, "Search");
            UpdateDatesUsingHolidays_ReplaceOnly(xlsFileSearchReport, "Paid Date");

            string patientNameValueFromSearch = string.Empty;
            string hicValueFromSearch = string.Empty;
            string sLocValueFromSearch = string.Empty;
            string cancelledDateFromSearch = string.Empty;
            string startDateValueFromSearch = string.Empty;
            string paidDateValueFromSearch = string.Empty;
            string reimbValueFromSearch = string.Empty;
            string TOBValueFromSearch = string.Empty;
            string patientNameValueFromAgency = string.Empty;
            string hicValueFromAgency = string.Empty;
            string startDateValueFromAgency = string.Empty;
            string paidDateValueFromAgency = string.Empty;
            DataTable csvData = new DataTable();
            DataTable resultFromSearch = new DataTable();
            Tuple<int, int> indexOfCell = new Tuple<int, int>(0, 0);
            string alphaOfNameCell = string.Empty;
            string alphaOfStatusCell = string.Empty;

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string changesQuery = GetSqlQueryByQueryName(csvData, "search table by TOB");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("C:\\Automation\\Files\\AgenciesHospice\\", "").Replace(" HH", "").Replace(".xlsx", "");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                resultFromSearch = ExecuteExcelQuery(xlsFileSearchReport, changesQuery);
                resultFromSearch = ConvertDatesToDateOnly(resultFromSearch);
                DataView dv = resultFromSearch.DefaultView;
                dv.Sort = " [Paid Date] desc";
                resultFromSearch = dv.ToTable();
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                int FuturepaymentTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "Future payment");
                int PaymentSummaryTabInd = GetSheetIndexByName(sheetNamesOfAgencyFile, "Payment Summary");

                bool TOB32Gor327 = false;
                bool TOB32Ior329 = false;

                Console.WriteLine($"Starting work on {agencyName}");

                if (resultFromSearch.Rows.Count > 0)
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromSearch.Rows.Count; indexOfPatients++)
                    {
                        patientNameValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Patient Name"].ToString();
                        patientNameValueFromSearch = RemoveMiddleName(patientNameValueFromSearch);
                        startDateValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Start Date"].ToString();
                        paidDateValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Paid Date"].ToString();
                        reimbValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Reimb"].ToString();
                        TOBValueFromSearch = resultFromSearch.Rows[indexOfPatients]["TOB"].ToString();
                        hicValueFromSearch = resultFromSearch.Rows[indexOfPatients]["HIC/MBI"].ToString();
                        sLocValueFromSearch = resultFromSearch.Rows[indexOfPatients]["S/Loc"].ToString();
                        cancelledDateFromSearch = resultFromSearch.Rows[indexOfPatients]["Cancelled Date"].ToString();

                        if (cancelledDateFromSearch == "")
                        {
                            string modifiedQuery = GetSqlQueryByQueryName(csvData, "select all future payment");
                            DataTable resultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], modifiedQuery);
                            resultFromAgency = ConvertDatesToDateOnly(resultFromAgency);

                            string columnNameofFoundDate = FindNextStartOrSOCDateColumn(hicValueFromSearch, resultFromAgency, startDateValueFromSearch);

                            if (columnNameofFoundDate == null && sLocValueFromSearch.ToLower().StartsWith("p", StringComparison.OrdinalIgnoreCase))
                            {
                                if (TOBValueFromSearch == "812" || TOBValueFromSearch == "814" || TOBValueFromSearch == "813" || TOBValueFromSearch == "811" || TOBValueFromSearch == "817" || TOBValueFromSearch == "81G" || TOBValueFromSearch == "81I" || TOBValueFromSearch == "81g" || TOBValueFromSearch == "81i")
                                {
                                    string lastRowNumberFound = FindNearestRowWithDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[FuturepaymentTabInd], paidDateValueFromSearch, logFilePath);

                                    if (lastRowNumberFound != null && reimbValueFromSearch != "0.0000")
                                    {
                                        InsertRowAndData(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[FuturepaymentTabInd], Int32.Parse(lastRowNumberFound), patientNameValueFromSearch, startDateValueFromSearch, paidDateValueFromSearch, reimbValueFromSearch, hicValueFromSearch);
                                        TOB32Ior329 = true;
                                    }
                                }
                            }
                        }
                    }

                    for (int indexOfPatients = 0; indexOfPatients < resultFromSearch.Rows.Count; indexOfPatients++)
                    {
                        patientNameValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Patient Name"].ToString();
                        patientNameValueFromSearch = RemoveMiddleName(patientNameValueFromSearch);
                        startDateValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Start Date"].ToString();
                        paidDateValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Paid Date"].ToString();
                        reimbValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Reimb"].ToString();
                        TOBValueFromSearch = resultFromSearch.Rows[indexOfPatients]["TOB"].ToString();
                        hicValueFromSearch = resultFromSearch.Rows[indexOfPatients]["HIC/MBI"].ToString();
                        cancelledDateFromSearch = resultFromSearch.Rows[indexOfPatients]["Cancelled Date"].ToString();

                        if (cancelledDateFromSearch == "")
                        {
                            string modifiedQuery = GetSqlQueryByQueryName(csvData, "select all future payment");
                            DataTable resultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], modifiedQuery);
                            resultFromAgency = ConvertDatesToDateOnly(resultFromAgency);

                            if (TOBValueFromSearch == "817" || TOBValueFromSearch == "81G" || TOBValueFromSearch == "81I" || TOBValueFromSearch == "81g" || TOBValueFromSearch == "81i")
                            {
                                decimal reimbFromPaymentSummary = FindReimbursementAmount(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[PaymentSummaryTabInd], startDateValueFromSearch, hicValueFromSearch);

                                decimal notFound = DeductReimbursement(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[FuturepaymentTabInd], startDateValueFromSearch, hicValueFromSearch, reimbFromPaymentSummary);

                                if (notFound < 0) { LogError($"There are no results in the agency file {agencyName} future payment tab for patient : {patientNameValueFromSearch}", logFilePath); }
                                else TOB32Gor327 = true;
                            }
                        }

                    }
                    SortExcelRowsByPaidDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[PaymentSummaryTabInd], "Paid Date", "desc");
                    SortExcelRowsByPaidDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[FuturepaymentTabInd], "Paid Date", "asc");
                }
            }

            UndoModifyExcelFile(originalChangesFilename);
        }

        public static void FinalProcess(string xlsFileSearchReport, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Final Process Initiated");

            Directory.CreateDirectory("ErrorLogs");
            ArchiveFiles(xlsFileSearchReport, "SearchReport");

            string logFileName = $"logFinalProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = Path.Combine("ErrorLogs", logFileName);

            string originalSearchFilename = ModifyExcelFile(xlsFileSearchReport, "Search");

            List<string> sheetNamesOfSearchFile = GetSheetNames(originalSearchFilename);
            ChangePatientsNameToProperCase(xlsFileSearchReport, sheetNamesOfSearchFile[0]);

            DataTable csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                string agencyFilePath = xlsfilePathAgency[indexOfAgencies];

                string agencyName = agencyFilePath
                    .Replace(@"C:\Automation\Files\AgenciesHospice\", "")
                    .Replace(" HH", "")
                    .Replace(".xlsx", "")
                    .Trim();

                Console.WriteLine($"Starting Final work on {agencyName}");

                string searchQuery = GetSqlQueryByQueryName(csvData, "Final All Search Rows");

                DataTable allSearchRows = ExecuteExcelQuery(xlsFileSearchReport, searchQuery);
                allSearchRows = ConvertDatesToDateOnly(allSearchRows);

                DataTable resultFromSearch = RCMHospiceHelpers.FilterFinalSearchReportRows(allSearchRows, agencyName);

                if (resultFromSearch.Rows.Count == 0)
                {
                    Console.WriteLine($"No Final rows found for {agencyName}");
                    continue;
                }

                ProcessAgencyFileDates(agencyFilePath);

                string newFilePath = "RecalculatedFile.xlsx";
                File.Copy(agencyFilePath, newFilePath, true);

                using (ExcelPackage package = new ExcelPackage(new FileInfo(newFilePath)))
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromSearch.Rows.Count; indexOfPatients++)
                    {
                        DataRow searchRow = resultFromSearch.Rows[indexOfPatients];

                        string patientName = RCMHospiceHelpers.GetDataRowValue(searchRow, "Patient Name", "E");
                        string hicMbi = RCMHospiceHelpers.GetDataRowValue(searchRow, "HIC/MBI", "D");
                        string startDateValue = RCMHospiceHelpers.GetDataRowValue(searchRow, "Start Date", "G");
                        string throughDateValue = RCMHospiceHelpers.GetDataRowValue(searchRow, "Through Date", "H");
                        string reimbValue = RCMHospiceHelpers.GetDataRowValue(searchRow, "Reimb", "AF");
                        string tob = RCMHospiceHelpers.GetDataRowValue(searchRow, "TOB", "L");
                        string sLoc = RCMHospiceHelpers.GetDataRowValue(searchRow, "S/LOC", "M");

                        if (string.IsNullOrWhiteSpace(hicMbi))
                            continue;

                        if (!DateTime.TryParse(startDateValue, out DateTime startDate))
                        {
                            Console.WriteLine($"Could not parse Start Date for {patientName} ({hicMbi})");
                            LogError($"Could not parse Start Date for {patientName} ({hicMbi})", logFilePath);
                            continue;
                        }

                        DateTime? throughDate = null;
                        if (DateTime.TryParse(throughDateValue, out DateTime parsedThroughDate))
                            throughDate = parsedThroughDate;

                        double reimb = RCMHospiceHelpers.ParseDoubleSafe(reimbValue);

                        bool startsWithP = sLoc.StartsWith("P", StringComparison.OrdinalIgnoreCase);
                        bool startsWithS = sLoc.StartsWith("S", StringComparison.OrdinalIgnoreCase);

                        bool isDeathDischargeTob = tob.Contains("811") || tob.Contains("814");
                        bool isRegularDischargeTob = tob.Contains("81B");
                        bool isDischargeTob = isDeathDischargeTob || isRegularDischargeTob;

                        if (!startsWithP && !startsWithS && !isDischargeTob)
                            continue;

                        var (capYear, sheetName) = RCMHospiceHelpers.GetCapYearAndSheet(startDate);

                        ExcelWorksheet ws;

                        try
                        {
                            ws = RCMHospiceHelpers.EnsureClaimsYearSheetExists(package, sheetName);
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                            LogError(ex.Message, logFilePath);
                            continue;
                        }

                        int hicMbiCol = RCMHospiceHelpers.GetColumnByHeader(ws, "HIC/MBI");
                        int dcStatusCol = RCMHospiceHelpers.GetColumnByHeader(ws, "DC Status");
                        int endDateCol = RCMHospiceHelpers.GetColumnByHeader(ws, "End Date");

                        if (hicMbiCol <= 0)
                        {
                            Console.WriteLine($"HIC/MBI column was not found on {sheetName} for {agencyName}");
                            LogError($"HIC/MBI column was not found on {sheetName} for {agencyName}", logFilePath);
                            continue;
                        }

                        List<int> monthColumns = RCMHospiceHelpers.GetMonthColumns(ws);

                        if (monthColumns.Count == 0)
                        {
                            Console.WriteLine($"No Month columns found on {sheetName} for {agencyName}");
                            LogError($"No Month columns found on {sheetName} for {agencyName}", logFilePath);
                            continue;
                        }

                        int rowIndex = RCMHospiceHelpers.FindRowByHicAndDate(ws, hicMbi, startDate, hicMbiCol, monthColumns);

                        if (rowIndex <= 0)
                        {
                            Console.WriteLine($"Could not find HIC/MBI {hicMbi} with Start Date {startDate:M/d/yyyy} in {agencyName}, sheet {sheetName}");
                            LogError($"Could not find HIC/MBI {hicMbi} with Start Date {startDate:M/d/yyyy} in {agencyName}, sheet {sheetName}", logFilePath);
                            continue;
                        }

                        int monthCol = RCMHospiceHelpers.FindMatchingMonthColumn(ws, rowIndex, startDate, monthColumns);

                        if (monthCol <= 0)
                        {
                            Console.WriteLine($"Could not find matching month date {startDate:M/d/yyyy} for HIC/MBI {hicMbi} in {agencyName}, sheet {sheetName}");
                            LogError($"Could not find matching month date {startDate:M/d/yyyy} for HIC/MBI {hicMbi} in {agencyName}, sheet {sheetName}", logFilePath);
                            continue;
                        }

                        int paidCol = monthCol + 1;
                        int monthStatusCol = monthCol + 2;

                        if (startsWithP)
                        {
                            ws.Cells[rowIndex, paidCol].Value = reimb;
                            ws.Cells[rowIndex, monthStatusCol].Value = "Paid";

                            Console.WriteLine($"Paid posted for {patientName} ({hicMbi}) on {startDate:M/d/yyyy}, amount {reimb}");
                            LogError($"Paid posted for {patientName} ({hicMbi}) on {startDate:M/d/yyyy}, amount {reimb}", logFilePath);
                        }

                        if (startsWithS)
                        {
                            ws.Cells[rowIndex, monthStatusCol].Value = "In Process";

                            Console.WriteLine($"In Process posted for {patientName} ({hicMbi}) on {startDate:M/d/yyyy}");
                            LogError($"In Process posted for {patientName} ({hicMbi}) on {startDate:M/d/yyyy}", logFilePath);
                        }

                        if (isDischargeTob)
                        {
                            if (dcStatusCol <= 0 || endDateCol <= 0)
                            {
                                Console.WriteLine($"DC Status or End Date column was not found on {sheetName} for {agencyName}");
                                LogError($"DC Status or End Date column was not found on {sheetName} for {agencyName}", logFilePath);
                                continue;
                            }

                            string searchReportColZ = RCMHospiceHelpers.GetDataRowValueByExcelColumn(searchRow, "Z");
                            string searchReportColAA = RCMHospiceHelpers.GetDataRowValueByExcelColumn(searchRow, "AA");

                            bool isDeath = false;

                            if (isDeathDischargeTob)
                                isDeath = RCMHospiceHelpers.HasDeathCode55(searchReportColZ, searchReportColAA);

                            ws.Cells[rowIndex, dcStatusCol].Value = isDeath ? "Death" : "DC";

                            if (throughDate.HasValue)
                            {
                                ws.Cells[rowIndex, endDateCol].Value = throughDate.Value;
                                ws.Cells[rowIndex, endDateCol].Style.Numberformat.Format = "m/d/yyyy";
                            }

                            RCMHospiceHelpers.ClearAfterStatusAndMarkDc(ws, rowIndex, monthStatusCol);

                            Console.WriteLine($"Discharge updated for {patientName} ({hicMbi}) as {(isDeath ? "Death" : "DC")} with End Date {(throughDate.HasValue ? throughDate.Value.ToString("M/d/yyyy") : "blank")} | TOB={tob} | Z={searchReportColZ} | AA={searchReportColAA}");
                            LogError($"Discharge updated for {patientName} ({hicMbi}) as {(isDeath ? "Death" : "DC")} with End Date {(throughDate.HasValue ? throughDate.Value.ToString("M/d/yyyy") : "blank")} | TOB={tob} | Z={searchReportColZ} | AA={searchReportColAA}", logFilePath);
                        }
                    }

                    package.Save();
                }

                File.Copy(newFilePath, agencyFilePath, true);
                File.Delete(newFilePath);

                ProcessAgencyFileDates(agencyFilePath);
            }

            UndoModifyExcelFile(originalSearchFilename);

            Console.WriteLine("Final Process Completed");
        }
    }
 }
