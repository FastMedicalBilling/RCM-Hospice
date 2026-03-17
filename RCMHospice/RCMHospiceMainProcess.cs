using System.Data;
using System.Globalization;
using static RCMHospice.RCMHospiceHelpers;

namespace RCMHospice
{
    public static class RCMHospiceMainProcess
    {
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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                Console.WriteLine($"Starting work on {agencyName}");
                string statusCheck = CheckStatuses(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[4], agencyName);

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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                resultFromSuspense = ExecuteExcelQuery(xlsFileSuspenseReport, changesQuery);
                resultFromSuspense = ConvertDatesToDateOnly(resultFromSuspense);
                ClearTableLeaveHeader(xlsfilePathAgency[indexOfAgencies]);

                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);

                Console.WriteLine($"Starting work on {agencyName}");
                ClearColumnsDandE(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[1]);

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
                                UpdatePaymentsSheetAggregate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[1], newDate, reimbValueFromSuspense, "suspense projection");
                                UpdateTotalsFromSuspense(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[1]);
                                WriteToInProcessTab(xlsfilePathAgency[indexOfAgencies], patientNameFromSuspense, startDateFromSuspense, submitDateValueFromSuspense, reimbValueFromSuspense);
                            }
                        }
                        else if (TOBValueFromSuspense == "811" || TOBValueFromSuspense == "812" || TOBValueFromSuspense == "813" || TOBValueFromSuspense == "814")
                        {                           
                            WriteToInProcessTab(xlsfilePathAgency[indexOfAgencies], patientNameFromSuspense, startDateFromSuspense, submitDateValueFromSuspense, reimbValueFromSuspense);
                        }
                    }
                }
                SortExcelRowsByPaidDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[5], "Submit Date", "asc");
                ChangePatientsNameToProperCase(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[5]);
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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                Console.WriteLine($"Starting work on {agencyName}");
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfAgencyFile[1]);
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
                                    formattedAmountPASTDUE = filteredresultFromPayment.Columns[index + 1].ColumnName.Replace('#','.');
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
                                if(listOfCells.Length != listOfYes.Length)
                                {
                                    isListMatch = false;
                                }

                                for (int i=0; i < listOfCells.Length; i++)
                                {
                                    if (listOfYes[i].ToLower().Contains("yes") && listOfCells[i]!="" && isListMatch == true)
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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                changesQuery = changesQuery.Replace("worksheet", sheetNamesOfHighLevel[0]);
                resultFromPaymentSummary = ExecuteExcelQuery(xlsFilePaymentSummary, changesQuery);
                resultFromPaymentSummary = ConvertDatesToDateOnly(resultFromPaymentSummary);
                DataView dv = resultFromPaymentSummary.DefaultView;
                dv.Sort = " [Pay Date] asc";
                resultFromPaymentSummary = dv.ToTable();
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                Console.WriteLine($"Starting work on {agencyName}");

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
                            UpdatePaymentsSheet(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[1], payDateValueFromPaymentSummary, checkAmountFromPaymentSummary, "received");
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
                                UpdatePaymentsSheet(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[1], payDateValueFromPaymentSummary, projectedAmountFromPaymentSummary, "projected");
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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                Console.WriteLine($"Starting work on {agencyName}");
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                CopyRowToSummary(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[3], sheetNamesOfAgencyFile[2]);
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
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                resultFromSearch = ExecuteExcelQuery(xlsFileSearchReport, changesQuery);
                resultFromSearch = ConvertDatesToDateOnly(resultFromSearch);
                DataView dv = resultFromSearch.DefaultView;
                dv.Sort = " [Paid Date] desc";
                resultFromSearch = dv.ToTable();
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
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
                                    string lastRowNumberFound = FindNearestRowWithDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[3], paidDateValueFromSearch, logFilePath);

                                    if (lastRowNumberFound != null && reimbValueFromSearch != "0.0000")
                                    {
                                        InsertRowAndData(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[3], Int32.Parse(lastRowNumberFound), patientNameValueFromSearch, startDateValueFromSearch, paidDateValueFromSearch, reimbValueFromSearch, hicValueFromSearch);
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
                                decimal reimbFromPaymentSummary = FindReimbursementAmount(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[2], startDateValueFromSearch, hicValueFromSearch);

                                decimal notFound = DeductReimbursement(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[3], startDateValueFromSearch, hicValueFromSearch, reimbFromPaymentSummary);

                                if (notFound < 0) { LogError($"There are no results in the agency file {agencyName} future payment tab for patient : {patientNameValueFromSearch}", logFilePath); }
                                else TOB32Gor327 = true;
                            }
                        }

                    }
                    SortExcelRowsByPaidDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[2], "Paid Date", "desc");
                    SortExcelRowsByPaidDate(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[3], "Paid Date", "asc");
                }
            }

            UndoModifyExcelFile(originalChangesFilename);
        }

        public static void FinalProcess(string xlsFileSearchReport/*, string xlsFileSearchReportS*/, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Final Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            ArchiveFiles(xlsFileSearchReport, "SearchReport");
            //ArchiveFiles(xlsFileSearchReportS, "SearchReport");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logFinalProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);

            string originalChangesFilename = ModifyExcelFile(xlsFileSearchReport, "Search");
            List<string> sheetNamesOfSearchFile = GetSheetNames(originalChangesFilename);
            ChangePatientsNameToProperCase(xlsFileSearchReport, sheetNamesOfSearchFile[0]);

            string patientNameValueFromSearch = string.Empty;
            string startDateValueFromSearch = string.Empty;
            string reimbValueFromSearch = string.Empty;
            string hicValueFromSearch = string.Empty;
            string tobValueFromSearch = string.Empty;
            string sLocValueFromSearch = string.Empty;
            string patientNameValueFromAgency = string.Empty;
            string startDateValueFromAgency = string.Empty;
            string hicValueFromAgency = string.Empty;
            string cancelledDateFromSearch = string.Empty;
            DataTable csvData = new DataTable();
            DataTable resultFromSearch = new DataTable();
            DataTable resultFromSearchS = new DataTable();
            Tuple<int, int> indexOfCell = new Tuple<int, int>(0, 0);
            string alphaOfNameCell = string.Empty;
            string alphaOfStatusCell = string.Empty;

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                ProcessAgencyFileDates(xlsfilePathAgency[indexOfAgencies]);
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                ChangePatientsNameToProperCase(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0]);

                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string changesQuery = GetSqlQueryByQueryName(csvData, "get all relevant search table rows");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                Console.WriteLine($"Starting work on {agencyName}");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                resultFromSearch = ExecuteExcelQuery(xlsFileSearchReport, changesQuery);
                resultFromSearch = ConvertDatesToDateOnly(resultFromSearch);

                if (resultFromSearch.Rows.Count > 0)
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromSearch.Rows.Count; indexOfPatients++)
                    {
                        patientNameValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Patient Name"].ToString();
                        startDateValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Start Date"].ToString();
                        reimbValueFromSearch = resultFromSearch.Rows[indexOfPatients]["Reimb"].ToString();
                        hicValueFromSearch = resultFromSearch.Rows[indexOfPatients]["HIC/MBI"].ToString();
                        tobValueFromSearch = resultFromSearch.Rows[indexOfPatients]["TOB"].ToString();
                        sLocValueFromSearch = resultFromSearch.Rows[indexOfPatients]["S/Loc"].ToString();
                        cancelledDateFromSearch = resultFromSearch.Rows[indexOfPatients]["Cancelled Date"].ToString();

                        string modifiedQuery = GetSqlQueryByQueryName(csvData, "filter agency rows with start date");
                        modifiedQuery = modifiedQuery.Replace("PatientName", patientNameValueFromSearch);
                        DataTable resultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], modifiedQuery);
                        resultFromAgency = ConvertDatesToDateOnly(resultFromAgency);
                        var twistName = patientNameValueFromSearch.Split(",");
                        var patientNameValueFromSearchTwisted = twistName[1].Substring(1) + ", " + twistName[0];

                        string columnNameofFoundDate = FindNextStartOrSOCDateColumn(hicValueFromSearch, resultFromAgency, startDateValueFromSearch);
                        
                        if (cancelledDateFromSearch == "")
                        {
                            if (tobValueFromSearch == "329")
                            {
                                if (sLocValueFromSearch.ToLower().StartsWith("s", StringComparison.OrdinalIgnoreCase))
                                {
                                    if (columnNameofFoundDate != null)
                                    {
                                        LogError($"Column date found {columnNameofFoundDate} for patient {patientNameValueFromSearch} with S status in S\\Loc, inserting in progress status", logFilePath);
                                        Console.WriteLine($"Column date found {columnNameofFoundDate} for patient {patientNameValueFromSearch} with S status in S\\Loc, inserting in progress status");
                                        indexOfCell = FindCellLocation(xlsfilePathAgency[indexOfAgencies], hicValueFromSearch, columnNameofFoundDate, startDateValueFromSearch);
                                        alphaOfNameCell = FindCellLocationAlpha(xlsfilePathAgency[indexOfAgencies], indexOfCell);

                                        switch (columnNameofFoundDate)
                                        {
                                            case "SOC":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 4)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "In Process", "string");
                                                break;
                                            case "Start 2":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 7)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "In Process", "string");
                                                break;
                                            case "Start 3":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 10)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "In Process", "string");
                                                break;
                                            case "Start 4":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 13)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "In Process", "string");
                                                break;
                                            case "Start 5":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 16)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "In Process", "string");
                                                break;
                                            case "Start 6":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 19)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "In Process", "string");
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                    else if(sLocValueFromSearch.ToLower().StartsWith("p", StringComparison.OrdinalIgnoreCase))
                                    {
                                        InsertPatientNotFound(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], hicValueFromSearch, startDateValueFromSearch, patientNameValueFromSearch, "Continued");
                                        ProcessAgencyFileDates(xlsfilePathAgency[indexOfAgencies]);
                                        LogError($"Column date not found for patient {patientNameValueFromSearch} for date {startDateValueFromSearch} adding a line as Continued", logFilePath);
                                        Console.WriteLine($"Column date not found for patient {patientNameValueFromSearch} for date {startDateValueFromSearch}");
                                    }
                                }
                                else
                                {
                                    if (columnNameofFoundDate != null)
                                    {
                                        indexOfCell = FindCellLocation(xlsfilePathAgency[indexOfAgencies], hicValueFromSearch, columnNameofFoundDate, startDateValueFromSearch);
                                        alphaOfNameCell = FindCellLocationAlpha(xlsfilePathAgency[indexOfAgencies], indexOfCell);
                                        switch (columnNameofFoundDate)
                                        {
                                            case "SOC":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 3)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 4)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                                break;
                                            case "Start 2":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 6)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 7)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                                break;
                                            case "Start 3":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 9)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 10)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                                break;
                                            case "Start 4":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 12)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 13)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                                break;
                                            case "Start 5":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 15)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 16)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                                break;
                                            case "Start 6":
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 18)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 19)[0]);
                                                ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                                break;
                                            default:
                                                break;
                                        }
                                        ProcessAgencyFileDates(xlsfilePathAgency[indexOfAgencies]);
                                        LogError($"Column date found {columnNameofFoundDate} for patient {patientNameValueFromSearch} for date {startDateValueFromSearch} and Reimb value {reimbValueFromSearch}", logFilePath);
                                        Console.WriteLine($"Column date found {columnNameofFoundDate} for patient {patientNameValueFromSearch} for date {startDateValueFromSearch} and Reimb value {reimbValueFromSearch}");
                                    }
                                }
                            }
                        }
                      
                        if (cancelledDateFromSearch == "")
                        {
                            if (tobValueFromSearch == "327" || tobValueFromSearch == "32G" || tobValueFromSearch == "32I")
                            {
                                Double reimbValueFromSearchInt = Double.Parse(reimbValueFromSearch);
                                if (columnNameofFoundDate != null && reimbValueFromSearchInt != 0)
                                {
                                    LogError($"Column date found {columnNameofFoundDate} for patient {patientNameValueFromSearch} for date {startDateValueFromSearch} and Reimb value {reimbValueFromSearch}", logFilePath);
                                    Console.WriteLine($"Column date found {columnNameofFoundDate} for patient {patientNameValueFromSearch} for date {startDateValueFromSearch} and Reimb value {reimbValueFromSearch}");
                                    indexOfCell = FindCellLocation(xlsfilePathAgency[indexOfAgencies], hicValueFromSearch, columnNameofFoundDate, startDateValueFromSearch);
                                    alphaOfNameCell = FindCellLocationAlpha(xlsfilePathAgency[indexOfAgencies], indexOfCell);

                                    switch (columnNameofFoundDate)
                                    {
                                        case "SOC":
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 3)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 4)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                            break;
                                        case "Start 2":
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 6)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 7)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                            break;
                                        case "Start 3":
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 9)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 10)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                            break;
                                        case "Start 4":
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 12)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 13)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                            break;
                                        case "Start 5":
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 15)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 16)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                            break;
                                        case "Start 6":
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 18)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, reimbValueFromSearch, "double");
                                            alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 19)[0]);
                                            ModifyCell(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Paid", "string");
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            UndoModifyExcelFile(originalChangesFilename);
        }

        //public static void MainNOAProcess(string xlsfilePathChanges, string[] xlsfilePathAgency, string csvFilePath)
        //{
        //    Console.WriteLine($"Main NOA Process Initiated");
        //    // Ensure the log directory exists.
        //    Directory.CreateDirectory("ErrorLogs");

        //    // Generate a unique log file name using a timestamp.
        //    string logFileName = $"logNOAProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
        //    string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);
        //    ArchiveFiles(xlsfilePathChanges, "NOAStatus");

        //    string originalChangesFilename = ModifyExcelFile(xlsfilePathChanges, "Change");
        //    for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
        //    {
        //        ArchiveFiles(xlsfilePathAgency[indexOfAgencies], "Agencies");
        //    }

        //    string patientNameValueFromChanges = string.Empty;
        //    string admitDateValueFromChanges = string.Empty;
        //    string reasonCodeFromChanges = string.Empty;
        //    string hicValueFromChanges = string.Empty;
        //    string patientNameValueFromAgency = string.Empty;
        //    string hicValueFromAgency = string.Empty;
        //    string admitDateValueFromAgency = string.Empty;
        //    List<discrepencies> listOfDiscrepancies = new List<discrepencies>();
        //    DataTable csvData = new DataTable();
        //    DataTable resultFromChanges = new DataTable();
        //    DataTable resultFromChangesWithSorReason = new DataTable();

        //    for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
        //    {
        //        DataTable tableOfPatientsToInsert = new DataTable();
        //        string newFilePath = "RecalculatedFile.xlsx";
        //        File.Copy(xlsfilePathAgency[indexOfAgencies], newFilePath, true);

        //        csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
        //        string changesQuery = GetSqlQueryByQueryName(csvData, "Filter relevant rows for agency");
        //        string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
        //        Console.WriteLine($"Starting work on {agencyName}");
        //        changesQuery = changesQuery.Replace("AgencyName", agencyName);
        //        resultFromChanges = ExecuteExcelQuery(xlsfilePathChanges, changesQuery);
        //        resultFromChanges = ConvertDatesToDateOnly(resultFromChanges);
        //        resultFromChangesWithSorReason = RemoveRowsFromTable(resultFromChanges, 1);
        //        List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
        //        bool insertPatientAccepted = false;

        //        if (resultFromChangesWithSorReason.Rows.Count > 0)
        //        {
        //            for (int indexOfPatients = 0; indexOfPatients < resultFromChangesWithSorReason.Rows.Count; indexOfPatients++)
        //            {
        //                insertPatientAccepted = false;
        //                patientNameValueFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["Patient Name"].ToString();
        //                patientNameValueFromChanges = RemoveMiddleName(patientNameValueFromChanges);
        //                admitDateValueFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["Admit Date"].ToString();
        //                reasonCodeFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["Reason Code(s)"].ToString();
        //                hicValueFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["HIC/MBI"].ToString();

        //                var manipulateLastName = patientNameValueFromChanges.Split(",");
        //                patientNameValueFromChanges = TrimName(manipulateLastName[0]) + "%," + manipulateLastName[1];

        //                //RecalculateFormulas(xlsfilePathAgency[indexOfAgencies]);
        //                string modifiedQuery = GetSqlQueryByQueryName(csvData, "Get patient name and SOC");
        //                modifiedQuery = modifiedQuery.Replace("HICnumber", hicValueFromChanges);
        //                modifiedQuery = modifiedQuery.Replace("worksheet", sheetNamesOfAgencyFile[0]);
        //                DataTable resultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], modifiedQuery);
        //                patientNameValueFromChanges = patientNameValueFromChanges.Replace("%", "");
        //                if (resultFromAgency.Rows.Count > 0)
        //                {
        //                    resultFromAgency = ConvertDatesToDateOnly(resultFromAgency);

        //                    for (int indOfPatients = 0; indOfPatients < resultFromAgency.Rows.Count; indOfPatients++)
        //                    {
        //                        Tuple<int, int> indexOfCell = new Tuple<int, int>(0, 0);
        //                        string alphaOfNameCell = string.Empty;
        //                        string alphaOfStatusCell = string.Empty;

        //                        if (resultFromAgency.Rows.Count > 0)
        //                        {
        //                            // Get the value in the "Patient Name" column for the first row.
        //                            patientNameValueFromAgency = resultFromAgency.Rows[indOfPatients]["Patients Name"].ToString();
        //                            hicValueFromAgency = resultFromAgency.Rows[indOfPatients]["HIC/MBI"].ToString();
        //                            patientNameValueFromAgency = RemoveMiddleName(patientNameValueFromAgency);

        //                            admitDateValueFromAgency = resultFromAgency.Rows[indOfPatients]["SOC"].ToString();

        //                            if (DateTime.TryParseExact(admitDateValueFromAgency, "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
        //                            {
        //                                admitDateValueFromAgency = parsedDate.ToString("M/d/yyyy");
        //                            }

        //                            if (DateTime.TryParseExact(admitDateValueFromChanges, "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out parsedDate))
        //                            {
        //                                admitDateValueFromChanges = parsedDate.ToString("M/d/yyyy");
        //                            }

        //                            indexOfCell = FindCellLocation(xlsfilePathAgency[indexOfAgencies], hicValueFromAgency, admitDateValueFromAgency);
        //                            if (indexOfCell != null)
        //                            {
        //                                alphaOfNameCell = FindCellLocationAlpha(xlsfilePathAgency[indexOfAgencies], hicValueFromAgency, admitDateValueFromAgency);
        //                                alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 1)[0]);
        //                            }

        //                            if (hicValueFromAgency == hicValueFromChanges && admitDateValueFromAgency == admitDateValueFromChanges)
        //                            {
        //                                insertPatientAccepted = false;
        //                                if (reasonCodeFromChanges != "37263")
        //                                {
        //                                    ModifyCell(newFilePath, sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Accepted", "string");
        //                                    Console.WriteLine($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Accepted");
        //                                }
        //                                else
        //                                {
        //                                    ModifyCell(newFilePath, sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Approved", "string");
        //                                    Console.WriteLine($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Approved"); ;
        //                                }
        //                            }
        //                            else
        //                            {
        //                                // Add items to the list
        //                                insertPatientAccepted = true;
        //                            }
        //                        }
        //                        else
        //                        {
        //                            //name was not found in agency file
        //                            listOfDiscrepancies.Add(new discrepencies { Agency = agencyName, PatientName = patientNameValueFromChanges, AdmitDate = admitDateValueFromChanges });
        //                            LogError($"The following Descrepency was found:\n, Agency: Patient Name: {patientNameValueFromAgency} Admit Date: {admitDateValueFromAgency}\n, Status File: Patient Name: {patientNameValueFromChanges} Admit Date: {admitDateValueFromChanges}\n", logFilePath);
        //                        }
        //                    }
        //                    if (insertPatientAccepted)
        //                        tableOfPatientsToInsert = AddOrUpdateRow(tableOfPatientsToInsert, hicValueFromChanges, admitDateValueFromChanges, patientNameValueFromChanges);
        //                }
        //                else
        //                {
        //                    InsertPatientNotFound(newFilePath, sheetNamesOfAgencyFile[0], hicValueFromChanges, admitDateValueFromChanges, patientNameValueFromChanges, "Accepted");
        //                    Console.WriteLine($"Patient : {patientNameValueFromChanges} in Agency : {agencyName} was not found and inserted with status Accepted");
        //                    LogError($"Patient : {patientNameValueFromChanges} in Agency : {agencyName} was not found and inserted with status Accepted", logFilePath);
        //                }
        //            }
        //        }

        //        if (insertPatientAccepted)
        //        {
        //            for (int indexOfInserts = 0; indexOfInserts < tableOfPatientsToInsert.Rows.Count; indexOfInserts++)
        //            {
        //                InsertPatientNotFound(newFilePath, sheetNamesOfAgencyFile[0], tableOfPatientsToInsert.Rows[indexOfInserts]["hicValueFromChanges"].ToString(), tableOfPatientsToInsert.Rows[indexOfInserts]["admitDateValueFromChanges"].ToString(), tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString(), "Accepted");
        //                Console.WriteLine($"Patient : {tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString()} in Agency : {agencyName} was not found and inserted with status Accepted");
        //                LogError($"Patient : {tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString()} in Agency : {agencyName} was not found and inserted with status Accepted", logFilePath);
        //            }
        //        }

        //        File.Delete(xlsfilePathAgency[indexOfAgencies]);
        //        File.Move(newFilePath, xlsfilePathAgency[indexOfAgencies]);
        //    }

        //    List<discrepencies> patientsMoreThan3Days = new List<discrepencies>();
        //    for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
        //    {
        //        //check if "in process" status is over 3 days
        //        csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
        //        string inProcessQuery = GetSqlQueryByQueryName(csvData, "Get all in process patients");
        //        DataTable inProcessResultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], inProcessQuery);
        //        string agencyName = GetFileNameWithoutExtension(xlsfilePathAgency[indexOfAgencies]);
        //        patientsMoreThan3Days.AddRange(GetPatientsWith3DaysOrMoreDifference(inProcessResultFromAgency, agencyName));
        //    }

        //    //send email of discrepencies if available
        //    if (listOfDiscrepancies.Count > 0 || patientsMoreThan3Days.Count > 0)
        //    {
        //        DataTable configurationInfo = ReadCsvFile("ConfigFile.csv", "SmtpInfo", "Details");
        //        string sendToEmail = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToEmail");
        //        string sendToName = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToName");
        //        SendNOAEmailWithTemplate(sendToName, sendToEmail, listOfDiscrepancies, patientsMoreThan3Days);
        //    }

        //    UndoModifyExcelFile(originalChangesFilename);
        //} 
        
        public static void MainNOAProcess(string xlsfilePathChanges, string[] xlsfilePathAgency, string csvFilePath)
        {
            Console.WriteLine($"Main NOA Process Initiated");
            // Ensure the log directory exists.
            Directory.CreateDirectory("ErrorLogs");

            // Generate a unique log file name using a timestamp.
            string logFileName = $"logNOAProcess_{DateTime.Now:yyyyMMddHHmmssfff}.txt";
            string logFilePath = System.IO.Path.Combine("ErrorLogs", logFileName);
            ArchiveFiles(xlsfilePathChanges, "NOAStatus");

            string originalChangesFilename = ModifyExcelFile(xlsfilePathChanges, "Change");
            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                ArchiveFiles(xlsfilePathAgency[indexOfAgencies], "Agencies");
            }

            string patientNameValueFromChanges = string.Empty;
            string admitDateValueFromChanges = string.Empty;
            string reasonCodeFromChanges = string.Empty;
            string hicValueFromChanges = string.Empty;
            string patientNameValueFromAgency = string.Empty;
            string hicValueFromAgency = string.Empty;
            string admitDateValueFromAgency = string.Empty;
            List<discrepencies> listOfDiscrepancies = new List<discrepencies>();
            DataTable csvData = new DataTable();
            DataTable resultFromChanges = new DataTable();
            DataTable resultFromChangesWithSorReason = new DataTable();

            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                DataTable tableOfPatientsToInsert = new DataTable();
                string newFilePath = "RecalculatedFile.xlsx";
                File.Copy(xlsfilePathAgency[indexOfAgencies], newFilePath, true);

                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string changesQuery = GetSqlQueryByQueryName(csvData, "Filter relevant rows for agency");
                string agencyName = xlsfilePathAgency[indexOfAgencies].Replace("..\\Agencies\\", "").Replace(" HH", "").Replace(".xlsx", "");
                Console.WriteLine($"Starting work on {agencyName}");
                changesQuery = changesQuery.Replace("AgencyName", agencyName);
                resultFromChanges = ExecuteExcelQuery(xlsfilePathChanges, changesQuery);
                resultFromChanges = ConvertDatesToDateOnly(resultFromChanges);
                resultFromChangesWithSorReason = RemoveRowsFromTable(resultFromChanges, 1);
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                bool insertPatientAccepted = false;

                if (resultFromChangesWithSorReason.Rows.Count > 0)
                {
                    for (int indexOfPatients = 0; indexOfPatients < resultFromChangesWithSorReason.Rows.Count; indexOfPatients++)
                    {
                        insertPatientAccepted = false;
                        patientNameValueFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["Patient Name"].ToString();
                        patientNameValueFromChanges = RemoveMiddleName(patientNameValueFromChanges);
                        admitDateValueFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["Admit Date"].ToString();
                        reasonCodeFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["Reason Code(s)"].ToString();
                        hicValueFromChanges = resultFromChangesWithSorReason.Rows[indexOfPatients]["HIC/MBI"].ToString();

                        var manipulateLastName = patientNameValueFromChanges.Split(",");
                        patientNameValueFromChanges = TrimName(manipulateLastName[0]) + "%," + manipulateLastName[1];

                        //RecalculateFormulas(xlsfilePathAgency[indexOfAgencies]);
                        string modifiedQuery = GetSqlQueryByQueryName(csvData, "Get patient name and SOC");
                        modifiedQuery = modifiedQuery.Replace("HICnumber", hicValueFromChanges);
                        modifiedQuery = modifiedQuery.Replace("worksheet", sheetNamesOfAgencyFile[0]);
                        DataTable resultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], modifiedQuery);
                        patientNameValueFromChanges = patientNameValueFromChanges.Replace("%", "");
                        if (resultFromAgency.Rows.Count > 0)
                        {
                            resultFromAgency = ConvertDatesToDateOnly(resultFromAgency);

                            for (int indOfPatients = 0; indOfPatients < resultFromAgency.Rows.Count; indOfPatients++)
                            {
                                Tuple<int, int> indexOfCell = new Tuple<int, int>(0, 0);
                                string alphaOfNameCell = string.Empty;
                                string alphaOfStatusCell = string.Empty;

                                if (resultFromAgency.Rows.Count > 0)
                                {
                                    // Get the value in the "Patient Name" column for the first row.
                                    patientNameValueFromAgency = resultFromAgency.Rows[indOfPatients]["Patients Name"].ToString();
                                    hicValueFromAgency = resultFromAgency.Rows[indOfPatients]["HIC/MBI"].ToString();
                                    patientNameValueFromAgency = RemoveMiddleName(patientNameValueFromAgency);

                                    admitDateValueFromAgency = resultFromAgency.Rows[indOfPatients]["SOC"].ToString();

                                    if (DateTime.TryParseExact(admitDateValueFromAgency, "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out DateTime parsedDate))
                                    {
                                        admitDateValueFromAgency = parsedDate.ToString("M/d/yyyy");
                                    }

                                    if (DateTime.TryParseExact(admitDateValueFromChanges, "MM/dd/yyyy", null, System.Globalization.DateTimeStyles.None, out parsedDate))
                                    {
                                        admitDateValueFromChanges = parsedDate.ToString("M/d/yyyy");
                                    }

                                    indexOfCell = FindCellLocation(xlsfilePathAgency[indexOfAgencies], hicValueFromAgency, admitDateValueFromAgency);
                                    if (indexOfCell != null)
                                    {
                                        alphaOfNameCell = FindCellLocationAlpha(xlsfilePathAgency[indexOfAgencies], hicValueFromAgency, admitDateValueFromAgency);
                                        alphaOfStatusCell = ReplaceLetterInIndex(alphaOfNameCell, GetExcelColumnName(indexOfCell.Item2 + 1)[0]);
                                    }

                                    if (hicValueFromAgency == hicValueFromChanges && admitDateValueFromAgency == admitDateValueFromChanges)
                                    {
                                        insertPatientAccepted = false;
                                        if (reasonCodeFromChanges != "37263" || reasonCodeFromChanges == "")
                                        {
                                            ModifyCell(newFilePath, sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Accepted", "string");
                                            LogError($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Accepted", logFilePath);
                                            Console.WriteLine($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Accepted");
                                        }
                                        else
                                        {
                                            ModifyCell(newFilePath, sheetNamesOfAgencyFile[0], alphaOfStatusCell, "Approved", "string");
                                            LogError($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Approved", logFilePath);
                                            Console.WriteLine($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Approved"); ;
                                        }
                                    }
                                    else
                                    {
                                        // Add items to the list
                                        insertPatientAccepted = true;
                                    }
                                }
                                else
                                {
                                    //name was not found in agency file
                                    listOfDiscrepancies.Add(new discrepencies { Agency = agencyName, PatientName = patientNameValueFromChanges, AdmitDate = admitDateValueFromChanges });
                                    LogError($"The following Descrepency was found:\n, Agency: Patient Name: {patientNameValueFromAgency} Admit Date: {admitDateValueFromAgency}\n, Status File: Patient Name: {patientNameValueFromChanges} Admit Date: {admitDateValueFromChanges}\n", logFilePath);
                                }
                            }
                            if (insertPatientAccepted)
                                tableOfPatientsToInsert = AddOrUpdateRow(tableOfPatientsToInsert, hicValueFromChanges, admitDateValueFromChanges, patientNameValueFromChanges, reasonCodeFromChanges);
                        }
                        else
                        {
                            InsertPatientNotFound(newFilePath, sheetNamesOfAgencyFile[0], hicValueFromChanges, admitDateValueFromChanges, patientNameValueFromChanges, "Accepted");
                            Console.WriteLine($"Patient : {patientNameValueFromChanges} in Agency : {agencyName} was not found and inserted with status Accepted");
                            LogError($"Patient : {patientNameValueFromChanges} in Agency : {agencyName} was not found and inserted with status Accepted", logFilePath);
                        }
                    }
                }

                if (insertPatientAccepted)
                {
                    for (int indexOfInserts = 0; indexOfInserts < tableOfPatientsToInsert.Rows.Count; indexOfInserts++)
                    {

                        if (tableOfPatientsToInsert.Rows[indexOfInserts]["reasonCodeFromChanges"].ToString() == "37263")
                        {
                            InsertPatientNotFound(newFilePath, sheetNamesOfAgencyFile[0], tableOfPatientsToInsert.Rows[indexOfInserts]["hicValueFromChanges"].ToString(), tableOfPatientsToInsert.Rows[indexOfInserts]["admitDateValueFromChanges"].ToString(), tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString(), "Approved");
                            LogError($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Approved", logFilePath);
                            Console.WriteLine($"Patient : {patientNameValueFromAgency} in Agency : {agencyName} was found and modified to Approved");
                        }
                        else
                        {
                            InsertPatientNotFound(newFilePath, sheetNamesOfAgencyFile[0], tableOfPatientsToInsert.Rows[indexOfInserts]["hicValueFromChanges"].ToString(), tableOfPatientsToInsert.Rows[indexOfInserts]["admitDateValueFromChanges"].ToString(), tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString(), "Accepted");
                            Console.WriteLine($"Patient : {tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString()} in Agency : {agencyName} was not found and inserted with status Accepted");
                            LogError($"Patient : {tableOfPatientsToInsert.Rows[indexOfInserts]["patientNameValueFromChanges"].ToString()} in Agency : {agencyName} was not found and inserted with status Accepted", logFilePath);
                        }
                    }
                }

                File.Delete(xlsfilePathAgency[indexOfAgencies]);
                File.Move(newFilePath, xlsfilePathAgency[indexOfAgencies]);
            }

            List<discrepencies> patientsMoreThan3Days = new List<discrepencies>();
            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                //check if "in process" status is over 3 days
                csvData = ReadCsvFile(csvFilePath, "QueryName", "SqlQuery");
                string inProcessQuery = GetSqlQueryByQueryName(csvData, "Get all in process patients");
                DataTable inProcessResultFromAgency = ExecuteExcelQuery(xlsfilePathAgency[indexOfAgencies], inProcessQuery);
                string agencyName = GetFileNameWithoutExtension(xlsfilePathAgency[indexOfAgencies]);
                patientsMoreThan3Days.AddRange(GetPatientsWith3DaysOrMoreDifference(inProcessResultFromAgency, agencyName));
            }

            //send email of discrepencies if available
            if (listOfDiscrepancies.Count > 0 || patientsMoreThan3Days.Count > 0)
            {
                DataTable configurationInfo = ReadCsvFile("ConfigFile.csv", "SmtpInfo", "Details");
                string sendToEmail = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToEmail");
                string sendToName = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "sendToName");
                SendNOAEmailWithTemplate(sendToName, sendToEmail, listOfDiscrepancies, patientsMoreThan3Days);
            }

            UndoModifyExcelFile(originalChangesFilename);
        }
    }
}
