using static RCMHospice.RCMHospiceMainProcess;

namespace RCMHospice
{
    class RCMHospiceProcess
    {
        static void Main(string[] args)
        {
            #region prep
            Console.Title = "RCM Process";
            Console.ForegroundColor = ConsoleColor.Green;
            if (OperatingSystem.IsWindows())
            {
                Console.WindowHeight = 40;
            }

            if (args.Length < 7)
            {
                Console.WriteLine("Usage: RCM Process <xlsfilePathChanges.xlsx> <xlsfilePathAgency.xlsx> <QueryFile.csv> <SearchReport.xlsx> <PaymentSummary.xlsx> <AgencyList.xlsx> <Suspense.xlsx>");
                return;
            }

            string agencyDirectoryPath = args[0];
            string csvFilePath = args[1];
            string xlsFileSearchReportP = args[2];
            string xlsFilePaymentSummary = args[3];
            string xlsFileAgencyList = args[4];
            string xlsFileSuspenseReport = args[6];

            if (!Directory.Exists(agencyDirectoryPath) || !File.Exists(csvFilePath) || !File.Exists(xlsFileSearchReportP) /*|| !File.Exists(xlsFileSearchReportS)*/ || !File.Exists(xlsFilePaymentSummary) || !File.Exists(xlsFileAgencyList) || !File.Exists(xlsFileSuspenseReport))
            {
                Console.WriteLine("All input files must exist.\n\n Usage: RCM Process <xlsfilePathAgency> <QueryFile.csv> <SearchReport.xlsx> <PaymentSummary.xlsx> <AgencyList.xlsx> <Suspense.xlsx>");
                return;
            }

            // Get a list of XLSX files in the specified directory
            string[] agencyXlsxFiles = Directory.GetFiles(agencyDirectoryPath, "*.xlsx");

            if (agencyXlsxFiles.Length == 0)
            {
                Console.WriteLine("No XLSX files found in the specified directory.");
                return;
            }


            //HIC Process
            /*string xlsxHICFile = string.Empty;
            if (args.Length >= 8) // Check if there are at least 8 arguments (0-based index)
            {
                xlsxHICFile = args[7];

                if (!File.Exists(xlsxHICFile))
                {
                    Console.WriteLine("HIC File Not Found");
                    return;
                }

                HICPullProcess(agencyXlsxFiles, xlsxHICFile, csvFilePath);
            }*/
            #endregion

            RCMHospiceHelpers.RunStep("Main NOA Process", () => MainNOAProcess(xlsFileSearchReportP, agencyXlsxFiles, csvFilePath));
            RCMHospiceHelpers.RunStep("Final Process", () => FinalProcess(xlsFileSearchReportP, agencyXlsxFiles, csvFilePath));
            RCMHospiceHelpers.RunStep("Future Payment Process", () => FuturePaymentProcess(xlsFileSearchReportP, agencyXlsxFiles, csvFilePath));
            RCMHospiceHelpers.RunStep("Payment Summary Process", () => PaymentSummaryProcess(xlsFilePaymentSummary, agencyXlsxFiles, csvFilePath));
            RCMHospiceHelpers.RunStep("Suspense Process", () => SuspenseProcess(xlsFileSuspenseReport, agencyXlsxFiles, csvFilePath));
            RCMHospiceHelpers.RunStep("Move Future Payment To Summary Process", () => MoveFuturePaymentToSummaryProcess(agencyXlsxFiles, csvFilePath));

            #region change to proper case
            RCMHospiceHelpers.NamesToProperCaseOnAllAgencies(agencyXlsxFiles);
            #endregion

            if (DateTime.Today.DayOfWeek != DayOfWeek.Saturday && DateTime.Today.DayOfWeek != DayOfWeek.Sunday)
            {
                RCMHospiceHelpers.RunStep("Email Agency List Process", () => EmailAgencyList(xlsFileAgencyList, agencyXlsxFiles, csvFilePath));
            }

            RCMHospiceHelpers.RunStep("Future Summary Process", () => FutureSummaryProcess(xlsFileSearchReportP, agencyXlsxFiles, csvFilePath));
            Environment.Exit(0);
        }
    }
}