using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Data;
using System.Data.OleDb;
using ClosedXML.Excel;
using System.Text.RegularExpressions;
using System.Text;
using MimeKit;
using MailKit.Net.Smtp;
using System.Globalization;
using Excel = Microsoft.Office.Interop.Excel;
using MailKit.Security;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Util.Store;
using Google.Apis.Gmail.v1;
using Google.Apis.Services;
using Google.Apis.Util;
using Twilio;
using Twilio.Rest.Api.V2010.Account;
using Twilio.Types;
using System.Runtime.InteropServices;
using Path = System.IO.Path;
using System.Diagnostics;
using Microsoft.VisualBasic.FileIO;


namespace RCMHospice
{
    public static class RCMHospiceHelpers
    {
        public static int GetSheetIndexByName(List<string> sheetNamesOfAgencyFile, string sheetName)
        {
            if (sheetNamesOfAgencyFile == null)
                throw new ArgumentNullException(nameof(sheetNamesOfAgencyFile));

            if (string.IsNullOrWhiteSpace(sheetName))
                throw new ArgumentException("Sheet name cannot be empty.", nameof(sheetName));

            for (int i = 0; i < sheetNamesOfAgencyFile.Count; i++)
            {
                if (sheetNamesOfAgencyFile[i].Trim().Equals(sheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                    return i;
            }

            throw new Exception($"Worksheet '{sheetName}' was not found in sheetNamesOfAgencyFile.");
        }

        public static (int capYear, string sheetName) GetCapYearAndSheet(DateTime date)
        {
            int capYear = date.Month >= 10 ? date.Year + 1 : date.Year;
            string sheetName = $"Claims FY{capYear.ToString().Substring(2)}";

            return (capYear, sheetName);
        }

        public static ExcelWorksheet EnsureClaimsYearSheetExists(ExcelPackage package, string sheetName)
        {
            ExcelWorksheet existingSheet = package.Workbook.Worksheets[sheetName];

            if (existingSheet != null)
                return existingSheet;

            ExcelWorksheet templateSheet = package.Workbook.Worksheets["Claims"];

            if (templateSheet == null)
                throw new Exception($"Template sheet 'Claims' was not found. Cannot create {sheetName}.");

            ExcelWorksheet newSheet = package.Workbook.Worksheets.Add(sheetName, templateSheet);
            newSheet.Hidden = eWorkSheetHidden.Visible;

            ClearClaimsTemplateData(newSheet);

            return newSheet;
        }

        public static void ClearClaimsTemplateData(ExcelWorksheet ws)
        {
            if (ws.Dimension == null)
                return;

            int lastRow = ws.Dimension.End.Row;
            int lastCol = ws.Dimension.End.Column;

            for (int row = 2; row <= lastRow; row++)
            {
                for (int col = 1; col <= lastCol; col++)
                {
                    ws.Cells[row, col].Value = null;
                }
            }
        }

        public static int GetColumnByHeader(ExcelWorksheet ws, string headerName)
        {
            if (ws.Dimension == null)
                return -1;

            int lastCol = ws.Dimension.End.Column;

            for (int col = 1; col <= lastCol; col++)
            {
                string header = ws.Cells[1, col].Text.Trim();

                if (string.Equals(header, headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            return -1;
        }

        public static List<int> GetMonthColumns(ExcelWorksheet ws)
        {
            List<int> monthColumns = new List<int>();

            if (ws.Dimension == null)
                return monthColumns;

            int lastCol = ws.Dimension.End.Column;

            for (int col = 1; col <= lastCol; col++)
            {
                string header = ws.Cells[1, col].Text.Trim();

                if (header.StartsWith("Month", StringComparison.OrdinalIgnoreCase))
                    monthColumns.Add(col);
            }

            return monthColumns;
        }

        public static int FindRowByHicMbi(ExcelWorksheet ws, string hicMbi, int hicMbiCol)
        {
            if (ws.Dimension == null || hicMbiCol <= 0)
                return -1;

            int lastRow = ws.Dimension.End.Row;

            for (int row = 2; row <= lastRow; row++)
            {
                string existingHic = ws.Cells[row, hicMbiCol].Text.Trim();

                if (string.Equals(existingHic, hicMbi.Trim(), StringComparison.OrdinalIgnoreCase))
                    return row;
            }

            return -1;
        }

        public static int FindRowByHicAndDate(ExcelWorksheet ws, string hicMbi, DateTime dateToFind, int hicMbiCol, List<int> monthColumns)
        {
            if (ws.Dimension == null || hicMbiCol <= 0)
                return -1;

            int lastRow = ws.Dimension.End.Row;

            for (int row = 2; row <= lastRow; row++)
            {
                string existingHic = ws.Cells[row, hicMbiCol].Text.Trim();

                if (!string.Equals(existingHic, hicMbi.Trim(), StringComparison.OrdinalIgnoreCase))
                    continue;

                foreach (int monthCol in monthColumns)
                {
                    if (DateTime.TryParse(ws.Cells[row, monthCol].Text, out DateTime cellDate))
                    {
                        if (cellDate.Date == dateToFind.Date)
                            return row;
                    }
                }
            }

            return -1;
        }

        public static int FindMatchingMonthColumn(ExcelWorksheet ws, int row, DateTime dateToFind, List<int> monthColumns)
        {
            foreach (int monthCol in monthColumns)
            {
                if (DateTime.TryParse(ws.Cells[row, monthCol].Text, out DateTime cellDate))
                {
                    if (cellDate.Date == dateToFind.Date)
                        return monthCol;
                }
            }

            return -1;
        }

        public static DataTable FilterFinalSearchReportRows(DataTable searchReportTable, string agencyName)
        {
            DataTable filteredTable = searchReportTable.Clone();

            string[] validTobs = { "811", "812", "813", "814", "817", "81G", "81I" };

            foreach (DataRow row in searchReportTable.Rows)
            {
                string agency = row.Table.Columns.Contains("Agency") ? row["Agency"].ToString().Trim() : "";
                string tob = row.Table.Columns.Contains("TOB") ? row["TOB"].ToString().Trim() : "";

                bool agencyMatches = agency.IndexOf(agencyName, StringComparison.OrdinalIgnoreCase) >= 0;
                bool tobMatches = validTobs.Any(x => tob.IndexOf(x, StringComparison.OrdinalIgnoreCase) >= 0);

                if (agencyMatches && tobMatches)
                    filteredTable.ImportRow(row);
            }

            return filteredTable;
        }

        public static double ParseDoubleSafe(string value)
        {
            if (double.TryParse(value, out double result))
                return result;

            return 0;
        }

        public static DataTable FilterNOESearchReportRows(DataTable searchReportTable, string agencyName)
        {
            DataTable filteredTable = searchReportTable.Clone();

            foreach (DataRow row in searchReportTable.Rows)
            {
                string agency = row.Table.Columns.Contains("Agency")
                    ? row["Agency"].ToString().Trim()
                    : "";

                string tob = row.Table.Columns.Contains("TOB")
                    ? row["TOB"].ToString().Trim()
                    : "";

                bool agencyMatches = agency.IndexOf(agencyName, StringComparison.OrdinalIgnoreCase) >= 0;

                bool tobMatches =
                    tob.IndexOf("81A", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    tob.IndexOf("81C", StringComparison.OrdinalIgnoreCase) >= 0;

                if (agencyMatches && tobMatches)
                {
                    filteredTable.ImportRow(row);
                }
            }

            return filteredTable;
        }

        public static void ForceColumnToText(string excelFilePath, string columnHeaderName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            var file = new FileInfo(excelFilePath);

            using (var package = new ExcelPackage(file))
            {
                var ws = package.Workbook.Worksheets[0];
                if (ws == null || ws.Dimension == null)
                    return;

                int startRow = ws.Dimension.Start.Row;
                int endRow = ws.Dimension.End.Row;
                int startCol = ws.Dimension.Start.Column;
                int endCol = ws.Dimension.End.Column;

                // 1. Find the column index by header name
                int targetCol = -1;

                for (int col = startCol; col <= endCol; col++)
                {
                    string header = ws.Cells[startRow, col].Text.Trim();

                    if (string.Equals(header, columnHeaderName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetCol = col;
                        break;
                    }
                }

                if (targetCol == -1)
                {
                    Console.WriteLine($"Column '{columnHeaderName}' not found.");
                    return;
                }

                // 2. Force column format to text
                ws.Cells[startRow, targetCol, endRow, targetCol].Style.Numberformat.Format = "@";

                // 3. Rewrite values as text (this is the KEY step)
                for (int row = startRow + 1; row <= endRow; row++)
                {
                    var cell = ws.Cells[row, targetCol];

                    string text = cell.Text?.Trim();

                    if (!string.IsNullOrEmpty(text))
                    {
                        cell.Value = text;
                    }
                }

                package.Save();
            }
        }
        
        public static int GetNextPatientRow(ExcelWorksheet ws, int patientsNameCol)
        {
            if (ws.Dimension == null)
                return 2;

            int lastRow = ws.Dimension.End.Row;

            for (int row = 2; row <= lastRow; row++)
            {
                string patientName = ws.Cells[row, patientsNameCol].Text.Trim();

                if (string.IsNullOrWhiteSpace(patientName))
                    return row;
            }

            return lastRow + 1;
        }

        public static List<DateTime> GenerateMonthDates(DateTime socDate, int capYear)
        {
            List<DateTime> months = new List<DateTime>();
            months.Add(socDate);

            DateTime nextMonth = new DateTime(socDate.Year, socDate.Month, 1).AddMonths(1);
            DateTime endMonth = new DateTime(capYear, 9, 1);

            while (nextMonth <= endMonth)
            {
                months.Add(nextMonth);
                nextMonth = nextMonth.AddMonths(1);
            }

            return months;
        }

        public static void WriteMonthDatesToMonthColumns(ExcelWorksheet ws, int row, List<DateTime> monthDates, List<int> monthColumns)
        {
            int count = Math.Min(monthDates.Count, monthColumns.Count);

            for (int i = 0; i < count; i++)
            {
                ws.Cells[row, monthColumns[i]].Value = monthDates[i];
                ws.Cells[row, monthColumns[i]].Style.Numberformat.Format = "m/d/yyyy";
            }
        }

        public static bool HasAnyScheduledNumberForToday(DataTable resultFromPaymentSummary)
        {
            if (resultFromPaymentSummary == null)
                throw new ArgumentNullException(nameof(resultFromPaymentSummary));

            if (!resultFromPaymentSummary.Columns.Contains("Pay Date"))
                throw new Exception("Column 'Pay Date' was not found.");

            if (!resultFromPaymentSummary.Columns.Contains("Scheduled"))
                throw new Exception("Column 'Scheduled' was not found.");

            DateTime today = DateTime.Today;

            foreach (DataRow row in resultFromPaymentSummary.Rows)
            {
                if (row == null)
                    continue;

                string payDateText = row["Pay Date"]?.ToString()?.Trim();
                string scheduledText = row["Scheduled"]?.ToString()?.Trim();

                if (string.IsNullOrWhiteSpace(payDateText) || string.IsNullOrWhiteSpace(scheduledText))
                    continue;

                if (!DateTime.TryParse(payDateText, out DateTime payDate))
                    continue;

                if (payDate.Date != today)
                    continue;

                string cleanedScheduled = scheduledText.Replace("$", "").Replace(",", "").Trim();

                if (decimal.TryParse(cleanedScheduled, out decimal scheduledAmount) && scheduledAmount != 0)
                    return true;
            }

            return false;
        }

        public static void DeleteCsvFiles(string countFilePath, string totalFilePath)
        {
            DeleteFileSafe(countFilePath);
            DeleteFileSafe(totalFilePath);
        }

        private static void DeleteFileSafe(string filePath)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(filePath))
                    return;

                if (File.Exists(filePath))
                {
                    File.Delete(filePath);
                    Console.WriteLine($"Deleted file: {filePath}");
                }
                else
                {
                    Console.WriteLine($"File not found (skip delete): {filePath}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to delete file '{filePath}': {ex.Message}");
            }
        }

        public static string GetCompanyFileNameFromMappingFile()
        {
            if (!File.Exists(RCMHospiceProcess.MappingFilePath))
                throw new FileNotFoundException("MappingFile.csv was not found.", RCMHospiceProcess.MappingFilePath);

            var lines = File.ReadAllLines(RCMHospiceProcess.MappingFilePath);

            foreach (var line in lines)
            {
                if (line.StartsWith("companyFileName|", StringComparison.OrdinalIgnoreCase))
                {
                    var parts = line.Split('|');

                    if (parts.Length < 2)
                        throw new Exception("companyFileName line is malformed.");

                    return parts[1].Trim();
                }
            }

            throw new Exception("companyFileName| line not found in MappingFile.csv");
        }

        public static void PopulateCapTabFromCsvFiles(string xlsxFilePath, string csvFilePath1, string csvFilePath2)
        {
            if (string.IsNullOrWhiteSpace(xlsxFilePath))
                throw new ArgumentException("xlsxFilePath is required.", nameof(xlsxFilePath));

            if (string.IsNullOrWhiteSpace(csvFilePath1))
                throw new ArgumentException("csvFilePath1 is required.", nameof(csvFilePath1));

            if (string.IsNullOrWhiteSpace(csvFilePath2))
                throw new ArgumentException("csvFilePath2 is required.", nameof(csvFilePath2));

            if (!File.Exists(xlsxFilePath))
                throw new FileNotFoundException("Excel file not found.", xlsxFilePath);

            if (!File.Exists(csvFilePath1))
                throw new FileNotFoundException("CSV file not found.", csvFilePath1);

            if (!File.Exists(csvFilePath2))
                throw new FileNotFoundException("CSV file not found.", csvFilePath2);

            string countCsvPath = null;
            string totalCsvPath = null;

            foreach (string csvPath in new[] { csvFilePath1, csvFilePath2 })
            {
                string fileName = Path.GetFileName(csvPath);

                if (fileName.IndexOf("count", StringComparison.OrdinalIgnoreCase) >= 0)
                    countCsvPath = csvPath;
                else if (fileName.IndexOf("total", StringComparison.OrdinalIgnoreCase) >= 0)
                    totalCsvPath = csvPath;
            }

            if (string.IsNullOrWhiteSpace(countCsvPath))
                throw new Exception("Could not identify the COUNT CSV. One filename must contain 'count'.");

            if (string.IsNullOrWhiteSpace(totalCsvPath))
                throw new Exception("Could not identify the TOTAL CSV. One filename must contain 'total'.");

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(xlsxFilePath)))
            {
                ExcelWorksheet ws = package.Workbook.Worksheets["Cap"];
                if (ws == null)
                    throw new Exception("Worksheet 'Cap' was not found.");

                int yearColumn = GetWorksheetColumnByHeader(ws, "Year");
                int psrColumn = GetWorksheetColumnByHeader(ws, "PS&R");
                int totalColumn = GetWorksheetColumnByHeader(ws, "Total");

                Dictionary<int, int> yearToRowMap = BuildYearToRowMap(ws, yearColumn);

                PopulateCapFromCountCsv(ws, yearToRowMap, psrColumn, countCsvPath);
                PopulateCapFromTotalCsv(ws, yearToRowMap, totalColumn, totalCsvPath);

                package.Save();
            }
        }

        private static void PopulateCapFromCountCsv(ExcelWorksheet ws, Dictionary<int, int> yearToRowMap, int destinationColumn, string countCsvPath)
        {
            List<Dictionary<string, string>> rows = ReadCsvAsDictionaries(countCsvPath);

            if (rows.Count == 0)
                throw new Exception("Count CSV is empty: " + countCsvPath);

            foreach (var row in rows)
            {
                string capYearText = GetCsvValue(row, "Cap Year");
                string beneficiaryCountText = GetCsvValue(row, "Total Beneficiary Count");

                if (!int.TryParse(capYearText, out int capYear))
                    continue;

                if (!TryParseDecimal(beneficiaryCountText, out decimal beneficiaryCount))
                    continue;

                if (!yearToRowMap.TryGetValue(capYear, out int targetRow))
                    continue;

                ws.Cells[targetRow, destinationColumn].Value = beneficiaryCount;
            }
        }

        private static void PopulateCapFromTotalCsv(ExcelWorksheet ws, Dictionary<int, int> yearToRowMap, int destinationColumn, string totalCsvPath)
        {
            List<Dictionary<string, string>> rows = ReadCsvAsDictionaries(totalCsvPath);

            if (rows.Count == 0)
                throw new Exception("Total CSV is empty: " + totalCsvPath);

            Dictionary<int, decimal> totalsByCapYear = new Dictionary<int, decimal>();

            foreach (var row in rows)
            {
                string serviceFromText = GetCsvValue(row, "Service From");
                string serviceThroughText = GetCsvValue(row, "Service Through");
                string amountText = GetCsvValue(row, "Net Reimbursement");
                string revenueCodeText = GetCsvValue(row, "Revenue Code");

                if (string.IsNullOrWhiteSpace(serviceFromText) ||
                    string.IsNullOrWhiteSpace(serviceThroughText) ||
                    string.IsNullOrWhiteSpace(amountText))
                    continue;

                if (!TryParseDate(serviceFromText, out DateTime serviceFrom))
                    continue;

                if (!TryParseDate(serviceThroughText, out DateTime serviceThrough))
                    continue;

                if (!TryParseDecimal(amountText, out decimal amount))
                    continue;

                if (!string.Equals(revenueCodeText, "**SUM**", StringComparison.OrdinalIgnoreCase))
                    continue;

                int capYear = GetCapYearFromRange(serviceFrom, serviceThrough);
                if (capYear == 0)
                    continue;

                if (!totalsByCapYear.ContainsKey(capYear))
                    totalsByCapYear[capYear] = 0m;

                totalsByCapYear[capYear] += amount;
            }

            foreach (var kvp in totalsByCapYear)
            {
                if (!yearToRowMap.TryGetValue(kvp.Key, out int targetRow))
                    continue;

                ws.Cells[targetRow, destinationColumn].Value = kvp.Value;
            }
        }

        private static int GetCapYearFromRange(DateTime fromDate, DateTime throughDate)
        {
            if (fromDate.Month == 10 &&
                fromDate.Day == 1 &&
                throughDate.Month == 9 &&
                throughDate.Day == 30 &&
                throughDate.Year == fromDate.Year + 1)
            {
                return throughDate.Year;
            }

            return 0;
        }

        private static Dictionary<int, int> BuildYearToRowMap(ExcelWorksheet ws, int yearColumn)
        {
            var map = new Dictionary<int, int>();

            if (ws.Dimension == null)
                throw new Exception("Worksheet 'Cap' is empty.");

            int startRow = ws.Dimension.Start.Row + 1;
            int endRow = ws.Dimension.End.Row;

            for (int row = startRow; row <= endRow; row++)
            {
                string yearText = ws.Cells[row, yearColumn].Text?.Trim();

                if (int.TryParse(yearText, out int year))
                {
                    if (!map.ContainsKey(year))
                        map.Add(year, row);
                }
            }

            return map;
        }

        private static int GetWorksheetColumnByHeader(ExcelWorksheet ws, string headerName)
        {
            if (ws.Dimension == null)
                throw new Exception("Worksheet has no data.");

            int headerRow = ws.Dimension.Start.Row;
            int startCol = ws.Dimension.Start.Column;
            int endCol = ws.Dimension.End.Column;

            for (int col = startCol; col <= endCol; col++)
            {
                string text = ws.Cells[headerRow, col].Text?.Trim();
                if (string.Equals(text, headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }

            throw new Exception($"Header '{headerName}' was not found in worksheet '{ws.Name}'.");
        }

        private static List<Dictionary<string, string>> ReadCsvAsDictionaries(string csvPath)
        {
            var result = new List<Dictionary<string, string>>();

            using (var parser = new TextFieldParser(csvPath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");
                parser.HasFieldsEnclosedInQuotes = true;
                parser.TrimWhiteSpace = false;

                if (parser.EndOfData)
                    return result;

                string[] headers = parser.ReadFields();
                if (headers == null || headers.Length == 0)
                    return result;

                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields() ?? Array.Empty<string>();
                    var row = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

                    for (int i = 0; i < headers.Length; i++)
                    {
                        string header = headers[i]?.Trim() ?? "";
                        string value = i < fields.Length ? fields[i]?.Trim() ?? "" : "";
                        row[header] = value;
                    }

                    result.Add(row);
                }
            }

            return result;
        }

        private static string GetCsvValue(Dictionary<string, string> row, string headerName)
        {
            foreach (var kvp in row)
            {
                if (string.Equals(kvp.Key?.Trim(), headerName, StringComparison.OrdinalIgnoreCase))
                    return kvp.Value?.Trim() ?? "";
            }

            return "";
        }

        private static bool TryParseDate(string input, out DateTime date)
        {
            date = DateTime.MinValue;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            string[] formats =
            {
        "M/d/yyyy",
        "MM/dd/yyyy",
        "M/d/yy",
        "MM/dd/yy"
    };

            return DateTime.TryParseExact(
                input.Trim(),
                formats,
                CultureInfo.InvariantCulture,
                DateTimeStyles.None,
                out date);
        }

        private static bool TryParseDecimal(string input, out decimal value)
        {
            value = 0m;

            if (string.IsNullOrWhiteSpace(input))
                return false;

            input = input.Replace("$", "").Replace(",", "").Trim();

            return decimal.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
        }

        public static bool RunEIDMReportsDownloader(out string errorMessage)
        {
            errorMessage = string.Empty;

            try
            {
                if (!File.Exists(RCMHospiceProcess.EIDMReportsDllPath))
                {
                    errorMessage = $"DLL not found: {RCMHospiceProcess.EIDMReportsDllPath}";
                    return false;
                }

                string vstestPath = GetVsTestConsolePath();

                StringBuilder outputBuilder = new StringBuilder();
                StringBuilder errorBuilder = new StringBuilder();

                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = vstestPath,
                        Arguments = $"\"{RCMHospiceProcess.EIDMReportsDllPath}\" /Tests:LoginAndDownloadReports",
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        UseShellExecute = false,
                        CreateNoWindow = true
                    },
                    EnableRaisingEvents = true
                };

                process.OutputDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(e.Data))
                    {
                        Console.WriteLine(e.Data);
                        outputBuilder.AppendLine(e.Data);
                    }
                };

                process.ErrorDataReceived += (s, e) =>
                {
                    if (!string.IsNullOrWhiteSpace(e.Data))
                    {
                        Console.WriteLine("ERR: " + e.Data);
                        errorBuilder.AppendLine(e.Data);
                    }
                };

                process.Start();
                process.BeginOutputReadLine();
                process.BeginErrorReadLine();

                int timeoutMilliseconds = 10 * 60 * 1000; // 10 minutes

                bool exited = process.WaitForExit(timeoutMilliseconds);

                if (!exited)
                {
                    try
                    {
                        Console.WriteLine("EIDM downloader is stuck. Killing vstest process tree...");

                        process.Kill(entireProcessTree: true);
                        process.WaitForExit();
                    }
                    catch (Exception killEx)
                    {
                        Console.WriteLine("Failed to kill EIDM downloader process: " + killEx.Message);
                    }

                    errorMessage = "EIDM downloader timed out and was killed.";
                    return false;
                }

                // Important: lets async output/error readers finish flushing
                process.WaitForExit();

                if (process.ExitCode != 0)
                {
                    string output = outputBuilder.ToString();
                    string errors = errorBuilder.ToString();

                    errorMessage =
                        $"vstest failed with exit code {process.ExitCode}" +
                        Environment.NewLine +
                        errors +
                        Environment.NewLine +
                        output;

                    return false;
                }

                return true;
            }
            catch (Exception ex)
            {
                errorMessage = ex.Message;
                return false;
            }
        }

        private static string GetVsTestConsolePath()
        {
            string vswherePath = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86),
                @"Microsoft Visual Studio\Installer\vswhere.exe");

            if (!File.Exists(vswherePath))
                throw new Exception("vswhere.exe not found.");

            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = vswherePath,
                    Arguments = "-latest -products * -requires Microsoft.VisualStudio.PackageGroup.TestTools.Core -find Common7\\IDE\\Extensions\\TestPlatform\\vstest.console.exe",
                    RedirectStandardOutput = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                }
            };

            process.Start();
            string output = process.StandardOutput.ReadToEnd().Trim();
            process.WaitForExit();

            if (string.IsNullOrWhiteSpace(output) || !File.Exists(output))
                throw new Exception("Could not locate vstest.console.exe.");

            return output;
        }

        public static void WriteCompanyCodeToMappingFile(string ccnNumber)
        {
            if (string.IsNullOrWhiteSpace(ccnNumber))
                throw new ArgumentException("CCN number cannot be null or empty.", nameof(ccnNumber));

            if (!File.Exists(RCMHospiceProcess.MappingFilePath))
                throw new FileNotFoundException("MappingFile.csv was not found.", RCMHospiceProcess.MappingFilePath);

            var lines = File.ReadAllLines(RCMHospiceProcess.MappingFilePath);
            bool found = false;

            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].StartsWith("companyCode|", StringComparison.OrdinalIgnoreCase))
                {
                    lines[i] = $"companyCode|{ccnNumber}";
                    found = true;
                    break;
                }
            }

            if (!found)
                throw new Exception("companyCode line not found in MappingFile.csv");

            File.WriteAllLines(RCMHospiceProcess.MappingFilePath, lines);
        }

        public static string GetCcnFromCapTab(string xlsxFilePath)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(xlsxFilePath)))
            {
                var worksheet = package.Workbook.Worksheets["Cap"];
                if (worksheet == null)
                    throw new Exception("Worksheet 'Cap' was not found.");

                int startRow = worksheet.Dimension?.Start.Row ?? 1;
                int endRow = worksheet.Dimension?.End.Row ?? 1;
                int startCol = worksheet.Dimension?.Start.Column ?? 1;
                int endCol = worksheet.Dimension?.End.Column ?? 1;

                for (int row = startRow; row <= endRow; row++)
                {
                    for (int col = startCol; col <= endCol; col++)
                    {
                        string cellValue = worksheet.Cells[row, col].Text?.Trim();

                        if (string.Equals(cellValue, "CCN", StringComparison.OrdinalIgnoreCase))
                        {
                            for (int nextCol = col + 1; nextCol <= endCol; nextCol++)
                            {
                                string ccnValue = worksheet.Cells[row, nextCol].Text?.Trim();

                                if (!string.IsNullOrWhiteSpace(ccnValue))
                                    return ccnValue;
                            }

                            return string.Empty;
                        }
                    }
                }

                return string.Empty;
            }
        }

        public static void NamesToProperCaseOnAllAgencies(string[] xlsfilePathAgency)
        {
            for (int indexOfAgencies = 0; indexOfAgencies < xlsfilePathAgency.Length; indexOfAgencies++)
            {
                List<string> sheetNamesOfAgencyFile = GetSheetNames(xlsfilePathAgency[indexOfAgencies]);
                //change all Patien's name to proper case on all relevant worksheets
                ChangePatientsNameToProperCase(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[0]);
                ChangePatientsNameToProperCase(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[2]);
                ChangePatientsNameToProperCase(xlsfilePathAgency[indexOfAgencies], sheetNamesOfAgencyFile[3]);
            }
        }

        public static void ClearTableLeaveHeader(string excelFilePath)
        {
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException(excelFilePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(excelFilePath));

            var ws = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.Equals("In Process", StringComparison.OrdinalIgnoreCase));
            if (ws == null)
                throw new Exception("Worksheet 'In Process' not found.");

            // CLEAR ALL DATA BELOW HEADER
            if (ws.Dimension != null && ws.Dimension.End.Row > 1)
            {
                ws.Cells[2, 1, ws.Dimension.End.Row, ws.Dimension.End.Column].Clear();
            }

            package.Save();
        }

        public static void WriteToInProcessTab(string excelFilePath, string patientNameFromSuspense, string startDateFromSuspense, string submitDateValueFromSuspense, string reimbValueFromSuspense)
        {
            if (!File.Exists(excelFilePath))
                throw new FileNotFoundException(excelFilePath);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(new FileInfo(excelFilePath));

            var ws = package.Workbook.Worksheets.FirstOrDefault(w => w.Name.Equals("In Process", StringComparison.OrdinalIgnoreCase));
            if (ws == null)
                throw new Exception("Worksheet 'In Process' not found.");

            const int headerRow = 1;

            int patientCol = GetColumn(ws, headerRow, "PATIENTS NAME");
            int startCol = GetColumn(ws, headerRow, "Start Date");
            int submitCol = GetColumn(ws, headerRow, "Submit Date");
            int reimbCol = GetColumn(ws, headerRow, "Projected Reimb");
            int paidDateCol = GetColumn(ws, headerRow, "Projected Paid Date");
            int notesCol = GetColumn(ws, headerRow, "Notes");

            int row = ws.Dimension?.End.Row >= 2 ? ws.Dimension.End.Row + 1 : 2;

            ws.Cells[row, patientCol].Value = patientNameFromSuspense;
            ws.Cells[row, startCol].Value = startDateFromSuspense;
            ws.Cells[row, submitCol].Value = submitDateValueFromSuspense;
            ws.Cells[row, reimbCol].Value = reimbValueFromSuspense;

            DateTime today = DateTime.Today;

            if (DateTime.TryParse(submitDateValueFromSuspense, out DateTime submitDate))
            {
                DateTime projectedPaid = submitDate.AddDays(14);

                if (today > projectedPaid)
                {
                    ws.Cells[row, paidDateCol].Value = "";
                    ws.Cells[row, notesCol].Value = "Past Due";
                }
                else
                {
                    ws.Cells[row, paidDateCol].Value = projectedPaid;
                    ws.Cells[row, paidDateCol].Style.Numberformat.Format = "MM/dd/yyyy";
                }
            }

            package.Save();
        }

        private static int GetColumn(ExcelWorksheet ws, int headerRow, string header)
        {
            for (int c = 1; c <= ws.Dimension.End.Column; c++)
                if ((ws.Cells[headerRow, c].Text ?? "").Trim().Equals(header, StringComparison.OrdinalIgnoreCase))
                    return c;

            throw new Exception($"Header '{header}' not found.");
        }

        public static void ArchiveFiles(string xlsfilePath, string folderName)
        {
            if (!Directory.Exists($"..\\{folderName}\\Archive\\"))
            {
                // Create the folder
                Directory.CreateDirectory($"..\\{folderName}\\Archive\\");
                Console.WriteLine("Folder created successfully.");
            }
  
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            FileInfo originalFileInfo = new FileInfo(xlsfilePath);
            string originalFileName = originalFileInfo.Name;
            string archiveFolder = $"..\\{folderName}\\Archive\\{originalFileInfo.Name}_orignal{timestamp}{originalFileInfo.Extension}";
            File.Copy(xlsfilePath, archiveFolder);
        }

        public static string FindNextStartOrSOCDateColumn(string hicValue, DataTable dataTable, string dateString)
        {
            // Find all rows that contain the patient name.
            DataRow[] rows = dataTable.Select($"[HIC/MBI] LIKE '{hicValue}'");
            if (rows.Length == 0)
            {
                return null;
            }

            // Try to parse the date string into a DateTime object.
            DateTime dateToFind;
            if (!DateTime.TryParseExact(dateString, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateToFind))
            {
                // If the date string cannot be parsed, return null.
                if (!DateTime.TryParseExact(dateString, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out dateToFind))
                {
                    return null;
                }
            }

            // Iterate through the filtered rows for the patient name.
            foreach (DataRow row in rows)
            {
                foreach (DataColumn column in dataTable.Columns)
                {
                    string inputDateString = row[column].ToString();
                    string dateWithoutTime = inputDateString.Split(' ')[0]; // Remove the time component
                    if (DateTime.TryParseExact(dateWithoutTime, "M/d/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime rowDate))
                    {
                        if (rowDate == dateToFind)
                        {
                            return dataTable.Columns[column.Ordinal].ColumnName;
                        }
                    }
                }
            }

            // Return null if the specific date is not found.
            return null;
        }


        public static DataTable RemoveRowsFromTable(DataTable dataTable, int rowIndexInTable)
        {
            DataTable filteredDataTable = dataTable.Clone(); // Create a new DataTable with the same structure

            foreach (DataRow row in dataTable.Rows)
            {
                string valueS = row[1].ToString();
                string valueReason = row[4].ToString();

                if (valueS.StartsWith("S", StringComparison.OrdinalIgnoreCase))
                {
                    // Add the row to the filteredDataTable if it starts with "S"
                    filteredDataTable.ImportRow(row);
                }

                if (valueReason.Contains("37263", StringComparison.OrdinalIgnoreCase))
                {
                    // Add the row to the filteredDataTable if it has reasoncode 37263
                    filteredDataTable.ImportRow(row);
                }
            }
            return filteredDataTable;
        }

        public static string GetFileNameWithoutExtension(string filePath)
        {
            try
            {
                // Use Path.GetFileNameWithoutExtension to extract the file name without extension
                string fileNameWithoutExtension = System.IO.Path.GetFileNameWithoutExtension(filePath);
                return fileNameWithoutExtension;
            }
            catch (Exception ex)
            {
                // Handle any exceptions that may occur during the process
                Console.WriteLine($"Error extracting file name without extension: {ex.Message}");
                return null;
            }
        }

        public static List<discrepencies> GetPatientsWith3DaysOrMoreDifference(DataTable inProcessResultFromAgency, string agencyName)
        {
            List<discrepencies> patientsWith3DaysOrMoreDifference = new List<discrepencies>();

            // Assuming "NOA" is the column name for the SOC date in your DataTable
            string socColumnName = "NOA";

            foreach (DataRow row in inProcessResultFromAgency.Rows)
            {
                if (row[socColumnName] != DBNull.Value && row[socColumnName] is DateTime)
                {
                    DateTime socDate = (DateTime)row[socColumnName];
                    DateTime currentDate = DateTime.Now;

                    // Calculate the difference in days
                    TimeSpan difference = currentDate - socDate;
                    int daysDifference = (int)difference.TotalDays;

                    // Check if the difference is 3 days or more
                    if (daysDifference >= 3)
                    {
                        // Assuming "patient name" is the column name for the patient name in your DataTable
                        string patientName = row["PATIENTS NAME"].ToString();
                        patientsWith3DaysOrMoreDifference.Add(new discrepencies { Agency = agencyName, PatientName = patientName });
                    }
                }
            }
            return patientsWith3DaysOrMoreDifference;
        }

        public static void ChangePatientsNameToProperCase(string fileName, string worksheetName)
        {
            using (var excelPackage = new ExcelPackage(new System.IO.FileInfo(fileName)))
            {
                ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets[worksheetName];

                if (worksheet == null)
                {
                    Console.WriteLine("Worksheet not found!");
                    return;
                }

                int patientNameColumnIndex = FindColumnIndexByHeader(worksheet, "Patient Name");

                if (patientNameColumnIndex == -1)
                {
                    patientNameColumnIndex = FindColumnIndexByHeader(worksheet, "PATIENTS NAME");
                    if (patientNameColumnIndex == -1)
                    {
                        Console.WriteLine("Column 'Patient Name' not found!");
                        return;
                    }
                }

                // Loop through each cell in the column (including header row)
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    var cell = worksheet.Cells[row, patientNameColumnIndex];
                    string text = cell.Text;

                    // Check if the cell is not empty and skip the header row
                    if (row != 1 && !string.IsNullOrEmpty(text))
                    {
                        string[] nameParts = text.Split(',');
                        for (int i = 0; i < nameParts.Length; i++)
                        {
                            string[] nameWords = nameParts[i].Trim().Split(' ');
                            for (int j = 0; j < nameWords.Length; j++)
                            {
                                if (!string.IsNullOrWhiteSpace(nameWords[j]))
                                {
                                    char[] wordChars = nameWords[j].ToCharArray();
                                    if (wordChars.Length > 0)
                                        wordChars[0] = char.ToUpper(wordChars[0]);
                                    nameWords[j] = new string(wordChars);
                                }
                            }
                            nameParts[i] = string.Join(" ", nameWords);
                        }

                        string formattedName = string.Join(", ", nameParts);
                        cell.Value = formattedName;
                    }
                }

                excelPackage.Save();
                Console.WriteLine("Values in 'Patient Name' column changed to proper case.");
            }
        }

        // Helper method to find the column index by header text
        public static int FindColumnIndexByHeader(ExcelWorksheet worksheet, string columnHeader)
        {
            int columnIndex = -1;

            for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
            {
                if (worksheet.Cells[1, col].Text.Equals(columnHeader, StringComparison.OrdinalIgnoreCase))
                {
                    columnIndex = col;
                    break;
                }
            }

            return columnIndex;
        }

        public static string ModifyExcelFile(string excelFilePath, string typeFile)
        {
            // Rename the original file to include "_original" before the file extension.
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");
            FileInfo originalFileInfo = new FileInfo(excelFilePath);
            string originalFileName = originalFileInfo.Name;
            string originalFilePathWithoutExtension = System.IO.Path.Combine(originalFileInfo.Directory.FullName, System.IO.Path.GetFileNameWithoutExtension(originalFileName));
            string renamedFilePath = $"{originalFilePathWithoutExtension}_original{timestamp}{originalFileInfo.Extension}";

            if (File.Exists(renamedFilePath))
            {
                File.Delete(renamedFilePath);
            }

            File.Move(excelFilePath, renamedFilePath);

            // Create a copy of the renamed file.
            File.Copy(renamedFilePath, excelFilePath);
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you want to work with the first worksheet.

                if (typeFile.Contains("Change") || typeFile.Contains("Summary"))
                {
                    // Remove the first two rows (header rows).
                    worksheet.DeleteRow(1, 2);
                }

                if (typeFile.Contains("Search"))
                {
                    // Remove the first two rows (header rows).
                    worksheet.DeleteRow(1, 2);
                }

                int patientNameColumnIndex = 0;

                if (typeFile == "Agency")
                {
                    int columnCount = worksheet.Dimension.End.Column;

                    for (int col = 1; col <= columnCount; col++)
                    {
                        string cellValue = worksheet.Cells[1, col].Text.Trim();

                        if (cellValue.Equals("PATIENTS NAME", StringComparison.OrdinalIgnoreCase))
                        {
                            patientNameColumnIndex = col;
                            break;
                        }
                    }

                    // Iterate through the rows starting from the second row (data rows).
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        // Get the cell in the "Patient Name" column for the current row.
                        ExcelRangeBase cell = worksheet.Cells[row, patientNameColumnIndex];

                        // Replace dots from the cell text.
                        cell.Value = cell.Text.Replace(".", string.Empty);
                    }
                }
                else if (typeFile.Contains("Change") || typeFile.Contains("Search"))
                {
                    // Get the column index of "Patient Name" assuming it's in the first row (header row).


                    patientNameColumnIndex = worksheet.Cells["1:1"].First(c => c.Text == "Patient Name").Start.Column;

                }

                if (typeFile != "Summary")
                {
                    // Iterate through the rows starting from the third row (data rows).
                    for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                    {
                        // Get the cell in the "Patient Name" column for the current row.
                        ExcelRangeBase cell = worksheet.Cells[row, patientNameColumnIndex];

                        // Convert the cell text to lowercase.
                        cell.Value = cell.Text.ToLower();
                    }
                }

                // Save the modified Excel file.
                package.Save();
            }

            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you want to work with the first worksheet.
                ExcelRangeBase firstRow = worksheet.Cells["1:1"]; // Get the first row as a range.

                foreach (var cell in firstRow)
                {
                    if (cell.Text.Contains("*"))
                    {
                        // Remove asterisks from the cell text.
                        cell.Value = cell.Text.Replace("*", string.Empty);
                    }

                    if (cell.Text.Contains("."))
                    {
                        // Remove dots from the cell text.
                        cell.Value = cell.Text.Replace(".", string.Empty);
                    }
                }

                // Save the modified Excel file.
                package.Save();
            }

            if (typeFile.Contains("Search"))
            {
                using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you are working with the first worksheet

                    // Searching for the column index with the header "TOB"
                    var headerRow = 1; // Assuming the header is in the first row
                    int columnTOB = 0;
                    for (int col = 1; col <= worksheet.Dimension.End.Column; col++)
                    {
                        if (worksheet.Cells[headerRow, col].Text.Equals("TOB", StringComparison.OrdinalIgnoreCase))
                        {
                            columnTOB = col;
                            break;
                        }
                    }

                    // Changing the format of the entire column to "Text"
                    if (columnTOB > 0)
                    {
                        for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                        {
                            worksheet.Cells[row, columnTOB].Style.Numberformat.Format = "@";
                        }
                    }

                    package.Save(); // Save changes to the Excel file
                }
            }

            return renamedFilePath;
        }

        public static void UndoModifyExcelFile(string originalFileName)
        {
            try
            {
                // Check if the file with "_original + datetime stamp" exists.
                string modifiedFilePath = String.Empty;
                int lastUnderscoreIndex = originalFileName.LastIndexOf('_');
                if (lastUnderscoreIndex >= 0)
                {
                    modifiedFilePath = originalFileName.Substring(0, lastUnderscoreIndex) + ".xlsx";
                }
                else
                {
                    // Handle the case where there is no underscore before the last backslash.
                    Console.WriteLine("No underscore found before the last backslash in the input string.");
                }


                if (File.Exists(originalFileName))
                {
                    // Delete the modified file.
                    File.Delete(modifiedFilePath);

                    // Rename the original file to remove "_original + datetime stamp".
                    File.Move(originalFileName, modifiedFilePath);

                    Console.WriteLine("File modifications undone successfully.");
                }
                else
                {
                    Console.WriteLine("The modified file does not exist.");
                }
            }
            catch (Exception ex)
            {
                // Handle any exceptions that may occur during the process
                Console.WriteLine($"Error undoing file modifications: {ex.Message}");
            }
        }


        public static DataTable ExecuteExcelQuery(string excelFilePath, string sqlQuery)
        {
            string connectionString = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={excelFilePath};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\"";

            if (OperatingSystem.IsWindows())
            {
                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    connection.Open();

                    using (OleDbCommand command = new OleDbCommand(sqlQuery, connection))
                    {
                        using (OleDbDataAdapter adapter = new OleDbDataAdapter(command))
                        {
                            DataTable dataTable = new DataTable();

                            try
                            {
                                adapter.Fill(dataTable);
                                return dataTable;
                            }
                            catch (OleDbException ex)
                            {
                                Console.WriteLine($"There were no results for that query");
                                return new DataTable();
                            }
                        }
                    }
                }
            }
            else return new DataTable();
        }

        public static DataTable ReadCsvFile(string filePath, string rowName1, string rowName2)
        {
            DataTable dataTable = new DataTable();
            dataTable.Columns.Add(rowName1);
            dataTable.Columns.Add(rowName2);

            try
            {
                using (StreamReader reader = new StreamReader(filePath))
                {
                    string line;
                    while ((line = reader.ReadLine()) != null)
                    {
                        // Split the line using the pipe character as a delimiter
                        string[] parts = line.Split('|');

                        // Check if the line has the expected format (2 parts)
                        if (parts.Length == 2)
                        {
                            // Add a new row to the DataTable
                            DataRow row = dataTable.NewRow();
                            row[rowName1] = parts[0];
                            row[rowName2] = parts[1];
                            dataTable.Rows.Add(row);
                        }
                        else
                        {
                            Console.WriteLine($"Skipping line: {line}. Invalid format.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error reading CSV file: {ex.Message}");
            }

            return dataTable;
        }

        public static string GetSqlQueryByQueryName(DataTable dataTable, string queryName)
        {
            DataRow[] foundRows = dataTable.Select($"QueryName = '{queryName}'");

            if (foundRows.Length > 0)
            {
                return foundRows[0]["SqlQuery"].ToString();
            }
            else
            {
                Console.WriteLine($"QueryName '{queryName}' not found in the DataTable.");
                return null; // Or return an empty string or handle it as appropriate for your application.
            }
        }

        public static string GetCSVInfoByInfoName(DataTable dataTable, string rowName1, string rowName2, string infoName)
        {
            DataRow[] foundRows = dataTable.Select($"{rowName1} = '{infoName}'");

            if (foundRows.Length > 0)
            {
                return foundRows[0][rowName2].ToString();
            }
            else
            {
                Console.WriteLine($"The value '{infoName}' not found in the DataTable.");
                return null; // Or return an empty string or handle it as appropriate for your application.
            }
        }

        public static DataTable ConvertDatesToDateOnly(DataTable dataTable)
        {
            DataTable modifiedDataTable = new DataTable();

            foreach (DataColumn column in dataTable.Columns)
            {
                DataColumn newColumn = new DataColumn(column.ColumnName);

                if (column.DataType == typeof(DateTime))
                {
                    // Set the data type of the new column to string for date conversion.
                    newColumn.DataType = typeof(string);
                }
                else
                {
                    // Use the same data type for non-date columns.
                    newColumn.DataType = column.DataType;
                }

                modifiedDataTable.Columns.Add(newColumn);
            }

            foreach (DataRow row in dataTable.Rows)
            {
                DataRow newRow = modifiedDataTable.NewRow();

                foreach (DataColumn column in dataTable.Columns)
                {
                    if (column.DataType == typeof(DateTime))
                    {
                        DateTime? dateValue = row.Field<DateTime?>(column);

                        if (dateValue.HasValue)
                        {
                            // Convert the date to "MM/dd/yyyy" format as a string.
                            newRow[column.ColumnName] = dateValue.Value.ToString("MM/dd/yyyy");
                        }
                    }
                    else
                    {
                        // Copy non-date columns as-is.
                        newRow[column.ColumnName] = row[column];
                    }
                }

                modifiedDataTable.Rows.Add(newRow);
            }

            return modifiedDataTable;
        }
                
        public static Tuple<int, int> FindCellLocation(string excelFilePath, string valueToFind, string dateToFind)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you want to search the first worksheet.

                // Find the column index for the "SOC" header.
                int dateColumnIndex = -1;
                int nameColumnIndex = -1;
                foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                {
                    if (cell.Text == "SOC")
                    {
                        dateColumnIndex = cell.Start.Column;
                    }

                    if (cell.Text == "PATIENTS NAME")
                    {
                        nameColumnIndex = cell.Start.Column;
                    }
                }

                if (dateColumnIndex == -1)
                {
                    // "SOC" header not found, handle appropriately (throw an exception, return null, etc.).
                    throw new InvalidOperationException("Column with header 'SOC' not found.");
                }

                // Search for the value within the entire worksheet.
                foreach (var cell in worksheet.Cells)
                {
                    // Check if both name and date match.
                    if (cell.Text.Contains(valueToFind) && worksheet.Cells[cell.Start.Row, dateColumnIndex].Text == dateToFind)
                    {
                        // Return the location [row, col] as a Tuple.
                        return System.Tuple.Create(cell.Start.Row, nameColumnIndex);
                    }
                }
            }

            // If the value is not found, return null.
            return null;
        }       

        public static string FindCellLocationAlpha(string excelFilePath, string hicToFind, string dateToFind)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming you want to search the first worksheet.

                // Find the column index for the "SOC" header (date column).
                int dateColumnIndex = -1;
                foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                {
                    if (cell.Text == "SOC")
                    {
                        dateColumnIndex = cell.Start.Column;
                        break;
                    }
                }

                int hicolumnIndex = -1;
                foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                {
                    if (cell.Text == "HIC/MBI")
                    {
                        hicolumnIndex = cell.Start.Column;
                        break;
                    }
                }

                int nameColumnIndex = -1;
                foreach (var cell in worksheet.Cells[1, 1, 1, worksheet.Dimension.Columns])
                {
                    if (cell.Text == "PATIENTS NAME")
                    {
                        nameColumnIndex = cell.Start.Column;
                        break;
                    }
                }

                // Return null if the "SOC" header is not found.
                if (dateColumnIndex == -1 || nameColumnIndex == -1)
                    return null;

                // Search for both the name and date in all rows following the header.
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    var cell = worksheet.Cells[row, dateColumnIndex];

                    // Check if the cell in the "SOC" column contains the specified date.
                    if (cell.Text == dateToFind)
                    {
                        // Check if the cell in the same row and different column contains the specified name.
                        var hicCell = worksheet.Cells[row, hicolumnIndex];
                        if (hicCell.Text.Contains(hicToFind))
                        {
                            // Convert the row index to numeric value and column index to alphabetical value.
                            int rowNumericValue = row;
                            string columnAlphabeticalValue = GetExcelColumnName(nameColumnIndex);

                            // Create a string in the format "ColumnNameRowNumber" and return it.
                            string location = columnAlphabeticalValue + rowNumericValue;
                            return location;
                        }
                    }
                }
            }

            // If the name and date are not found, return null.
            return null;
        }

        // Helper method to convert column index to Excel alphabetical format (A, B, C, ..., Z, AA, AB, AC, ...)
        public static string GetExcelColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        public static void ModifyCell(string excelFilePath, string sheetName, string cellIndex, string newValue, string cellType)
        {
            using (var workbook = new XLWorkbook(excelFilePath))
            {
                var worksheet = workbook.Worksheet(sheetName);

                if (worksheet != null)
                {
                    var cell = worksheet.Cell(cellIndex);

                    if (cell != null)
                    {
                        if (cellType == "string")
                            cell.Value = newValue;
                        else if (cellType == "double")
                            cell.Value = Convert.ToDouble(newValue);
                        System.Threading.Thread.Sleep(1000);
                        workbook.Save();
                    }
                    else
                    {
                        Console.WriteLine($"Cell {cellIndex} not found in sheet {sheetName}.");
                    }
                }
                else
                {
                    Console.WriteLine($"Sheet {sheetName} not found in the Excel file.");
                }
            }
        }

        public static List<string> GetSheetNames(string excelFilePath)
        {
            List<string> sheetNames = new List<string>();

            try
            {
                using (var workbook = new XLWorkbook(excelFilePath))
                {
                    foreach (IXLWorksheet worksheet in workbook.Worksheets)
                    {
                        sheetNames.Add(worksheet.Name);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error retrieving sheet names: {ex.Message}");
            }

            return sheetNames;
        }

        public static string ReplaceLetterInIndex(string inputString, char newLetter)
        {
            if (inputString.Length >= 2 && char.IsLetter(inputString[0]))
            {
                char[] charArray = inputString.ToCharArray();
                charArray[0] = newLetter;
                return new string(charArray);
            }
            else
            {
                // Invalid input string format, return the input string as is.
                return inputString;
            }
        }

        public static void SendNOAEmailWithTemplate(string recipientName, string recipientEmails, List<discrepencies> listOfDiscrepancies, List<discrepencies> patientsMoreThan3Days)
        {
            try
            {
                DataTable configurationInfo = ReadCsvFile("ConfigFile.csv", "SmtpInfo", "Details");
                string smtpUser = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "smtpUsername");
                string smtpName = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "SenderName");
                string smtpServer = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "smtpServer");
                int smtpPort = Int32.Parse(GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "smtpPort"));
                string smtpPassword = GetCSVInfoByInfoName(configurationInfo, "SmtpInfo", "Details", "smtpPassword");

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(smtpName, smtpUser)); // Replace with your name and Gmail address

                var recipientEmailList = recipientEmails.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var recipientEmail in recipientEmailList)
                {
                    message.To.Add(new MailboxAddress("", recipientEmail)); // Leave the first argument as an empty string or specify a recipient name
                }

                // Read the HTML template from a file
                string htmlBody = File.ReadAllText("HTMLTemplate.html");

                // Replace "Admin" with recipientName in the HTML content
                htmlBody = htmlBody.Replace("Admin", recipientName);

                // Extract the email subject from the HTML content
                string emailSubject = ExtractTitleFromHtml(htmlBody) + $" For All Agencies";

                // Append today's date without a timestamp
                string formattedDate = DateTime.Now.ToString("MMMM dd, yyyy");
                message.Subject = $"{emailSubject} - {formattedDate}"; // Append the date to the subject

                // Generate the list of discrepancies.
                StringBuilder discrepanciesList = new StringBuilder();
                int counter = 1;
                foreach (var discrepancy in listOfDiscrepancies)
                {
                    discrepanciesList.Append("<li>").Append("Discrepancy " + counter + "</li>");
                    discrepanciesList.Append("<ul>");
                    discrepanciesList.Append("<li>").Append("Agency: ").Append(discrepancy.Agency).Append("</li>");
                    discrepanciesList.Append("<li>").Append("Patient Name: ").Append(discrepancy.PatientName).Append("</li>");
                    discrepanciesList.Append("<li>").Append("Admit Date: ").Append(discrepancy.AdmitDate).Append("</li>");
                    discrepanciesList.Append("</ul>");
                    counter++;
                }

                if (patientsMoreThan3Days.Count > 0)
                {
                    // Generate the list of patientsMoreThan3DaysList.
                    counter = 1;
                    StringBuilder patientsMoreThan3DaysList = new StringBuilder();
                    foreach (var patient in patientsMoreThan3Days)
                    {
                        patientsMoreThan3DaysList.Append("<li>").Append("Overdue Patient " + counter + "</li>");
                        patientsMoreThan3DaysList.Append("<ul>");
                        patientsMoreThan3DaysList.Append("<li>").Append("Agency: ").Append(patient.Agency).Append("</li>");
                        patientsMoreThan3DaysList.Append("<li>").Append("Patient Name: ").Append(patient.PatientName).Append("</li>");
                        patientsMoreThan3DaysList.Append("</ul>");
                        counter++;
                    }

                    // Replace the {patientsMoreThan3DaysList} placeholder in the HTML template with the generated list
                    htmlBody = htmlBody.Replace("{patientsMoreThan3DaysList}", patientsMoreThan3DaysList.ToString());
                }
                else
                    htmlBody = htmlBody.Replace("{patientsMoreThan3DaysList}", "");

                // Replace the {DiscrepanciesList} placeholder in the HTML template with the generated list
                htmlBody = htmlBody.Replace("{DiscrepanciesList}", discrepanciesList.ToString());

                var builder = new BodyBuilder();
                builder.HtmlBody = htmlBody;

                message.Body = builder.ToMessageBody();

                using (var client = new SmtpClient())
                {
                    // Connect to the SMTP server (Gmail's SMTP server).
                    client.Connect(smtpServer, smtpPort, true);

                    // Authenticate with your Gmail account.
                    client.Authenticate(smtpUser, smtpPassword); // Replace with your Gmail address and password

                    // Send the email.
                    client.Send(message);

                    // Disconnect from the SMTP server.
                    client.Disconnect(true);
                }

                Console.WriteLine("Email sent successfully using MailKit.");
            }
            catch (Exception ex)
            {
                LogException(ex);
                Console.WriteLine($"Error sending email: {ex.Message}");
            }
        }

        public static string BuildHtmlEmailFromBody(
        string textBody,
        string emailSubject,
        DateTime runDate,
        string htmlTemplatePath)
        {
            if (string.IsNullOrWhiteSpace(textBody))
                throw new ArgumentNullException(nameof(textBody));

            htmlTemplatePath = System.IO.Path.GetFullPath(htmlTemplatePath);

            if (!File.Exists(htmlTemplatePath))
                throw new FileNotFoundException("HTML template file not found", htmlTemplatePath);

            // 1 - Load HTML template from file
            string html = File.ReadAllText(htmlTemplatePath);

            // 2 - Extract pieces from text body
            string clientName = ExtractClientName(textBody);
            string paymentRows = ExtractPaymentRows(textBody);
            string totalPastDue = ExtractTotalPastDue(textBody);

            // 3 - Replace placeholders
            html = html.Replace("{{EMAIL_SUBJECT}}", emailSubject)
                       .Replace("{{EMAIL_TITLE}}", emailSubject)
                       .Replace("{{CLIENT_NAME}}", clientName)
                       .Replace("{{RUN_DATE}}", runDate.ToString("MM/dd/yyyy"))
                       .Replace("{{PAYMENT_ROWS}}", paymentRows)
                       .Replace("{{TOTAL_PAST_DUE}}", totalPastDue);

            return html;
        }

        public static string ExtractClientName(string textBody)
        {
            // Example line: "Good Afternoon A&A Home Health,"
            var regex = new Regex(@"Good Afternoon\s+(?<Name>.+?),",
                RegexOptions.IgnoreCase);

            var match = regex.Match(textBody);
            if (match.Success)
                return match.Groups["Name"].Value.Trim();

            return "Valued Client";
        }

        public static string ExtractPaymentRows(string textBody)
        {
            // Money token: 1,234 or 1,234.56 or 1234.56 (optional $, optional empty)
            const string Money = @"\$?\s*(?:[0-9]{1,3}(?:,[0-9]{3})*(?:\.\d{1,2})?|[0-9]+(?:\.\d{1,2})?)?";

            var paymentRegex = new Regex(
                @"^\s*Pay Date\s*:\s*(?<PayDate>[^,]+)\s*,\s*" +
                @"Deposit Date\s*:\s*(?<DepositDate>[^,]+)\s*,\s*" +
                @"Amount\s*:\s*(?<Amount>" + Money + @")" +
                @"(?:\s*,\s*Amount In Process\s*:\s*(?<AmountInProcess>" + Money + @"))?" +
                @"(?:\s*,\s*Total\s*:\s*(?<Total>" + Money + @"))?" +
                @"(?:\s*,\s*Notes\s*:\s*(?<Notes>.*))?$",
                RegexOptions.IgnoreCase);

            var lines = textBody.Replace("\r", "").Split('\n');
            var sb = new StringBuilder();
            bool alternate = false;

            foreach (var raw in lines)
            {
                var line = raw.Trim();
                if (string.IsNullOrEmpty(line)) continue;

                // be tolerant to "Pay Date :" etc.
                if (!Regex.IsMatch(line, @"^\s*Pay Date\s*:", RegexOptions.IgnoreCase))
                    continue;

                var match = paymentRegex.Match(line);
                if (!match.Success)
                    continue;

                string payDate = match.Groups["PayDate"].Value.Trim();
                string depositDate = match.Groups["DepositDate"].Value.Trim();

                // Amounts may be blank ("", "$", etc.) - only format when there is a value
                string amountRaw = match.Groups["Amount"].Value.Trim();
                string amount = string.IsNullOrWhiteSpace(amountRaw) || amountRaw == "$"
                    ? ""
                    : FormatMoney(amountRaw);

                string inProcRaw = match.Groups["AmountInProcess"].Success
                    ? match.Groups["AmountInProcess"].Value.Trim()
                    : "";
                string amountInProc = string.IsNullOrWhiteSpace(inProcRaw) || inProcRaw == "$"
                    ? ""
                    : FormatMoney(inProcRaw);

                string totalRaw = match.Groups["Total"].Success
                    ? match.Groups["Total"].Value.Trim()
                    : "";
                string total = string.IsNullOrWhiteSpace(totalRaw) || totalRaw == "$"
                    ? ""
                    : FormatMoney(totalRaw);

                string notes = match.Groups["Notes"].Success
                    ? match.Groups["Notes"].Value.Trim()
                    : "";

                string bg = alternate ? "#f8fafc" : "#ffffff";
                alternate = !alternate;

                sb.AppendLine($@"
                <tr style=""background-color:{bg};"">
                  <td style=""padding:8px; border:1px solid #d9e2ec;"">{payDate}</td>
                  <td style=""padding:8px; border:1px solid #d9e2ec;"">{depositDate}</td>
                  <td style=""padding:8px; border:1px solid #d9e2ec; text-align:right;"">{amount}</td>
                  <td style=""padding:8px; border:1px solid #d9e2ec; text-align:right;"">{amountInProc}</td>
                  <td style=""padding:8px; border:1px solid #d9e2ec; text-align:right;"">{total}</td>
                  <td style=""padding:8px; border:1px solid #d9e2ec;"">{notes}</td>
                </tr>");
            }

            return sb.ToString();
        }

        public static string ExtractTotalPastDue(string textBody)
        {
            var lines = textBody.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

            foreach (var rawLine in lines)
            {
                var line = rawLine.Trim();
                if (line.StartsWith("Total Amount In Process Past Due:", StringComparison.OrdinalIgnoreCase))
                {
                    int idx = line.IndexOf(':');
                    if (idx >= 0 && idx + 1 < line.Length)
                    {
                        string value = line.Substring(idx + 1).Trim();
                        return FormatMoney(value);
                    }
                }
            }

            return "$0.00";
        }

        public static string FormatMoney(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
                return string.Empty;

            raw = raw.Trim();

            // Match either:
            // - digits with proper thousands separators: 5,913 or 5,913.80
            // - plain digits: 5913 or 5913.80
            var m = Regex.Match(raw, @"\$?\s*(?<dollars>(?:\d{1,3}(?:,\d{3})+|\d+))(?:\.(?<cents>\d+))?");
            if (!m.Success)
                return string.Empty;

            var dollarsStr = m.Groups["dollars"].Value.Replace(",", "");
            if (!long.TryParse(dollarsStr, NumberStyles.None, CultureInfo.InvariantCulture, out var dollars))
                return string.Empty;

            var centsRaw = m.Groups["cents"].Success ? m.Groups["cents"].Value : "";
            var cents2 =
                centsRaw.Length >= 2 ? centsRaw.Substring(0, 2) :
                centsRaw.Length == 1 ? (centsRaw + "0") :
                "00";

            var dollarsFormatted = dollars.ToString("N0", CultureInfo.InvariantCulture);
            return $"${dollarsFormatted}.{cents2}";
        }

        public static void SendEmail(string recipientName, string recipientEmails, string subject, string body, string logFilePath, string attachmentFilePath = null, bool isHtml = false)
        {
            try
            {
                string clientSecretPath = "client_secret.json";
                string[] scopes = { "https://mail.google.com/" };
                string applicationName = "RCMProcess";

                var credential = AuthenticateAndGetAccessToken(clientSecretPath, scopes, applicationName);

                var message = new MimeMessage();
                message.From.Add(new MailboxAddress(recipientName, "office@hhabilling.com"));

                var recipientEmailList = recipientEmails.Split(new[] { ',', ';', ' ' }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var email in recipientEmailList)
                {
                    message.To.Add(MailboxAddress.Parse(email));
                }

                message.Subject = subject;

                var builder = new BodyBuilder();

                if (isHtml)
                {
                    builder.HtmlBody = body;

                    // Embed header image via CID (local file next to exe)
                    string headerImagePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "FMBLetterhead.png");
                    if (File.Exists(headerImagePath))
                    {
                        var headerImage = builder.LinkedResources.Add(headerImagePath);
                        headerImage.ContentId = "fmbheader"; // must match <img src="cid:fmbheader">
                                                             // Do NOT set ContentType here - your MimeKit version makes it read-only (and it's not needed).
                    }
                    else
                    {
                        Console.WriteLine($"Header image not found at '{headerImagePath}'.");
                        LogError($"Header image not found at '{headerImagePath}'.", logFilePath);
                    }
                }
                else
                {
                    builder.TextBody = body;
                }

                if (!string.IsNullOrEmpty(attachmentFilePath))
                {
                    if (File.Exists(attachmentFilePath))
                    {
                        builder.Attachments.Add(attachmentFilePath);
                    }
                    else
                    {
                        Console.WriteLine($"Attachment file '{attachmentFilePath}' does not exist.");
                        LogError($"Attachment file '{attachmentFilePath}' does not exist.", logFilePath);
                    }
                }

                message.Body = builder.ToMessageBody();

                using (var client = new SmtpClient())
                {
                    client.Connect("smtp.gmail.com", 587, SecureSocketOptions.StartTls);

                    var oauth2 = new SaslMechanismOAuth2("office@hhabilling.com", credential.Token.AccessToken);
                    client.Authenticate(oauth2);

                    client.Send(message);
                    client.Disconnect(true);
                }

                Console.WriteLine("Email sent successfully.");
                LogError("Email sent successfully.", logFilePath);
            }
            catch (Exception ex)
            {
                LogException(ex);
                Console.WriteLine($"Error sending email: {ex.Message}");
                LogError($"Error sending email: {ex.Message}", logFilePath);
            }
        }

        public static UserCredential AuthenticateAndGetAccessToken(string clientSecretPath, string[] scopes, string applicationName)
        {
            // Step 1: Use a temporary ID to authenticate and get email
            string tempTokenPath = "temp-token-store";
            UserCredential tempCredential;

            using (var stream = new FileStream(clientSecretPath, FileMode.Open, FileAccess.Read))
            {
                tempCredential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    "temp-user",
                    CancellationToken.None,
                    new FileDataStore(tempTokenPath, true)).Result;
            }

            // Step 2: Get the actual authenticated email address using Gmail API
            var gmailService = new GmailService(new BaseClientService.Initializer
            {
                HttpClientInitializer = tempCredential,
                ApplicationName = applicationName
            });

            var profile = gmailService.Users.GetProfile("me").Execute();
            string authenticatedEmail = profile.EmailAddress;
            Console.WriteLine("Authentication successful for: " + authenticatedEmail);

            // Step 3: Re-authorize using the actual email as the user ID for token storage
            string realTokenPath = "token-store";
            UserCredential credential;

            using (var stream = new FileStream(clientSecretPath, FileMode.Open, FileAccess.Read))
            {
                credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    scopes,
                    authenticatedEmail,
                    CancellationToken.None,
                    new FileDataStore(realTokenPath, true)).Result;
            }

            // Step 4: Refresh token if needed
            if (credential.Token.IsExpired(SystemClock.Default))
            {
                credential.RefreshTokenAsync(CancellationToken.None).Wait();
            }

            return credential;
        }

        public static void LogError(string errorMessage, string logFilePath)
        {
            try
            {
                // Open the log file in append mode, create if it doesn't exist.
                using (StreamWriter writer = new StreamWriter(logFilePath, true))
                {
                    // Write the error message along with a timestamp.
                    string logEntry = $"[{DateTime.Now}] {errorMessage}";
                    writer.WriteLine(logEntry);
                }

                //Console.WriteLine($"Error logged to {logFilePath}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error logging: {ex.Message}");
            }
        }

        public static void LogException(Exception ex)
        {
            try
            {
                // Create a log file folder (if it doesn't exist) in the current directory
                string logFolder = System.IO.Path.Combine(Directory.GetCurrentDirectory(), "LogFiles");
                Directory.CreateDirectory(logFolder);

                // Generate a unique log file name with a timestamp
                string logFileName = $"error_log_{DateTime.Now:yyyyMMdd_HHmmss}.txt";

                // Combine the folder path and log file name
                string logFilePath = System.IO.Path.Combine(logFolder, logFileName);

                // Create or open the log file and append the exception message
                using (StreamWriter writer = File.AppendText(logFilePath))
                {
                    writer.WriteLine($"[{DateTime.Now}] Exception: {ex.Message}");
                    writer.WriteLine($"StackTrace: {ex.StackTrace}");
                    writer.WriteLine(); // Add an empty line for separation
                }
            }
            catch (Exception logEx)
            {
                // Handle any exceptions that occur while logging the exception
                Console.WriteLine($"Error logging exception: {logEx.Message}");
            }
        }

        public static string ExtractTitleFromHtml(string html)
        {
            string titlePattern = @"<title>(.*?)<\/title>";
            System.Text.RegularExpressions.Match match = Regex.Match(html, titlePattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
            if (match.Success)
            {
                return match.Groups[1].Value.Trim();
            }
            return "Default Subject"; // Provide a default subject if the title tag is not found.
        }        

        //method to move a row down and insert a new row with the specified data.
        public static void InsertRowAndData(string filePath, string worksheetName, int rowNumber, string patientName, string startDate, string paidDate, string reimb, string hicValue)
        {
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];

                if (worksheet == null)
                {
                    // Handle the case when the worksheet is not found
                    throw new ArgumentException("Worksheet not found.");
                }

                // Parse string dates to DateTime objects without timestamp
                DateTime parsedStartDate = DateTime.ParseExact(startDate, "M/d/yyyy", CultureInfo.InvariantCulture).Date;

                // Adjust paid date if it's a weekend and the paid date is today
                DateTime parsedPaidDate = DateTime.ParseExact(paidDate, "M/d/yyyy", CultureInfo.InvariantCulture).Date;
                if (parsedPaidDate.Date == DateTime.Today)
                {
                    parsedPaidDate = GetNextBusinessDay(parsedPaidDate);
                }

                // Convert DateTime objects back to string in the desired format
                string formattedStartDate = parsedStartDate.ToString("M/d/yyyy");
                string formattedPaidDate = parsedPaidDate.ToString("M/d/yyyy");

                // Check if the combination of patientName and formattedStartDate already exists in the table
                for (int row = 1; row <= worksheet.Dimension.End.Row; row++)
                {
                    if (worksheet.Cells[row, 1].Value?.ToString() == patientName)
                    {
                        string existingDateString = worksheet.Cells[row, 2].GetValue<DateTime>().ToString("M/d/yyyy", CultureInfo.InvariantCulture);
                        if (existingDateString == formattedStartDate)
                        {
                            // Data already exists, so return without inserting
                            return;
                        }
                    }
                }

                // Shift the rows below the specified row number down
                worksheet.InsertRow(rowNumber + 1, 1);

                double reimbToDouble = Convert.ToDouble(reimb, CultureInfo.InvariantCulture);

                // Insert the data into the specified row
                worksheet.Cells[rowNumber + 1, 1].Value = patientName;
                worksheet.Cells[rowNumber + 1, 2].Value = formattedStartDate;
                worksheet.Cells[rowNumber + 1, 3].Value = reimbToDouble;
                worksheet.Cells[rowNumber + 1, 4].Value = formattedPaidDate;
                worksheet.Cells[rowNumber + 1, 6].Value = hicValue;

                // Set the horizontal alignment to right for the cells
                worksheet.Cells[rowNumber + 1, 2, rowNumber + 1, 4].Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;


                package.Save();
            }
        }

        public static DateTime GetNextBusinessDay(DateTime date)
        {
            do
            {
                date = date.AddDays(1);
            } while (date.DayOfWeek == DayOfWeek.Saturday || date.DayOfWeek == DayOfWeek.Sunday);

            return date;
        }
       
        public static string FindNearestRowWithDate(string filePath, string worksheetName, string paidDate, string logFilePath)
        {
            FileInfo file = new FileInfo(filePath);
            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = null;
                try
                {
                    worksheet = package.Workbook.Worksheets[worksheetName];
                }
                catch (Exception ex)
                {
                    LogError(ex.Message.ToString(), logFilePath);
                }

                if (worksheet == null)
                {
                    throw new ArgumentException("Worksheet not found.");
                }

                int rowCount = worksheet.Dimension.Rows;
                int paidDateColumnIndex = -1;

                // Find the column index for "Paid Date"
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    if (worksheet.Cells[1, col].Value?.ToString() == "Paid Date")
                    {
                        paidDateColumnIndex = col;
                        break;
                    }
                }

                if (paidDateColumnIndex == -1)
                {
                    throw new ArgumentException("Column 'Paid Date' not found.");
                }

                DateTime targetDate = DateTime.ParseExact(paidDate, "M/d/yyyy", System.Globalization.CultureInfo.InvariantCulture);

                int lastRow = -1;

                for (int row = rowCount; row > 0; row--)
                {
                    if (DateTime.TryParse(worksheet.Cells[row, paidDateColumnIndex].Value?.ToString(), out DateTime date) && date.Date < targetDate.Date)
                    {
                        lastRow = row;
                        break;
                    }
                }

                if (lastRow != -1)
                {
                    return lastRow.ToString();
                }

                // If the date is after all other dates in the sheet, return the last row
                return rowCount.ToString();
            }
        }

        public static void SortExcelRowsByPaidDate(string filePath, string worksheetName, string paidDateColumnName, string direction)
        {
            FileInfo fileInfo = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];

                if (worksheet == null)
                    throw new ArgumentException($"Worksheet {worksheetName} not found in the file.");

                if (worksheet.Dimension == null)
                {
                    Console.WriteLine($"Worksheet '{worksheetName}' is empty. Skipping sort.");
                    return;
                }

                int paidDateColumnIndex = GetColumnNumber(worksheet, paidDateColumnName);

                var start = worksheet.Dimension.Start;
                var end = worksheet.Dimension.End;

                for (int row = start.Row + 1; row <= end.Row; row++)
                {
                    var cell = worksheet.Cells[row, paidDateColumnIndex];
                    var cellValue = cell.Value;
                    var cellText = cell.Text?.Trim();

                    if (cellValue == null || string.IsNullOrWhiteSpace(cellText))
                    {
                        Console.WriteLine($"Skipping date parsing for row {row} in worksheet '{worksheet.Name}' because '{paidDateColumnName}' is empty.");
                        continue;
                    }

                    if (cellValue is string)
                    {
                        if (DateTime.TryParse(cellValue.ToString().Trim(), out DateTime parsedDate))
                        {
                            cell.Value = parsedDate;
                        }
                        else
                        {
                            Console.WriteLine($"Skipping date parsing for row {row} in worksheet '{worksheet.Name}' because '{paidDateColumnName}' is invalid. Value: '{cellText}'");
                        }
                    }
                }

                var rows = new List<object[]>();

                for (int row = start.Row + 1; row <= end.Row; row++)
                {
                    var rowData = new object[end.Column];

                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        rowData[col - 1] = worksheet.Cells[row, col].Value;
                    }

                    rows.Add(rowData);
                }

                if (direction.Equals("desc", StringComparison.OrdinalIgnoreCase))
                {
                    rows = rows
                        .OrderBy(row => IsEmptyDateValue(row[paidDateColumnIndex - 1]) ? 1 : 0)
                        .ThenByDescending(row => ParseToDate(row[paidDateColumnIndex - 1]))
                        .ToList();
                }
                else
                {
                    rows = rows
                        .OrderBy(row => IsEmptyDateValue(row[paidDateColumnIndex - 1]) ? 1 : 0)
                        .ThenBy(row => ParseToDate(row[paidDateColumnIndex - 1]))
                        .ToList();
                }

                for (int row = start.Row + 1; row <= end.Row; row++)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        worksheet.Cells[row, col].Value = null;
                    }
                }

                int currentRow = start.Row + 1;

                foreach (var rowData in rows)
                {
                    for (int col = start.Column; col <= end.Column; col++)
                    {
                        worksheet.Cells[currentRow, col].Value = rowData[col - 1];
                    }

                    currentRow++;
                }

                package.Save();
            }
        }

        public static bool IsEmptyDateValue(object value)
        {
            if (value == null)
                return true;

            if (value is string text && string.IsNullOrWhiteSpace(text))
                return true;

            return false;
        }

        // A helper function to ensure the cell value is treated as a proper DateTime
        public static DateTime? ParseToDate(object cellValue)
        {
            if (cellValue == null) return null;

            if (cellValue is DateTime dateValue)
            {
                return dateValue;  // Already a DateTime object
            }
            else if (double.TryParse(cellValue.ToString(), out double serialDate))
            {
                // Handle Excel serial number date format
                return DateTime.FromOADate(serialDate);
            }
            else if (DateTime.TryParse(cellValue.ToString().Trim(), out DateTime parsedDate))
            {
                return parsedDate;  // Try to parse the date as a string
            }

            return null;  // Return null for invalid dates
        }

        public static int GetColumnNumber(ExcelWorksheet ws, string columnName)
        {
            if (ws == null)
                throw new ArgumentNullException(nameof(ws));

            if (ws.Dimension == null)
                throw new ArgumentException($"Worksheet '{ws.Name}' is empty.");

            int colCount = ws.Dimension.End.Column;

            string Normalize(string value)
            {
                return value?
                    .Replace("\u00A0", " ")
                    .Replace("\r", " ")
                    .Replace("\n", " ")
                    .Trim()
                    .ToLower() ?? string.Empty;
            }

            string targetColumnName = Normalize(columnName);

            for (int i = 1; i <= colCount; i++)
            {
                string rawValue = ws.Cells[1, i].Value?.ToString() ?? "";
                string textValue = ws.Cells[1, i].Text ?? "";
                string normalizedValue = Normalize(rawValue);

                Console.WriteLine($"Column {i}: Value='{rawValue}' | Text='{textValue}' | Normalized='{normalizedValue}'");

                if (normalizedValue == targetColumnName)
                    return i;
            }

            throw new ArgumentException($"Column '{columnName}' not found in worksheet '{ws.Name}'.");
        }

        public static decimal FindReimbursementAmount(string fileName, string sheetName, string startDate, string hicValue)
        {
            startDate = ConvertDateFormat(startDate);

            FileInfo fileInfo = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                int rowCount = worksheet.Dimension.Rows;

                int hicColumn = -1;
                int startDateColumn = -1;
                int reimbColumn = -1;

                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    if (worksheet.Cells[1, col].Value != null)
                    {
                        if (worksheet.Cells[1, col].Value.ToString() == "HIC/MBI")
                        {
                            hicColumn = col;
                        }
                        else if (worksheet.Cells[1, col].Value.ToString() == "Start Date")
                        {
                            startDateColumn = col;
                        }
                        else if (worksheet.Cells[1, col].Value.ToString() == "Reimb")
                        {
                            reimbColumn = col;
                        }
                    }
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, hicColumn].Value != null && worksheet.Cells[row, startDateColumn].Value != null)
                    {
                        string dateString = worksheet.Cells[row, startDateColumn].Value.ToString();
                        string formattedDate = string.Empty;
                        if (DateTime.TryParse(dateString, out DateTime originalDate))
                        {
                            // No time component, use the original date string
                            formattedDate = originalDate.ToString("M/d/yyyy");
                            // Do something with the formattedDate if needed
                        }
                        else
                            formattedDate = dateString;

                        if (worksheet.Cells[row, hicColumn].Value.ToString() == hicValue && formattedDate == startDate)
                        {
                            if (worksheet.Cells[row, reimbColumn].Value != null)
                            {
                                return Convert.ToDecimal(worksheet.Cells[row, reimbColumn].Value);
                            }
                        }
                    }
                }
            }
            return 0; // Return 0 if the patient and start date are not found.
        }

        public static decimal DeductReimbursement(string fileName, string sheetName, string startDate, string hicValue, decimal reimbFromPaymentSummary)
        {
            startDate = ConvertDateFormat(startDate);

            FileInfo fileInfo = new FileInfo(fileName);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];
                int rowCount = worksheet.Dimension.Rows;

                int hicColumn = -1;
                int startDateColumn = -1;
                int reimbColumn = -1;

                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    if (worksheet.Cells[1, col].Value != null)
                    {
                        if (worksheet.Cells[1, col].Value.ToString() == "HIC/MBI")
                        {
                            hicColumn = col;
                        }
                        else if (worksheet.Cells[1, col].Value.ToString() == "Start Date")
                        {
                            startDateColumn = col;
                        }
                        else if (worksheet.Cells[1, col].Value.ToString() == "Reimb")
                        {
                            reimbColumn = col;
                        }
                    }
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    if (worksheet.Cells[row, hicColumn].Value != null && worksheet.Cells[row, startDateColumn].Value != null)
                    {
                        string dateString = worksheet.Cells[row, startDateColumn].Value.ToString();
                        string formattedDate = string.Empty;
                        if (DateTime.TryParse(dateString, out DateTime originalDate))
                        {
                            // No time component, use the original date string
                            formattedDate = originalDate.ToString("M/d/yyyy");
                            // Do something with the formattedDate if needed
                        }
                        else
                            formattedDate = dateString;

                        if (worksheet.Cells[row, hicColumn].Value.ToString() == hicValue && formattedDate == startDate)
                        {
                            if (worksheet.Cells[row, reimbColumn].Value != null)
                            {
                                decimal originalReimbursement = Convert.ToDecimal(worksheet.Cells[row, reimbColumn].Value);
                                decimal updatedReimbursement = originalReimbursement - reimbFromPaymentSummary;
                                worksheet.Cells[row, reimbColumn].Value = updatedReimbursement;
                                worksheet.Cells[row, reimbColumn + 2].Value = "Adjustment";
                                package.Save();
                                return updatedReimbursement;
                            }
                        }
                    }
                }
            }
            return -1; // Return 0 if the patient and start date are not found.
        }

        public static void UpdatePaymentsSheet(string agencyFile, string worksheetName, string payDate, string amount, string type)
        {
            FileInfo fileInfo = new FileInfo(agencyFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];

                int rowCount = worksheet.Dimension.Rows;
                int payDateColumn = -1;

                // Find the column index of "Pay Date"
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    if (worksheet.Cells[1, col].Value.ToString() == "DATE")
                    {
                        payDateColumn = col;
                        break;
                    }
                }

                // If "Pay Date" column is found
                if (payDateColumn != -1)
                {
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Try parsing the cell value to DateTime for comparison
                        if (DateTime.TryParseExact(worksheet.Cells[row, payDateColumn].Value?.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime cellPayDate))
                        {
                            // Compare the parsed dates
                            if (cellPayDate.Date == DateTime.ParseExact(payDate, "M/d/yyyy", CultureInfo.InvariantCulture).Date)
                            {
                                // Input the amount received in the next column
                                int amountColumn = payDateColumn + 2;
                                if (decimal.TryParse(amount, out decimal amountValue))
                                {
                                    worksheet.Cells[row, amountColumn].Value = amountValue;
                                }

                                int notesColumn = GetColumnIndex(worksheet, "NOTES");
                                // If type is "projected", input "Projected" in the next column
                                if (type == "projected")
                                {
                                    worksheet.Cells[row, notesColumn].Value = "Projected";
                                }
                                else if (type == "suspense projection")
                                {
                                    worksheet.Cells[row, notesColumn].Value = "Projected";
                                }
                                else
                                {
                                    worksheet.Cells[row, notesColumn].Value = "";
                                }

                                break;
                            }
                        }
                    }
                }

                package.Save();
            }
        }

        public static void CopyRowToSummary(string agencyFile, string futurePaymentsSheetName, string summaryPaymentsSheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial; // Set the license context

            // Load the Excel file
            FileInfo fileInfo = new FileInfo(agencyFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorkbook workbook = package.Workbook;
                if (workbook != null)
                {
                    // Get the Future Payments sheet
                    ExcelWorksheet futurePaymentsSheet = workbook.Worksheets[futurePaymentsSheetName];
                    if (futurePaymentsSheet != null)
                    {
                        // Get the Summary Payments sheet
                        ExcelWorksheet summaryPaymentsSheet = workbook.Worksheets[summaryPaymentsSheetName];
                        if (summaryPaymentsSheet != null)
                        {
                            int currentRow = 1; // Skip the header row

                            // Find the column index of the "Paid Date" column in the "Future Payments" sheet
                            int dateColumnIndex = 0;
                            foreach (var firstRowCell in futurePaymentsSheet.Cells[1, 1, 1, futurePaymentsSheet.Dimension.End.Column])
                            {
                                if (string.Equals(firstRowCell.Text, "Paid Date", StringComparison.OrdinalIgnoreCase))
                                {
                                    dateColumnIndex = firstRowCell.Start.Column;
                                    break;
                                }
                            }

                            if (dateColumnIndex > 0)
                            {
                                List<int> rowIndexes = new List<int>();
                                for (int row = 2; row <= futurePaymentsSheet.Dimension.End.Row; row++)
                                {
                                    DateTime cellDate;
                                    if (DateTime.TryParse(futurePaymentsSheet.Cells[row, dateColumnIndex].Value?.ToString(), out cellDate))
                                    {
                                        if (cellDate.Date == DateTime.Today)
                                        {
                                            // Shift existing rows down
                                            summaryPaymentsSheet.InsertRow(currentRow + 1, 1);

                                            // Copy the entire row to the "Summary Payments" sheet
                                            for (int col = 1; col <= futurePaymentsSheet.Dimension.End.Column; col++)
                                            {
                                                var cellValue = futurePaymentsSheet.Cells[row, col].Value;
                                                summaryPaymentsSheet.Cells[currentRow + 1, col].Value = cellValue;

                                                // Change the format of the "Paid Date" cell to date
                                                if (col == dateColumnIndex)
                                                {
                                                    summaryPaymentsSheet.Cells[currentRow, col].Style.Numberformat.Format = "m/d/yyyy";
                                                }
                                            }
                                            currentRow++;

                                            rowIndexes.Add(row);
                                        }
                                    }
                                }

                                // Delete the row in the "Future Payments" sheet
                                for (int index = 0; index < rowIndexes.Count; index++)
                                {
                                    futurePaymentsSheet.DeleteRow(2, 1);  
                                }
                            }
                            else
                            {
                                throw new InvalidOperationException("Column 'Paid Date' not found in the 'Future Payments' sheet.");
                            }
                        }
                    }

                    // Save the changes
                    package.Save();
                }
            }
        }

        public static string ConvertDateFormat(string inputDate)
        {
            if (DateTime.TryParseExact(inputDate, "MM/dd/yyyy", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime result))
            {
                return result.ToString("M/d/yyyy");
            }
            return null;
        }        
        
        public static string RemoveMiddleName(string fullName)
        {
            // Split the full name into individual words
            string[] words = fullName.Split(' ');

            // Find the middle name (assumes middle name is one letter)
            string middleName = Array.Find(words, name => name.Length == 1 || (name.Length == 2 && name.EndsWith(".")));

            if (!string.IsNullOrEmpty(middleName))
            {
                // Remove the middle name (including dot, if present)
                words = Array.FindAll(words, name => name != middleName);

                // Concatenate the modified words to form the new name
                string modifiedName = string.Join(" ", words);

                // Remove any extra spaces resulting from the removal
                modifiedName = modifiedName.Trim();

                return modifiedName;
            }
            else
            {
                // Handle the case where the input doesn't have a middle name
                return fullName;
            }
        }

        public static string TrimName(string name)
        {
            // Truncate the name to 15 characters if it's longer
            if (name.Length > 15)
            {
                name = name.Substring(0, 15).Trim();
            }

            return name;
        }

        public static void ProcessAgencyFileDates(string excelFileName)
        {
            FileInfo file = new FileInfo(excelFileName);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0]; // Assuming data is in the first worksheet

                int socColumnIndex = -1;
                int startColumnIndex = -1;

                // Find SOC and Start column indices
                for (int i = 1; i <= worksheet.Dimension.Columns; i++)
                {
                    string header = worksheet.Cells[1, i].Text.Trim();

                    if (header.Equals("SOC", StringComparison.OrdinalIgnoreCase))
                    {
                        socColumnIndex = i;
                    }
                    else if (header.StartsWith("Start", StringComparison.OrdinalIgnoreCase))
                    {
                        startColumnIndex = i;
                        break; // Stop after finding the first "Start" column
                    }
                }

                if (socColumnIndex == -1 || startColumnIndex == -1)
                {
                    Console.WriteLine("SOC or Start columns not found in the file.");
                    return;
                }

                // Find the last "Start" column index
                int lastStartColumnIndex = worksheet.Dimension.Columns;
                for (int i = lastStartColumnIndex; i > startColumnIndex; i--)
                {
                    string header = worksheet.Cells[1, i].Text.Trim();
                    if (header.StartsWith("Start", StringComparison.OrdinalIgnoreCase))
                    {
                        lastStartColumnIndex = i;
                        break;
                    }
                }

                // Process each row starting from the second row (skipping header)
                for (int row = 2; row <= worksheet.Dimension.Rows; row++)
                {
                    string cellContent = worksheet.Cells[row, socColumnIndex].Text;

                    DateTime socDate;
                    if (DateTime.TryParseExact(
                        cellContent,
                        "M/d/yyyy",
                        CultureInfo.InvariantCulture,
                        DateTimeStyles.None,
                        out socDate))
                    {
                        DateTime startDate = socDate;

                        // Update Start columns with dates if the cell is empty
                        for (int i = startColumnIndex; i <= lastStartColumnIndex; i++)
                        {
                            string startColumnName = worksheet.Cells[1, i].Text.Trim();
                            if (worksheet.Cells[row, i].Value == null && startColumnName.StartsWith("Start", StringComparison.OrdinalIgnoreCase))
                            {
                                startDate = startDate.AddDays(30);
                                worksheet.Cells[row, i].Value = startDate.ToString("M/d/yyyy");
                            }
                        }
                    }
                    else
                    {
                        
                    }
                }

                package.Save();
            }
        }

        public static DataTable DeleteRowsByName(DataTable dataTable, string columnName, string nameToKeep)
        {
            // Clone the original DataTable
            DataTable filteredDataTable = dataTable.Clone();

            // Get rows to keep in the new DataTable
            DataRow[] rowsToKeep = dataTable.Select($"{columnName} = '{nameToKeep}'");

            // Import the rows into the new DataTable
            foreach (DataRow row in rowsToKeep)
            {
                filteredDataTable.ImportRow(row);
            }

            return filteredDataTable;
        }

        public static void UpdateHICNumber(string excelFileName, string worksheetName, string patientsName, string HICNumber)
        {
            // Load the Excel file using EPPlus
            using (var package = new ExcelPackage(new FileInfo(excelFileName)))
            {
                // Get the first worksheet
                var worksheet = package.Workbook.Worksheets[worksheetName];

                // Find the column index for "PATIENTS NAME" and "HIC/MBI"
                var patientsNameColumnIndex = GetColumnIndex(worksheet, "PATIENTS NAME");
                var HICColumnIndex = GetColumnIndex(worksheet, "HIC/MBI");

                // Find the row indexes where the patient name matches
                var rowIndexes = FindRowIndex(worksheet, patientsNameColumnIndex, patientsName);

                if (rowIndexes.Length > 0)
                {
                    // Update the value in the "HIC/MBI" column for the found rows
                    foreach (var rowIndex in rowIndexes)
                    {
                        var currentHIC = worksheet.Cells[rowIndex, HICColumnIndex].Value;
                        if (currentHIC == null)
                        {
                            worksheet.Cells[rowIndex, HICColumnIndex].Value = HICNumber;
                            Console.WriteLine($"HIC number for Patient name '{patientsName}' Inserted");
                        }
                        else
                            Console.WriteLine($"Skipping HIC Insert For: '{patientsName}' Already Inserted");
                    }

                    // Save the changes back to the Excel file
                    package.Save();
                }
                else
                {
                    Console.WriteLine($"Patient name '{patientsName}' not found in the worksheet.");
                }
            }            
        }

        public static int GetColumnIndex(ExcelWorksheet worksheet, string columnName)
        {
            // Find the column index based on the column name
            var headerRow = worksheet.Cells["1:1"];
            var column = headerRow.FirstOrDefault(c => c.Text == columnName);

            if (column != null)
            {
                return column.Start.Column;
            }

            // Return -1 if the column is not found
            return -1;
        }

        public static int[] FindRowIndex(ExcelWorksheet worksheet, int columnIndex, string value)
        {
            value = value.ToLower();
            // Find the row indexes where the specified value is located in the given column
            var column = worksheet.Cells[2, columnIndex, worksheet.Dimension.Rows, columnIndex];
            var cells = column.Where(c => c.Text.ToLower() == value);

            if (cells.Any())
            {
                return cells.Select(c => c.Start.Row).ToArray();
            }

            // Return an empty array if the value is not found in the specified column
            return new int[0];
        }

        public static void CopyDataTableToWorksheet(string xlsxFile, DataTable dataTable, string worksheetName)
        {
            using (ExcelPackage package = new ExcelPackage(new FileInfo(xlsxFile)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];
                if (worksheet == null)
                {
                    worksheet = package.Workbook.Worksheets.Add(worksheetName);
                }

                int nameColIndex = dataTable.Columns.IndexOf("PATIENTS NAME");
                int excelNameColIndex = worksheet.Cells["1:1"].First(c => c.Value.ToString() == "PATIENTS NAME").Start.Column;
                int hicMbiColIndex = dataTable.Columns.IndexOf("HIC/MBI");

                // If either column is not found, exit the method
                if (nameColIndex == -1 || excelNameColIndex == 0 || hicMbiColIndex == -1)
                {
                    throw new ArgumentException("Column 'PATIENTS NAME' or 'HIC/MBI' not found in DataTable or worksheet.");
                }

                // Iterate through each row in the DataTable
                foreach (DataRow row in dataTable.Rows)
                {
                    string patientName = row.Field<string>("PATIENTS NAME");
                    object hicMbiValue = row["HIC/MBI"];

                    // Find all rows in the worksheet where the patient name matches
                    var excelMatchingRows = worksheet.Cells["A:A"].Where(c => string.Equals(c.Text, patientName, StringComparison.OrdinalIgnoreCase));

                    // Iterate through each matching row in the worksheet
                    foreach (var excelCell in excelMatchingRows)
                    {
                        // Get the row index of the matching cell
                        int rowIndex = excelCell.Start.Row;

                        // Insert the value from the DataTable to the corresponding "HIC/MBI" column in the worksheet
                        worksheet.Cells[rowIndex, excelNameColIndex + (hicMbiColIndex - nameColIndex)].Value = hicMbiValue;
                    }
                }

                package.Save();
            }
        }

        public static void CompareAndInsert(DataTable distinctResultsHICFile, DataTable resultFromAgencyFile)
        {
            foreach (DataRow rowHICFile in distinctResultsHICFile.Rows)
            {
                string patientName = rowHICFile["Patient Name"].ToString();
                string hicMbiValue = rowHICFile["HIC/MBI"].ToString();

                foreach (DataRow rowAgencyFile in resultFromAgencyFile.Rows)
                {
                    string patientNameAgency = rowAgencyFile["PATIENTS NAME"].ToString();
                    patientName = RemoveMiddleName(patientName);
                    patientNameAgency = RemoveMiddleName(patientNameAgency);
                    if (patientName.ToLower() == patientNameAgency.ToLower())
                    {
                        if (string.IsNullOrEmpty(rowAgencyFile["HIC/MBI"].ToString()))
                        {
                            rowAgencyFile["HIC/MBI"] = hicMbiValue;
                            Console.WriteLine($"Inserted HIC/MBI value '{hicMbiValue}' for patient '{patientNameAgency}'");
                        }
                        else
                        {
                            Console.WriteLine($"Skipping insertion for patient '{patientNameAgency}' as HIC/MBI already exists.");
                        }
                    }
                }
            }
        }

        public static void InsertPatientNotFound(string fileName, string worksheetName, string hic, string date, string patientName, string status)
        {
            FileInfo file = new FileInfo(fileName);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];

                if (worksheet == null)
                {
                    // Worksheet with the specified name doesn't exist
                    return;
                }

                // Find the column index for headers
                int hicMbiColumn = FindColumnIndex(worksheet, "HIC/MBI");
                int socColumn = FindColumnIndex(worksheet, "SOC");
                int patientNameColumn = FindColumnIndex(worksheet, "PATIENTS NAME");
                int statusColumn = FindColumnIndex(worksheet, "Status");

                // Find the last row in the worksheet under "PATIENTS NAME" column that has a value
                int lastRow = FindLastRowWithValue(worksheet, patientNameColumn);

                // Insert data into respective columns in the row after the last row with a value under "PATIENTS NAME" column
                int newRow = lastRow + 1;

                worksheet.Cells[newRow, hicMbiColumn].Value = hic;
                worksheet.Cells[newRow, socColumn].Value = DateTime.Parse(date);
                worksheet.Cells[newRow, patientNameColumn].Value = patientName;
                worksheet.Cells[newRow, statusColumn].Value = status;
                //worksheet.Cells[newRow, statusColumn + 2].Value = "In Process";

                package.Save();
            }
        }

        public static int FindColumnIndex(ExcelWorksheet worksheet, string columnHeader)
        {
            int columnCount = worksheet.Dimension.Columns;
            for (int col = 1; col <= columnCount; col++)
            {
                if (worksheet.Cells[1, col].Value?.ToString().Trim() == columnHeader)
                {
                    return col;
                }
            }
            // If the column with the specified header is not found, return -1
            return -1;
        }        

        public static int FindLastRowWithValue(ExcelWorksheet worksheet, int column)
        {
            int rowCount = worksheet.Dimension.End.Row;
            for (int row = rowCount; row >= 1; row--)
            {
                if (worksheet.Cells[row, column].Value != null)
                {
                    return row;
                }
            }
            // If no value found in the column, return 0
            return 0;
        }

        public static DataTable AddOrUpdateRow(DataTable dataTable, string hicValue, string admitDateValue, string patientNameValue, string reasonCodeFromChanges)
        {
            // Check if the columns exist, if not, add them
            if (!dataTable.Columns.Contains("hicValueFromChanges"))
                dataTable.Columns.Add("hicValueFromChanges");
            if (!dataTable.Columns.Contains("admitDateValueFromChanges"))
                dataTable.Columns.Add("admitDateValueFromChanges");
            if (!dataTable.Columns.Contains("patientNameValueFromChanges"))
                dataTable.Columns.Add("patientNameValueFromChanges");
            if (!dataTable.Columns.Contains("reasonCodeFromChanges"))
                dataTable.Columns.Add("reasonCodeFromChanges");

            // Check if the combination of values already exists in the table
            foreach (DataRow row in dataTable.Rows)
            {
                if (row["hicValueFromChanges"].ToString() == hicValue &&
                    row["admitDateValueFromChanges"].ToString() == admitDateValue)
                {
                    // Combination already exists, return the original table
                    return dataTable;
                }
            }

            // Combination doesn't exist, add a new row
            DataRow newRow = dataTable.NewRow();
            newRow["hicValueFromChanges"] = hicValue;
            newRow["admitDateValueFromChanges"] = admitDateValue;
            newRow["patientNameValueFromChanges"] = patientNameValue;
            newRow["reasonCodeFromChanges"] = reasonCodeFromChanges;
            dataTable.Rows.Add(newRow);

            return dataTable;
        }

        public static string AddDaysToDate(string date, int numberOfDays)
        {
            // Parse the input date string to a DateTime object
            DateTime parsedDate = DateTime.ParseExact(date, "MM/dd/yyyy", CultureInfo.InvariantCulture);

            // Add the specified number of days to the date
            DateTime newDate = parsedDate.AddDays(numberOfDays);

            // Return the new date as a string in "MM/dd/yyyy" format
            return newDate.ToString("MM/dd/yyyy");
        }        

        public static void UpdatePaymentsSheetAggregate(string agencyFile, string worksheetName, string payDate, string amount, string type)
        {
            bool isPastDue = false;
            FileInfo fileInfo = new FileInfo(agencyFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[worksheetName];

                int rowCount = worksheet.Dimension.Rows;
                int payDateColumn = -1;

                // Find the column index of "Pay Date"
                for (int col = 1; col <= worksheet.Dimension.Columns; col++)
                {
                    if (worksheet.Cells[1, col].Value.ToString() == "DATE")
                    {
                        payDateColumn = col;
                        break;
                    }
                }

                // If "Pay Date" column is found
                if (payDateColumn != -1)
                {
                    // Get the column index for "NOTES" using the helper method
                    int notesColumn = GetColumnIndex(worksheet, "NOTES");
                    decimal amountValue = 0;
                    for (int row = 2; row <= rowCount; row++)
                    {
                        // Try parsing the cell value to DateTime for comparison
                        if (DateTime.TryParseExact(worksheet.Cells[row, payDateColumn].Value?.ToString(), "M/d/yyyy h:mm:ss tt", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime cellPayDate))
                        {
                            // Compare the parsed dates
                            if (cellPayDate.Date == DateTime.ParseExact(payDate, "M/d/yyyy", CultureInfo.InvariantCulture).Date)
                            {
                                if (decimal.TryParse(amount.Replace("$", "").Replace(",", ""), out decimal amountValueOG))
                                {
                                    amountValue = amountValueOG;
                                    // Get the SUSPENSE or Past Due field and aggregate with the new Reimb.
                                    DateTime parsedDate;
                                    decimal suspenseOriginal;

                                    if (DateTime.TryParseExact(payDate, "MM/dd/yyyy",
                                        CultureInfo.InvariantCulture, DateTimeStyles.None, out parsedDate))
                                    {
                                        if (parsedDate > DateTime.Today)
                                        {
                                            isPastDue = false;
                                            Console.WriteLine("Pay Date is not in the past.");
                                            
                                            // Get the SUSPENSE field and aggregate with the new Reimb.
                                            int suspenseColumn = GetColumnIndex(worksheet, "SUSPENSE");
                                            suspenseOriginal = Convert.ToDecimal(worksheet.Cells[row, suspenseColumn].Value);
                                            worksheet.Cells[row, suspenseColumn].Value = suspenseOriginal + amountValue;
                                        }
                                        else
                                        {
                                            // Date is in the past
                                            isPastDue = true;
                                            int pastDueColumn = GetColumnIndex(worksheet, "PAST DUE");
                                            decimal reimDec = Convert.ToDecimal(amount.Replace(",", "").Replace("$", ""));
                                            decimal currentCalc = 0;

                                            if (worksheet.Cells[1, pastDueColumn + 1].Text == "")
                                            {
                                                currentCalc = 0;
                                            }
                                            else
                                                currentCalc = Convert.ToDecimal(worksheet.Cells[1, pastDueColumn + 1].Text.Replace(",", "").Replace("$", ""));

                                            worksheet.Cells[1, pastDueColumn+1].Value = currentCalc + reimDec;
                                        }
                                    }
                                }
                                // Handle the "type" input and store it in the "NOTES" column

                                if (notesColumn != -1 && !isPastDue)
                                {

                                    if (type == "projected")
                                    {
                                        worksheet.Cells[row, notesColumn].Value = "Projected";
                                    }
                                    if (type == "suspense projection")
                                    {
                                        worksheet.Cells[row, notesColumn].Value = "Projected";
                                    }
                                    else
                                    {
                                        worksheet.Cells[row, notesColumn].Value = "";
                                    }
                                }

                                break;
                            }
                        }
                    }                    
                }

                package.Save();
            }
        }

        public static bool TryParseDecimalCell(ExcelRangeBase cell, out decimal value)
        {
            value = 0m;
            if (cell == null) return false;

            // 1) If Excel stored a number, use it directly
            if (cell.Value != null)
            {
                switch (cell.Value)
                {
                    case decimal d: value = d; return true;
                    case double d: value = (decimal)d; return true;
                    case float f: value = (decimal)f; return true;
                    case long l: value = l; return true;
                    case int i: value = i; return true;
                }
            }

            // 2) Otherwise parse from text or value string
            string s = cell.Text;
            if (string.IsNullOrWhiteSpace(s) && cell.Value != null)
                s = cell.Value.ToString();

            if (string.IsNullOrWhiteSpace(s))
                return false;

            // Normalize: remove ALL whitespace (incl. Unicode), $ and commas
            s = new string(s.Where(ch => !char.IsWhiteSpace(ch)).ToArray())
                 .Replace("$", "")
                 .Replace(",", "");

            // Accounting negatives "(123.45)"
            bool neg = s.StartsWith("(") && s.EndsWith(")");
            if (neg) s = s.Trim('(', ')');

            if (decimal.TryParse(s, NumberStyles.AllowLeadingSign | NumberStyles.AllowDecimalPoint,
                                 CultureInfo.InvariantCulture, out var parsed))
            {
                value = neg ? -parsed : parsed;
                return true;
            }
            return false;
        }

        // Look within a field's own span [startCol, endColExclusive)
        public static bool TryGetFieldValueInSpan(ExcelWorksheet ws, int row, int startCol, int endColExclusive, out decimal value)
        {
            value = 0m;

            for (int col = startCol; col < endColExclusive; col++)
            {
                var cell = ws.Cells[row, col];

                // Try numeric directly
                if (TryParseDecimalCell(cell, out var v))
                {
                    value = v;
                    return true;
                }

                // If this cell is a lone "$", peek one cell to the right (still inside span)
                var t = cell?.Text?.Trim();
                if (t == "$" && col + 1 < endColExclusive)
                {
                    var right = ws.Cells[row, col + 1];
                    if (TryParseDecimalCell(right, out var vr))
                    {
                        value = vr;
                        return true;
                    }
                }
            }

            return false;
        }

        public static void UpdateTotalsFromSuspense(string agencyFile, string worksheetName)
        {
            FileInfo fileInfo = new FileInfo(agencyFile);
            using (ExcelPackage package = new ExcelPackage(fileInfo))
            {
                var ws = package.Workbook.Worksheets[worksheetName];
                if (ws?.Dimension == null) return;

                int rows = ws.Dimension.Rows;

                int amountCol = GetColumnIndex(ws, "AMOUNT");
                int suspenseCol = GetColumnIndex(ws, "SUSPENSE");
                int totalCol = GetColumnIndex(ws, "TOTAL");

                if (amountCol == -1 || suspenseCol == -1 || totalCol == -1)
                    throw new Exception("Required columns not found (AMOUNT, SUSPENSE, TOTAL).");

                for (int r = 2; r <= rows; r++)
                {
                    // SUSPENSE is required: search ONLY inside [suspenseCol, totalCol)
                    if (!TryGetFieldValueInSpan(ws, r, suspenseCol, totalCol, out var suspenseVal))
                        continue;

                    // AMOUNT optional: search ONLY inside [amountCol, suspenseCol)
                    decimal amountVal = 0m;
                    TryGetFieldValueInSpan(ws, r, amountCol, suspenseCol, out amountVal);

                    ws.Cells[r, totalCol].Value = suspenseVal + amountVal;

                    // Optional formatting
                    // ws.Cells[r, totalCol].Style.Numberformat.Format = "$#,##0.00";
                }

                package.Save();
            }
        }

        public static void RunStep(string stepName, Action stepAction)
        {
            try
            {
                stepAction();
            }
            catch (Exception ex)
            {
                if (!RCMHospiceProcess.DebugMode)
                    SendSms("18184312850", "18182906235", $"{stepName} has failed during RCM Process");
                else
                    Console.WriteLine($"{stepName} has failed during RCM Process - Debug mode on");
                throw;
            }
        }

        public static void SendSms(string fromNumber, string toNumber, string messageBody)
        {
            try
            {
                TwilioClient.Init("YOUR_TWILIO_ACCOUNT_SID", "YOUR_TWILIO_AUTH_TOKEN");

                var message = MessageResource.Create(
                    body: messageBody,
                    from: new PhoneNumber(fromNumber),
                    to: new PhoneNumber(toNumber)
                );

                Console.WriteLine($"SMS sent! SID: {message.Sid}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to send SMS: {ex.Message}");
            }
        }

        public static void ClearColumnsDandE(string filePath, string sheetName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            FileInfo file = new FileInfo(filePath);
            if (!file.Exists)
                throw new FileNotFoundException("Excel file not found.", filePath);

            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[sheetName];
                if (worksheet == null)
                    throw new ArgumentException($"Worksheet '{sheetName}' not found.", nameof(sheetName));

                // Determine the used range
                var lastRow = worksheet.Dimension?.End.Row ?? 0;
                var lastCol = worksheet.Dimension?.End.Column ?? 0;

                // --- 1. Clear Columns D and E (skip header row) ---
                for (int row = 2; row <= lastRow; row++)
                {
                    worksheet.Cells[row, 4].Clear(); // Column D
                    worksheet.Cells[row, 5].Clear(); // Column E
                }

                // --- 2. Find "PAST DUE" header and clear the cell to its right ---
                for (int col = 1; col <= lastCol; col++)
                {
                    var headerValue = worksheet.Cells[1, col].Text?.Trim();
                    if (string.Equals(headerValue, "PAST DUE", System.StringComparison.OrdinalIgnoreCase))
                    {
                        worksheet.Cells[1, col + 1].Clear();
                        break; // Stop after first match
                    }
                }

                package.Save();
            }
        }

        public static string TruncateMessageBody(string body)
        {
            if (string.IsNullOrWhiteSpace(body))
                return body;

            var lines = body.Replace("\r", "").Split('\n').ToList();
            var output = new List<string>();

            // Remove the "Attached is..." line and blank line after it
            for (int i = 0; i < lines.Count; i++)
            {
                if (lines[i].Contains("Attached is the updated spreadsheet for today"))
                {
                    lines.RemoveAt(i);

                    if (i < lines.Count && string.IsNullOrWhiteSpace(lines[i]))
                        lines.RemoveAt(i);

                    break;
                }
            }

            bool insidePaymentSection = false;

            foreach (var line in lines)
            {
                // Remove Past Due line
                if (line.StartsWith("Total Amount In Process Past Due"))
                    continue;

                // Start of payment lines
                if (line.StartsWith("Here are the upcoming payments"))
                {
                    insidePaymentSection = true;
                    output.Add(line);
                    continue;
                }

                // Handle payment lines
                if (insidePaymentSection && line.StartsWith("Pay Date:"))
                {
                    // Deposit Date: 11/25/2025
                    var depositMatch = Regex.Match(line, @"Deposit Date:\s*([\d/]+)");

                    // Amount: $11,940 or $3,584.46 etc
                    // This pattern:
                    // - 1 to 3 digits
                    // - optional groups of ,XXX
                    // - optional .decimals
                    // - stops when it hits a comma or end of line
                    var amountMatch = Regex.Match(
                        line,
                        @"Amount:\s*\$?([0-9]{1,3}(?:,[0-9]{3})*(?:\.\d+)?)(?=,|$)"
                    );

                    string deposit = depositMatch.Success
                        ? $"Deposit Date: {depositMatch.Groups[1].Value}"
                        : "";

                    string amount = amountMatch.Success
                        ? $"Amount: ${amountMatch.Groups[1].Value}"
                        : "";

                    output.Add($"{deposit}, {amount}");
                    continue;
                }

                // Exit payment section once other text appears
                if (insidePaymentSection && !line.StartsWith("Pay Date:") && !string.IsNullOrWhiteSpace(line))
                {
                    insidePaymentSection = false;
                }

                // Keep normal lines
                output.Add(line);
            }

            // Cleanup double blank lines
            for (int i = output.Count - 1; i > 0; i--)
            {
                if (string.IsNullOrWhiteSpace(output[i]) && string.IsNullOrWhiteSpace(output[i - 1]))
                    output.RemoveAt(i);
            }

            return string.Join("\n", output);
        }

        public static string CheckStatuses(string filePath, string sheetName, string agencyName)
        {
            // Resolve to full absolute path so Excel Interop does not get confused by ".."
            string fullPath = System.IO.Path.GetFullPath(filePath);

            if (!File.Exists(fullPath))
            {
                throw new FileNotFoundException($"File not found at resolved path: {fullPath}");
            }

            // 1) Use Excel Interop to force recalculation and save cached values
            Excel.Application excelApp = null;
            Excel.Workbook excelWb = null;

            try
            {
                excelApp = new Excel.Application
                {
                    Visible = false,
                    DisplayAlerts = false
                };

                // Important: use the full path, no extra quotes
                excelWb = excelApp.Workbooks.Open(
                    Filename: fullPath,
                    ReadOnly: false,
                    UpdateLinks: 0
                );

                // Force full recalculation and rebuild dependencies
                excelApp.CalculateFullRebuild();

                // Save workbook so cached values are written into the file
                excelWb.Save();
            }
            finally
            {
                // Clean up COM objects
                if (excelWb != null)
                {
                    excelWb.Close(false);
                    Marshal.ReleaseComObject(excelWb);
                    excelWb = null;
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                    excelApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

            // 2) Now use EPPlus to read the cached values
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using (var package = new ExcelPackage(new FileInfo(fullPath)))
            {
                var workbook = package.Workbook;
                var worksheet = workbook.Worksheets[sheetName];
                if (worksheet == null)
                    throw new Exception($"Sheet '{sheetName}' not found.");

                int colStatus1 = -1;
                int colStatus2 = -1;

                // Find headers in row 1
                int totalColumns = worksheet.Dimension.End.Column;
                for (int col = 1; col <= totalColumns; col++)
                {
                    var header = worksheet.Cells[1, col].Text?.Trim();

                    if (header.Equals("Status 1", StringComparison.OrdinalIgnoreCase))
                        colStatus1 = col;

                    if (header.Equals("Status 2", StringComparison.OrdinalIgnoreCase))
                        colStatus2 = col;
                }

                if (colStatus1 == -1 || colStatus2 == -1)
                    throw new Exception("One or both required columns ('Status 1', 'Status 2') were not found.");

                int row = 2; // start after header

                while (true)
                {
                    // Read the text shown in Excel after recalc
                    string val1 = worksheet.Cells[row, colStatus1].Text?.Trim();
                    string val2 = worksheet.Cells[row, colStatus2].Text?.Trim();

                    // If both empty -> end of data
                    if (string.IsNullOrWhiteSpace(val1) && string.IsNullOrWhiteSpace(val2))
                    {
                        string msg = $"All lines are Equal in Agency: {agencyName}";
                        Console.WriteLine(msg);
                        return "";
                    }

                    // Check for "Not Equal"
                    if (val1.Equals("Not Equal", StringComparison.OrdinalIgnoreCase) ||
                        val2.Equals("Not Equal", StringComparison.OrdinalIgnoreCase))
                    {
                        string msg = $"Not Equal found in Agency: {agencyName}";
                        Console.WriteLine(msg);
                        return agencyName;
                    }

                    row++;
                }
            }
        }

        public static void ApplyHolidayAggregation_HighLevelPaymentSummary(string excelFilePath, string columnName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 1) Load Holidays.xlsx and build holiday -> newDate map (DateTime -> DateTime)
            var holidayMap = LoadHolidayMap();

            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var ws = package.Workbook.Worksheets[0];
                if (ws.Dimension == null) return;

                int lastRow = ws.Dimension.End.Row;
                int lastCol = ws.Dimension.End.Column;

                int providerCol = GetColumnIndex(ws, "Provider", 1, lastCol);
                int dateCol = GetColumnIndex(ws, columnName, 1, lastCol);
                int scheduledCol = GetColumnIndex(ws, "Scheduled", 1, lastCol);
                int projectedCol = GetColumnIndex(ws, "Projected", 1, lastCol);

                if (providerCol < 1 || dateCol < 1 || scheduledCol < 1 || projectedCol < 1)
                    throw new InvalidOperationException("Missing required columns (Provider, Pay Date, Scheduled, Projected).");

                // 2) Read all rows into a structure so we can aggregate and delete safely
                // Grouping key: Provider + EffectiveDate (EffectiveDate = holidayMap[date] if date is holiday else date)
                var groups = new Dictionary<string, List<RowInfo>>();

                for (int row = 2; row <= lastRow; row++)
                {
                    string provider = ws.Cells[row, providerCol].Text?.Trim() ?? "";
                    if (string.IsNullOrWhiteSpace(provider))
                        continue;

                    if (!TryGetDateFromCell(ws.Cells[row, dateCol].Value, out DateTime payDate))
                        continue;

                    DateTime effectiveDate = holidayMap.TryGetValue(payDate.Date, out DateTime newDate)
                        ? newDate.Date
                        : payDate.Date;

                    string key = provider + "|" + effectiveDate.ToString("yyyy-MM-dd");

                    if (!groups.TryGetValue(key, out var list))
                    {
                        list = new List<RowInfo>();
                        groups[key] = list;
                    }

                    list.Add(new RowInfo
                    {
                        RowNumber = row,
                        Provider = provider,
                        OriginalDate = payDate.Date,
                        EffectiveDate = effectiveDate.Date,
                        IsHolidayRow = holidayMap.ContainsKey(payDate.Date)
                    });
                }

                // 3) For each group where there are holiday rows mapping into an effective date:
                // Add Scheduled/Projected from holiday rows into keeper row (row already on effective date if present),
                // then delete holiday rows.
                var rowsToDelete = new List<int>();

                foreach (var kvp in groups)
                {
                    var rows = kvp.Value;

                    // Only act when there is at least one holiday row in the group
                    var holidayRows = rows.Where(r => r.IsHolidayRow).ToList();
                    if (holidayRows.Count == 0)
                        continue;

                    // Find keeper row: prefer the row that already has the effective date (non-holiday row)
                    var keeper = rows.FirstOrDefault(r => !r.IsHolidayRow && r.OriginalDate == r.EffectiveDate);

                    // If there is no existing row with the target effective date,
                    // keep the first holiday row as the keeper and change its date to the effective date.
                    if (keeper == null)
                    {
                        keeper = holidayRows[0];
                        ws.Cells[keeper.RowNumber, dateCol].Value = keeper.EffectiveDate;
                        ws.Cells[keeper.RowNumber, dateCol].Style.Numberformat.Format = "mm/dd/yyyy";

                        // This keeper is no longer treated as a deletion target.
                        holidayRows.RemoveAt(0);
                    }

                    // Sum Scheduled and Projected from the remaining holiday rows (and only those rows)
                    decimal addScheduled = 0m;
                    decimal addProjected = 0m;

                    foreach (var hr in holidayRows)
                    {
                        addScheduled += GetDecimal(ws.Cells[hr.RowNumber, scheduledCol].Value);
                        addProjected += GetDecimal(ws.Cells[hr.RowNumber, projectedCol].Value);
                    }

                    // Add into keeper row, preserving existing values
                    decimal keeperScheduled = GetDecimal(ws.Cells[keeper.RowNumber, scheduledCol].Value);
                    decimal keeperProjected = GetDecimal(ws.Cells[keeper.RowNumber, projectedCol].Value);

                    ws.Cells[keeper.RowNumber, scheduledCol].Value = keeperScheduled + addScheduled;
                    ws.Cells[keeper.RowNumber, projectedCol].Value = keeperProjected + addProjected;

                    // Delete the holiday rows we aggregated from
                    foreach (var hr in holidayRows)
                    {
                        rowsToDelete.Add(hr.RowNumber);
                    }
                }

                // 4) Delete rows bottom-up
                rowsToDelete = rowsToDelete.Distinct().OrderByDescending(r => r).ToList();
                foreach (int r in rowsToDelete)
                {
                    ws.DeleteRow(r);
                }

                package.Save();
            }
        }

        private class RowInfo
        {
            public int RowNumber { get; set; }
            public string Provider { get; set; } = "";
            public DateTime OriginalDate { get; set; }
            public DateTime EffectiveDate { get; set; }
            public bool IsHolidayRow { get; set; }
        }

        public static Dictionary<DateTime, DateTime> LoadHolidayMap()
        {
            var map = new Dictionary<DateTime, DateTime>();

            string holidaysPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\Holidays\Holidays.xlsx");
            using (var pkg = new ExcelPackage(new FileInfo(holidaysPath)))
            {
                var ws = pkg.Workbook.Worksheets[0];
                if (ws.Dimension == null) return map;

                int lastRow = ws.Dimension.End.Row;

                for (int row = 2; row <= lastRow; row++)
                {
                    if (!TryGetDateFromCell(ws.Cells[row, 1].Value, out DateTime holiday))
                        continue;

                    if (!TryGetDateFromCell(ws.Cells[row, 2].Value, out DateTime newDate))
                        continue;

                    map[holiday.Date] = newDate.Date;
                }
            }

            return map;
        }

        public static int GetColumnIndex(ExcelWorksheet ws, string headerName, int headerRow, int lastCol)
        {
            for (int col = 1; col <= lastCol; col++)
            {
                if (string.Equals(ws.Cells[headerRow, col].Text?.Trim(), headerName, StringComparison.OrdinalIgnoreCase))
                    return col;
            }
            return -1;
        }

        public static bool TryGetDateFromCell(object value, out DateTime date)
        {
            if (value is DateTime dt)
            {
                date = dt.Date;
                return true;
            }

            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(s))
            {
                date = default;
                return false;
            }

            // Excel OADate sometimes comes as double-like
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double oa))
            {
                try
                {
                    date = DateTime.FromOADate(oa).Date;
                    return true;
                }
                catch { }
            }

            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
            {
                date = parsed.Date;
                return true;
            }

            date = default;
            return false;
        }

        public static decimal GetDecimal(object value)
        {
            if (value == null) return 0m;

            if (value is decimal dec) return dec;
            if (value is double d) return Convert.ToDecimal(d);
            if (value is int i) return i;
            if (value is long l) return l;

            // Handles "$ 3,844.73", "3,844.73", "-", etc.
            string s = Convert.ToString(value, CultureInfo.InvariantCulture)?.Trim() ?? "";
            if (string.IsNullOrWhiteSpace(s)) return 0m;
            if (s == "-" || s.Equals("null", StringComparison.OrdinalIgnoreCase)) return 0m;

            s = s.Replace("$", "").Replace(",", "").Trim();

            if (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal result))
                return result;

            return 0m;
        }

        public static void UpdateDatesUsingHolidays_ReplaceOnly(string excelFilePath, string columnName)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            // 1. Load Holidays.xlsx (..\Holidays\Holidays.xlsx) into a DataTable
            var holidaysTable = new DataTable();
            holidaysTable.Columns.Add("Holiday", typeof(string)); // MM/DD/YYYY
            holidaysTable.Columns.Add("NewDate", typeof(string)); // MM/DD/YYYY

            string holidaysPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, @"..\Holidays\Holidays.xlsx");

            using (var holidaysPackage = new ExcelPackage(new FileInfo(holidaysPath)))
            {
                var holidaysWs = holidaysPackage.Workbook.Worksheets[0]; // first worksheet
                if (holidaysWs.Dimension == null)
                    return;

                int lastRow = holidaysWs.Dimension.End.Row;

                for (int row = 2; row <= lastRow; row++) // assume headers in row 1
                {
                    string holidayStr = HolidayReplacer_GetDateCellAsString(holidaysWs.Cells[row, 1].Value);
                    if (string.IsNullOrWhiteSpace(holidayStr))
                        continue;

                    string newDateStr = HolidayReplacer_GetDateCellAsString(holidaysWs.Cells[row, 2].Value);
                    holidaysTable.Rows.Add(holidayStr, newDateStr);
                }
            }

            // 2. Build a lookup dictionary from holiday date to new date
            var holidayLookup = new Dictionary<DateTime, string>();

            foreach (DataRow row in holidaysTable.Rows)
            {
                if (DateTime.TryParseExact(row["Holiday"]?.ToString(),
                                           "MM/dd/yyyy",
                                           CultureInfo.InvariantCulture,
                                           DateTimeStyles.None,
                                           out DateTime holidayDate))
                {
                    holidayLookup[holidayDate.Date] = row["NewDate"]?.ToString() ?? "";
                }
            }

            // 3. Open the target Excel file and find the requested column
            using (var package = new ExcelPackage(new FileInfo(excelFilePath)))
            {
                var ws = package.Workbook.Worksheets[0]; // first worksheet
                if (ws.Dimension == null)
                    return;

                int lastRow = ws.Dimension.End.Row;
                int lastCol = ws.Dimension.End.Column;

                int targetColIndex = -1;
                for (int col = 1; col <= lastCol; col++)
                {
                    if (string.Equals(ws.Cells[1, col].Text?.Trim(),
                                      columnName?.Trim(),
                                      StringComparison.OrdinalIgnoreCase))
                    {
                        targetColIndex = col;
                        break;
                    }
                }

                if (targetColIndex == -1)
                    throw new ArgumentException($"Column '{columnName}' was not found in the first worksheet.", nameof(columnName));

                // 4. Scan that column and replace dates that match a holiday
                for (int row = 2; row <= lastRow; row++)
                {
                    object cellValue = ws.Cells[row, targetColIndex].Value;
                    if (cellValue == null)
                        continue;

                    if (!HolidayReplacer_TryGetDateFromCell(cellValue, out DateTime currentDate))
                        continue;

                    if (holidayLookup.TryGetValue(currentDate.Date, out string newDateStr) &&
                        DateTime.TryParseExact(newDateStr,
                                               "MM/dd/yyyy",
                                               CultureInfo.InvariantCulture,
                                               DateTimeStyles.None,
                                               out DateTime newDate))
                    {
                        // Write the new date as a real date with the proper format
                        ws.Cells[row, targetColIndex].Value = newDate;
                        ws.Cells[row, targetColIndex].Style.Numberformat.Format = "mm/dd/yyyy";
                    }
                }

                package.Save();
            }
        }

        public static string HolidayReplacer_GetDateCellAsString(object value)
        {
            if (value == null)
                return null;

            if (value is DateTime dt)
                return dt.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);

            // EPPlus may give OADate as double
            if (double.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                                NumberStyles.Any,
                                CultureInfo.InvariantCulture,
                                out double oa))
            {
                try
                {
                    return DateTime.FromOADate(oa).ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
                }
                catch
                {
                    // fall through
                }
            }

            // Last try: regular parse
            if (DateTime.TryParse(Convert.ToString(value, CultureInfo.InvariantCulture),
                                  CultureInfo.InvariantCulture,
                                  DateTimeStyles.None,
                                  out DateTime parsed))
            {
                return parsed.ToString("MM/dd/yyyy", CultureInfo.InvariantCulture);
            }

            // If it is already a string, return it
            return value.ToString();
        }

        public static bool HolidayReplacer_TryGetDateFromCell(object value, out DateTime date)
        {
            if (value is DateTime dt)
            {
                date = dt.Date;
                return true;
            }

            string s = Convert.ToString(value, CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(s))
            {
                date = default;
                return false;
            }

            // Excel OADate sometimes comes as numeric string/double
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double oa))
            {
                try
                {
                    date = DateTime.FromOADate(oa).Date;
                    return true;
                }
                catch
                {
                    // fall through
                }
            }

            if (DateTime.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsed))
            {
                date = parsed.Date;
                return true;
            }

            date = default;
            return false;
        }
    }

    public class discrepencies
        {
            public string PatientName { get; set; }
            public string AdmitDate { get; set; }
            public string Agency { get; internal set; }
        }
}
