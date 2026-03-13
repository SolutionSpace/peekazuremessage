using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PeekAzureMessage
{
    public class ExtractMessages
    {
        public async static Task extractMessages()
        {
            string projectDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\.."));
            string excelFolderPath = Path.Combine(projectDirectory, "ExcelSheets");

            string searchResultsPath = Path.Combine(excelFolderPath, "SearchResults");
            // Create folder if not exists
            if (!Directory.Exists(searchResultsPath))
                Directory.CreateDirectory(searchResultsPath);

            string safeTimestamp = DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss");
            string outputFileName = $"SearchResults_{safeTimestamp}.xlsx";
            string outputFilePath = Path.Combine(searchResultsPath, outputFileName);

            //List of string to be searched in the file.
            Console.WriteLine("Enter the search terms (comma separated):");
            string input = Console.ReadLine() ?? string.Empty;

            // Convert input to list of trimmed, non-empty search terms
            List<string> searchTerms = input
                .Split(',', StringSplitOptions.RemoveEmptyEntries)
                .Select(s => s.Trim())
                .ToList();

            if (searchTerms.Count == 0)
            {
                Console.WriteLine("No search terms provided.");
                return;
            }

            await SearchAndExtractRowsAsync(excelFolderPath, searchTerms, outputFilePath);

        }

        public static async Task SearchAndExtractRowsAsync(string folderPath, List<string> searchTerms, string outputFile)
        {
            if (!Directory.Exists(folderPath))
            {
                Console.WriteLine("Folder path not found.");
                return;
            }

            if (searchTerms == null || searchTerms.Count == 0)
            {
                Console.WriteLine("No search terms provided.");
                return;
            }

            await Task.Run(() =>
            {
                var excelFiles = Directory.GetFiles(folderPath, "*.xlsx", SearchOption.TopDirectoryOnly);
                if (excelFiles.Length == 0)
                {
                    Console.WriteLine("No Excel files found in the folder.");
                    return;
                }

                using var resultWorkbook = new XLWorkbook();
                var resultSheet = resultWorkbook.Worksheets.Add("Results");

                bool headerAdded = false;
                int resultRow = 1;
                int totalColumns = 0;

                // Keep track of added Message IDs to avoid duplicates
                HashSet<string> addedMessageIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                foreach (var file in excelFiles)
                {
                    Console.WriteLine($"Processing file: {Path.GetFileName(file)}");

                    try
                    {
                        using var workbook = new XLWorkbook(file);

                        foreach (var sheet in workbook.Worksheets)
                        {
                            var headerRow = sheet.FirstRowUsed();
                            if (headerRow == null) continue;

                            totalColumns = headerRow.LastCellUsed().Address.ColumnNumber;
                            var headerCells = headerRow.Cells(1, totalColumns).Select(c => c.GetValue<string>().Trim()).ToList();

                            int msgTextColIndex = headerCells.FindIndex(h => h.Equals("Message Text", StringComparison.OrdinalIgnoreCase)) + 1;
                            int msgIdColIndex = headerCells.FindIndex(h => h.Equals("Message ID", StringComparison.OrdinalIgnoreCase)) + 1;

                            if (msgTextColIndex == 0 || msgIdColIndex == 0)
                            {
                                Console.WriteLine($"Skipping sheet '{sheet.Name}' — required columns missing.");
                                continue;
                            }

                            if (!headerAdded)
                            {
                                for (int col = 1; col <= totalColumns; col++)
                                    resultSheet.Cell(resultRow, col).Value = headerRow.Cell(col).GetValue<string>();
                                resultRow++;
                                headerAdded = true;
                            }

                            var allRows = sheet.RowsUsed().Skip(1).ToList();

                            foreach (var row in allRows)
                            {
                                string messageText = row.Cell(msgTextColIndex).GetValue<string>();

                                // Search the "Message Text" column only
                                if (searchTerms.Any(term => messageText.Contains(term, StringComparison.OrdinalIgnoreCase)))
                                {
                                    // Get the Message ID of the matched record
                                    string messageId = row.Cell(msgIdColIndex).GetValue<string>();

                                    if (!string.IsNullOrWhiteSpace(messageId))
                                    {
                                        // Find all rows that have this Message ID
                                        var matchingIdRows = allRows
                                            .Where(r => r.Cell(msgIdColIndex).GetValue<string>()
                                            .Equals(messageId, StringComparison.OrdinalIgnoreCase))
                                            .ToList();

                                        foreach (var idRow in matchingIdRows)
                                        {
                                            string idValue = idRow.Cell(msgIdColIndex).GetValue<string>();


                                            for (int col = 1; col <= totalColumns; col++)
                                                resultSheet.Cell(resultRow, col).Value = idRow.Cell(col).Value;

                                            resultRow++;
                                            addedMessageIds.Add(idValue);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine($"Error processing {file}: {ex.Message}");
                    }
                }

                if (resultRow > 1)
                {
                    resultWorkbook.SaveAs(outputFile);
                    Console.WriteLine($"Search complete. Results saved to: {outputFile}");
                }
                else
                {
                    Console.WriteLine("No matches found in any file.");
                }
            });
        }
        //end of function
    }
}
