using Azure.Messaging.ServiceBus;
using ClosedXML.Excel;

namespace PeekAzureMessage
{
    public class ReadMessages
    {
        public async static Task readMessages(string connectionString, string queueName)
        {
            // 1️ Get Date Range Input
            Console.Write("Enter FROM date (yyyy-MM-dd): ");
            DateTime fromDate = DateTime.Parse(Console.ReadLine()).Date; // 00:00:00

            Console.Write("Enter TO date (yyyy-MM-dd) [press Enter for today]: ");
            string toDateInput = Console.ReadLine();
            DateTime toDate = string.IsNullOrWhiteSpace(toDateInput)
                ? DateTime.Today.AddDays(1).AddSeconds(-1)  // today 23:59:59
                : DateTime.Parse(toDateInput).Date.AddDays(1).AddSeconds(-1);

            Console.Write("Enter starting Sequence Number [press Enter for 0]: ");
            string seqInput = Console.ReadLine();
            long startSequence = string.IsNullOrWhiteSpace(seqInput) ? 0 : long.Parse(seqInput);

            Console.WriteLine($"\nProcess Starting time " + DateTime.Now + "\n\n");
            Console.WriteLine($"\nSearching messages from {fromDate:G} to {toDate:G}\n");

            // 2️ Setup Service Bus Client
            await using var client = new ServiceBusClient(connectionString);
            var receiver = client.CreateReceiver(queueName);

            int batchSize = 500;
            long sequenceNumber = startSequence;
            var messagesInRange = new List<ServiceBusReceivedMessage>();

            // 3️ Peek Messages Batch-wise
            while (true)
            {
                var messages = await receiver.PeekMessagesAsync(batchSize, sequenceNumber);
                if (messages.Count == 0)
                    break;

                foreach (var msg in messages)
                {
                    sequenceNumber = msg.SequenceNumber + 1;
                    // If message time exceeds toDate → stop further processing
                    if (msg.EnqueuedTime > toDate)
                    {
                        Console.WriteLine("Reached messages beyond the requested date range. Stopping...");
                        break;
                    }

                    // Filter by EnqueuedTimeUtc
                    if (msg.EnqueuedTime >= fromDate && msg.EnqueuedTime <= toDate)
                    {
                        messagesInRange.Add(msg);
                    }
                }
                // If we reached the date limit, stop reading more batches
                if (messages.Count > 0 && messages[^1].EnqueuedTime > toDate)
                    break;
            }

            Console.WriteLine($"Total messages found in range: {messagesInRange.Count}");

            // 4️ Export to Excel (only if there are messages)
            if (messagesInRange.Count > 0)
            {
                string projectDirectory = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory, @"..\..\.."));
                string excelFolderPath = Path.Combine(projectDirectory, "ExcelSheets");

                // Create folder if not exists
                if (!Directory.Exists(excelFolderPath))
                    Directory.CreateDirectory(excelFolderPath);

                string fileName = $"Messages_{fromDate:yyyy-MM-dd}_to_{toDate:yyyy-MM-dd}.xlsx";
                string filePath = Path.Combine(excelFolderPath, fileName);

                using var workbook = new XLWorkbook();
                var worksheet = workbook.Worksheets.Add("Logs");

                // Header row
                worksheet.Cell(1, 1).Value = "Sequence Number";
                worksheet.Cell(1, 2).Value = "Message ID";
                worksheet.Cell(1, 3).Value = "Enqueued Time";
                worksheet.Cell(1, 4).Value = "State";
                worksheet.Cell(1, 5).Value = "Label / Subject";
                worksheet.Cell(1, 6).Value = "Message Text";

                int row = 2;
                foreach (var msg in messagesInRange)
                {
                    worksheet.Cell(row, 1).Value = msg.SequenceNumber;
                    worksheet.Cell(row, 2).Value = msg.MessageId;
                    worksheet.Cell(row, 3).Value = msg.EnqueuedTime.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss");
                    worksheet.Cell(row, 4).Value = msg.State.ToString();
                    worksheet.Cell(row, 5).Value = msg.Subject ?? "";
                    string bodyText = msg.Body.ToString();
                    if (bodyText.Length > 32767)
                    {
                        // Save full message to separate text file
                        string txtDir = Path.Combine(excelFolderPath, "FullMessages");
                        Directory.CreateDirectory(txtDir);

                        string txtFile = Path.Combine(txtDir, $"Message_{msg.SequenceNumber}.txt");
                        await File.WriteAllTextAsync(txtFile, bodyText);

                        // Store truncated preview + note about full file
                        worksheet.Cell(row, 6).Value = bodyText.Substring(0, 30000) +
                            $"... [Full message saved in {Path.GetFileName(txtFile)}]";
                    }
                    else
                    {
                        worksheet.Cell(row, 6).Value = bodyText;
                    }
                    row++;
                }

                // Auto-fit and style headers
                worksheet.Columns().AdjustToContents();
                var headerRange = worksheet.Range("A1:F1");
                headerRange.Style.Font.Bold = true;
                headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

                workbook.SaveAs(filePath);

                Console.WriteLine($"\nExcel file saved successfully at:\n{filePath}");
            }
            else
            {
                Console.WriteLine("No messages found within the given date range.");
            }
        }
    }
}
