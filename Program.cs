using Azure.Messaging.ServiceBus;
using ClosedXML.Excel;
using PeekAzureMessage;

class Program
{
    private const string connectionString = "add the connection string";
    private const string queueName = "archive";

    static async Task Main(string[] args)
    {
        Console.Write("Enter your choice  ( 1 to Read Messages and  2 to Extract Messages ): ");
        var choice = Console.ReadLine();
        if (choice != null)
        {
            if (choice == "1")
               await ReadMessages.readMessages(connectionString, queueName);
            else
                await ExtractMessages.extractMessages();
        }
    }
}
