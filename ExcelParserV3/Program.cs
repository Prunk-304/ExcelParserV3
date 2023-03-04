using Microsoft.Extensions.DependencyInjection;

namespace ExcelParserV3
{
    class Program
    {
        static void Main(string[] args)
        {
            var startup = new Startup();

            var services = new ServiceCollection();
            startup.ConfigureServices(services);
            var serviceProvider = services.BuildServiceProvider();

            var parser = serviceProvider.GetService<ExcelParser>();
            parser.ParseExcelFile();

            var printToConsole = serviceProvider.GetService<PrintToConsole>();
            printToConsole.PrintLists();
        }
    }
}