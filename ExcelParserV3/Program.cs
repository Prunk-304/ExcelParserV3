using Microsoft.Extensions.DependencyInjection;

namespace ExcelParserV3
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
              Console.WriteLine("Ошибка запуска программы. Введите путь к excel файлу");
              return;
            }
            string filepath = args[0];
            var startup = new Startup();

            var services = new ServiceCollection();
            startup.ConfigureServices(services);
            var serviceProvider = services.BuildServiceProvider();

            var parser = serviceProvider.GetService<ExcelParser>();
            parser.ParseExcelFile(filepath);

            var printToConsole = serviceProvider.GetService<PrintToConsole>();
            printToConsole.PrintLists();
        }
    }
}