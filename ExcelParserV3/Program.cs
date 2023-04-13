using Microsoft.Extensions.DependencyInjection;

namespace ExcelParserV3
{
    class Program
    {
        static void Main(string[] args)
        {
            if (args.Length == 0)
            {
              Console.WriteLine("Ошибка запуска программы. Введите путь к excel файлу." +
                                "Для подробной справки введите --help или -h");
              return;
            }

            if (args.Length >0 && (args[0] == "--help"|| args[0] == "-h"))
            {
                Console.WriteLine("Для работы программы введите адрес файла в формате:" +
                                  " -p D:\\Solutions\\Example.xlsx");
                return;
            }
            if (args.Length == 2 && (args[0]=="--parse" || args[0]=="-p"))
            {
                string filepath = args[1];
                var startup = new Startup();

                var services = new ServiceCollection();
                startup.ConfigureServices(services);
                var serviceProvider = services.BuildServiceProvider();

                var parser = serviceProvider.GetService<ExcelParser>();
                parser.ParseExcelFile(filepath);

                var printToConsole = serviceProvider.GetService<PrintToConsole>();
                printToConsole.PrintLists();
            }
            else
            {
                Console.WriteLine("Ошибка запуска программы. Для подробной справки введите --help или -h");
            }
        }
    }
}