using Microsoft.Extensions.Configuration;
using NLog;
using NLog.Config;
using NLog.Extensions.Logging;

namespace ExcelParserV3
{
    public class NLogConfigReader
    {
        private readonly string _filename;

        public NLogConfigReader(string filename)
        {
            _filename = filename;
        }
        public LoggingConfiguration GetNLogConfiguration()
        {
            var builder = new ConfigurationBuilder()
                    .SetBasePath(Directory.GetCurrentDirectory())
                    .AddJsonFile(_filename, optional: false, reloadOnChange: true);
                var configuration = builder.Build();

            try
            {
                var nlogConfig = new NLogLoggingConfiguration(configuration.GetSection("NLog"));
                try
                {
                    LogManager.Configuration = nlogConfig;
                }
                catch (NLogConfigurationException ex)
                {
                    Console.WriteLine($"Ошибка установки конфигурации NLog: {ex.Message}");
                }
                return LogManager.Configuration;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Ошибка чтения конфигурационного файла: {ex.Message}");
                return null;
            }
        }
    }
}