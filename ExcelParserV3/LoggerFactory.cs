using NLog;

namespace ExcelParserV3;

public class LoggerFactory
{
    private readonly NLogConfigReader _configReader;

    public LoggerFactory(NLogConfigReader configReader)
    {
        _configReader = configReader;
    }

    public ILogger CreateLogger()
    {
        var nlogConfig = _configReader.GetNLogConfiguration();
        LogManager.Configuration = nlogConfig;
        return LogManager.GetCurrentClassLogger();
    }
}