using NLog;

namespace ExcelParserV3;

public class PrintToConsole
{
    private readonly ColumsListBase _excelParse;
    private readonly ILogger _logger;

    public PrintToConsole(ColumsListBase excelParse, LoggerFactory loggerFactory)
    {
        _excelParse = excelParse;
        _logger = loggerFactory.CreateLogger();
    }

    public void PrintLists()
    {
        _logger.Info("Печать списка номеров:");
        foreach (var item in _excelParse.number)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка названий показателей:");
        foreach (var item in _excelParse.indicator)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка единиц измерения:");
        foreach (var item in _excelParse.metric)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка отчетов за позапрошлый год:");
        foreach (var item in _excelParse.reportBeforeLast)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка отчетов за прошлый год:");
        foreach (var item in _excelParse.reportLast)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка значений показателей за этот год:");
        foreach (var item in _excelParse.evaluationIndicator)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка прогнозируемых значений показателей за следующий год(Консервативный):");
        foreach (var item in _excelParse.forecastOneYearConservative)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка прогнозируемых значений показателей за следующий год(Базовый):");
        foreach (var item in _excelParse.forecastOneYearBase)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка прогнозируемых значений показателей за последующий год(Консервативный):");
        foreach (var item in _excelParse.forecastTwoYearConservative)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка прогнозируемых значений показателей за последующий год(Базовый):");
        foreach (var item in _excelParse.forecastTwoYearBase)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка прогнозируемых значений показателей за третий год(Консервативный):");
        foreach (var item in _excelParse.forecastThreeYearConservative)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка прогнозируемых значений показателей за третий год(Базовый):");
        foreach (var item in _excelParse.forecastThreeYearBase)
        {
            _logger.Info(item);
        }

        _logger.Info("Печать списка названий обобщенных показателей:");
        foreach (var item in _excelParse.generalizedIndicator)
        {
            _logger.Info(item);
        }
    }
}
