using NLog;
using OfficeOpenXml;

namespace ExcelParserV3;

public class ExcelParser
{
    private readonly ColumsListBase _excelParse;
    private readonly ILogger _logger;

    public ExcelParser(ColumsListBase excelParse, LoggerFactory loggerFactory)
    {
        _excelParse = excelParse;
        _logger = loggerFactory.CreateLogger();
    }

    public void ParseExcelFile(string filePath)
    {
        try
        {   string check;
            byte[] bin = File.ReadAllBytes(filePath);
            using (MemoryStream stream = new MemoryStream(bin))
            using (ExcelPackage excelPackage = new ExcelPackage(stream))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    for (int j = worksheet.Dimension.Start.Column; j <= worksheet.Dimension.End.Column; j++)
                    {
                        for (int i = worksheet.Dimension.Start.Row + 10; i <= worksheet.Dimension.End.Row; i++)
                        {
                            if (j == 2 && (i == 11 || i == 22 || i == 26 || i == 66 || i == 73 || i == 78 || i == 87 ||
                                           i == 97 || i == 101 || i == 117 || i == 154 || i == 161))
                            {
                                _excelParse.generalizedIndicator.Add(worksheet.Cells[i, j].Value.ToString());
                                i++;
                            }

                            if (worksheet.Cells[i, j].Value != null)
                            {
                                check = worksheet.Cells[i, j].Value.ToString();
                                if (check == "Примечание:")
                                {
                                    break;
                                }

                                switch (j)
                                {
                                    case 1:
                                        _excelParse.number.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 2:
                                        _excelParse.indicator.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 3:
                                        _excelParse.metric.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 4:
                                        _excelParse.reportBeforeLast.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 5:
                                        _excelParse.reportLast.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 6:
                                        _excelParse.evaluationIndicator.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 7:
                                        _excelParse.forecastOneYearConservative.Add(worksheet.Cells[i, j].Value
                                            .ToString());
                                        break;
                                    case 8:
                                        _excelParse.forecastOneYearBase.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 9:
                                        _excelParse.forecastTwoYearConservative.Add(worksheet.Cells[i, j].Value
                                            .ToString());
                                        break;
                                    case 10:
                                        _excelParse.forecastTwoYearBase.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                    case 11:
                                        _excelParse.forecastThreeYearConservative.Add(worksheet.Cells[i, j].Value
                                            .ToString());
                                        break;
                                    case 12:
                                        _excelParse.forecastThreeYearBase.Add(worksheet.Cells[i, j].Value.ToString());
                                        break;
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            _logger.Error(ex, "Ошибка при попытке парсинга Excel файла");
        }
    }
}