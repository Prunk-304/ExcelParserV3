namespace ExcelParserV3;
public class ColumsListBase
{
    public List<string> number { get; set; }
    public List<string> indicator { get; set; }
    public List<string> metric { get; set; }
    public List<string> reportBeforeLast { get; set; }
    public List<string> reportLast { get; set; }
    public List<string> evaluationIndicator { get; set; }
    public List<string> forecastOneYearConservative { get; set; }
    public List<string> forecastOneYearBase { get; set; }
    public List<string> forecastTwoYearConservative { get; set; }
    public List<string> forecastTwoYearBase { get; set; }
    public List<string> forecastThreeYearConservative { get; set; }
    public List<string> forecastThreeYearBase { get; set; }
    public List<string> generalizedIndicator { get; set; }

    public ColumsListBase()
    {
        number = new List<string>();
        indicator = new List<string>();
        metric = new List<string>();
        reportBeforeLast = new List<string>();
        reportLast = new List<string>();
        evaluationIndicator = new List<string>();
        forecastOneYearConservative = new List<string>();
        forecastOneYearBase = new List<string>();
        forecastTwoYearConservative = new List<string>();
        forecastTwoYearBase = new List<string>();
        forecastThreeYearConservative = new List<string>();
        forecastThreeYearBase = new List<string>();
        generalizedIndicator = new List<string>();
    }
}