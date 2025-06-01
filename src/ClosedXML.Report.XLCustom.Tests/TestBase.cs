using ClosedXML.Excel;
using System.Diagnostics;
using System.Text;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests;

/// <summary>
/// Base class for XLCustom test classes with common helper methods
/// </summary>
public abstract class TestBase : IDisposable
{
    protected readonly ITestOutputHelper _output;

    protected TestBase(ITestOutputHelper output)
    {
        _output = output ?? throw new ArgumentNullException(nameof(output));

        XLCustomRegistry.ResetAll();
        _output.WriteLine("Test initialization: Registry reset");
    }

    public void Dispose()
    {
        XLCustomRegistry.ResetAll();
        _output.WriteLine("Test cleanup: Registry reset");
    }

    /// <summary>
    /// Logs the result of template generation to the test output
    /// </summary>
    protected void LogResult(XLGenerateResult result)
    {
        if (result == null)
        {
            _output.WriteLine("Generation result is null");
            return;
        }

        _output.WriteLine($"Template generation completed. HasErrors: {result.HasErrors}");

        if (result.HasErrors)
        {
            foreach (var error in result.ParsingErrors)
            {
                _output.WriteLine($"ParsingError: {error.Message}, Range: {error.Range}");
            }
        }
    }

    /// <summary>
    /// Logs the cell values and formats of a range for debugging
    /// </summary>
    protected void LogCellInfo(IXLRange range, string title = null)
    {
        if (range == null) return;

        if (!string.IsNullOrEmpty(title))
        {
            _output.WriteLine(title);
            _output.WriteLine(new string('-', title.Length));
        }

        foreach (var row in range.Rows())
        {
            var sb = new StringBuilder();
            foreach (var cell in row.Cells())
            {
                string value = cell.GetString();
                string address = cell.Address.ToString();
                string format = cell.Style.NumberFormat.Format;

                sb.Append($"{address}='{value}' (Format: {format}) | ");
            }

            _output.WriteLine(sb.ToString());
        }
    }

    /// <summary>
    /// Creates a sample template with standard test data
    /// </summary>
    protected IXLWorkbook CreateSampleTemplate()
    {
        var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("SampleData");

        // Add header row
        sheet.Cell("A1").Value = "Product";
        sheet.Cell("B1").Value = "Price";
        sheet.Cell("C1").Value = "Quantity";
        sheet.Cell("D1").Value = "Total";

        // Style header row
        var headerRange = sheet.Range("A1:D1");
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;

        // Add sample data rows with template tags
        sheet.Cell("A2").Value = "<<Range(Products)>>";
        sheet.Cell("A3").Value = "{{item.Name}}";
        sheet.Cell("B3").Value = "{{item.Price:C}}";
        sheet.Cell("C3").Value = "{{item.Quantity}}";
        sheet.Cell("D3").Value = "{{item.Price * item.Quantity:C}}";
        sheet.Cell("A4").Value = "<<EndRange>>";

        // Add total row
        sheet.Cell("C5").Value = "Total:";
        sheet.Cell("C5").Style.Font.Bold = true;
        sheet.Cell("D5").Value = "{{Products.Sum(p => p.Price * p.Quantity):C}}";
        sheet.Cell("D5").Style.Font.Bold = true;

        return workbook;
    }

    /// <summary>
    /// Creates a list of sample products for testing
    /// </summary>
    protected List<ProductModel> CreateSampleProducts()
    {
        return new List<ProductModel>
        {
            new ProductModel { Name = "Product A", Price = 19.99m, Quantity = 3 },
            new ProductModel { Name = "Product B", Price = 29.99m, Quantity = 2 },
            new ProductModel { Name = "Product C", Price = 9.99m, Quantity = 5 }
        };
    }

    /// <summary>
    /// Sample product model for testing
    /// </summary>
    protected class ProductModel
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Quantity { get; set; }
        public bool InStock { get; set; }
    }
}

/// <summary>
/// Trace listener that redirects output to XUnit's test output
/// </summary>
public class XUnitTraceListener : TraceListener
{
    private readonly ITestOutputHelper _output;

    public XUnitTraceListener(ITestOutputHelper output)
    {
        _output = output;
    }

    public override void Write(string message)
    {
        try
        {
            _output.WriteLine(message);
        }
        catch (Exception)
        {
            // 테스트 컨텍스트 외부에서 호출되는 경우 무시
            // 필요하다면 여기서 콘솔이나 다른 로그 메커니즘으로 출력 가능
            Console.Write(message);
        }
    }

    public override void WriteLine(string message)
    {
        try
        {
            _output.WriteLine(message);
        }
        catch (Exception)
        {
            // 테스트 컨텍스트 외부에서 호출되는 경우 무시
            // 필요하다면 여기서 콘솔이나 다른 로그 메커니즘으로 출력 가능
            Console.WriteLine(message);
        }
    }
}