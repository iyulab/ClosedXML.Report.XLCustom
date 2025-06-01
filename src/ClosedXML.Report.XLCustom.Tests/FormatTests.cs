using ClosedXML.Excel;
using FluentAssertions;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests;

public class FormatTests : TestBase
{
    public FormatTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void StandardNumericFormats_ShouldApplyCorrectly()
    {
        // Arrange
        var testModel = new FormatTestModel
        {
            Number = 1234.56m
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("NumericFormatTest");

        // Test various numeric formats
        sheet.Cell("A1").Value = "Currency (C):";
        sheet.Cell("B1").Value = "{{Number:C}}";

        sheet.Cell("A2").Value = "Number (N2):";
        sheet.Cell("B2").Value = "{{Number:N2}}";

        sheet.Cell("A3").Value = "Percent (P):";
        sheet.Cell("B3").Value = "{{Number:P}}";

        sheet.Cell("A4").Value = "Fixed-point (F3):";
        sheet.Cell("B4").Value = "{{Number:F3}}";

        sheet.Cell("A5").Value = "Exponential (E2):";
        sheet.Cell("B5").Value = "{{Number:E2}}";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms).Preprocess();
        template.AddVariable(testModel);
        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");
        var ws = template.Workbook.Worksheet("NumericFormatTest");

        // Verify cell values are properly formatted
        ws.Cell("B1").GetFormattedString().Should().Contain("1,234.56"); // Currency format
        ws.Cell("B2").GetFormattedString().Should().Be("1,234.56"); // Number format with 2 decimal places
        ws.Cell("B3").GetFormattedString().Should().Contain("123,456"); // Percent format (converts to percentage)
        ws.Cell("B4").GetFormattedString().Should().Be("1234.560"); // Fixed-point with 3 decimal places
        ws.Cell("B5").GetFormattedString().Should().Match("*E+*"); // Exponential format
    }

    [Fact]
    public void CustomNumericFormats_ShouldApplyCorrectly()
    {
        // Arrange
        var testModel = new FormatTestModel
        {
            Number = 1234.56m
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("CustomNumericFormatTest");

        // Test custom numeric formats
        sheet.Cell("A1").Value = "Custom Format (#,##0.000):";
        sheet.Cell("B1").Value = "{{Number:#,##0.000}}";

        sheet.Cell("A2").Value = "Custom Format (0.00):";
        sheet.Cell("B2").Value = "{{Number:0.00}}";

        sheet.Cell("A3").Value = "Custom Format (#,##0):";
        sheet.Cell("B3").Value = "{{Number:#,##0}}";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms).Preprocess();
        template.AddVariable(testModel);
        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");
        var ws = template.Workbook.Worksheet("CustomNumericFormatTest");

        // Verify cell values are properly formatted
        ws.Cell("B1").GetFormattedString().Should().Be("1,234.560"); // Custom format with 3 decimal places
        ws.Cell("B2").GetFormattedString().Should().Be("1234.56"); // Custom format with 2 decimal places
        ws.Cell("B3").GetFormattedString().Should().Be("1,235"); // Custom format with no decimal places (rounded)
    }

    [Fact]
    public void DateFormats_ShouldApplyCorrectly()
    {
        // Arrange
        var testModel = new FormatTestModel
        {
            Date = new DateTime(2025, 5, 8, 13, 45, 30)
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("DateFormatTest");

        // Test various date formats
        sheet.Cell("A1").Value = "Date (d):";
        sheet.Cell("B1").Value = "{{Date:d}}";

        sheet.Cell("A2").Value = "Custom Date (yyyy-MM-dd):";
        sheet.Cell("B2").Value = "{{Date:yyyy-MM-dd}}";

        sheet.Cell("A3").Value = "Long Date (D):";
        sheet.Cell("B3").Value = "{{Date:D}}";

        sheet.Cell("A4").Value = "Date and Time (f):";
        sheet.Cell("B4").Value = "{{Date:f}}";

        sheet.Cell("A5").Value = "Custom Time (HH:mm:ss):";
        sheet.Cell("B5").Value = "{{Date:HH:mm:ss}}";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms).Preprocess();
        template.AddVariable(testModel);
        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");
        var ws = template.Workbook.Worksheet("DateFormatTest");

        // Verify date formats - exact strings will depend on culture, so we check for expected patterns
        ws.Cell("B2").GetFormattedString().Should().Be("2025-05-08"); // Custom ISO date format
        ws.Cell("B5").GetFormattedString().Should().Be("13:45:30"); // Custom time format
    }

    [Fact]
    public void AlternatingFormats_ShouldPreferFormatTags()
    {
        // Arrange
        var testModel = new FormatTestModel
        {
            Number = 1234.56m,
            Date = new DateTime(2025, 5, 8)
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("AlternatingFormatTest");

        // Set cell formats in Excel
        sheet.Cell("A1").Value = "Excel Format vs Tag Format:";
        sheet.Cell("B1").Value = "{{Number:C}}";
        sheet.Cell("B1").Style.NumberFormat.Format = "0.0000"; // Excel format should be overridden by the tag

        sheet.Cell("A2").Value = "Date Excel Format vs Tag Format:";
        sheet.Cell("B2").Value = "{{Date:yyyy-MM-dd}}";
        sheet.Cell("B2").Style.DateFormat.Format = "MM/dd/yyyy"; // Excel format should be overridden by the tag

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms).Preprocess();
        template.AddVariable(testModel);
        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");
        var ws = template.Workbook.Worksheet("AlternatingFormatTest");

        // Verify that format tags override Excel cell formats
        ws.Cell("B1").GetFormattedString().Should().Contain("1,234.56"); // Should use currency format, not 0.0000
        ws.Cell("B2").GetFormattedString().Should().Be("2025-05-08"); // Should use ISO format, not MM/dd/yyyy
    }

    [Fact]
    public void FormatWithCalculations_ShouldFormatAfterCalculation()
    {
        // Arrange
        var testModel = new CalculationModel
        {
            Price = 19.99m,
            Quantity = 3
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("CalcFormatTest");

        // Test formatting of calculated values
        sheet.Cell("A1").Value = "Unit Price:";
        sheet.Cell("B1").Value = "{{Price:C}}";

        sheet.Cell("A2").Value = "Quantity:";
        sheet.Cell("B2").Value = "{{Quantity}}";

        sheet.Cell("A3").Value = "Total (Currency):";
        sheet.Cell("B3").Value = "{{Price * Quantity:C}}"; // Format after calculation

        sheet.Cell("A4").Value = "Total (Fixed):";
        sheet.Cell("B4").Value = "{{Price * Quantity:F2}}"; // Format after calculation

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms).Preprocess();
        template.AddVariable(testModel);
        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");
        var ws = template.Workbook.Worksheet("CalcFormatTest");

        // Verify formatted calculation results
        var calculatedValue = testModel.Price * testModel.Quantity; // 59.97
        ws.Cell("B3").GetFormattedString().Should().Contain("59.97"); // Should be formatted as currency
        ws.Cell("B4").GetFormattedString().Should().Be("59.97"); // Should be formatted with 2 decimal places
    }

    [Fact]
    public void FormatInRanges_ShouldApplyToAllItems()
    {
        // Arrange
        var products = new List<ProductModel>
        {
            new ProductModel { Name = "Product 1", Price = 19.99m, Quantity = 3 },
            new ProductModel { Name = "Product 2", Price = 29.99m, Quantity = 2 },
            new ProductModel { Name = "Product 3", Price = 9.99m, Quantity = 5 }
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("RangeFormatTest");

        // Template expressions in row 3
        sheet.Cell("A3").Value = "{{item.Name}}";
        sheet.Cell("B3").Value = "{{item.Price:C}}";
        sheet.Cell("C3").Value = "{{item.Quantity}}";
        sheet.Cell("D3").Value = "{{item.Price * item.Quantity:C}}";

        // Service row in row 4
        sheet.Cell("A4").Value = ""; // Can be empty or contain service tags
        sheet.Cell("B4").Value = "";
        sheet.Cell("C4").Value = "";
        sheet.Cell("D4").Value = "";

        // Service column required for vertical tables
        sheet.Cell("E3").Value = "";

        // Define range including both template row and service row
        var productRange = sheet.Range("A3:E4");
        workbook.DefinedNames.Add("Products", productRange);

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms).Preprocess();
        template.AddVariable("Products", products);
        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");
        var ws = template.Workbook.Worksheet("RangeFormatTest");

        // Verify formatted values in range
        for (int i = 0; i < products.Count; i++)
        {
            int row = i + 3; // Starting from row 3 where the template is defined
            var product = products[i];

            ws.Cell($"A{row}").GetString().Should().Be(product.Name);
            ws.Cell($"B{row}").GetFormattedString().Should().Contain(product.Price.ToString("0.00"));
            ws.Cell($"C{row}").GetValue<int>().Should().Be(product.Quantity);
            ws.Cell($"D{row}").GetFormattedString().Should().Contain((product.Price * product.Quantity).ToString("0.00"));
        }
    }

    public class FormatTestModel
    {
        public string Text { get; set; }
        public decimal Number { get; set; }
        public DateTime Date { get; set; }
    }

    public class CalculationModel
    {
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }

    public class ProductModel
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
        public int Quantity { get; set; }
    }
}