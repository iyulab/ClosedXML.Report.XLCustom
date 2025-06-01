using ClosedXML.Excel;
using FluentAssertions;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests;

public class BasicTests : TestBase
{
    public BasicTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void SimpleProperties_ShouldBindCorrectly()
    {
        // Arrange
        var testModel = new TestModel
        {
            StringProperty = "Test String Value",
            IntProperty = 42,
            DecimalProperty = 123.45m,
            DateProperty = new DateTime(2025, 5, 6),
            BoolProperty = true
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("PropertyTest");

        // Setup property bindings
        sheet.Cell("A1").Value = "String:";
        sheet.Cell("B1").Value = "{{StringProperty}}";
        sheet.Cell("A2").Value = "Integer:";
        sheet.Cell("B2").Value = "{{IntProperty}}";
        sheet.Cell("A3").Value = "Decimal:";
        sheet.Cell("B3").Value = "{{DecimalProperty}}";
        sheet.Cell("A4").Value = "Date:";
        sheet.Cell("B4").Value = "{{DateProperty}}";
        sheet.Cell("A5").Value = "Boolean:";
        sheet.Cell("B5").Value = "{{BoolProperty}}";

        // Save template to memory stream
        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("StringProperty", testModel.StringProperty);
        template.AddVariable("IntProperty", testModel.IntProperty);
        template.AddVariable("DecimalProperty", testModel.DecimalProperty);
        template.AddVariable("DateProperty", testModel.DateProperty);
        template.AddVariable("BoolProperty", testModel.BoolProperty);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        var ws = template.Workbook.Worksheet("PropertyTest");
        ws.Cell("B1").GetString().Should().Be("Test String Value");
        ws.Cell("B2").GetValue<int>().Should().Be(42);
        ws.Cell("B3").GetValue<double>().Should().BeApproximately(123.45, 0.001);
        ws.Cell("B4").GetDateTime().Should().Be(new DateTime(2025, 5, 6));
        ws.Cell("B5").GetBoolean().Should().BeTrue();
    }

    [Fact]
    public void NestedProperties_ShouldBindCorrectly()
    {
        // Arrange
        var nestedModel = new NestedModel
        {
            Parent = new ParentModel
            {
                Name = "Parent Name",
                Child = new ChildModel
                {
                    Name = "Child Name",
                    Age = 10
                }
            }
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("NestedTest");

        // Setup nested property bindings
        sheet.Cell("A1").Value = "Parent:";
        sheet.Cell("B1").Value = "{{model.Parent.Name}}";
        sheet.Cell("A2").Value = "Child:";
        sheet.Cell("B2").Value = "{{model.Parent.Child.Name}}";
        sheet.Cell("A3").Value = "Child Age:";
        sheet.Cell("B3").Value = "{{model.Parent.Child.Age}}";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("model", nestedModel);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        var ws = template.Workbook.Worksheet("NestedTest");
        ws.Cell("B1").GetString().Should().Be("Parent Name");
        ws.Cell("B2").GetString().Should().Be("Child Name");
        ws.Cell("B3").GetValue<int>().Should().Be(10);
    }

    [Fact]
    public void FormatSpecifiers_ShouldApplyCorrectly()
    {
        // Arrange
        var formatModel = new FormatModel
        {
            Currency = 1234.56m,
            Percentage = 0.75m,
            LargeNumber = 1234567,
            Date = new DateTime(2025, 5, 6)
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("FormatTest");

        // Setup cells with format
        sheet.Cell("A1").Value = "Currency:";
        sheet.Cell("B1").Value = "{{Currency}}";
        sheet.Cell("B1").Style.NumberFormat.Format = "$#,##0.00";

        sheet.Cell("A2").Value = "Percentage:";
        sheet.Cell("B2").Value = "{{Percentage}}";
        sheet.Cell("B2").Style.NumberFormat.Format = "0.00%";

        sheet.Cell("A3").Value = "Large Number:";
        sheet.Cell("B3").Value = "{{LargeNumber}}";
        sheet.Cell("B3").Style.NumberFormat.Format = "#,##0";

        sheet.Cell("A4").Value = "Date:";
        sheet.Cell("B4").Value = "{{Date}}";
        sheet.Cell("B4").Style.DateFormat.Format = "yyyy-MM-dd";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("Currency", formatModel.Currency);
        template.AddVariable("Percentage", formatModel.Percentage);
        template.AddVariable("LargeNumber", formatModel.LargeNumber);
        template.AddVariable("Date", formatModel.Date);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        var ws = template.Workbook.Worksheet("FormatTest");
        ws.Cell("B1").GetValue<double>().Should().BeApproximately(1234.56, 0.001);
        ws.Cell("B2").GetValue<double>().Should().BeApproximately(0.75, 0.001);
        ws.Cell("B3").GetValue<int>().Should().Be(1234567);
        ws.Cell("B4").GetDateTime().Should().Be(new DateTime(2025, 5, 6));

        // Check format was applied (by checking the format string)
        ws.Cell("B1").Style.NumberFormat.Format.Should().Be("$#,##0.00");
        ws.Cell("B2").Style.NumberFormat.Format.Should().Be("0.00%");
        ws.Cell("B3").Style.NumberFormat.Format.Should().Be("#,##0");
        ws.Cell("B4").Style.DateFormat.Format.Should().Be("yyyy-MM-dd");
    }

    [Fact]
    public void CalculatedFields_ShouldComputeCorrectly()
    {
        // Arrange
        var product = new ProductModel
        {
            Name = "Test Product",
            UnitPrice = 25.99m,
            Quantity = 3,
            DiscountRate = 0.10m
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("CalcTest");

        // Setup fields with expressions
        sheet.Cell("A1").Value = "Product:";
        sheet.Cell("B1").Value = "{{Name}}";

        sheet.Cell("A2").Value = "Unit Price:";
        sheet.Cell("B2").Value = "{{UnitPrice}}";
        sheet.Cell("B2").Style.NumberFormat.Format = "$#,##0.00";

        sheet.Cell("A3").Value = "Quantity:";
        sheet.Cell("B3").Value = "{{Quantity}}";

        sheet.Cell("A4").Value = "Discount Rate:";
        sheet.Cell("B4").Value = "{{DiscountRate}}";
        sheet.Cell("B4").Style.NumberFormat.Format = "0.00%";

        sheet.Cell("A5").Value = "Subtotal:";
        sheet.Cell("B5").Value = "{{UnitPrice * Quantity}}";
        sheet.Cell("B5").Style.NumberFormat.Format = "$#,##0.00";

        sheet.Cell("A6").Value = "Discount:";
        sheet.Cell("B6").Value = "{{UnitPrice * Quantity * DiscountRate}}";
        sheet.Cell("B6").Style.NumberFormat.Format = "$#,##0.00";

        sheet.Cell("A7").Value = "Total:";
        sheet.Cell("B7").Value = "{{(UnitPrice * Quantity) - (UnitPrice * Quantity * DiscountRate)}}";
        sheet.Cell("B7").Style.NumberFormat.Format = "$#,##0.00";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("Name", product.Name);
        template.AddVariable("UnitPrice", product.UnitPrice);
        template.AddVariable("Quantity", product.Quantity);
        template.AddVariable("DiscountRate", product.DiscountRate);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        var ws = template.Workbook.Worksheet("CalcTest");

        // Check field values
        ws.Cell("B1").GetString().Should().Be("Test Product");
        ws.Cell("B2").GetValue<double>().Should().BeApproximately(25.99, 0.001);
        ws.Cell("B3").GetValue<int>().Should().Be(3);
        ws.Cell("B4").GetValue<double>().Should().BeApproximately(0.10, 0.001);

        // Check calculated values
        double subtotal = (double)(product.UnitPrice * product.Quantity);
        double discount = (double)(product.UnitPrice * product.Quantity * product.DiscountRate);
        double total = subtotal - discount;

        ws.Cell("B5").GetValue<double>().Should().BeApproximately(subtotal, 0.001);
        ws.Cell("B6").GetValue<double>().Should().BeApproximately(discount, 0.001);
        ws.Cell("B7").GetValue<double>().Should().BeApproximately(total, 0.001);
    }

    public class TestModel
    {
        public string StringProperty { get; set; }
        public int IntProperty { get; set; }
        public decimal DecimalProperty { get; set; }
        public DateTime DateProperty { get; set; }
        public bool BoolProperty { get; set; }
    }

    public class ParentModel
    {
        public string Name { get; set; }
        public ChildModel Child { get; set; }
    }

    public class ChildModel
    {
        public string Name { get; set; }
        public int Age { get; set; }
    }

    public class NestedModel
    {
        public ParentModel Parent { get; set; }
    }

    public class FormatModel
    {
        public decimal Currency { get; set; }
        public decimal Percentage { get; set; }
        public int LargeNumber { get; set; }
        public DateTime Date { get; set; }
    }

    public class ProductModel
    {
        public string Name { get; set; }
        public decimal UnitPrice { get; set; }
        public int Quantity { get; set; }
        public decimal DiscountRate { get; set; }
    }
}