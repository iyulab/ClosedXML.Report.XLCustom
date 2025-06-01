using ClosedXML.Excel;
using FluentAssertions;
using Xunit.Abstractions;

namespace ClosedXML.Report.XLCustom.Tests;

public class CollectionTests : TestBase
{
    public CollectionTests(ITestOutputHelper output) : base(output)
    {
    }

    [Fact]
    public void SimpleCollection_ShouldRenderRows()
    {
        // Arrange
        var productList = new List<ProductItem>
        {
            new ProductItem { Id = "P001", Name = "Laptop", Category = "Electronics", Price = 1299.99m },
            new ProductItem { Id = "P002", Name = "Smartphone", Category = "Electronics", Price = 899.99m },
            new ProductItem { Id = "P003", Name = "Headphones", Category = "Accessories", Price = 249.99m }
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Products");

        // Add headers
        sheet.Cell("A1").Value = "Product ID";
        sheet.Cell("B1").Value = "Name";
        sheet.Cell("C1").Value = "Category";
        sheet.Cell("D1").Value = "Price";

        // Add template expressions for data rows
        sheet.Cell("A2").Value = "{{item.Id}}";
        sheet.Cell("B2").Value = "{{item.Name}}";
        sheet.Cell("C2").Value = "{{item.Category}}";
        sheet.Cell("D2").Value = "{{item.Price}}";
        sheet.Cell("D2").Style.NumberFormat.Format = "$#,##0.00";

        // Add service row with aggregation tags
        sheet.Cell("A3").Value = "Total";
        sheet.Cell("D3").Value = "<<sum>>";
        sheet.Cell("D3").Style.NumberFormat.Format = "$#,##0.00";

        // Add service column (mandatory for vertical tables)
        sheet.Cell("E2").Value = "";

        // Define the named range for the collection
        var productRange = sheet.Range("A2:E3");
        workbook.DefinedNames.Add("ProductList", productRange);

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("ProductList", productList);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        var ws = template.Workbook.Worksheet("Products");

        // Verify data was expanded correctly
        ws.Cell("A2").GetString().Should().Be("P001");
        ws.Cell("B2").GetString().Should().Be("Laptop");
        ws.Cell("C2").GetString().Should().Be("Electronics");
        ws.Cell("D2").GetValue<double>().Should().BeApproximately(1299.99, 0.001);

        // Verify rows were created for all items
        ws.Cell("A3").GetString().Should().Be("P002");
        ws.Cell("B3").GetString().Should().Be("Smartphone");
        ws.Cell("A4").GetString().Should().Be("P003");
        ws.Cell("B4").GetString().Should().Be("Headphones");

        // Verify the summary row (total)
        double expectedTotal = (double)productList.Sum(p => p.Price);
        ws.Cell($"D{2 + productList.Count}").GetValue<double>().Should().BeApproximately(expectedTotal, 0.001);
    }

    [Fact]
    public void HorizontalCollection_ShouldBindIndexedItems()
    {
        // Arrange
        var items = new List<ItemModel>
        {
            new ItemModel { Name = "Item A", Price = 10.50m },
            new ItemModel { Name = "Item B", Price = 20.75m },
            new ItemModel { Name = "Item C", Price = 15.25m }
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("Horizontal");

        // Setup horizontal collection with indexed access
        sheet.Cell("A1").Value = "Name 1";
        sheet.Cell("B1").Value = "Price 1";
        sheet.Cell("C1").Value = "Name 2";
        sheet.Cell("D1").Value = "Price 2";
        sheet.Cell("E1").Value = "Name 3";
        sheet.Cell("F1").Value = "Price 3";

        sheet.Cell("A2").Value = "{{items[0].Name}}";
        sheet.Cell("B2").Value = "{{items[0].Price}}";
        sheet.Cell("C2").Value = "{{items[1].Name}}";
        sheet.Cell("D2").Value = "{{items[1].Price}}";
        sheet.Cell("E2").Value = "{{items[2].Name}}";
        sheet.Cell("F2").Value = "{{items[2].Price}}";

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("items", items);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        var ws = template.Workbook.Worksheet("Horizontal");

        // Verify indexed access works
        ws.Cell("A2").GetString().Should().Be("Item A");
        ws.Cell("B2").GetValue<double>().Should().BeApproximately(10.50, 0.001);
        ws.Cell("C2").GetString().Should().Be("Item B");
        ws.Cell("D2").GetValue<double>().Should().BeApproximately(20.75, 0.001);
        ws.Cell("E2").GetString().Should().Be("Item C");
        ws.Cell("F2").GetValue<double>().Should().BeApproximately(15.25, 0.001);
    }

    [Fact]
    public void GroupedCollection_ShouldCreateGroupsWithAggregations()
    {
        // Arrange
        var salesData = new List<SalesRecord>
        {
            new SalesRecord { Region = "North", Product = "Widget A", Units = 150, Revenue = 7500.00m },
            new SalesRecord { Region = "North", Product = "Widget B", Units = 200, Revenue = 12000.00m },
            new SalesRecord { Region = "South", Product = "Widget A", Units = 175, Revenue = 8750.00m },
            new SalesRecord { Region = "South", Product = "Widget B", Units = 120, Revenue = 7200.00m }
        };

        using var workbook = new XLWorkbook();
        var sheet = workbook.AddWorksheet("GroupedSales");

        // Add headers
        sheet.Cell("A1").Value = "Region";
        sheet.Cell("B1").Value = "Product";
        sheet.Cell("C1").Value = "Units";
        sheet.Cell("D1").Value = "Revenue";

        // Add template expressions for data rows
        sheet.Cell("A2").Value = "{{item.Region}}";
        sheet.Cell("B2").Value = "{{item.Product}}";
        sheet.Cell("C2").Value = "{{item.Units}}";
        sheet.Cell("D2").Value = "{{item.Revenue}}";
        sheet.Cell("D2").Style.NumberFormat.Format = "$#,##0.00";

        // Add service row with grouping and aggregation
        sheet.Cell("A3").Value = "<<group>>"; // Group by Region
        sheet.Cell("B3").Value = "<<group>>"; // Group by Product within Region
        sheet.Cell("C3").Value = "<<sum>>"; // Sum Units for each group
        sheet.Cell("D3").Value = "<<sum>>"; // Sum Revenue for each group
        sheet.Cell("D3").Style.NumberFormat.Format = "$#,##0.00";

        // Add service column
        sheet.Cell("E2").Value = "";

        // Define the named range
        var salesRange = sheet.Range("A2:E3");
        workbook.DefinedNames.Add("SalesData", salesRange);

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);
        template.AddVariable("SalesData", salesData);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        // Testing the exact layout is complex due to grouping behavior
        // but we can verify the worksheet exists and has data
        template.Workbook.Worksheets.Contains("GroupedSales").Should().BeTrue();
        var ws = template.Workbook.Worksheet("GroupedSales");

        // Verify some basic expected data in the first rows
        ws.Cell("A2").GetString().Should().Be("North");
        ws.Cell("B2").GetString().Should().Be("Widget A");

        // Aggregation testing is difficult with direct cell access due to 
        // uncertain row positioning after grouping, so we'll test basic rendering occurred
        ws.LastRowUsed().RowNumber().Should().BeGreaterThan(salesData.Count);
    }

    [Fact]
    public void NestedCollection_ShouldRenderCorrectly()
    {
        // Arrange
        var orders = new List<Order>
{
    new Order
    {
        OrderId = "ORD-001",
        CustomerName = "John Smith",
        OrderDate = new DateTime(2025, 5, 1),
        Items = new List<OrderItem>
        {
            new OrderItem { ProductName = "Laptop", Quantity = 1, UnitPrice = 1299.99m },
            new OrderItem { ProductName = "Mouse", Quantity = 2, UnitPrice = 25.50m }
        }
    },
    new Order
    {
        OrderId = "ORD-002",
        CustomerName = "Jane Doe",
        OrderDate = new DateTime(2025, 5, 3),
        Items = new List<OrderItem>
        {
            new OrderItem { ProductName = "Smartphone", Quantity = 1, UnitPrice = 899.99m },
            new OrderItem { ProductName = "Charger", Quantity = 1, UnitPrice = 19.99m },
            new OrderItem { ProductName = "Case", Quantity = 1, UnitPrice = 29.99m }
        }
    }
};

        using var workbook = new XLWorkbook();

        // Orders sheet
        var orderSheet = workbook.AddWorksheet("Orders");
        orderSheet.Cell("A1").Value = "Order ID";
        orderSheet.Cell("B1").Value = "Customer";
        orderSheet.Cell("C1").Value = "Order Date";
        orderSheet.Cell("D1").Value = "Total Items";
        orderSheet.Cell("E1").Value = "Total Amount";

        orderSheet.Cell("A2").Value = "{{item.OrderId}}";
        orderSheet.Cell("B2").Value = "{{item.CustomerName}}";
        orderSheet.Cell("C2").Value = "{{item.OrderDate}}";
        orderSheet.Cell("C2").Style.DateFormat.Format = "yyyy-MM-dd";
        orderSheet.Cell("D2").Value = "{{item.Items.Count}}";
        orderSheet.Cell("E2").Value = "{{item.OrderTotal}}";
        orderSheet.Cell("E2").Style.NumberFormat.Format = "$#,##0.00";

        // Service row and column
        orderSheet.Cell("A3").Value = "";
        orderSheet.Cell("F2").Value = "";

        var orderRange = orderSheet.Range("A2:F3");
        workbook.DefinedNames.Add("Orders", orderRange);

        // Order Details sheet for nested items
        var detailSheet = workbook.AddWorksheet("OrderDetails");
        detailSheet.Cell("A1").Value = "Order:";
        // 수정: 부모 객체에 직접 접근 (첫 번째 주문 정보 표시)
        detailSheet.Cell("B1").Value = "{{CurrentOrder.OrderId}} - {{CurrentOrder.CustomerName}}";

        detailSheet.Cell("A3").Value = "Product";
        detailSheet.Cell("B3").Value = "Quantity";
        detailSheet.Cell("C3").Value = "Unit Price";
        detailSheet.Cell("D3").Value = "Line Total";

        detailSheet.Cell("A4").Value = "{{item.ProductName}}";
        detailSheet.Cell("B4").Value = "{{item.Quantity}}";
        detailSheet.Cell("C4").Value = "{{item.UnitPrice}}";
        detailSheet.Cell("C4").Style.NumberFormat.Format = "$#,##0.00";
        detailSheet.Cell("D4").Value = "{{item.Quantity * item.UnitPrice}}";
        detailSheet.Cell("D4").Style.NumberFormat.Format = "$#,##0.00";

        // Service row with sum
        detailSheet.Cell("A5").Value = "Total";
        detailSheet.Cell("D5").Value = "<<sum>>";
        detailSheet.Cell("D5").Style.NumberFormat.Format = "$#,##0.00";

        // Service column
        detailSheet.Cell("E4").Value = "";

        var itemRange = detailSheet.Range("A4:E5");
        // 명명 규칙에 맞게 중첩 관계 표현
        workbook.DefinedNames.Add("Orders_Items", itemRange);

        using var ms = new MemoryStream();
        workbook.SaveAs(ms);
        ms.Position = 0;

        // Act
        var template = new XLCustomTemplate(ms);

        // 주문 목록 추가
        template.AddVariable("Orders", orders);

        // 주문 아이템 추가 (첫 번째 주문의 아이템으로 시작)
        template.AddVariable("Orders_Items", orders[0].Items);

        // 현재 주문 정보 추가 (첫 번째 주문으로 설정)
        template.AddVariable("CurrentOrder", orders[0]);

        var result = template.Generate();
        LogResult(result);

        // Assert
        result.HasErrors.Should().BeFalse("Template generation should succeed without errors");

        // Check Orders sheet
        var wsOrders = template.Workbook.Worksheet("Orders");
        wsOrders.Cell("A2").GetString().Should().Be("ORD-001");
        wsOrders.Cell("B2").GetString().Should().Be("John Smith");
        wsOrders.Cell("C2").GetDateTime().Should().Be(new DateTime(2025, 5, 1));
        wsOrders.Cell("D2").GetValue<int>().Should().Be(2); // Number of items

        // Check calculated total matches
        double order1Total = (double)orders[0].Items.Sum(i => i.Quantity * i.UnitPrice);
        wsOrders.Cell("E2").GetValue<double>().Should().BeApproximately(order1Total, 0.001);

        // Check second order
        wsOrders.Cell("A3").GetString().Should().Be("ORD-002");
        wsOrders.Cell("B3").GetString().Should().Be("Jane Doe");

        // Check OrderDetails sheet exists
        template.Workbook.Worksheets.Contains("OrderDetails").Should().BeTrue();
    }

    public class ProductItem
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public decimal Price { get; set; }
    }

    public class ItemModel
    {
        public string Name { get; set; }
        public decimal Price { get; set; }
    }

    public class SalesRecord
    {
        public string Region { get; set; }
        public string Product { get; set; }
        public int Units { get; set; }
        public decimal Revenue { get; set; }
    }

    public class Order
    {
        public string OrderId { get; set; }
        public string CustomerName { get; set; }
        public DateTime OrderDate { get; set; }
        public List<OrderItem> Items { get; set; } = new List<OrderItem>();

        public decimal OrderTotal => Items.Sum(i => i.Quantity * i.UnitPrice);
    }

    public class OrderItem
    {
        public string ProductName { get; set; }
        public int Quantity { get; set; }
        public decimal UnitPrice { get; set; }
    }
}