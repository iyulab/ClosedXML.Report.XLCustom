using ClosedXML.Excel;
using ClosedXML.Report.XLCustom;
using DocumentFormat.OpenXml.ExtendedProperties;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace TestConsoleApp;

internal class Program
{
    static void Main(string[] args)
    {
        Console.WriteLine("Starting XLCustomTemplate Test Console App");
        Console.WriteLine("=========================================\n");

        // Create demo data
        var company = CreateDemoCompany();
        var products = CreateDemoProducts();

        // Create a template workbook
        using var workbook = new XLWorkbook();

        // Add three worksheets for different examples
        CreateBasicPropertiesSheet(workbook);
        CreateProductListSheet(workbook);
        CreateOrderDetailsSheet(workbook);

        // Save the template
        string templatePath = Path.GetFullPath("template.xlsx");
        workbook.SaveAs(templatePath);
        Console.WriteLine($"Template saved to: {templatePath}");

        // Open the template file
        OpenFile(templatePath);
        Console.WriteLine("Template file opened in default application.");

        // Generate the final report
        try
        {
            Console.WriteLine("\nGenerating report from template...");
            using var template = new XLCustomTemplate(templatePath);

            // Register built-in functions
            template.RegisterBuiltInFunctions();

            // Register a custom function
            template.RegisterFunction("productStatus", (cell, value, parameters) =>
            {
                if (value is bool inStock)
                {
                    string status = inStock ? "In Stock" : "Out of Stock";
                    cell.SetValue(status);
                    cell.Style.Font.FontColor = inStock ? XLColor.Green : XLColor.Red;
                    cell.Style.Font.Bold = true;
                }
            });

            // Add variables
            template.AddVariable(company); // Add all properties of company object
            template.AddVariable("Products", products); // Add collection

            // Generate the report
            var result = template.Generate();
            if (!result.HasErrors)
            {
                string outputPath = Path.GetFullPath("output.xlsx");
                template.Workbook.SaveAs(outputPath);
                Console.WriteLine($"Report generated successfully: {outputPath}\n");

                // Open the generated file
                OpenFile(outputPath);
                Console.WriteLine("Generated file opened in default application.");

                // Display any messages
                foreach (var error in result.ParsingErrors)
                {
                    Console.WriteLine($"Message: {error.Message}, {error.Range}");
                }
            }
            else
            {
                Console.WriteLine("Errors occurred during report generation:");
                foreach (var error in result.ParsingErrors)
                {
                    Console.WriteLine($"Message: {error.Message}, {error.Range}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Exception occurred: {ex.Message}");
            Console.WriteLine(ex.StackTrace);
        }

        Console.WriteLine("\nPress any key to exit...");
        Console.ReadKey();
    }

    private static Company CreateDemoCompany()
    {
        return new Company
        {
            Name = "Awesome Tech Solutions",
            Founded = new DateTime(2010, 5, 8),
            Revenue = 12456789.50m,
            IsPublic = true,
            EmployeeCount = 256,
            WebsiteUrl = "https://www.awesometech.example.com",
            Address = new Address
            {
                Street = "123 Innovation Drive",
                City = "Tech City",
                State = "CA",
                ZipCode = "94043",
                Country = "USA"
            },
            CEO = new Person
            {
                Name = "Jane Smith",
                Title = "Chief Executive Officer",
                Email = "jane.smith@awesometech.example.com",
                Phone = "555-123-4567",
                ProfileImageUrl = "https://placehold.co/60x60?text=CEO"
            }
        };
    }

    private static List<Product> CreateDemoProducts()
    {
        return new List<Product>
        {
            new Product
            {
                Id = "P001",
                Name = "Premium Laptop",
                Category = "Electronics",
                Price = 1299.99m,
                InStock = true,
                Description = "High-performance laptop with advanced features",
                ImageUrl = "https://placehold.co/60x60?text=Laptop",
                Rating = 4.8,
                ReleaseDate = new DateTime(2023, 9, 15)
            },
            new Product
            {
                Id = "P002",
                Name = "Wireless Headphones",
                Category = "Audio",
                Price = 199.99m,
                InStock = true,
                Description = "Premium noise-cancelling headphones",
                ImageUrl = "https://placehold.co/60x60?text=Headphones",
                Rating = 4.7,
                ReleaseDate = new DateTime(2024, 2, 10)
            },
            new Product
            {
                Id = "P003",
                Name = "Professional Camera",
                Category = "Photography",
                Price = 899.99m,
                InStock = false,
                Description = "High-resolution digital camera for professionals",
                ImageUrl = "https://placehold.co/60x60?text=Camera",
                Rating = 4.9,
                ReleaseDate = new DateTime(2024, 3, 22)
            },
            new Product
            {
                Id = "P004",
                Name = "Smart Watch",
                Category = "Wearables",
                Price = 349.99m,
                InStock = true,
                Description = "Feature-packed smart watch with health monitoring",
                ImageUrl = "https://placehold.co/60x60?text=SmartWatch",
                Rating = 4.5,
                ReleaseDate = new DateTime(2024, 1, 5)
            },
            new Product
            {
                Id = "P005",
                Name = "Portable Speaker",
                Category = "Audio",
                Price = 129.99m,
                InStock = true,
                Description = "Waterproof portable Bluetooth speaker",
                ImageUrl = "https://placehold.co/60x60?text=Speaker",
                Rating = 4.6,
                ReleaseDate = new DateTime(2023, 11, 10)
            }
        };
    }

    private static void CreateBasicPropertiesSheet(XLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Company Info");

        // Title and styling
        sheet.Cell("A1").Value = "Company Profile";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Cell("A1").Style.Font.FontSize = 14;
        sheet.Range("A1:E1").Merge();

        // Basic property bindings
        sheet.Cell("A3").Value = "Name:";
        sheet.Cell("B3").Value = "{{Name}}";

        sheet.Cell("A4").Value = "Founded:";
        sheet.Cell("B4").Value = "{{Founded}}";
        sheet.Cell("B4").Style.DateFormat.Format = "yyyy-MM-dd";

        sheet.Cell("A5").Value = "Revenue:";
        sheet.Cell("B5").Value = "{{Revenue:C}}";

        sheet.Cell("A6").Value = "Employee Count:";
        sheet.Cell("B6").Value = "{{EmployeeCount}}";

        sheet.Cell("A7").Value = "Public Company:";
        sheet.Cell("B7").Value = "{{IsPublic}}";

        sheet.Cell("A8").Value = "Website:";
        sheet.Cell("B8").Value = "{{WebsiteUrl|link(Visit Website)}}";

        // Nested properties
        sheet.Cell("A10").Value = "Address Information";
        sheet.Cell("A10").Style.Font.Bold = true;
        sheet.Range("A10:C10").Merge();

        sheet.Cell("A11").Value = "Street:";
        sheet.Cell("B11").Value = "{{Address.Street}}";

        sheet.Cell("A12").Value = "City:";
        sheet.Cell("B12").Value = "{{Address.City}}";

        sheet.Cell("A13").Value = "State:";
        sheet.Cell("B13").Value = "{{Address.State}}";

        sheet.Cell("A14").Value = "Zip Code:";
        sheet.Cell("B14").Value = "{{Address.ZipCode}}";

        sheet.Cell("A15").Value = "Country:";
        sheet.Cell("B15").Value = "{{Address.Country}}";

        // CEO information
        sheet.Cell("D3").Value = "CEO Information";
        sheet.Cell("D3").Style.Font.Bold = true;
        sheet.Range("D3:F3").Merge();

        sheet.Cell("D4").Value = "Name:";
        sheet.Cell("E4").Value = "{{CEO.Name|bold}}";

        sheet.Cell("D5").Value = "Title:";
        sheet.Cell("E5").Value = "{{CEO.Title}}";

        sheet.Cell("D6").Value = "Email:";
        sheet.Cell("E6").Value = "{{CEO.Email|color(Blue)}}";

        sheet.Cell("D7").Value = "Phone:";
        sheet.Cell("E7").Value = "{{CEO.Phone}}";

        sheet.Cell("D8").Value = "Profile Image:";
        sheet.Cell("E8").Value = "{{CEO.ProfileImageUrl|image(60)}}";

        sheet.Row(8).Height = 100;

        // Apply styling to the whole worksheet
        sheet.Columns().AdjustToContents();
    }

    private static void CreateProductListSheet(XLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Products");

        // Title and styling
        sheet.Cell("A1").Value = "Product Catalog";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Cell("A1").Style.Font.FontSize = 14;
        sheet.Range("A1:H1").Merge();

        // Table headers with styling
        sheet.Cell("A2").Value = "ID";
        sheet.Cell("B2").Value = "Image";
        sheet.Cell("C2").Value = "Name";
        sheet.Cell("D2").Value = "Category";
        sheet.Cell("E2").Value = "Price";
        sheet.Cell("F2").Value = "Status";
        sheet.Cell("G2").Value = "Rating";
        sheet.Cell("H2").Value = "Released";

        // Style the header row
        var headerRange = sheet.Range("A2:H2");
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
        headerRange.Style.Font.Bold = true;
        headerRange.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        // Template row
        sheet.Cell("A3").Value = "{{item.Id}}";
        sheet.Cell("B3").Value = "{{item.ImageUrl|image(30)}}";
        sheet.Cell("C3").Value = "{{item.Name|bold}}";
        sheet.Cell("D3").Value = "{{item.Category}}";
        sheet.Cell("E3").Value = "{{item.Price:C}}";
        sheet.Cell("F3").Value = "{{item.InStock|productStatus}}";
        sheet.Cell("G3").Value = "{{item.Rating:F1}}";
        sheet.Cell("H3").Value = "{{item.ReleaseDate}}";
        sheet.Cell("H3").Style.DateFormat.Format = "yyyy-MM-dd";

        // Set row height for the image
        sheet.Row(3).Height = 100;

        // Add alternating row styling and formula demonstration
        sheet.Cell("A4").Value = ""; // Service row for aggregations

        // Add service column
        sheet.Cell("I3").Value = "";

        // Define named range using DefinedNames
        var productsRange = sheet.Range("A3:I4");
        workbook.DefinedNames.Add("Products", productsRange);

        // Add totals section
        sheet.Cell("D6").Value = "Total Products:";
        sheet.Cell("E6").Value = "{{Products.Count}}";
        sheet.Cell("D6").Style.Font.Bold = true;

        sheet.Cell("D7").Value = "Average Price:";
        sheet.Cell("E7").Value = "<<average>>"; // This will be replaced with average of the Price column
        sheet.Cell("E7").Style.NumberFormat.Format = "$#,##0.00";
        sheet.Cell("D7").Style.Font.Bold = true;

        // Add column for calculated total
        sheet.Cell("J2").Value = "Total Value";
        sheet.Cell("J2").Style.Fill.BackgroundColor = XLColor.LightGray;
        sheet.Cell("J2").Style.Font.Bold = true;
        sheet.Cell("J2").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;

        sheet.Cell("J3").Value = "{{item.Price * (item.InStock ? 1 : 0):C}}"; // Only count value if in stock
        sheet.Cell("J4").Value = "<<sum>>";
        sheet.Cell("J4").Style.NumberFormat.Format = "$#,##0.00";

        // Adjust columns
        sheet.Columns().AdjustToContents();
    }

    private static void CreateOrderDetailsSheet(XLWorkbook workbook)
    {
        var sheet = workbook.AddWorksheet("Order Example");

        // Title
        sheet.Cell("A1").Value = "Order Example with Calculated Fields";
        sheet.Cell("A1").Style.Font.Bold = true;
        sheet.Cell("A1").Style.Font.FontSize = 14;
        sheet.Range("A1:F1").Merge();

        // Order details section
        sheet.Cell("A3").Value = "Order ID:";
        sheet.Cell("B3").Value = "ORD-2025-001";

        sheet.Cell("D3").Value = "Customer:";
        sheet.Cell("E3").Value = "{{Name}}"; // Using company name as customer

        sheet.Cell("A4").Value = "Order Date:";
        sheet.Cell("B4").Value = "{{Today}}";
        sheet.Cell("B4").Style.DateFormat.Format = "yyyy-MM-dd";

        // Products table
        sheet.Cell("A6").Value = "Product";
        sheet.Cell("B6").Value = "Unit Price";
        sheet.Cell("C6").Value = "Quantity";
        sheet.Cell("D6").Value = "Discount";
        sheet.Cell("E6").Value = "Total";

        // Style headers
        var headerRange = sheet.Range("A6:E6");
        headerRange.Style.Fill.BackgroundColor = XLColor.LightGray;
        headerRange.Style.Font.Bold = true;

        // First product row - We'll use the first product from our collection
        sheet.Cell("A7").Value = "{{Products[0].Name}}";
        sheet.Cell("B7").Value = "{{Products[0].Price}}";
        sheet.Cell("B7").Style.NumberFormat.Format = "$#,##0.00";
        sheet.Cell("C7").Value = "2"; // Hardcoded quantity for demonstration
        sheet.Cell("D7").Value = "10%";
        sheet.Cell("E7").Value = "{{Products[0].Price * 2 * 0.9}}"; // Price * Quantity * (1 - Discount)
        sheet.Cell("E7").Style.NumberFormat.Format = "$#,##0.00";

        // Second product row
        sheet.Cell("A8").Value = "{{Products[1].Name}}";
        sheet.Cell("B8").Value = "{{Products[1].Price}}";
        sheet.Cell("B8").Style.NumberFormat.Format = "$#,##0.00";
        sheet.Cell("C8").Value = "1"; // Hardcoded quantity
        sheet.Cell("D8").Value = "0%";
        sheet.Cell("E8").Value = "{{Products[1].Price * 1 * 1.0}}"; // No discount
        sheet.Cell("E8").Style.NumberFormat.Format = "$#,##0.00";

        // Third product row
        sheet.Cell("A9").Value = "{{Products[4].Name}}";
        sheet.Cell("B9").Value = "{{Products[4].Price}}";
        sheet.Cell("B9").Style.NumberFormat.Format = "$#,##0.00";
        sheet.Cell("C9").Value = "3"; // Hardcoded quantity
        sheet.Cell("D9").Value = "5%";
        sheet.Cell("E9").Value = "{{Products[4].Price * 3 * 0.95}}"; // 5% discount
        sheet.Cell("E9").Style.NumberFormat.Format = "$#,##0.00";

        // Total section
        sheet.Cell("D11").Value = "Subtotal:";
        sheet.Cell("D11").Style.Font.Bold = true;
        sheet.Cell("E11").Value = "{{Products[0].Price * 2 * 0.9 + Products[1].Price * 1 + Products[4].Price * 3 * 0.95}}";
        sheet.Cell("E11").Style.NumberFormat.Format = "$#,##0.00";

        sheet.Cell("D12").Value = "Tax (8%):";
        sheet.Cell("D12").Style.Font.Bold = true;
        sheet.Cell("E12").Value = "{{(Products[0].Price * 2 * 0.9 + Products[1].Price * 1 + Products[4].Price * 3 * 0.95) * 0.08}}";
        sheet.Cell("E12").Style.NumberFormat.Format = "$#,##0.00";

        sheet.Cell("D13").Value = "Total:";
        sheet.Cell("D13").Style.Font.Bold = true;
        sheet.Cell("E13").Value = "{{(Products[0].Price * 2 * 0.9 + Products[1].Price * 1 + Products[4].Price * 3 * 0.95) * 1.08}}";
        sheet.Cell("E13").Style.NumberFormat.Format = "$#,##0.00";
        sheet.Cell("E13").Style.Font.Bold = true;

        // Adjust columns
        sheet.Columns().AdjustToContents();
    }
    /// <summary>
    /// Opens a file with the default associated application
    /// </summary>
    /// <param name="path">Full path to the file</param>
    private static void OpenFile(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                // Check the operating system
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    // Windows
                    Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    // macOS
                    Process.Start("open", path);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    // Linux
                    Process.Start("xdg-open", path);
                }
                else
                {
                    Console.WriteLine($"Cannot open file: Unsupported operating system");
                }
            }
            else
            {
                Console.WriteLine($"Cannot open file: File does not exist at {path}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Error opening file: {ex.Message}");
        }
    }

    public class Company
    {
        public string Name { get; set; }
        public DateTime Founded { get; set; }
        public decimal Revenue { get; set; }
        public int EmployeeCount { get; set; }
        public bool IsPublic { get; set; }
        public string WebsiteUrl { get; set; }
        public Address Address { get; set; }
        public Person CEO { get; set; }
    }

    public class Address
    {
        public string Street { get; set; }
        public string City { get; set; }
        public string State { get; set; }
        public string ZipCode { get; set; }
        public string Country { get; set; }
    }

    public class Person
    {
        public string Name { get; set; }
        public string Title { get; set; }
        public string Email { get; set; }
        public string Phone { get; set; }
        public string ProfileImageUrl { get; set; }
    }

    public class Product
    {
        public string Id { get; set; }
        public string Name { get; set; }
        public string Category { get; set; }
        public decimal Price { get; set; }
        public bool InStock { get; set; }
        public string Description { get; set; }
        public string ImageUrl { get; set; }
        public double Rating { get; set; }
        public DateTime ReleaseDate { get; set; }
    }
}