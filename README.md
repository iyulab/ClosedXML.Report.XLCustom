# ClosedXML.Report.XLCustom

> **Based on [ClosedXML.Report](https://github.com/ClosedXML/ClosedXML.Report)**  
> This extension maintains full compatibility with the original library while adding enhanced expression handling capabilities.

## What's New

ClosedXML.Report.XLCustom extends the original library with:

- **Format expressions**: `{{Value:F2}}`, `{{Date:yyyy-MM-dd}}`
- **Function expressions**: `{{Value|bold}}`, `{{ImageUrl|image(150)}}`
- **Custom function registration**
- **Built-in functions** for styling and formatting

## Installation

```bash
Install-Package ClosedXML.Report.XLCustom
```

## Quick Start

```csharp
var template = new XLCustomTemplate("template.xlsx");
template.RegisterBuiltIns(); // Add built-in functions

// Standard ClosedXML.Report usage
template.AddVariable("Title", "Sales Report");
template.AddVariable("Products", productList);

// Custom function registration
template.RegisterFunction("highlight", (cell, value, parameters) => {
    cell.SetValue(value);
    cell.Style.Fill.BackgroundColor = XLColor.Yellow;
});

template.Generate();
template.SaveAs("result.xlsx");
```

## Expression Syntax

### Standard Variables (Original ClosedXML.Report)
```
{{VariableName}}
{{Object.Property}}
```

### Format Expressions (New)
```
{{Price:C}}           // Currency format
{{Date:yyyy-MM-dd}}   // Date format
{{Value:F2}}          // 2 decimal places
```

### Function Expressions (New)
```
{{Text|bold}}                    // Make bold
{{Text|color(Red)}}             // Set color
{{ImagePath|image(100,100)}}    // Insert image
{{Url|link(Click here)}}        // Create hyperlink
```

## Built-in Functions

Register with `template.RegisterBuiltIns()`:

| Function | Usage | Description |
|----------|-------|-------------|
| `bold` | `{{Text|bold}}` | Bold text |
| `italic` | `{{Text|italic}}` | Italic text |
| `color` | `{{Text|color(Red)}}` | Text color |
| `link` | `{{Url|link(Display)}}` | Hyperlink |
| `image` | `{{Path|image(150)}}` | Insert image |

## Built-in Global Variables

Automatically available:

- `{{Today:d}}` - Current date
- `{{Year}}` - Current year  
- `{{MachineName}}` - Computer name
- `{{UserName}}` - User name