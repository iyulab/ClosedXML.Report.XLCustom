namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Configuration options for XLCustomTemplate
/// </summary>
public class XLCustomTemplateOptions
{
    /// <summary>
    /// Controls whether to use the global registry for functions and variables
    /// </summary>
    public bool UseGlobalRegistry { get; set; } = true;

    /// <summary>
    /// Controls whether to automatically register built-in functions
    /// </summary>
    public bool RegisterBuiltInFunctions { get; set; } = true;

    /// <summary>
    /// Controls whether to automatically register built-in global variables
    /// </summary>
    public bool RegisterBuiltInGlobalVariables { get; set; } = true;

    /// <summary>
    /// Controls whether to preprocess the template immediately after creation
    /// </summary>
    public bool PreprocessImmediately { get; set; } = false;

    /// <summary>
    /// Creates a default options instance
    /// </summary>
    public static XLCustomTemplateOptions Default => new XLCustomTemplateOptions();
}