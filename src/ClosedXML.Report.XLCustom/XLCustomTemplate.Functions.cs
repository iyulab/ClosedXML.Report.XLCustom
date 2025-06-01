namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Partial class containing function registration functionality
/// </summary>
public partial class XLCustomTemplate
{
    /// <summary>
    /// Registers a custom function processor that can manipulate cells
    /// </summary>
    public XLCustomTemplate RegisterFunction(string functionName, XLFunctionHandler function)
    {
        _functionRegistry.Register(functionName, function);
        _preprocessed = false; // Need to reprocess after registering functions

        // 등록 확인 로깅
        Log.Debug($"Function '{functionName}' registered in {(_useGlobalRegistry ? "global" : "local")} registry");
        return this;
    }

    /// <summary>
    /// Registers all built-in functions and global variables
    /// </summary>
    public XLCustomTemplate RegisterBuiltIns()
    {
        RegisterBuiltInFunctions();
        RegisterBuiltInGlobalVariables();
        return this;
    }

    /// <summary>
    /// Registers built-in functions
    /// </summary>
    public XLCustomTemplate RegisterBuiltInFunctions()
    {
        Log.Debug("Before registering built-in functions");

        // Register built-in functions to appropriate registry
        if (_useGlobalRegistry)
        {
            // Register to global registry
            XLCustomRegistry.Instance.FunctionRegistry.RegisterBuiltInFunctions();
        }
        else
        {
            // Register to local registry
            (_functionRegistry as FunctionRegistry)?.RegisterBuiltInFunctions();
        }

        // Reset preprocessing flag to ensure reprocessing
        _preprocessed = false;

        // Log registered functions for debugging
        Log.Debug("After registering built-in functions");
        foreach (var name in _functionRegistry.GetFunctionNames())
        {
            Log.Debug($"Registered function: {name}");
        }

        // Verify specific functions are registered
        Log.Debug($"bold function registered: {_functionRegistry.IsRegistered("bold")}");
        Log.Debug($"italic function registered: {_functionRegistry.IsRegistered("italic")}");
        Log.Debug($"color function registered: {_functionRegistry.IsRegistered("color")}");
        Log.Debug($"link function registered: {_functionRegistry.IsRegistered("link")}");
        Log.Debug($"image function registered: {_functionRegistry.IsRegistered("image")}");

        return this;
    }

    /// <summary>
    /// Debug method: Process an expression and return the result
    /// </summary>
    public string DebugExpression(string expression, IXLCell? cell = null)
    {
        if (cell == null)
        {
            // Create temporary workbook and cell for processing
            using var tempWorkbook = new XLWorkbook();
            var tempWorksheet = tempWorkbook.AddWorksheet("Temp");
            cell = tempWorksheet.Cell(1, 1);
        }

        // Ensure preprocessing is completed before processing expressions
        EnsurePreprocessed();

        var expressionProcessor = new XLExpressionProcessor(_functionRegistry, _globalVariables);
        return expressionProcessor.ProcessExpression(expression, cell);
    }

    /// <summary>
    /// Gets the function registry for testing purposes
    /// </summary>
    public FunctionRegistry GetFunctionRegistryForTest()
    {
        return _functionRegistry as FunctionRegistry;
    }
}