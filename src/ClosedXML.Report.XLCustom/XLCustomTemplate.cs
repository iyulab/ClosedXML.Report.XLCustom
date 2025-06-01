using ClosedXML.Report.Options;
using ClosedXML.Report.XLCustom.Tags;
using DocumentFormat.OpenXml.Vml.Office;
using System.Collections;
using System.Reflection;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Extends the XLTemplate class with enhanced expression handling capabilities
/// </summary>
public partial class XLCustomTemplate : IXLTemplate
{
    private XLTemplate _baseTemplate = null!;
    private readonly FunctionRegistry _functionRegistry;
    private readonly GlobalVariableRegistry _globalVariables;
    private readonly IXLWorkbook _workbook;
    private bool _preprocessed = false;
    private bool _templateCreated = false;
    private readonly bool _useGlobalRegistry;
    private readonly XLCustomTemplateOptions _options;

    static XLCustomTemplate()
    {
        try
        {
            // 기능 태그만 등록
            TagsRegister.Add<CustomFunctionTag>("customfunction", 140);
            TagsRegister.Add<FormatTag>("format", 120);

            // GlobalVariableValueTag 제거됨

            Log.Debug("Custom tags registered successfully");
        }
        catch (Exception ex)
        {
            Log.Debug($"Tag registration error: {ex.Message}");
        }
    }

    /// <summary>
    /// Creates a new instance of XLCustomTemplate from file path
    /// </summary>
    public XLCustomTemplate(string filePath, XLCustomTemplateOptions? options = null)
    {
        _options = options ?? XLCustomTemplateOptions.Default;
        _useGlobalRegistry = _options.UseGlobalRegistry;

        // Use global registry or create a local one
        _functionRegistry = _useGlobalRegistry
            ? XLCustomRegistry.Instance.FunctionRegistry
            : new FunctionRegistry();

        _globalVariables = _useGlobalRegistry
            ? XLCustomRegistry.Instance.GlobalVariables
            : new GlobalVariableRegistry();

        if (!_useGlobalRegistry && _options.RegisterBuiltInGlobalVariables)
        {
            Log.Debug("Created XLCustomTemplate with local registries");
            // Register built-in global variables in local registry
            _globalVariables.RegisterBuiltIns();
        }

        // Load workbook
        _workbook = new XLWorkbook(filePath);

        // Register built-ins if configured
        InitializeTemplate();
    }

    /// <summary>
    /// Creates a new instance of XLCustomTemplate from a stream
    /// </summary>
    public XLCustomTemplate(Stream stream, XLCustomTemplateOptions? options = null)
    {
        _options = options ?? XLCustomTemplateOptions.Default;
        _useGlobalRegistry = _options.UseGlobalRegistry;

        // Use global registry or create a local one
        _functionRegistry = _useGlobalRegistry
            ? XLCustomRegistry.Instance.FunctionRegistry
            : new FunctionRegistry();

        _globalVariables = _useGlobalRegistry
            ? XLCustomRegistry.Instance.GlobalVariables
            : new GlobalVariableRegistry();

        if (!_useGlobalRegistry && _options.RegisterBuiltInGlobalVariables)
        {
            Log.Debug("Created XLCustomTemplate with local registries");
            // Register built-in global variables in local registry
            _globalVariables.RegisterBuiltIns();
        }

        // Load workbook
        _workbook = new XLWorkbook(stream);

        // Register built-ins if configured
        InitializeTemplate();
    }

    /// <summary>
    /// Creates a new instance of XLCustomTemplate from an existing workbook
    /// </summary>
    public XLCustomTemplate(IXLWorkbook workbook, XLCustomTemplateOptions? options = null)
    {
        _options = options ?? XLCustomTemplateOptions.Default;
        _useGlobalRegistry = _options.UseGlobalRegistry;

        // Use global registry or create a local one
        _functionRegistry = _useGlobalRegistry
            ? XLCustomRegistry.Instance.FunctionRegistry
            : new FunctionRegistry();

        _globalVariables = _useGlobalRegistry
            ? XLCustomRegistry.Instance.GlobalVariables
            : new GlobalVariableRegistry();

        if (!_useGlobalRegistry && _options.RegisterBuiltInGlobalVariables)
        {
            Log.Debug("Created XLCustomTemplate with local registries");
            // Register built-in global variables in local registry
            _globalVariables.RegisterBuiltIns();
        }

        // Store workbook reference
        _workbook = workbook;

        // Register built-ins if configured
        InitializeTemplate();
    }

    /// <summary>
    /// Initialize the template based on options
    /// </summary>
    private void InitializeTemplate()
    {
        // Register built-in functions if configured
        if (_options.RegisterBuiltInFunctions)
        {
            RegisterBuiltInFunctions();
        }

        // Register built-in global variables if configured
        if (_options.RegisterBuiltInGlobalVariables && _useGlobalRegistry)
        {
            RegisterBuiltInGlobalVariables();
        }

        // Preprocess immediately if configured
        if (_options.PreprocessImmediately)
        {
            EnsurePreprocessed();
        }
    }

    /// <summary>
    /// Gets the underlying workbook
    /// </summary>
    public IXLWorkbook Workbook => _templateCreated ? _baseTemplate.Workbook : _workbook;

    /// <summary>
    /// Force workbook preprocessing
    /// </summary>
    public XLCustomTemplate Preprocess()
    {
        EnsurePreprocessed();
        return this;
    }

    /// <summary>
    /// Generates the template with data and returns the result
    /// </summary>
    public XLGenerateResult Generate()
    {
        try
        {
            // Ensure the template is created from workbook
            EnsureTemplateCreated();

            // Register all global variables to the template
            AddGlobalVariablesToTemplate();

            // Generate the template using base XLTemplate
            var result = _baseTemplate.Generate();

            // Log any errors encountered during generation
            if (result.HasErrors)
            {
                foreach (var error in result.ParsingErrors)
                {
                    Log.Debug($"Generation error: {error.Message}, Range: {error.Range}");
                }
            }

            // Process all image markers in the workbook
            Functions.ImageFunction.ProcessAllImageMarkers(_baseTemplate.Workbook);

            // Apply special formatting for date cells post-generation
            ProcessDateFormats();

            return result;
        }
        catch (Exception ex)
        {
            // Log detailed error information
            Log.Debug($"Error generating template: {ex.Message}");
            Log.Debug($"Stack trace: {ex.StackTrace}");

            // Return an error result
            var errors = new XLCustomTemplateErrors
            {
                new TemplateError(ex.Message, null)
            };
            var errorResult = new XLGenerateResult(errors);
            return errorResult;
        }
    }

    private void ProcessDateFormats()
    {
        foreach (var worksheet in _baseTemplate.Workbook.Worksheets)
        {
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell == null) continue;

                // 날짜 셀에 대한 추가 처리
                if (cell.Value.Type == XLDataType.DateTime)
                {
                    string format = cell.Style.DateFormat.Format;
                    if (format == "d")
                    {
                        // 짧은 날짜 형식의 경우 시간 부분 제거
                        DateTime dateValue = (DateTime)cell.Value;
                        cell.Value = dateValue.Date;
                        Log.Debug($"Processed date cell {cell.Address} to remove time part, format: {format}");
                    }
                }
            }
        }
    }

    private void AddGlobalVariablesToTemplate()
    {
        int addedCount = 0;
        Log.Debug("Starting to add global variables to template");

        foreach (var variableName in _globalVariables.GetNames())
        {
            try
            {
                // 현재 글로벌 변수의 값을 가져옴
                var value = _globalVariables.GetValue(variableName);

                if (value == null)
                {
                    Log.Debug($"Global variable {variableName} has null value, skipping");
                    continue;
                }

                // 템플릿에 직접 추가
                _baseTemplate.AddVariable(variableName, value);
                addedCount++;

                Log.Debug($"Added global variable to template: {variableName} = {value}, Type: {value.GetType().Name}");
            }
            catch (Exception ex)
            {
                // 개별 변수 추가 중 오류가 발생해도 계속 진행
                Log.Debug($"Error adding global variable {variableName}: {ex.Message}");
            }
        }

        Log.Debug($"Finished adding global variables, total added: {addedCount}");
    }

    public void AddVariable(object value)
    {
        EnsureTemplateCreated();

        if (value is IDictionary dictionary)
        {
            foreach (DictionaryEntry entry in dictionary)
            {
                AddVariable(entry.Key.ToString()!, entry.Value);
            }
        }
        else
        {
            var type = value.GetType();
            var fields = type.GetFields(BindingFlags.Public | BindingFlags.Instance).Where(f => f.IsPublic)
                .Select(f => new { f.Name, val = f.GetValue(value), type = f.FieldType })
                .Concat(type.GetProperties(BindingFlags.Public | BindingFlags.Instance).Where(f => f.CanRead)
                    .Select(f => new { f.Name, val = f.GetValue(value, new object[] { }), type = f.PropertyType }));

            foreach (var field in fields)
            {
                AddVariable(field.Name, field.val);
            }
        }
    }

    public void AddVariable(string alias, object? value)
    {
        EnsureTemplateCreated();
        _baseTemplate.AddVariable(alias, value);
    }

    public void Dispose()
    {
        if (_templateCreated)
        {
            _baseTemplate.Dispose();
        }
        else
        {
            _workbook.Dispose();
        }
    }

    private void EnsurePreprocessed()
    {
        if (!_preprocessed)
        {
            // 간소화된 접근 방식을 사용하는 표현식 프로세서 생성
            var expressionProcessor = new XLExpressionProcessor(_functionRegistry, _globalVariables);

            PreprocessWorkbook(_workbook, expressionProcessor);
            _preprocessed = true;
        }
    }

    private void EnsureTemplateCreated()
    {
        if (!_templateCreated)
        {
            EnsurePreprocessed(); // 전처리 확인

            _baseTemplate = new XLTemplate(_workbook);
            _templateCreated = true;

            Log.Debug("Base XLTemplate created from processed workbook");
        }
    }

    private void PreprocessWorkbook(IXLWorkbook workbook, XLExpressionProcessor expressionProcessor)
    {
        Log.Debug("Starting workbook preprocessing...");

        // Process all worksheets in the workbook
        foreach (var worksheet in workbook.Worksheets)
        {
            Log.Debug($"Processing worksheet: {worksheet.Name}");

            // Process all used cells that may contain custom expressions
            foreach (var cell in worksheet.CellsUsed())
            {
                if (cell.HasFormula) continue; // Skip formula cells

                if (cell.DataType == XLDataType.Text)
                {
                    var value = cell.GetString();
                    if (string.IsNullOrEmpty(value)) continue;

                    // Process cell content and replace with compatible tags if needed
                    var newValue = expressionProcessor.ProcessExpression(value, cell);
                    if (newValue != value)
                    {
                        Log.Debug($"Cell {cell.Address}: Replaced '{value}' with '{newValue}'");
                        cell.Value = newValue;
                    }
                }
            }
        }

        Log.Debug("Workbook preprocessing completed");
    }
}