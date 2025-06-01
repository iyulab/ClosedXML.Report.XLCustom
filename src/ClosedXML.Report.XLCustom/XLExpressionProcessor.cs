namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Processes custom expressions and converts them to compatible tags
/// </summary>
public class XLExpressionProcessor
{
    private readonly FunctionRegistry _functionRegistry;
    private readonly GlobalVariableRegistry _globalVariables;

    // Format expression: {{variable:format}} 
    private static readonly Regex FormatExpressionRegex =
        new Regex(@"\{\{\s*([^{}:]+)\s*:\s*([^{}]+)\s*\}\}", RegexOptions.Compiled);

    // Function expression: {{variable|function}} or {{variable|function(parameters)}}
    private static readonly Regex FunctionExpressionRegex =
        new Regex(@"\{\{\s*([^{}|]+)\s*\|\s*([^{}(]+)\s*(?:\(\s*([^{}]*)\s*\))?\s*\}\}", RegexOptions.Compiled);

    // Simple variable expression: {{variable}} - 이제 처리하지 않고 ClosedXML에서 처리하도록 함
    // 하지만 글로벌 변수가 있는지 확인하기 위해 패턴 유지
    private static readonly Regex SimpleVariableRegex =
        new Regex(@"\{\{\s*([^{}|:]+)\s*\}\}", RegexOptions.Compiled);

    public XLExpressionProcessor(FunctionRegistry functionRegistry, GlobalVariableRegistry globalVariables)
    {
        _functionRegistry = functionRegistry ?? throw new ArgumentNullException(nameof(functionRegistry));
        _globalVariables = globalVariables ?? throw new ArgumentNullException(nameof(globalVariables));
    }

    public string ProcessExpression(string value, IXLCell cell)
    {
        if (string.IsNullOrEmpty(value))
            return value;

        Log.Debug($"Processing expression: {value}");

        // 간소화된 접근 방식 - format과 function만 변환
        // 변수 표현식은 그대로 두고 ClosedXML.Report가 처리하도록 함
        return ProcessFormatAndFunctionExpressions(value);
    }

    private string ProcessFormatAndFunctionExpressions(string value)
    {
        bool isModified = false;

        // Format 표현식 처리: {{Value:format}}
        string result = FormatExpressionRegex.Replace(value, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var formatName = match.Groups[2].Value.Trim();
            Log.Debug($"Format expression found: {variableName}:{formatName}");

            // 변수가 실제로 사용 가능한지 확인
            if (!IsValidVariable(variableName))
            {
                Log.Debug($"Variable not found: {variableName}, keeping original expression");
                return match.Value;
            }

            // format 태그 생성
            isModified = true;
            return $"<<format name=\"{variableName}\" format=\"{formatName}\">>";
        });

        if (isModified)
        {
            value = result;
            Log.Debug($"After format processing: {value}");
            isModified = false;
        }

        // Function expression processing: {{Value|function}} or {{Value|function(params)}}
        result = FunctionExpressionRegex.Replace(value, match =>
        {
            var variableName = match.Groups[1].Value.Trim();
            var functionName = match.Groups[2].Value.Trim();
            var paramString = match.Groups.Count > 3 ? match.Groups[3].Value : "";

            Log.Debug($"Function expression found: {variableName}|{functionName}({paramString})");

            // Verify variable validity
            if (!IsValidVariable(variableName))
            {
                Log.Debug($"Variable not found: {variableName}, keeping original expression");
                return match.Value;
            }

            // Verify function registration with detailed logging
            bool functionRegistered = _functionRegistry.IsRegistered(functionName);
            Log.Debug($"Function {functionName} registration check: {functionRegistered}");

            if (!functionRegistered)
            {
                Log.Debug($"Function not registered: {functionName}, keeping original expression");
                Log.Debug($"Available functions: {string.Join(", ", _functionRegistry.GetFunctionNames())}");
                return match.Value;
            }

            // Generate tag parameters
            var tagParams = new StringBuilder();

            // Create function tag
            tagParams.Append($"<<customfunction name=\"{variableName}\" function=\"{functionName}\"");

            if (!string.IsNullOrEmpty(paramString))
            {
                var parameters = ParseParameters(paramString);
                tagParams.Append($" parameters=\"{EscapeParameter(string.Join(",", parameters))}\"");
            }

            tagParams.Append(">>");

            Log.Debug($"Created function tag: {tagParams}");
            isModified = true;

            return tagParams.ToString();
        });

        if (isModified)
        {
            value = result;
            Log.Debug($"After function processing: {value}");
        }

        return value;
    }

    /// <summary>
    /// Checks if a variable is valid for template processing
    /// </summary>
    private bool IsValidVariable(string variableName)
    {
        Log.Debug($"Checking validity of variable: {variableName}");

        // Check if it's a global variable
        if (_globalVariables.IsRegistered(variableName))
        {
            Log.Debug($"Variable {variableName} is a registered global variable");
            return true;
        }

        // Check for dot-separated variable names (item properties or nested objects)
        if (variableName.Contains('.'))
        {
            // Validate item.Property pattern
            string[] parts = variableName.Split('.');
            if (parts.Length > 1 && parts[0].Equals("item", StringComparison.OrdinalIgnoreCase))
            {
                // item.xxx format is considered valid
                Log.Debug($"Variable {variableName} is an item property reference");
                return true;
            }
        }

        // For template variables that we can't verify at this stage,
        // assume they are valid to allow processing
        // The actual validation will happen during template generation
        Log.Debug($"Variable {variableName} assumed to be valid (template variable)");
        return true;
    }

    /// <summary>
    /// 간소화된 변수 표현식({{Variable}})을 찾아서 기록
    /// </summary>
    public HashSet<string> GetSimpleVariableNames(string value)
    {
        var result = new HashSet<string>();

        if (string.IsNullOrEmpty(value))
            return result;

        // 간단한 변수 표현식 검색
        var matches = SimpleVariableRegex.Matches(value);
        foreach (Match match in matches)
        {
            if (match.Groups.Count > 1)
            {
                var variableName = match.Groups[1].Value.Trim();
                result.Add(variableName);
            }
        }

        return result;
    }

    /// <summary>
    /// Parses a comma-separated parameter string, properly handling parameters with commas and parentheses
    /// </summary>
    private string[] ParseParameters(string paramString)
    {
        if (string.IsNullOrEmpty(paramString))
            return Array.Empty<string>();

        var parameters = new List<string>();
        var currentParam = new StringBuilder();
        int parenLevel = 0;
        bool inQuote = false;
        char lastChar = '\0';

        for (int i = 0; i < paramString.Length; i++)
        {
            char c = paramString[i];

            // Handle quoted strings
            if (c == '\'' && lastChar != '\\')
            {
                inQuote = !inQuote;
                currentParam.Append(c);
            }
            // Handle parentheses (only count them if not in a quote)
            else if (c == '(' && !inQuote)
            {
                parenLevel++;
                currentParam.Append(c);
            }
            else if (c == ')' && !inQuote)
            {
                parenLevel--;
                currentParam.Append(c);
            }
            // Handle parameter separator (only if not in quotes or parentheses)
            else if (c == ',' && parenLevel == 0 && !inQuote)
            {
                // Add parameter when a separator is encountered
                parameters.Add(currentParam.ToString().Trim());
                currentParam.Clear();
            }
            else
            {
                currentParam.Append(c);
            }

            lastChar = c;
        }

        // Add final parameter
        if (currentParam.Length > 0)
        {
            parameters.Add(currentParam.ToString().Trim());
        }

        return parameters.ToArray();
    }

    /// <summary>
    /// Escapes a parameter value for safe inclusion in tag parameters
    /// </summary>
    private string EscapeParameter(string param)
    {
        // Wrap parameter in single quotes if it contains commas or parentheses
        if (param.Contains(',') || param.Contains('(') || param.Contains(')'))
        {
            // Escape existing single quotes
            param = param.Replace("'", "''");
            return $"'{param}'";
        }

        return param;
    }
}