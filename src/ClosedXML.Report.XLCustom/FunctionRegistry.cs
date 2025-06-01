using ClosedXML.Report.XLCustom.Functions;
using System.Collections.Concurrent;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Registry for custom function handlers with thread-safe operation
/// </summary>
public class FunctionRegistry
{
    private readonly ConcurrentDictionary<string, XLFunctionHandler> _functions =
        new ConcurrentDictionary<string, XLFunctionHandler>(StringComparer.OrdinalIgnoreCase);

    // 내장 함수 등록 여부를 추적하는 플래그
    private bool _builtInFunctionsRegistered = false;

    /// <summary>
    /// Registers a custom function
    /// Later registrations with the same name will override earlier ones
    /// </summary>
    public void Register(string functionName, XLFunctionHandler function)
    {
        if (function == null)
            throw new ArgumentNullException(nameof(function));

        _functions[functionName] = function;
        Log.Debug($"Registered function: {functionName}");
    }

    /// <summary>
    /// Checks if a function is registered
    /// </summary>
    public bool IsRegistered(string functionName)
    {
        return _functions.ContainsKey(functionName);
    }

    /// <summary>
    /// Gets a registered function by name
    /// </summary>
    public XLFunctionHandler? GetFunction(string functionName)
    {
        if (_functions.TryGetValue(functionName, out var function))
            return function;

        return null;
    }

    /// <summary>
    /// Removes a registered function
    /// </summary>
    public bool RemoveFunction(string functionName)
    {
        return _functions.TryRemove(functionName, out _);
    }

    /// <summary>
    /// Clears all registered functions
    /// </summary>
    public void Clear()
    {
        _functions.Clear();
        _builtInFunctionsRegistered = false;
        Log.Debug("Cleared all registered functions");
    }

    /// <summary>
    /// Gets all registered function names
    /// </summary>
    public IEnumerable<string> GetFunctionNames()
    {
        return _functions.Keys.ToList();
    }

    /// <summary>
    /// Registers built-in functions
    /// </summary>
    public void RegisterBuiltInFunctions()
    {
        // 이미 등록된 경우 중복 등록 방지
        if (_builtInFunctionsRegistered)
        {
            Log.Debug("Built-in functions already registered");
            return;
        }

        // 기본 함수 등록
        Register("bold", (cell, value, _) => {
            cell.SetValue(value);
            cell.Style.Font.Bold = true;
        });

        Register("italic", (cell, value, _) => {
            cell.SetValue(value);
            cell.Style.Font.Italic = true;
        });

        Register("color", (cell, value, parameters) => {
            cell.SetValue(value);
            var colorName = parameters.Length > 0 ? parameters[0] : "Black";
            try
            {
                var color = XLColor.FromName(colorName);
                cell.Style.Font.FontColor = color;
            }
            catch (Exception ex)
            {
                Log.Debug($"Error setting color: {ex.Message}");
                cell.Style.Font.FontColor = XLColor.Black;
            }
        });

        Register("link", (cell, value, parameters) => {
            var url = value?.ToString();
            if (string.IsNullOrEmpty(url)) return;

            string text = parameters.Length > 0 ? parameters[0] : url;
            cell.Value = text;
            cell.SetHyperlink(new XLHyperlink(url));
        });

        // 이미지 함수들을 등록
        Functions.ImageFunction.RegisterImageFunctions(this);

        _builtInFunctionsRegistered = true;
        Log.Debug("Registered built-in functions");
    }
}