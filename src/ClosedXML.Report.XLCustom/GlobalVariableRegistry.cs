using System.Collections.Concurrent;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Registry for global variables that can be used in expressions
/// Thread-safe implementation for concurrent access
/// </summary>
public class GlobalVariableRegistry
{
    private readonly ConcurrentDictionary<string, Func<object>> _variables =
        new ConcurrentDictionary<string, Func<object>>(StringComparer.OrdinalIgnoreCase);

    // 내장 변수 등록 여부를 추적하기 위한 플래그
    private bool _builtInsRegistered = false;

    /// <summary>
    /// Registers a global variable with a value provider function
    /// The function is called each time the variable is used, allowing for dynamic values
    /// </summary>
    public void Register(string name, Func<object> valueProvider)
    {
        if (string.IsNullOrEmpty(name))
            throw new ArgumentNullException(nameof(name));

        if (valueProvider == null)
            throw new ArgumentNullException(nameof(valueProvider));

        _variables[name] = valueProvider;
        Log.Debug($"Registered global variable: {name}");
    }

    /// <summary>
    /// Registers a global variable with a static value
    /// </summary>
    public void Register(string name, object value)
    {
        Register(name, () => value);
    }

    /// <summary>
    /// Checks if a global variable is registered
    /// </summary>
    public bool IsRegistered(string name)
    {
        return _variables.ContainsKey(name);
    }

    /// <summary>
    /// Gets the value of a global variable
    /// </summary>
    public object? GetValue(string name)
    {
        if (_variables.TryGetValue(name, out var valueProvider))
        {
            try
            {
                return valueProvider();
            }
            catch (Exception ex)
            {
                Log.Debug($"Error getting value for global variable '{name}': {ex.Message}");
                return null;
            }
        }

        return null;
    }

    /// <summary>
    /// Removes a global variable
    /// </summary>
    public bool Remove(string name)
    {
        return _variables.TryRemove(name, out _);
    }

    /// <summary>
    /// Clears all global variables
    /// </summary>
    public void Clear()
    {
        _variables.Clear();
        _builtInsRegistered = false;
        Log.Debug("Cleared all global variables");
    }

    /// <summary>
    /// Gets all registered global variable names
    /// </summary>
    public IEnumerable<string> GetNames()
    {
        return _variables.Keys.ToList();
    }

    /// <summary>
    /// Registers built-in global variables
    /// </summary>
    public void RegisterBuiltIns()
    {
        // 이미 등록된 경우 중복 등록 방지 (옵션)
        if (_builtInsRegistered)
        {
            Log.Debug("Built-in global variables already registered");
            return;
        }

        // Date and time variables
        Register("Today", () => DateTime.Today);
        Register("Now", () => DateTime.Now);
        Register("UtcNow", () => DateTime.UtcNow);

        // Current year, month, day
        Register("Year", () => DateTime.Today.Year);
        Register("Month", () => DateTime.Today.Month);
        Register("Day", () => DateTime.Today.Day);

        // System information
        Register("MachineName", Environment.MachineName);
        Register("UserName", Environment.UserName);
        Register("OSVersion", Environment.OSVersion.ToString());

        // Application information
        Register("AppDomain", AppDomain.CurrentDomain.FriendlyName);

        _builtInsRegistered = true;
        Log.Debug("Registered built-in global variables");
    }
}