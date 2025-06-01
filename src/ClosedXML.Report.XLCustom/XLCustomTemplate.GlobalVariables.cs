namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Partial class containing global variables related functionality
/// </summary>
public partial class XLCustomTemplate
{
    /// <summary>
    /// Registers a global variable with a dynamic value provider
    /// </summary>
    public XLCustomTemplate RegisterGlobalVariable(string name, Func<object> valueProvider)
    {
        _globalVariables.Register(name, valueProvider);
        _preprocessed = false; // Need to reprocess after registering variables
        return this;
    }

    /// <summary>
    /// Registers a global variable with a static value
    /// </summary>
    public XLCustomTemplate RegisterGlobalVariable(string name, object value)
    {
        _globalVariables.Register(name, () => value);
        _preprocessed = false; // Need to reprocess after registering variables
        return this;
    }

    /// <summary>
    /// Registers built-in global variables
    /// </summary>
    public XLCustomTemplate RegisterBuiltInGlobalVariables()
    {
        // 내장 변수 등록
        if (!_useGlobalRegistry)
        {
            // For local registry, register built-ins
            _globalVariables.RegisterBuiltIns();
        }
        else
        {
            // 전역 레지스트리 사용 시에도 내장 변수 등록이 필요할 수 있음
            // 이미 등록되어 있을 수 있지만, 명시적으로 호출된 경우 재등록
            XLCustomRegistry.Instance.GlobalVariables.RegisterBuiltIns();
        }

        // 등록된 변수 로깅
        Log.Debug("Built-in global variables registered.");
        foreach (var name in _globalVariables.GetNames())
        {
            Log.Debug($"Registered global variable: {name}");
        }

        _preprocessed = false; // Need to reprocess
        return this;
    }

    /// <summary>
    /// Gets the global variable registry for testing purposes
    /// </summary>
    public GlobalVariableRegistry GetGlobalVariablesForTest()
    {
        return _globalVariables;
    }
}