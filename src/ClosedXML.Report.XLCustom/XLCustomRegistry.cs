namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Singleton registry for accessing function handlers and global variables
/// Thread-safe implementation with explicit lifecycle control
/// </summary>
public class XLCustomRegistry
{
    // 고정된 Lazy<T> 인스턴스 대신 필요시 재생성할 수 있는 Lazy<T> 객체 사용
    private static Lazy<XLCustomRegistry> _lazyInstance =
        new Lazy<XLCustomRegistry>(() => new XLCustomRegistry(), LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>
    /// Gets the singleton instance of the registry
    /// </summary>
    public static XLCustomRegistry Instance => _lazyInstance.Value;

    /// <summary>
    /// The global function registry that persists across all template instances
    /// </summary>
    public FunctionRegistry FunctionRegistry { get; private set; }

    /// <summary>
    /// The global variable registry that persists across all template instances
    /// </summary>
    public GlobalVariableRegistry GlobalVariables { get; private set; }

    /// <summary>
    /// Explicitly resets the function registry
    /// This should only be called when the application needs to start fresh
    /// </summary>
    public static void ResetFunctionRegistry()
    {
        Log.Debug("Explicitly resetting the global function registry");
        Instance.FunctionRegistry = new FunctionRegistry();
    }

    /// <summary>
    /// Explicitly resets the global variable registry
    /// This should only be called when the application needs to start fresh
    /// </summary>
    public static void ResetGlobalVariables()
    {
        Log.Debug("Explicitly resetting the global variable registry");
        Instance.GlobalVariables = new GlobalVariableRegistry();

        // 재설정 후 내장 변수 다시 등록
        Instance.GlobalVariables.RegisterBuiltIns();
    }

    /// <summary>
    /// Resets all registries to their initial state
    /// </summary>
    public static void ResetAll()
    {
        Log.Debug("Resetting entire XLCustomRegistry");

        // Lazy<T> 인스턴스 자체를 새로 생성하여 완전히 초기화된 싱글톤 보장
        _lazyInstance = new Lazy<XLCustomRegistry>(() => new XLCustomRegistry(), LazyThreadSafetyMode.ExecutionAndPublication);
    }

    /// <summary>
    /// Creates a new instance with fresh registries
    /// </summary>
    private XLCustomRegistry()
    {
        FunctionRegistry = new FunctionRegistry();
        GlobalVariables = new GlobalVariableRegistry();

        // Register built-in global variables
        GlobalVariables.RegisterBuiltIns();

        Log.Debug("Created XLCustomRegistry singleton with new registries");
    }
}