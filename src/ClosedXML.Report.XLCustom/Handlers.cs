namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Delegate for custom function handlers
/// </summary>
public delegate void XLFunctionHandler(IXLCell cell, object value, string[] parameters);