using ClosedXML.Report.Utils;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Extension methods for ClosedXML objects
/// </summary>
public static class XLExtensions
{
    /// <summary>
    /// Sets a value to a cell with proper type handling
    /// </summary>
    public static void SetValue(this IXLCell cell, object value)
    {
        if (value is DateTime dateValue)
        {
            cell.Value = dateValue;
        }
        else if (value is TimeSpan timeValue)
        {
            cell.Value = timeValue;
        }
        else
        {
            cell.SetValue(XLCellValueConverter.FromObject(value));
        }
    }
}
