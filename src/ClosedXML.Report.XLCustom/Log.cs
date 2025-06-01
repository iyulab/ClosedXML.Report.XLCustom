using System.Diagnostics;

namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Simple logging utility
/// </summary>
internal class Log
{
    [Conditional("DEBUG")]
    public static void Debug(string message)
    {
        System.Diagnostics.Debug.WriteLine(message);
    }
}