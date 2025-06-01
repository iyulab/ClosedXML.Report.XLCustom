using System.Reflection;

namespace ClosedXML.Report.XLCustom;
public class XLCustomTemplateErrors : TemplateErrors
{
    private readonly FieldInfo _errorsField;

    public XLCustomTemplateErrors()
    {
        _errorsField = typeof(TemplateErrors).GetField("_errors",
            BindingFlags.NonPublic | BindingFlags.Instance);
    }

    internal void Add(TemplateError templateError)
    {
        var errors = (List<TemplateError>)_errorsField.GetValue(this);

        if (!errors.Exists(x => x.Range.Equals(templateError.Range) && x.Message == templateError.Message))
        {
            errors.Add(templateError);
        }
    }
}