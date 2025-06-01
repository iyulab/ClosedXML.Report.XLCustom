namespace ClosedXML.Report.XLCustom;

/// <summary>
/// Partial class containing IXLTemplate interface implementation methods
/// </summary>
public partial class XLCustomTemplate
{
    public void SaveAs(string file)
    {
        EnsureTemplateCreated();
        _baseTemplate.SaveAs(file);
    }

    public void SaveAs(string file, SaveOptions options)
    {
        EnsureTemplateCreated();
        _baseTemplate.SaveAs(file, options);
    }

    public void SaveAs(string file, bool validate, bool evaluateFormulae = false)
    {
        EnsureTemplateCreated();
        _baseTemplate.SaveAs(file, validate, evaluateFormulae);
    }

    public void SaveAs(Stream stream)
    {
        EnsureTemplateCreated();
        _baseTemplate.SaveAs(stream);
    }

    public void SaveAs(Stream stream, SaveOptions options)
    {
        EnsureTemplateCreated();
        _baseTemplate.SaveAs(stream, options);
    }

    public void SaveAs(Stream stream, bool validate, bool evaluateFormulae = false)
    {
        EnsureTemplateCreated();
        _baseTemplate.SaveAs(stream, validate, evaluateFormulae);
    }
}