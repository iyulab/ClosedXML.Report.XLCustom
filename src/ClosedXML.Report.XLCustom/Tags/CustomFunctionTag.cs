using ClosedXML.Report.Options;

namespace ClosedXML.Report.XLCustom.Tags;

/// <summary>
/// Tag that applies custom functions to cells
/// </summary>
public class CustomFunctionTag : OptionTag
{
    public override void Execute(ProcessingContext context)
    {
        var xlCell = Cell.GetXlCell(context.Range);
        var functionRegistry = XLCustomRegistry.Instance.FunctionRegistry;

        try
        {
            var variableName = GetParameter("name");
            var functionName = GetParameter("function");
            var parametersStr = GetParameter("parameters");

            Log.Debug($"CustomFunctionTag - variable: {variableName}, function: {functionName}, params: {parametersStr}");

            if (string.IsNullOrEmpty(variableName))
            {
                xlCell.Value = "Error: Missing variable name";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            if (string.IsNullOrEmpty(functionName))
            {
                xlCell.Value = "Error: Missing function name";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // 매개변수 파싱
            var parameters = new List<string>();
            if (!string.IsNullOrEmpty(parametersStr))
            {
                parameters.AddRange(parametersStr.Split(',').Select(p => UnescapeParameter(p.Trim())));
            }

            // 변수 평가 - 범위 내 item 객체에 대한 검증 추가
            object? value = null;
            try
            {
                // 변수 평가 시도
                // Evaluate 메서드가 실패해도 변환 가능하면 계속 진행
                try
                {
                    value = context.Evaluator.Evaluate(variableName, new Parameter("item", context.Value));
                    Log.Debug($"Evaluated variable {variableName} = {value ?? "null"}");
                }
                catch (Exception ex)
                {
                    // item 이 없는 경우 처리 시도
                    try
                    {
                        value = context.Evaluator.Evaluate(variableName);
                        Log.Debug($"Evaluated variable without item: {variableName} = {value ?? "null"}");
                    }
                    catch
                    {
                        // 두 가지 방법 모두 실패하면 원래 오류 표시
                        Log.Debug($"Error evaluating variable {variableName}: {ex.Message}");
                        xlCell.Value = $"Error: {ex.Message}";
                        xlCell.Style.Font.FontColor = XLColor.Red;
                        return;
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Debug($"Error evaluating variable {variableName}: {ex.Message}");
                xlCell.Value = $"Error: {ex.Message}";
                xlCell.Style.Font.FontColor = XLColor.Red;
                return;
            }

            // 함수 적용
            if (functionRegistry.IsRegistered(functionName))
            {
                try
                {
                    var function = functionRegistry.GetFunction(functionName);
                    if (function != null)
                    {
                        function(xlCell, value, parameters.ToArray());
                        Log.Debug($"Applied function {functionName} with {parameters.Count} parameters");
                    }
                    else
                    {
                        xlCell.Value = $"Function not found: {functionName}";
                        xlCell.Style.Font.FontColor = XLColor.Red;
                    }
                }
                catch (Exception ex)
                {
                    Log.Debug($"Error applying function {functionName}: {ex.Message}");
                    xlCell.Value = $"Function error: {ex.Message}";
                    xlCell.Style.Font.FontColor = XLColor.Red;
                }
            }
            else
            {
                xlCell.Value = $"Unknown function: {functionName}";
                xlCell.Style.Font.FontColor = XLColor.Red;
            }
        }
        catch (Exception ex)
        {
            Log.Debug($"General error in CustomFunctionTag: {ex.Message}");
            xlCell.Value = $"Error: {ex.Message}";
            xlCell.Style.Font.FontColor = XLColor.Red;
        }
    }

    /// <summary>
    /// Unescapes a parameter value from tag parameters
    /// </summary>
    private string UnescapeParameter(string param)
    {
        if (param.StartsWith("'") && param.EndsWith("'") && param.Length >= 2)
        {
            param = param.Substring(1, param.Length - 2).Replace("''", "'");
        }

        return param;
    }
}